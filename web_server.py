"""
Web Server for Voice Assistant
Flask backend with SocketIO for real-time communication
"""

import os
import sys
import threading
import logging
import time
import re
import subprocess
import signal
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv

from flask import Flask, render_template, request, jsonify
from flask_socketio import SocketIO, emit
from flask_cors import CORS

# Import voice assistant modules
from voice_assistant_v2 import VoiceAssistant
from stt_module import SpeechToText, WhisperSTT, WakeWordListener
from tts_module import SmartTTS

# Load environment variables
load_dotenv()

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger('web_server')

# Initialize Flask app
app = Flask(__name__)
app.config['SECRET_KEY'] = os.getenv('FLASK_SECRET_KEY', 'voice-assistant-secret-key')
CORS(app)

# Initialize SocketIO
socketio = SocketIO(app, cors_allowed_origins="*", async_mode='threading')

class VoiceAssistantWeb:
    """Voice Assistant wrapper for web interface with wake word support"""
    
    def __init__(self):
        self.logger = logging.getLogger('voice_assistant_web')
        self.assistant = None
        self.stt = None
        self.tts = None
        self.listening = False
        self.speaking = False
        self._stop_event = threading.Event()
        self._speaking_flag = False
        self._interrupt_requested = False
        self._current_tts_process = None
        self._listen_thread = None  # Track the listening thread
        
        # Wake word and conversation mode state
        self.wake_word_listener = None
        self.wake_word_enabled = False
        self.in_conversation_mode = False  # True after wake word, False after bye
        self._conversation_stop_words = ['bye', 'goodbye', 'stop', 'exit', 'quit', 'see you']
        
        # Command suggestions for rotation
        self.command_list = [
            "Open Chrome", "Search for Python", "Increase volume", "Decrease volume",
            "Open Notepad", "Check weather", "Tell me a joke", "What's the time?"
        ]
        self.command_index = 0
        
        self._init_components()
    
    def _init_components(self):
        """Initialize voice assistant components"""
        try:
            self.assistant = VoiceAssistant()
            self.logger.info("Voice assistant initialized successfully")
        except Exception as e:
            self.logger.error(f"Failed to initialize voice assistant: {e}")
            self.assistant = None
        
        # Initialize STT
        try:
            self.stt = SpeechToText()
            self.logger.info("STT initialized successfully")
        except Exception as e:
            self.logger.warning(f"Primary STT failed: {e}, trying Whisper...")
            try:
                self.stt = WhisperSTT(model_size="base")
                self.logger.info("Whisper STT initialized successfully")
            except Exception as fallback_e:
                self.logger.error(f"Both STT engines failed: {fallback_e}")
                self.stt = None
        
        # Initialize TTS
        try:
            self.tts = SmartTTS()
            if self.tts.is_available():
                self.logger.info("TTS initialized successfully")
            else:
                self.logger.warning("TTS not available")
        except Exception as e:
            self.logger.warning(f"TTS initialization failed: {e}")
            self.tts = None
        
        # Initialize wake word listener - share microphone with STT to avoid conflicts
        try:
            # Pass the STT's microphone to wake word listener to avoid conflicts
            shared_microphone = self.stt.microphone if self.stt and hasattr(self.stt, 'microphone') else None
            self.wake_word_listener = WakeWordListener(
                callback=self._on_wake_word_detected,
                microphone=shared_microphone
            )
            self.logger.info("Wake word listener initialized (with shared microphone)")
            # Auto-start wake word listening on startup
            self._start_wake_word_mode()
        except Exception as e:
            self.logger.warning(f"Wake word listener initialization failed: {e}")
            self.wake_word_listener = None
    
    def _start_wake_word_mode(self):
        """Start listening for wake word (initial state)"""
        self.logger.info(f"_start_wake_word_mode() called - wake_word_listener exists: {self.wake_word_listener is not None}")
        
        if self.wake_word_listener:
            if not self.wake_word_listener.is_running():
                self.logger.info("Starting wake word listener...")
                self.wake_word_listener.start()
                self.wake_word_enabled = True
                self.in_conversation_mode = False
                self.logger.info("✅ Wake word mode started - Say 'Hey Nova' or 'Nova' to activate")
                socketio.emit('status_update', {
                    'status': 'wake_word',
                    'message': '🎧 Say "Hey Nova" or "Nova" to wake me up'
                })
            else:
                self.logger.info("Wake word listener already running")
        else:
            self.logger.error("Wake word listener is None - cannot start wake word mode")
    
    def _stop_wake_word_mode(self):
        """Stop wake word listening"""
        if self.wake_word_listener:
            # Check if it's still running (it may have stopped itself after detection)
            if self.wake_word_listener.is_running():
                self.wake_word_listener.stop()
            self.wake_word_enabled = False
            self.logger.info("Wake word mode stopped")
    
    def _on_wake_word_detected(self, wake_word_text: str):
        """Callback when wake word is detected - enters conversation mode"""
        self.logger.info(f"✅ Wake word detected: '{wake_word_text}'")
        
        # Stop wake word listener to free microphone
        self.logger.info("Stopping wake word listener to free microphone...")
        self._stop_wake_word_mode()
        
        # Wait for wake word listener to fully stop - reduced delay
        time.sleep(0.5)
        
        # Enter conversation mode
        self.in_conversation_mode = True
        self.logger.info("Entered conversation mode")
        
        # Emit event to frontend
        socketio.emit('wake_word_detected', {
            'message': 'Hello! How can I help you?',
            'wake_word': wake_word_text
        })
        socketio.emit('status_update', {
            'status': 'listening',
            'message': '👋 Hello! I\'m listening. Say "bye" to end conversation.'
        })
        
        # Say hello greeting
        try:
            if self.tts:
                self._speak_text("Hello! How can I help you today?")
                # Small delay to let greeting start speaking
                time.sleep(0.5)
        except Exception as e:
            self.logger.warning(f"TTS failed for greeting: {e}")
        
        # Start continuous listening for conversation
        self.logger.info("Starting listening for conversation...")
        time.sleep(0.2)  # Reduced delay to let microphone settle after TTS (was 0.5)
        if not self.listening:
            success = self.start_listening()
            if success:
                self.logger.info("✅ Listening started successfully")
            else:
                self.logger.error("❌ Failed to start listening")
        else:
            self.logger.warning("Already listening, not starting again")
    
    def _check_conversation_end(self, text: str) -> bool:
        """Check if user said bye/goodbye to end conversation"""
        if not text:
            return False
        text_lower = text.lower().strip()
        return any(stop_word in text_lower for stop_word in self._conversation_stop_words)
    
    def _end_conversation(self):
        """End conversation mode and return to wake word listening"""
        self.logger.info("_end_conversation() called - ending conversation mode")
        self.in_conversation_mode = False
        
        # Stop active listening (don't wait for thread)
        if self.listening:
            self.logger.info("Stopping active listening...")
            self.stop_listening()
        
        # Update UI FIRST with the goodbye message (immediate feedback)
        socketio.emit('output_update', {
            'message': '👋 Goodbye! Say my name when you need me again.',
            'speak': False
        })
        
        # Then speak the SAME goodbye message
        try:
            if self.tts:
                self._speak_text("Goodbye! Say my name when you need me again.")
        except Exception as e:
            self.logger.warning(f"TTS failed for goodbye: {e}")
        
        # Update status to wake word mode
        socketio.emit('status_update', {
            'status': 'wake_word_mode',
            'message': '😴 Wake word mode active - Say "Hey Nova" to activate'
        })
        
        # Return to wake word mode after a short delay - reduced
        self.logger.info("Waiting 1.5 seconds before restarting wake word mode...")
        time.sleep(1.5)
        self.logger.info("Restarting wake word mode...")
        self._start_wake_word_mode()
    
    def toggle_wake_word(self, enabled: bool):
        """Enable or disable wake word listening (manual override)"""
        if not self.wake_word_listener:
            return {'success': False, 'message': 'Wake word not available'}
        
        if enabled:
            self._start_wake_word_mode()
            return {'success': True, 'message': 'Wake word enabled. Say "Hey Nova" or "Nova" to activate'}
        else:
            self._stop_wake_word_mode()
            if self.listening:
                self.stop_listening()
            return {'success': True, 'message': 'Wake word disabled'}
        
        return {'success': True, 'message': f'Wake word already {"enabled" if enabled else "disabled"}'}
    
    def get_status(self):
        """Get current status"""
        return {
            'initialized': self.assistant is not None,
            'stt_available': self.stt is not None,
            'tts_available': self.tts is not None and self.tts.is_available(),
            'listening': self.listening,
            'speaking': self.speaking,
            'wake_word_enabled': self.wake_word_enabled,
            'wake_word_available': self.wake_word_listener is not None,
            'current_model': self.assistant.get_current_model() if self.assistant else None,
            'available_models': self.assistant.get_available_models() if self.assistant else []
        }
    
    def get_next_command_hint(self):
        """Get next command suggestion"""
        if not self.command_list:
            return "Try saying: 'Open Chrome'"
        command = self.command_list[self.command_index]
        self.command_index = (self.command_index + 1) % len(self.command_list)
        return f"Try saying: '{command}'"
    
    def start_listening(self):
        """Start listening for voice commands"""
        self.logger.info("start_listening() called")
        
        if self.listening:
            self.logger.warning("Already listening, returning False")
            return False
        
        # Wait for any previous listening thread to fully terminate
        if self._listen_thread and self._listen_thread.is_alive():
            self.logger.info("Waiting for previous listening thread to terminate...")
            self._listen_thread.join(timeout=2.0)
        
        # Recreate STT to ensure clean microphone state
        try:
            # Small delay to let previous microphone context fully release
            time.sleep(0.5)
            self.logger.info("Creating new SpeechToText instance...")
            self.stt = SpeechToText()
            self.logger.info("✅ STT reinitialized for fresh listening session")
        except Exception as e:
            self.logger.error(f"❌ Failed to reinitialize STT: {e}")
            if not self.stt:
                return False
        
        self.listening = True
        self._stop_event.clear()
        
        # Start listening in background thread
        self.logger.info("Starting listen_loop in background thread...")
        self._listen_thread = threading.Thread(target=self._listen_loop, daemon=True)
        self._listen_thread.start()
        self.logger.info("✅ Listening thread started")
        
        # Emit event to frontend to update UI
        socketio.emit('listening_started', {'success': True})
        
        return True
    
    def stop_listening(self):
        """Stop listening"""
        self.listening = False
        self._stop_event.set()
        self._stop_all_tts()
        
        # Don't wait for thread - it may be blocked on microphone input
        # Just log and return - the thread will exit on its own when it times out
        if self._listen_thread and self._listen_thread.is_alive():
            self.logger.info("Listening thread still active - will exit on next timeout")
        
        # Emit stopped event immediately so UI updates
        socketio.emit('listening_stopped')
        
        return True
    
    def _listen_loop(self):
        """Main listening loop - handles conversation mode"""
        if self.in_conversation_mode:
            socketio.emit('status_update', {
                'status': 'listening',
                'message': '🎤 I\'m listening... Say "bye" to end conversation'
            })
        else:
            socketio.emit('status_update', {
                'status': 'listening',
                'message': '🎤 Listening... Speak your command!'
            })
        
        while self.listening and not self._stop_event.is_set():
            try:
                # Listen for command with optimized timeout
                text = self.stt.listen(timeout=3, phrase_time_limit=8)
                
                if text and self.listening:
                    self.logger.info(f"Speech recognized: {text}")
                    socketio.emit('speech_recognized', {'text': text})
                    self.logger.info(f"Emitted speech_recognized event to client")
                    
                    # In conversation mode, check for bye/goodbye first
                    if self.in_conversation_mode and self._check_conversation_end(text):
                        self._end_conversation()
                        break
                    
                    # Process command
                    should_exit = self._process_command(text)
                    
                    if should_exit:
                        self.stop_listening()
                        socketio.emit('status_update', {
                            'status': 'ready',
                            'message': '👋 Goodbye!'
                        })
                        break
                    
                    # In conversation mode, continue listening
                    if self.in_conversation_mode:
                        socketio.emit('status_update', {
                            'status': 'listening',
                            'message': '🎤 I\'m listening... Say "bye" to end conversation'
                        })
                    else:
                        # Not in conversation mode, stop listening after command
                        self.stop_listening()
                        break
                    
                elif self.listening:
                    # No speech detected - in conversation mode, keep listening
                    if self.in_conversation_mode:
                        socketio.emit('status_update', {
                            'status': 'listening',
                            'message': '🎤 Still listening... Say "bye" to end'
                        })
                    else:
                        socketio.emit('output_update', {
                            'message': '❌ No speech detected. Try again.',
                            'speak': False
                        })
                        self.stop_listening()
                        break
                    
            except Exception as e:
                if self.listening:
                    self.logger.error(f"Listening error: {e}")
                    socketio.emit('output_update', {
                        'message': f'❌ Error: {str(e)}',
                        'speak': False
                    })
                    if not self.in_conversation_mode:
                        self.stop_listening()
                        break
        
        # Listening stopped
        self.listening = False
        # Only emit if not already stopped (avoid double events)
        if not self._stop_event.is_set():
            socketio.emit('listening_stopped')
        
        # If we were in conversation mode but listening stopped unexpectedly, return to wake word mode
        if self.in_conversation_mode and not self._stop_event.is_set():
            self._end_conversation()
        elif not self.in_conversation_mode and self.wake_word_listener and not self.wake_word_listener.is_running():
            # Return to wake word mode if not already there
            socketio.emit('status_update', {
                'status': 'ready',
                'message': 'Ready'
            })
    
    def _process_command(self, command: str):
        """Process a voice command"""
        if not command:
            return False
        
        command_lower = command.strip().lower()
        
        # Check for exit commands
        if command_lower in ['quit', 'exit', 'stop', 'bye']:
            # Update UI first for immediate feedback
            socketio.emit('output_update', {
                'message': '👋 Goodbye!',
                'speak': False
            })
            # Then speak the goodbye
            self._speak_text("Goodbye!")
            return True
        
        # Check for help
        if command_lower == 'help':
            help_text = "Available commands: Open apps, search, volume control, weather, jokes, time, history"
            socketio.emit('output_update', {
                'message': f'💡 {help_text}',
                'speak': True
            })
            return False
        
        # Check for history
        if command_lower == 'history':
            self._show_history()
            return False
        
        # Check for clear history
        if any(cmd in command_lower for cmd in ['clear history', 'clear chat', 'delete history', 'reset history']):
            # Clear both command history and conversation memory
            try:
                if self.assistant:
                    self.assistant.clear_conversation_memory()
                socketio.emit('output_update', {
                    'message': '🗑️ Conversation memory cleared.',
                    'speak': True
                })
                self._speak_text("Conversation memory cleared.")
            except Exception as e:
                self.logger.error(f"Error clearing history: {e}")
                socketio.emit('output_update', {
                    'message': '❌ Failed to clear history.',
                    'speak': False
                })
            return False
        
        # Show thinking indicator
        socketio.emit('thinking_start')
        socketio.emit('status_update', {
            'status': 'processing',
            'message': 'Nova is thinking...'
        })
        
        try:
            if not self.assistant:
                socketio.emit('output_update', {
                    'message': '❌ Voice assistant not initialized',
                    'speak': False
                })
                return False
            
            result = self.assistant.process_command(command)
            
            if result['success']:
                message = result['message']
                full_response = result.get('response', message)
                
                # Check if this is a conversational response (not a command execution)
                is_conversation = full_response and not message.startswith('Opening') and not message.startswith('Searching') and not message.startswith('Volume') and not message.startswith('Created')
                
                if is_conversation:
                    # For conversational responses, send as 💬 (for sidebar) and speak it
                    socketio.emit('output_update', {
                        'message': f'💬 {full_response}',
                        'speak': True
                    })
                    self._speak_text(full_response)
                else:
                    # For command executions, send ✅ (short) and 💬 (full) if different
                    socketio.emit('output_update', {
                        'message': f'✅ {message}',
                        'speak': False
                    })
                    
                    if full_response and full_response != message:
                        socketio.emit('output_update', {
                            'message': f'💬 {full_response}',
                            'speak': True
                        })
                        self._speak_text(full_response)
                    else:
                        self._speak_text(message)
            else:
                # Handle failure with helpful response
                self._handle_command_failure(command, result['message'])
            
        except Exception as e:
            self.logger.error(f"Command processing error: {e}")
            socketio.emit('output_update', {
                'message': f'❌ Error: {str(e)}',
                'speak': False
            })
        
        socketio.emit('status_update', {
            'status': 'listening',
            'message': '🎤 Listening...'
        })
        return False
    
    def _handle_command_failure(self, command: str, error_message: str):
        """Handle command failures with helpful responses"""
        command_lower = command.lower()
        
        if "application not found" in error_message.lower():
            short_response = "I didn't understand."
            full_response = "I couldn't find that application. Try saying 'Open Chrome' or 'Open Notepad'."
        elif "volume" in command_lower:
            short_response = "I didn't understand."
            full_response = "I can help with volume control! Try saying 'Increase volume' or 'Decrease volume'."
        elif "search" in command_lower:
            short_response = "I didn't understand."
            full_response = "I can help you search! Try saying 'Search for Python tutorials'."
        elif any(word in command_lower for word in ['time', 'weather', 'joke']):
            short_response = "I didn't understand."
            full_response = "I can tell you the time, check weather, or tell jokes! Just ask me directly."
        else:
            short_response = "I didn't understand."
            full_response = f"I didn't quite understand '{command}'. Try simpler commands like 'Open Chrome' or 'What's the time?'"
        
        # Emit short version for sidebar (💡) and full version for speech/display (💬)
        socketio.emit('output_update', {
            'message': f'💡 {short_response}',
            'speak': False
        })
        socketio.emit('output_update', {
            'message': f'💬 {full_response}',
            'speak': True
        })
        self._speak_text(full_response)
    
    def _show_history(self):
        """Show command history"""
        try:
            if not self.assistant:
                return
            
            history = self.assistant.get_history(limit=5)
            if not history:
                socketio.emit('output_update', {
                    'message': '📝 No command history yet.',
                    'speak': True
                })
                return
            
            messages = ['📝 Recent Commands:']
            for entry in history:
                success = "✅" if entry['success'] else "❌"
                messages.append(f"  {success} {entry['input']}")
            
            full_message = '\n'.join(messages)
            socketio.emit('output_update', {
                'message': full_message,
                'speak': True
            })
            self._speak_text("Here are your recent commands")
            
        except Exception as e:
            self.logger.error(f"History error: {e}")
    
    def _speak_text(self, text: str):
        """Speak text using TTS"""
        if not text:
            return
        
        # Use PowerShell TTS for guaranteed output
        try:
            if self._speaking_flag:
                return
            
            self._speaking_flag = True
            self.speaking = True
            socketio.emit('tts_status', {'speaking': True})
            
            # Clean text for PowerShell - PRESERVE important punctuation for proper speech
            # Keep apostrophes, periods, commas, question marks, exclamation marks
            clean_text = re.sub(r'[^\x00-\x7F]+', '', text)  # Remove non-ASCII only
            # Replace problematic chars but keep speech punctuation
            clean_text = re.sub(r'["\'\'`]', '', clean_text)  # Remove quotes that break PowerShell
            clean_text = ' '.join(clean_text.split())  # Normalize spaces
            
            if not clean_text.strip():
                self._speaking_flag = False
                self.speaking = False
                socketio.emit('tts_status', {'speaking': False})
                return
            
            # Use PowerShell with temp file
            import tempfile
            with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False, encoding='utf-8') as temp_file:
                temp_file.write(clean_text)
                temp_file_path = temp_file.name
            
            try:
                ps_command = f'powershell -Command "$text = Get-Content \'{temp_file_path}\'; Add-Type -AssemblyName System.Speech; (New-Object System.Speech.Synthesis.SpeechSynthesizer).Speak($text)"'
                
                process = subprocess.Popen(
                    ps_command,
                    shell=True,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True
                )
                
                self._current_tts_process = process
                process.wait()
                
            finally:
                try:
                    os.unlink(temp_file_path)
                except:
                    pass
            
        except Exception as e:
            self.logger.error(f"TTS error: {e}")
        finally:
            self._speaking_flag = False
            self.speaking = False
            socketio.emit('tts_status', {'speaking': False})
            self._current_tts_process = None
    
    def _stop_all_tts(self):
        """Stop all TTS processes"""
        self._interrupt_requested = True
        
        if self._current_tts_process:
            try:
                self._current_tts_process.terminate()
                self._current_tts_process.wait(timeout=0.5)
            except:
                try:
                    self._current_tts_process.kill()
                except:
                    pass
        
        time.sleep(0.1)  # Reduced delay for faster response (was 0.2)
        self._interrupt_requested = False
        self._speaking_flag = False
        self.speaking = False
    
    def process_text_command(self, command: str):
        """Process a text command from web interface"""
        if not command:
            return {'success': False, 'message': 'Empty command'}
        
        socketio.emit('output_update', {
            'message': f'You typed: {command}',
            'speak': False
        })
        
        should_exit = self._process_command(command)
        return {'success': True, 'should_exit': should_exit}
    
    def change_model(self, model_name: str):
        """Change AI model"""
        if not self.assistant:
            return {'success': False, 'message': 'Assistant not initialized'}
        
        success = self.assistant.set_model(model_name)
        if success:
            return {'success': True, 'message': f'Switched to {model_name}'}
        else:
            return {'success': False, 'message': f'Failed to switch to {model_name}'}
    
    def get_history(self):
        """Get command history"""
        if not self.assistant:
            return []
        try:
            return self.assistant.get_history(limit=10)
        except:
            return []
    
    def cleanup(self):
        """Cleanup resources"""
        self.stop_listening()
        if self.wake_word_listener and self.wake_word_listener.is_running():
            self.wake_word_listener.stop()
        self._stop_all_tts()
        if self.assistant:
            self.assistant.cleanup()

# Global assistant instance
assistant = None

@app.route('/')
def index():
    """Serve the main page"""
    return render_template('index.html')

@app.route('/api/status')
def api_status():
    """Get current status"""
    global assistant
    if assistant:
        return jsonify(assistant.get_status())
    return jsonify({'error': 'Assistant not initialized'}), 500

@app.route('/api/history')
def api_history():
    """Get command history"""
    global assistant
    if assistant:
        return jsonify(assistant.get_history())
    return jsonify([])

@app.route('/api/command', methods=['POST'])
def api_command():
    """Process a text command"""
    global assistant
    if not assistant:
        return jsonify({'success': False, 'message': 'Assistant not initialized'}), 500
    
    data = request.get_json()
    command = data.get('command', '')
    
    result = assistant.process_text_command(command)
    return jsonify(result)

@app.route('/api/model', methods=['POST'])
def api_change_model():
    """Change AI model"""
    global assistant
    if not assistant:
        return jsonify({'success': False, 'message': 'Assistant not initialized'}), 500
    
    data = request.get_json()
    model_name = data.get('model', '')
    
    result = assistant.change_model(model_name)
    return jsonify(result)

@app.route('/api/wake_word', methods=['POST'])
def api_toggle_wake_word():
    """Toggle wake word listening"""
    global assistant
    if not assistant:
        return jsonify({'success': False, 'message': 'Assistant not initialized'}), 500
    
    data = request.get_json()
    enabled = data.get('enabled', False)
    
    result = assistant.toggle_wake_word(enabled)
    return jsonify(result)

@app.route('/api/assistant/start', methods=['POST'])
def api_start_assistant():
    """Start the voice assistant (start ALL listening - wake word + continuous)"""
    global assistant
    if not assistant:
        return jsonify({'success': False, 'message': 'Assistant not initialized'}), 500
    
    try:
        # Check if already running
        if assistant.listening or (assistant.wake_word_listener and assistant.wake_word_listener.is_running()):
            return jsonify({'success': True, 'message': 'Already running', 'listening': True})
        
        # Start wake word listener first
        if assistant.wake_word_listener:
            assistant.wake_word_listener.start()
            assistant.wake_word_enabled = True
        
        # Then start continuous listening
        assistant._start_continuous_listening()
        return jsonify({'success': True, 'message': 'All listening started', 'listening': True})
    except Exception as e:
        return jsonify({'success': False, 'message': f'Error: {str(e)}'}), 500

@app.route('/api/assistant/stop', methods=['POST'])
def api_stop_assistant():
    """Stop the voice assistant (stop ALL listening - wake word + continuous)"""
    global assistant
    if not assistant:
        return jsonify({'success': False, 'message': 'Assistant not initialized'}), 500
    
    try:
        # Stop wake word listener first
        if assistant.wake_word_listener and assistant.wake_word_listener.is_running():
            assistant.wake_word_listener.stop()
            assistant.wake_word_enabled = False
        
        # Then stop continuous listening
        if assistant.listening:
            assistant.stop_listening()
            return jsonify({'success': True, 'message': 'All listening stopped', 'listening': False})
        else:
            return jsonify({'success': True, 'message': 'Already stopped', 'listening': False})
    except Exception as e:
        return jsonify({'success': False, 'message': f'Error: {str(e)}'}), 500

@app.route('/api/assistant/status', methods=['GET'])
def api_assistant_status():
    """Get assistant running status"""
    global assistant
    if not assistant:
        return jsonify({'running': False, 'listening': False})
    
    # Consider assistant "running" only if continuous listening OR wake word is active
    is_running = assistant.listening or (assistant.wake_word_listener and assistant.wake_word_listener.is_running())
    
    return jsonify({
        'running': is_running,
        'listening': assistant.listening
    })

@app.route('/api/startup/check', methods=['GET'])
def api_check_startup():
    """Check if Nova is set to start with Windows"""
    try:
        from startup_manager import StartupManager
        manager = StartupManager()
        enabled, _ = manager.is_startup_enabled()
        return jsonify({'enabled': enabled})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/startup/toggle', methods=['POST'])
def api_toggle_startup():
    """Toggle Windows startup registration"""
    try:
        from startup_manager import StartupManager
        manager = StartupManager()
        success, msg = manager.toggle_startup()
        if success:
            enabled, _ = manager.is_startup_enabled()
            return jsonify({'success': True, 'enabled': enabled, 'message': msg})
        else:
            return jsonify({'success': False, 'message': msg}), 500
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500

# SocketIO events
@socketio.on('connect')
def handle_connect():
    """Handle client connection"""
    logger.info('Client connected')
    emit('connected', {'message': 'Connected to Voice Assistant'})
    
    # Send current status
    global assistant
    if assistant:
        emit('status_update', {
            'status': 'ready' if not assistant.listening else 'listening',
            'message': 'Ready' if not assistant.listening else '🎤 Listening...'
        })

@socketio.on('disconnect')
def handle_disconnect():
    """Handle client disconnection"""
    logger.info('Client disconnected')

@socketio.on('start_listening')
def handle_start_listening():
    """Start listening for voice commands - enters conversation mode like wake word"""
    global assistant
    if assistant:
        # Stop wake word listener first to free microphone
        if assistant.wake_word_listener and assistant.wake_word_listener.is_running():
            assistant._stop_wake_word_mode()
        
        # Enter conversation mode (same as wake word)
        assistant.in_conversation_mode = True
        
        # Emit greeting (same as wake word)
        emit('wake_word_detected', {
            'message': 'Hello! How can I help you?',
            'wake_word': 'manual'
        })
        emit('status_update', {
            'status': 'listening',
            'message': '👋 Hello! I\'m listening. Say "bye" to end conversation.'
        })
        
        # Speak greeting
        try:
            if assistant.tts:
                threading.Thread(target=assistant._speak_text, 
                               args=("Hello! How can I help you today?",), 
                               daemon=True).start()
        except Exception as e:
            assistant.logger.warning(f"TTS failed for greeting: {e}")
        
        # Start listening after a short delay - reduced
        time.sleep(0.2)  # Reduced delay for faster response (was 0.5)
        success = assistant.start_listening()
        emit('listening_started', {'success': success})
    else:
        emit('listening_started', {'success': False, 'error': 'Assistant not initialized'})

@socketio.on('stop_listening')
def handle_stop_listening():
    """Stop listening - acts as bye/goodbye if in conversation mode"""
    global assistant
    if assistant:
        # If in conversation mode, treat this as saying "bye"
        if assistant.in_conversation_mode:
            assistant.logger.info("Stop button pressed during conversation mode - treating as bye")
            assistant._end_conversation()
        else:
            assistant.stop_listening()
            emit('listening_stopped')

@socketio.on('get_command_hint')
def handle_get_command_hint():
    """Get next command hint"""
    global assistant
    if assistant:
        hint = assistant.get_next_command_hint()
        emit('command_hint', {'hint': hint})

@socketio.on('speak_text')
def handle_speak_text(data):
    """Speak text via TTS"""
    global assistant
    if assistant:
        text = data.get('text', '')
        threading.Thread(target=assistant._speak_text, args=(text,), daemon=True).start()

def initialize_assistant():
    """Initialize the voice assistant"""
    global assistant
    try:
        assistant = VoiceAssistantWeb()
        logger.info("Voice Assistant Web initialized successfully")
    except Exception as e:
        logger.error(f"Failed to initialize Voice Assistant: {e}")
        assistant = None

def cleanup():
    """Cleanup on exit"""
    global assistant
    if assistant:
        assistant.cleanup()
        logger.info("Voice Assistant cleaned up")

if __name__ == '__main__':
    # Initialize assistant
    initialize_assistant()
    
    try:
        # Get port from environment variable for production
        port = int(os.environ.get('PORT', 5000))
        host = os.environ.get('HOST', '0.0.0.0')
        debug = os.environ.get('FLASK_ENV', 'development') == 'development'
        
        # Run the server
        logger.info(f"Starting web server on http://{host}:{port}")
        socketio.run(app, host=host, port=port, debug=debug)
    except KeyboardInterrupt:
        logger.info("Shutting down...")
    finally:
        cleanup()

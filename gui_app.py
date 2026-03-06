"""
Eye-catching GUI Frontend for Voice Assistant
Modern design with colors and contrasts, no scrolling
"""

import customtkinter as ctk
from tkinter import messagebox
import threading
import os
import sys
import subprocess
import signal
from pathlib import Path
from dotenv import load_dotenv
import time

# Import our modules
from voice_assistant_v2 import VoiceAssistant
from stt_module import SpeechToText, WhisperSTT
from tts_module import SmartTTS
import logging
# Local exception definitions
class InitializationError(Exception):
    def __init__(self, message, user_message=None):
        super().__init__(message)
        self.user_message = user_message

class AudioError(Exception):
    def __init__(self, message, user_message=None):
        super().__init__(message)
        self.user_message = user_message

class CommandExecutionError(Exception):
    def __init__(self, message, user_message=None):
        super().__init__(message)
        self.user_message = user_message

# Load environment variables
load_dotenv()

# Set appearance mode and color theme
ctk.set_appearance_mode("dark")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("green")  # Themes: "blue" (standard), "green", "dark-blue"

class VoiceAssistantGUI:
    """Eye-catching GUI for the voice assistant with modern design"""

    def __init__(self, root):
        self.logger = logging.getLogger('gui')
        self.root = root
        self.root.title("🤖 AI Voice Assistant")
        self.root.geometry("800x600")  # Larger default size for better visibility
        self.root.resizable(True, True)  # Allow resizing

        # Initialize components
        self._init_components()

        # Control flags
        self.listening = False
        self.running = True
        self.speaking = False

        # TTS queue for main thread processing
        self.tts_queue = []

        # Set up interruption handling
        self._setup_interruption_handling()

        # Command loop
        self.command_list = [
            "Open Chrome", "Search for Python", "Increase volume", "Decrease volume",
            "Play music", "Pause", "Next song", "Previous song", "Set alarm",
            "Check weather", "Tell me a joke", "What's the time?", "Quit"
        ]
        self.command_index = 0

        # Set up the UI
        self._setup_ui()

        # Set up cleanup on window close
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # Start command loop
        self._start_command_loop()

    def _init_components(self):
        """Initialize voice assistant components"""
        try:
            self.assistant = VoiceAssistant()
            self.status = "Ready"
        except Exception as e:
            self.status = f"Error: {str(e)}"
            messagebox.showerror("Initialization Error", f"Failed to initialize voice assistant: {e}")
            return

        # Initialize STT with fallback
        try:
            self.stt = SpeechToText()
        except Exception as e:
            self.logger.warning(f"Primary STT initialization failed: {str(e)}, trying fallback...")
            try:
                self.stt = WhisperSTT(model_size="base")
                self.logger.info("Fallback STT (Whisper) initialized successfully")
            except Exception as fallback_e:
                error_msg = f"Both STT engines failed: Primary - {str(e)}, Fallback - {str(fallback_e)}"
                self.logger.error(error_msg)
                self.stt = None
                messagebox.showwarning("STT Initialization Warning",
                    "Speech recognition is unavailable. You can still use text commands.\n"
                    "To fix speech recognition, check your internet connection and microphone permissions.")

        # Initialize TTS
        try:
            self.tts = SmartTTS()
            if self.tts.is_available():
                self.logger.info("TTS initialized successfully")
            else:
                self.logger.warning("TTS initialized but not available, will use text-only output")
        except Exception as e:
            self.logger.warning(f"TTS initialization failed: {str(e)}, continuing without TTS")
            self.tts = None

    def _setup_interruption_handling(self):
        """Set up interruption handling for TTS"""
        if self.tts and self.stt:
            # Define interruption callback
            def interruption_callback():
                print("\n🛑 Speech interrupted by user")
                # Set interruption flag for PowerShell TTS
                self._interrupt_requested = True
                # Also try to interrupt the TTS engine if available
                if self.tts.is_speaking():
                    self.tts.interrupt_speech()
                # Don't stop the interruption listener - keep it running for continuous interruption capability

            self.interruption_callback = interruption_callback
            self._interrupt_requested = False

    def _setup_ui(self):
        """Set up the user interface with modern CustomTkinter widgets"""
        # Main frame
        main_frame = ctk.CTkFrame(self.root, fg_color="#1a1a2e")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Title
        title_label = ctk.CTkLabel(
            main_frame,
            text="🤖 AI Voice Assistant",
            font=ctk.CTkFont(size=28, weight="bold"),
            text_color="#00ffff"
        )
        title_label.pack(pady=(0, 15))

        # Status label
        self.status_label = ctk.CTkLabel(
            main_frame,
            text=f"Status: {self.status}",
            font=ctk.CTkFont(size=14),
            text_color="#32cd32"
        )
        self.status_label.pack(pady=(0, 10))

        # AI Model selector
        model_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        model_frame.pack(fill="x", pady=(0, 10))

        model_label = ctk.CTkLabel(
            model_frame,
            text="AI Model:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#ffffff"
        )
        model_label.pack(side="left", padx=(0, 10))

        # Get available models
        available_models = self.assistant.get_available_models()
        if not available_models:
            available_models = ["No models available"]

        self.model_var = ctk.StringVar(value=self.assistant.get_current_model() or available_models[0])

        self.model_selector = ctk.CTkOptionMenu(
            model_frame,
            values=available_models,
            variable=self.model_var,
            command=self.change_model,
            font=ctk.CTkFont(size=11),
            fg_color="#2b2b2b",
            button_color="#444444",
            button_hover_color="#555555",
            text_color="#ffffff",
            width=150
        )
        self.model_selector.pack(side="left")

        # Output section
        output_label = ctk.CTkLabel(
            main_frame,
            text="Output:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#ffffff"
        )
        output_label.pack(anchor="w", pady=(0, 5))

        self.output_textbox = ctk.CTkTextbox(
            main_frame,
            height=120,
            font=ctk.CTkFont(size=11),
            fg_color="#2b2b2b",
            text_color="#ffffff",
            border_width=2,
            border_color="#00d4ff"
        )
        self.output_textbox.pack(fill="x", pady=(0, 20))
        self.output_textbox.insert("0.0", "Welcome! Click 'Start Listening' to begin.\n")
        self.output_textbox.configure(state="disabled")

        # Welcome message with TTS
        try:
            if self.tts:
                self.tts.speak("Welcome to the AI voice assistant! Click start listening to begin.")
        except Exception as e:
            self.logger.warning(f"TTS failed for GUI welcome message: {str(e)}")

        # Button frame
        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        button_frame.pack(fill="x", pady=(0, 15))

        # Start/Stop button
        self.listen_button = ctk.CTkButton(
            button_frame,
            text="🎤 Start Listening",
            command=self.toggle_listening,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#00d4ff",
            hover_color="#00bfff",
            text_color="#000000",
            height=50,
            width=200
        )
        self.listen_button.pack(side="left", padx=(0, 15))

        # Clear button
        clear_button = ctk.CTkButton(
            button_frame,
            text="🗑️ Clear",
            command=self.clear_output,
            font=ctk.CTkFont(size=12),
            fg_color="#ff6b6b",
            hover_color="#ff4500",
            text_color="#000000",
            height=50,
            width=120
        )
        clear_button.pack(side="left", padx=(0, 15))

        # Help button
        help_button = ctk.CTkButton(
            button_frame,
            text="❓ Help",
            command=self.show_help,
            font=ctk.CTkFont(size=12),
            fg_color="#fdcb6e",
            hover_color="#ffa500",
            text_color="#000000",
            height=50,
            width=100
        )
        help_button.pack(side="left", padx=(0, 15))



        # Instructions
        instructions = ctk.CTkLabel(
            main_frame,
            text="Say commands like:\n• 'Open Chrome'\n• 'Search for Python'\n• 'Increase volume'\n• 'Quit' to exit",
            font=ctk.CTkFont(size=10),
            text_color="#87ceeb",
            justify="left"
        )
        instructions.pack(anchor="w", pady=(10, 5))

        # Command display loop
        self.command_display = ctk.CTkLabel(
            main_frame,
            text="Try saying: 'Open Chrome'",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#ff6b6b"
        )
        self.command_display.pack(anchor="w", pady=(0, 10))

    def toggle_listening(self):
        """Toggle listening state"""
        if not self.listening:
            self.start_listening()
        else:
            self.stop_listening()

    def start_listening(self):
        """Start listening for voice commands"""
        self.listening = True
        self.listen_button.configure(text="⏹️ Stop Listening", fg_color="#ff4444", hover_color="#cc0000", text_color="#ffffff")
        self.status_label.configure(text="Status: Listening...", text_color="#ffa500")
        # Speak listening message synchronously to ensure it completes before listening starts
        self._speak_text_sync("🎤 Listening... Speak your command!")
        self.update_output("🎤 Listening... Speak your command!", speak=False)  # Don't speak again

        # Create stop event for this listening session
        self._stop_event = threading.Event()



        # Start listening for a single command
        threading.Thread(target=self._listen_once, daemon=True).start()

    def stop_listening(self):
        """Stop listening"""
        self.listening = False
        # Set stop event to signal listening thread to stop
        if hasattr(self, '_stop_event'):
            self._stop_event.set()
        # Stop interruption listener
        if self.stt:
            try:
                self.stt.stop_interruption_listener()
            except:
                pass
        # Add a short delay to allow AI response TTS to complete before stopping
        import time
        time.sleep(1.5)  # Reduced delay - give TTS 1.5 seconds to finish speaking AI response (was 3s)
        # Stop any ongoing TTS
        self._stop_all_tts()
        self.listen_button.configure(text="🎤 Start Listening", fg_color="#1f77b4", hover_color="#1565a0")
        self.status_label.configure(text="Status: Ready", text_color="#00ff88")
        self.update_output("Stopped listening.")

    def _listen_once(self):
        """Listen continuously for voice commands until exit command"""
        # Check if STT is available
        if self.stt is None:
            self.update_output("❌ Speech recognition is not available. Please check your microphone and internet connection.")
            self.stop_listening()
            return

        while self.listening and not self._stop_event.is_set():
            try:
                # Listen for a command with optimized timeout
                text = self.stt.listen(timeout=3, phrase_time_limit=8)

                if text and self.listening:
                    self.logger.debug(f"Speech recognized: {text[:50]}...")
                    self.update_output(f"You said: {text}", speak=False)  # Don't speak user input

                    # Process command and check if it's an exit command
                    should_exit = self.process_command(text)
                    if should_exit:
                        # Exit command detected, stop listening
                        self.stop_listening()
                        break
                    # Continue listening for next command
                elif self.listening:
                    # No speech detected, continue listening
                    self.update_output("❌ No speech detected. Try again.", speak=False)
                    continue

            except AudioError as e:
                if self.listening:
                    error_msg = f"Audio error: {str(e)}"
                    self.logger.error(error_msg, exc_info=True)
                    self.update_output(f"❌ Audio Error: {e}", speak=False)
                    # Continue listening despite audio error
                    continue
            except Exception as e:
                if self.listening:
                    error_msg = f"Unexpected listening error: {str(e)}"
                    self.logger.error(error_msg, exc_info=True)
                    self.update_output(f"❌ Listening Error: {e}", speak=False)
                    # Continue listening despite error
                    continue

    def process_command(self, command: str):
        """Process a voice command"""
        if not command:
            return False

        command = command.strip().lower()

        # Check for exit commands
        if command in ['quit', 'exit', 'stop', 'bye']:
            self.update_output("👋 Goodbye!", speak=True)
            self.running = False
            self.root.after(1000, self.root.quit)
            return True  # Exit command detected

        # Check for special commands
        if command == 'help':
            self.update_output("Available commands: Open apps, web search, volume control, etc.")
            return False

        if command == 'history':
            self.show_history()
            return False

        # Process through assistant
        self.status_label.configure(text="Status: Processing...", text_color="#ffa500")
        self.update_output("🔄 Processing command...", speak=False)  # Don't speak processing status

        try:
            result = self.assistant.process_command(command)

            if result['success']:
                message = result['message']
                self.update_output(f"✅ {message}", speak=False)  # Don't speak the short message
                # Always speak the AI response
                full_response = result.get('response', message)
                if full_response:
                    # Display and speak the full conversational response
                    self.update_output(f"💬 {full_response}", speak=True)
                else:
                    # For simple commands, just speak the confirmation
                    self.update_output(message, speak=True)
            else:
                # Provide helpful responses instead of error messages
                self._handle_command_failure(command, result['message'])

        except Exception as e:
            # Provide helpful responses for unexpected errors
            self._handle_command_failure(command, str(e))

        self.status_label.configure(text="Status: Ready", text_color="#00ff88")
        return False  # Not an exit command

    def _handle_command_failure(self, command: str, error_message: str):
        """Handle command failures with helpful, positive responses"""
        command_lower = command.lower()

        # Application opening failures
        if "application not found" in error_message.lower() or "file specified" in error_message.lower():
            helpful_responses = [
                f"I couldn't find that application. Try saying 'Open Chrome' or 'Open Notepad' instead.",
                f"Sorry, I don't recognize that app. You can try opening common apps like Chrome, Firefox, or Word.",
                f"That application isn't available. Would you like me to help you find the right command?"
            ]
            response = helpful_responses[len(command) % len(helpful_responses)]
            self.update_output(f"💡 {response}")

        # Volume control failures
        elif "volume" in command_lower:
            self.update_output("💡 I can help with volume control! Try saying 'Increase volume' or 'Decrease volume'.")

        # Music control failures
        elif any(word in command_lower for word in ['music', 'play', 'pause', 'next', 'previous', 'song']):
            self.update_output("💡 For music control, try commands like 'Play music', 'Pause', 'Next song', or 'Previous song'.")

        # Search failures
        elif "search" in command_lower:
            self.update_output("💡 I can help you search! Try saying 'Search for Python tutorials' or 'Search for weather'.")

        # Time/weather failures
        elif any(word in command_lower for word in ['time', 'weather', 'joke']):
            self.update_output("💡 I can tell you the time, check weather, or tell jokes! Just ask me directly.")

        # Alarm failures
        elif "alarm" in command_lower:
            self.update_output("💡 For alarms, try saying 'Set alarm for 3 PM' or 'Set alarm for 30 minutes'.")

        # General failures - provide helpful suggestions
        else:
            general_responses = [
                f"I didn't quite understand '{command}'. Try simpler commands like 'Open Chrome' or 'What's the time?'",
                f"I'm still learning! For now, I work best with commands like 'Search for Python' or 'Increase volume'.",
                f"That's an interesting request! I currently handle app opening, web searches, volume control, and basic questions.",
                f"I want to help with that, but I need a clearer command. Try 'Help' to see what I can do!"
            ]
            response = general_responses[len(command) % len(general_responses)]
            self.update_output(f"🤖 {response}")

    def show_history(self):
        """Show command history"""
        try:
            history = self.assistant.get_history(limit=5)
            if not history:
                self.update_output("No command history yet.")
                return

            self.update_output("📝 Recent Commands:")
            for entry in history:
                timestamp = entry['timestamp']
                user_input = entry['input']
                success = "✅" if entry['success'] else "❌"
                self.update_output(f"  {success} {user_input}")
        except Exception as e:
            self.update_output(f"Error getting history: {str(e)}")

    def clear_output(self):
        """Clear the output text area"""
        self.output_textbox.configure(state="normal")
        self.output_textbox.delete("0.0", "end")
        self.output_textbox.insert("0.0", "Output cleared.\n")
        self.output_textbox.configure(state="disabled")

    def update_output(self, message: str, speak: bool = True):
        """Update the output text area and optionally speak the message"""
        self.output_textbox.configure(state="normal")
        self.output_textbox.insert("end", message + "\n")
        self.output_textbox.see("end")  # Scroll to end if needed, but since fixed size, it will overwrite
        self.output_textbox.configure(state="disabled")

        # Speak the message asynchronously if TTS is available and speaking is enabled
        if speak and self.tts and self.tts.is_available():
            try:
                # Start TTS in background thread to avoid blocking listening
                threading.Thread(target=self._speak_text_async, args=(message,), daemon=True).start()
            except Exception as e:
                self.logger.warning(f"TTS failed for message '{message[:50]}...': {str(e)}")
                # Fallback to text-only output
                print(f"🔊 [TTS]: {message}")

    def _speak_text_async(self, text: str):
        """Asynchronous wrapper for _speak_text to avoid blocking the main thread"""
        self._speak_text(text)

    def _speak_text_sync(self, text: str):
        """Synchronous wrapper for _speak_text to block until speaking completes"""
        print(f"DEBUG: _speak_text_sync called with text: '{text}'")

        try:
            # Set a flag to prevent overlapping speech
            if not hasattr(self, '_speaking_flag'):
                self._speaking_flag = False

            if self._speaking_flag:
                print("DEBUG: Speech already in progress, skipping sync speak")
                return

            self._speaking_flag = True
            print("DEBUG: Starting synchronous PowerShell TTS")

            # Clean text for PowerShell (remove problematic chars and punctuation)
            import re
            clean_text = re.sub(r'[^\x00-\x7F]+', '', text)  # Remove non-ASCII
            # Remove all punctuation marks (keep only letters, numbers, and spaces)
            clean_text = re.sub(r'[^\w\s]', '', clean_text)
            # Replace quotes with simpler alternatives (though punctuation is removed, keeping for safety)
            clean_text = clean_text.replace('"', ' quote ')
            clean_text = clean_text.replace("'", ' apostrophe ')

            print(f"DEBUG: Speaking text synchronously via PowerShell: '{clean_text}'")

            # Use PowerShell with a temporary file to avoid quote issues
            import tempfile
            import os

            with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False, encoding='utf-8') as temp_file:
                temp_file.write(clean_text)
                temp_file_path = temp_file.name

            try:
                # Use PowerShell to read from file and speak
                ps_command = f'powershell -Command "$text = Get-Content \'{temp_file_path}\'; Add-Type -AssemblyName System.Speech; (New-Object System.Speech.Synthesis.SpeechSynthesizer).Speak($text)"'
                print(f"DEBUG: PowerShell command: {ps_command}")

                # Start PowerShell process and wait for it to complete
                result = subprocess.run(
                    ps_command,
                    shell=True,
                    capture_output=True,
                    text=True
                )

                if result.returncode == 0:
                    print("DEBUG: Synchronous PowerShell TTS completed successfully")
                else:
                    print(f"DEBUG: Synchronous PowerShell TTS failed: {result.stderr}")

            finally:
                # Clean up temp file
                try:
                    os.unlink(temp_file_path)
                except:
                    pass

        except Exception as e:
            print(f"DEBUG: Synchronous TTS error: {str(e)}")
            self.logger.warning(f"Synchronous TTS error: {str(e)}")
        finally:
            self._speaking_flag = False
            print("DEBUG: Reset speaking flag after sync speak")

    def _stop_all_tts(self):
        """Stop all ongoing TTS processes"""
        print("DEBUG: Stopping all TTS processes")
        # Set interruption flag to stop any ongoing TTS
        self._interrupt_requested = True

        # Try to terminate any current TTS process
        if hasattr(self, '_current_tts_process') and self._current_tts_process:
            try:
                print("DEBUG: Terminating current TTS process")
                self._current_tts_process.terminate()
                # Wait for termination with timeout
                self._current_tts_process.wait(timeout=0.5)
                print("DEBUG: TTS process terminated successfully")
            except subprocess.TimeoutExpired:
                try:
                    print("DEBUG: Force killing TTS process")
                    self._current_tts_process.kill()
                    self._current_tts_process.wait(timeout=0.5)
                except:
                    print("DEBUG: Could not terminate TTS process")
            except Exception as e:
                print(f"DEBUG: Error terminating TTS process: {e}")

        # Wait a moment for threads to terminate - reduced delay
        time.sleep(0.1)

        # Reset the flag
        self._interrupt_requested = False

        # Reset speaking flag
        self._speaking_flag = False
        print("DEBUG: All TTS processes stopped")

    def _speak_text(self, text: str):
        """Speak text using Windows PowerShell TTS for guaranteed audio output"""
        print(f"DEBUG: _speak_text called with text: '{text}'")

        try:
            # Set a flag to prevent overlapping speech
            if not hasattr(self, '_speaking_flag'):
                self._speaking_flag = False

            if self._speaking_flag:
                print("DEBUG: Speech already in progress, skipping")
                return

            self._speaking_flag = True
            print("DEBUG: Starting PowerShell TTS")

            # Speak in a separate thread using PowerShell
            def speak_with_powershell():
                try:
                    # Clean text for PowerShell (remove problematic chars and punctuation)
                    import re
                    clean_text = re.sub(r'[^\x00-\x7F]+', '', text)  # Remove non-ASCII
                    # Remove all punctuation marks (keep only letters, numbers, and spaces)
                    clean_text = re.sub(r'[^\w\s]', '', clean_text)
                    # Replace quotes with simpler alternatives (though punctuation is removed, keeping for safety)
                    clean_text = clean_text.replace('"', ' quote ')
                    clean_text = clean_text.replace("'", ' apostrophe ')

                    print(f"DEBUG: Speaking text via PowerShell: '{clean_text}'")

                    # Use PowerShell with a temporary file to avoid quote issues
                    import tempfile
                    import os

                    with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False, encoding='utf-8') as temp_file:
                        temp_file.write(clean_text)
                        temp_file_path = temp_file.name

                    try:
                        # Use PowerShell to read from file and speak
                        ps_command = f'powershell -Command "$text = Get-Content \'{temp_file_path}\'; Add-Type -AssemblyName System.Speech; (New-Object System.Speech.Synthesis.SpeechSynthesizer).Speak($text)"'
                        print(f"DEBUG: PowerShell command: {ps_command}")

                        # Start PowerShell process
                        process = subprocess.Popen(
                            ps_command,
                            shell=True,
                            stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE,
                            text=True
                        )

                        # Store process reference for interruption
                        self._current_tts_process = process

                        # Monitor the process for interruption
                        while process.poll() is None:  # While process is still running
                            if self._interrupt_requested:
                                print("DEBUG: Interruption requested, terminating PowerShell TTS")
                                try:
                                    process.terminate()
                                    process.wait(timeout=0.5)  # Wait briefly for termination
                                except subprocess.TimeoutExpired:
                                    try:
                                        process.kill()  # Force kill if it doesn't terminate
                                        process.wait(timeout=0.5)
                                    except:
                                        pass  # Ignore if process is already dead
                                self._interrupt_requested = False
                                print("DEBUG: PowerShell TTS interrupted successfully")
                                break
                            time.sleep(0.02)  # Reduced delay to prevent busy waiting (was 0.05)

                        # Clear process reference
                        if hasattr(self, '_current_tts_process'):
                            delattr(self, '_current_tts_process')

                        # Get results
                        stdout, stderr = process.communicate()

                        if process.returncode == 0:
                            print("DEBUG: PowerShell TTS completed successfully")
                        elif process.returncode in [-signal.SIGTERM, -signal.SIGKILL] or self._interrupt_requested:
                            print("DEBUG: PowerShell TTS was interrupted")
                        else:
                            print(f"DEBUG: PowerShell TTS failed: {stderr}")
                            print(f"DEBUG: PowerShell stdout: {stdout}")

                    finally:
                        # Clean up temp file
                        try:
                            os.unlink(temp_file_path)
                        except:
                            pass

                except subprocess.TimeoutExpired:
                    print("DEBUG: PowerShell TTS timed out")
                except Exception as e:
                    print(f"DEBUG: PowerShell TTS failed: {str(e)}")
                    self.logger.warning(f"PowerShell TTS failed: {str(e)}")
                finally:
                    self._speaking_flag = False
                    print("DEBUG: Reset speaking flag")

            speech_thread = threading.Thread(target=speak_with_powershell, daemon=True)
            speech_thread.start()

        except Exception as e:
            print(f"DEBUG: TTS error: {str(e)}")
            self.logger.warning(f"TTS error: {str(e)}")
            self._speaking_flag = False

    def _start_command_loop(self):
        """Start the command display loop"""
        self._update_command_display()
        self.root.after(3000, self._start_command_loop)  # Update every 3 seconds

    def _update_command_display(self):
        """Update the command display with the next command"""
        if self.command_list:
            command = self.command_list[self.command_index]
            self.command_display.configure(text=f"Try saying: '{command}'")
            self.command_index = (self.command_index + 1) % len(self.command_list)

    def show_help(self):
        """Show help dialog"""
        help_text = """Available Commands:
• Open Chrome / Firefox / etc.
• Search for [query]
• Increase / Decrease volume
• Play / Pause music
• Next / Previous song
• Set alarm for [time]
• Check weather
• Tell me a joke
• What's the time?
• History - Show recent commands
• Help - Show this help
• Quit - Exit the application

Speak naturally and the AI will understand!"""
        messagebox.showinfo("Help - Available Commands", help_text)

    def change_model(self, model_name: str):
        """Change the active AI model"""
        if self.assistant.set_model(model_name):
            self.update_output(f"🤖 Switched to AI model: {model_name}")
            self.status_label.configure(text=f"Status: Ready ({model_name})", text_color="#00ff88")
        else:
            self.update_output(f"❌ Failed to switch to model: {model_name}")
            # Reset selector to current model
            current = self.assistant.get_current_model()
            if current:
                self.model_var.set(current)

    def _on_close(self):
        """Handle window close event with proper cleanup"""
        print("DEBUG: Window close requested, performing cleanup...")
        # Stop listening if active
        if self.listening:
            self.stop_listening()
        # Stop any ongoing TTS
        self._stop_all_tts()
        # Force destroy the window after a short delay to allow cleanup
        self.root.after(500, self.root.destroy)

    def run(self):
        """Run the GUI application"""
        self.root.mainloop()


def main():
    """Main entry point"""
    root = ctk.CTk()
    app = VoiceAssistantGUI(root)
    app.run()


if __name__ == "__main__":
    main()

"""
Complete Voice Assistant with STT, TTS, Gemini Flash 2.0 via Google AI Studio, and Knowledge Base
Main application file
"""

import os
import sys
import time
import threading
import logging
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv
import psutil  # For performance monitoring

# Import our modules
from voice_assistant_v2 import VoiceAssistant
from stt_module import SpeechToText, WhisperSTT, WakeWordListener
from tts_module import SmartTTS

# Local exception definitions
class InitializationError(Exception):
    def __init__(self, message, user_message=None):
        super().__init__(message)
        self.user_message = user_message

class SpeechRecognitionError(Exception):
    def __init__(self, message, user_message=None):
        super().__init__(message)
        self.user_message = user_message

class TextToSpeechError(Exception):
    def __init__(self, message, user_message=None):
        super().__init__(message)
        self.user_message = user_message

# Load environment variables
load_dotenv()

class VoiceAssistantApp:
    """Complete voice assistant application"""

    def __init__(self, use_whisper: bool = False):
        """
        Initialize the voice assistant

        Args:
            use_whisper: If True, use Whisper STT instead of Google
        """
        self.logger = logging.getLogger('main_app')

        # Check for API key
        api_key = os.getenv('OPENROUTER_API_KEY')
        if not api_key:
            self.logger.warning("OPENROUTER_API_KEY not found in environment variables")

        # Initialize components

        # 1. Knowledge Base & Gemini
        try:
            self.logger.debug("Initializing voice assistant...")
            self.assistant = VoiceAssistant()
            self.logger.info("Voice assistant initialized successfully")
        except InitializationError as e:
            self.logger.error(f"Voice assistant initialization failed: {str(e)}", exc_info=True)
            raise
        except Exception as e:
            error_msg = f"Unexpected error initializing voice assistant: {str(e)}"
            self.logger.error(error_msg, exc_info=True)
            raise InitializationError(error_msg, "Failed to initialize the voice assistant core.")

        # 2. Speech-to-Text
        try:
            self.logger.debug("Initializing speech-to-text...")
            if use_whisper:
                self.stt = WhisperSTT(model_size="base")
                self.logger.info("Whisper STT initialized successfully")
            else:
                self.stt = SpeechToText()
                self.logger.info("Google STT initialized successfully")
        except SpeechRecognitionError as e:
            self.logger.warning(f"Primary STT initialization failed: {str(e)}, trying fallback...")
            try:
                self.stt = SpeechToText()
                self.logger.info("Fallback STT (Google) initialized successfully")
            except Exception as fallback_e:
                error_msg = f"Both STT engines failed: Primary - {str(e)}, Fallback - {str(fallback_e)}"
                self.logger.error(error_msg, exc_info=True)
                raise InitializationError(error_msg, "Unable to initialize speech recognition. Please check your microphone and internet connection.")
        except Exception as e:
            error_msg = f"Unexpected error initializing STT: {str(e)}"
            self.logger.error(error_msg, exc_info=True)
            raise InitializationError(error_msg, "Failed to initialize speech recognition system.")

        # 3. Text-to-Speech
        try:
            self.logger.debug("Initializing text-to-speech...")
            self.tts = SmartTTS()
            self.logger.info("Text-to-speech initialized successfully")
        except TextToSpeechError as e:
            self.logger.warning(f"TTS initialization failed: {str(e)}, continuing without TTS")
            self.tts = None
        except Exception as e:
            error_msg = f"Unexpected error initializing TTS: {str(e)}"
            self.logger.warning(error_msg + ", continuing without TTS")
            self.tts = None

        # 4. Control flags
        self.running = False
        self.listening = False
        self.wake_word_active = False
        
        # 5. Wake word listener - initialized with shared microphone from STT
        self.wake_word_listener = None
        try:
            self.logger.debug("Initializing wake word listener...")
            # Share the microphone with the main STT to avoid conflicts
            self.wake_word_listener = WakeWordListener(
                callback=self._on_wake_word_detected,
                microphone=self.stt.microphone if hasattr(self.stt, 'microphone') else None
            )
            self.logger.info("Wake word listener initialized successfully")
        except Exception as e:
            self.logger.warning(f"Wake word listener initialization failed: {e}")
            self.wake_word_listener = None

        # 6. Performance monitoring
        self.performance_stats = {
            'start_time': time.time(),
            'commands_processed': 0,
            'stt_latencies': [],
            'tts_latencies': [],
            'total_processing_time': 0
        }

        self.logger.info("Voice assistant application initialized successfully")
        self._log_performance_stats()  # Log initial performance stats

        # Welcome message with TTS
        try:
            if self.tts:
                self.tts.speak("Voice assistant initialized and ready to help!")
        except TextToSpeechError as e:
            self.logger.warning(f"TTS failed for welcome message: {str(e)}")

        self._print_help()
        
        # Start wake word listener if available
        if self.wake_word_listener:
            self._start_wake_word_listener()

    def _log_performance_stats(self):
        """Log current performance statistics"""
        stats = self.performance_stats
        self.logger.info(f"Performance Stats - Start Time: {datetime.fromtimestamp(stats['start_time'])}, Commands Processed: {stats['commands_processed']}, Total Processing Time: {stats['total_processing_time']:.2f}s")

    def _setup_interruption_handling(self):
        """Set up interruption handling for TTS"""
        if self.tts and self.stt:
            # Define interruption callback
            def interruption_callback():
                if self.tts.is_speaking():
                    print("\n🛑 Speech interrupted by user")
                    self.tts.interrupt_speech()
                    self.stt.stop_interruption_listener()

            self.interruption_callback = interruption_callback

    def _on_wake_word_detected(self, wake_word_text: str):
        """Callback when wake word is detected"""
        self.logger.info(f"Wake word detected: '{wake_word_text}'")
        self.wake_word_active = True
        
        # Pause wake word listener while processing command
        if self.wake_word_listener and self.wake_word_listener.is_running():
            self.wake_word_listener.stop()
        
        # Play a sound or speak to acknowledge
        try:
            if self.tts:
                self.tts.speak("Yes? I'm listening.")
        except TextToSpeechError as e:
            self.logger.warning(f"TTS failed for wake word acknowledgment: {str(e)}")
        
        # Small delay to let microphone settle
        time.sleep(0.5)
        
        # Listen for the actual command
        try:
            print("\n🎤 Wake word detected! Listening for your command...")
            text = self.stt.listen(timeout=5, phrase_time_limit=10)
            if text:
                print(f"📝 You said: {text}")
                self.process_command(text)
            else:
                print("❌ No command detected after wake word")
                try:
                    if self.tts:
                        self.tts.speak("I didn't hear a command. Please try again.")
                except TextToSpeechError as e:
                    self.logger.warning(f"TTS failed for no command message: {str(e)}")
        except Exception as e:
            self.logger.error(f"Error processing command after wake word: {e}")
        finally:
            self.wake_word_active = False
            # Resume wake word listener
            time.sleep(0.5)
            if self.wake_word_listener and not self.wake_word_listener.is_running():
                self.wake_word_listener.start()
    
    def _start_wake_word_listener(self):
        """Start the wake word listener"""
        if self.wake_word_listener and not self.wake_word_listener.is_running():
            self.wake_word_listener.start()
            print("\n🎧 Wake word listener active - Say 'Hey Nova' or 'Nova' to activate")
            try:
                if self.tts:
                    self.tts.speak("Say Hey Nova to wake me up")
            except TextToSpeechError as e:
                self.logger.warning(f"TTS failed for wake word hint: {str(e)}")
    
    def _stop_wake_word_listener(self):
        """Stop the wake word listener"""
        if self.wake_word_listener and self.wake_word_listener.is_running():
            self.wake_word_listener.stop()
            print("\n🎧 Wake word listener stopped")

    def _print_help(self):
        """Print usage instructions"""
        help_text = """HOW TO USE
Commands:
  • Say 'Hey Nova' or 'Nova' to activate voice input
  • Press ENTER to activate voice input
  • Say 'quit' or 'exit' to stop
  • Type 'history' to see command history
  • Type 'help' to see this message

Example commands:
  • 'Open Chrome'
  • 'Search for Python tutorials'
  • 'Open Notepad'
  • 'Increase volume'
  • 'What's the weather like?'"""

        print("\n" + "=" * 60)
        print("📖 " + help_text)
        print("=" * 60)

        # Speak the help information
        try:
            if self.tts:
                self.tts.speak("Here are the instructions for using the voice assistant. " + help_text.replace('\n', ' ').replace('•', ''))
        except TextToSpeechError as e:
            self.logger.warning(f"TTS failed for help message: {str(e)}")
    
    def process_command(self, command: str):
        """Process a single command"""
        if not command:
            return

        command = command.strip().lower()

        # Check for exit commands
        if command in ['quit', 'exit', 'stop', 'bye']:
            print("\n👋 Goodbye!")
            try:
                if self.tts:
                    self.tts.speak("Goodbye! Have a great day!")
            except TextToSpeechError as e:
                self.logger.warning(f"TTS failed for goodbye message: {str(e)}")
            self.running = False
            return

        # Check for special commands
        if command == 'help':
            self._print_help()
            return

        if command == 'history':
            try:
                self._show_history()
            except Exception as e:
                self.logger.error(f"Failed to show history: {str(e)}", exc_info=True)
                print("\n❌ Error retrieving command history")
            return

        # Process through assistant
        print(f"\n🔄 Processing command...")
        try:
            if self.tts:
                self.tts.speak("Processing your command")
        except TextToSpeechError as e:
            self.logger.warning(f"TTS failed for processing message: {str(e)}")

        try:
            result = self.assistant.process_command(command)

            if result['success']:
                print(f"✅ {result['message']}")
                response = result.get('response', result['message'])
                try:
                    if self.tts:
                        # Start interruption listener before speaking
                        self.stt.start_interruption_listener(self.tts.interrupt_speech)
                        self.tts.speak(response)
                        # Stop interruption listener after speaking
                        self.stt.stop_interruption_listener()
                except TextToSpeechError as e:
                    self.logger.warning(f"TTS failed for response: {str(e)}")
                    self.stt.stop_interruption_listener()
            else:
                print(f"❌ {result['message']}")
                try:
                    if self.tts:
                        # Start interruption listener before speaking
                        self.stt.start_interruption_listener(self.tts.interrupt_speech)
                        self.tts.speak("Sorry, I couldn't complete that command")
                        # Stop interruption listener after speaking
                        self.stt.stop_interruption_listener()
                except TextToSpeechError as e:
                    self.logger.warning(f"TTS failed for error message: {str(e)}")
                    self.stt.stop_interruption_listener()

        except Exception as e:
            error_msg = f"Unexpected error processing command '{command}': {str(e)}"
            self.logger.error(error_msg, exc_info=True)
            print(f"\n❌ Error processing command: {str(e)}")
            try:
                if self.tts:
                    # Start interruption listener before speaking
                    self.stt.start_interruption_listener(self.tts.interrupt_speech)
                    self.tts.speak("Sorry, I encountered an error processing your command")
                    # Stop interruption listener after speaking
                    self.stt.stop_interruption_listener()
            except TextToSpeechError as tts_e:
                self.logger.warning(f"TTS failed for error response: {str(tts_e)}")
                self.stt.stop_interruption_listener()
    
    def _show_history(self):
        """Show command history"""
        history = self.assistant.get_history(limit=10)
        
        if not history:
            print("\n📝 No command history yet")
            return
        
        print("\n📝 Recent Commands:")
        print("-" * 60)
        for i, entry in enumerate(history, 1):
            timestamp = entry['timestamp']
            user_input = entry['input']
            success = "✅" if entry['success'] else "❌"
            print(f"{i}. [{timestamp}] {success} {user_input}")
        print("-" * 60)
    
    def run_interactive(self):
        """Run in interactive mode (press enter to speak)"""
        self.running = True

        print("\n🎤 Interactive Mode")
        print("Press ENTER to start listening, or type a command\n")

        while self.running:
            try:
                user_input = input(">>> ").strip()

                if not user_input:
                    # Empty input = voice mode
                    print("\n🎤 Listening... (speak now)")
                    try:
                        if self.tts:
                            self.tts.speak("I'm listening. Please speak your command.")
                    except TextToSpeechError as e:
                        self.logger.warning(f"TTS failed for listening prompt: {str(e)}")

                    try:
                        text = self.stt.listen(timeout=5, phrase_time_limit=10)

                        if text:
                            self.process_command(text)
                        else:
                            print("❌ No speech detected")
                            try:
                                if self.tts:
                                    self.tts.speak("I didn't hear anything. Please try again.")
                            except TextToSpeechError as e:
                                self.logger.warning(f"TTS failed for no speech message: {str(e)}")
                    except SpeechRecognitionError as e:
                        self.logger.error(f"Speech recognition failed: {str(e)}", exc_info=True)
                        print(f"❌ Speech recognition error: {str(e)}")
                        try:
                            if self.tts:
                                self.tts.speak("Sorry, I had trouble hearing you. Please try again.")
                        except TextToSpeechError as tts_e:
                            self.logger.warning(f"TTS failed for speech error message: {str(tts_e)}")
                    except Exception as e:
                        error_msg = f"Unexpected error during speech input: {str(e)}"
                        self.logger.error(error_msg, exc_info=True)
                        print(f"❌ Unexpected error: {str(e)}")
                        try:
                            if self.tts:
                                self.tts.speak("There was an unexpected error. Please try again.")
                        except TextToSpeechError as tts_e:
                            self.logger.warning(f"TTS failed for unexpected error message: {str(tts_e)}")
                else:
                    # Text input
                    try:
                        if self.tts:
                            self.tts.speak(f"You typed: {user_input}")
                    except TextToSpeechError as e:
                        self.logger.warning(f"TTS failed for text input echo: {str(e)}")
                    self.process_command(user_input)

                print()  # Blank line for readability

            except KeyboardInterrupt:
                print("\n\n👋 Shutting down...")
                self.running = False
                break

            except Exception as e:
                error_msg = f"Unexpected error in interactive mode: {str(e)}"
                self.logger.error(error_msg, exc_info=True)
                print(f"\n❌ Error: {str(e)}")

        self.cleanup()
    

    
    def cleanup(self):
        """Cleanup resources"""
        print("\n🧹 Cleaning up...")
        self._stop_wake_word_listener()
        self.assistant.cleanup()
        print("✅ Cleanup complete")


def main():
    """Main entry point"""
    import argparse

    # Set up basic logging for main function
    import logging
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    parser = argparse.ArgumentParser(description='AI Voice Assistant with Gemini')
    parser.add_argument(
        '--whisper',
        action='store_true',
        help='Use Whisper STT instead of Google (better accuracy, slower)'
    )

    args = parser.parse_args()

    try:
        logging.info("Starting voice assistant application...")
        # Create assistant
        app = VoiceAssistantApp(use_whisper=args.whisper)

        # Run in interactive mode only
        logging.info("Starting interactive mode")
        app.run_interactive()

    except KeyboardInterrupt:
        logging.info("Application interrupted by user")
        print("\n👋 Goodbye!")
        sys.exit(0)

    except InitializationError as e:
        logging.error(f"Application initialization failed: {str(e)}", exc_info=True)
        print(f"\n❌ Initialization failed: {str(e)}")
        sys.exit(1)

    except Exception as e:
        logging.error(f"Fatal error in main application: {str(e)}", exc_info=True)
        print(f"\n❌ Fatal error: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()

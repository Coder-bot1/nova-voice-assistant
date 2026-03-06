"""
Speech-to-Text Module
Handles voice input and converts to text
"""

import speech_recognition as sr
from typing import Optional
import threading
import time
import queue
import logging

# Local exception definitions
class SpeechRecognitionError(Exception):
    def __init__(self, message, user_message=None):
        super().__init__(message)
        self.user_message = user_message

class AudioError(Exception):
    def __init__(self, message, user_message=None):
        super().__init__(message)
        self.user_message = user_message

class NetworkError(Exception):
    def __init__(self, message, user_message=None):
        super().__init__(message)
        self.user_message = user_message

class ConfigurationError(Exception):
    def __init__(self, message, user_message=None):
        super().__init__(message)
        self.user_message = user_message

class SpeechToText:
    """Handle speech recognition"""

    def __init__(self):
        self.logger = logging.getLogger('stt')
        self.recognizer = sr.Recognizer()
        # MICROPHONE ENABLED
        self.microphone = sr.Microphone()

        # Initialize interruption listener attributes
        self.interruption_running = False
        self.interruption_callback = None
        self.interruption_thread = None
        self.interruption_keywords = ["stop", "cancel", "quit", "shut up", "be quiet"]

        # Adjust for ambient noise - optimized for speed
        self.logger.info("Initializing microphone...")
        try:
            with self.microphone as source:
                # Reduced duration for faster startup (0.05s instead of 0.1s)
                self.recognizer.adjust_for_ambient_noise(source, duration=0.05)

                # Try to pre-warm microphone (optional - don't fail if no audio available)
                try:
                    self.recognizer.listen(source, timeout=0.1, phrase_time_limit=0.1)
                    self.logger.info("Microphone pre-warmed successfully")
                except Exception as prewarm_error:
                    self.logger.warning(f"Microphone pre-warming failed (expected in test environments): {prewarm_error}")
                    # Don't raise error - pre-warming is optional



            self.logger.info("Microphone ready!")
        except Exception as e:
            error_msg = f"Microphone initialization failed: {str(e)}"
            self.logger.warning(error_msg)
            raise AudioError(error_msg, "Please check your microphone connection and permissions.")



    def listen(self, timeout: int = 1, phrase_time_limit: int = 8) -> Optional[str]:
        """
        Listen for voice input and convert to text

        Args:
            timeout: Seconds to wait for speech to start (reduced for faster response)
            phrase_time_limit: Maximum seconds for a phrase

        Returns:
            Recognized text or None
        """
        try:
            with self.microphone as source:
                audio = self.recognizer.listen(
                    source,
                    timeout=timeout,
                    phrase_time_limit=phrase_time_limit
                )

            # Try Google Speech Recognition (free, no API key needed)
            try:
                text = self.recognizer.recognize_google(audio)
                return text

            except sr.UnknownValueError:
                return None

            except sr.RequestError as e:
                error_msg = f"Speech recognition service error: {str(e)}"
                raise NetworkError(error_msg, "Speech recognition service is currently unavailable. Please check your internet connection.")

        except sr.WaitTimeoutError:
            return None

        except Exception as e:
            error_msg = f"Unexpected error during speech recognition: {str(e)}"
            raise SpeechRecognitionError(error_msg, "There was a problem with speech recognition. Please try again.")
    
    def start_interruption_listener(self, callback):
        """
        Start listening for interruption keywords in a separate thread.
        Currently disabled to prevent conflicts with main speech recognition.
        """
        # Interruption listener is disabled to avoid microphone conflicts
        # The callback is stored but not used in the current implementation
        self.interruption_callback = callback
        self.logger.debug("Interruption listener is disabled to prevent microphone conflicts")

    def stop_interruption_listener(self):
        """
        Stop the interruption listener thread.
        Currently a no-op since interruption listener is disabled.
        """
        self.interruption_callback = None






class WhisperSTT:
    """
    Alternative STT using Faster Whisper for better accuracy
    Requires: faster-whisper library
    """

    def __init__(self, model_size: str = "base"):
        """
        Initialize Whisper model

        Args:
            model_size: tiny, base, small, medium, large
        """
        try:
            from faster_whisper import WhisperModel

            self.model = WhisperModel(model_size, device="cpu", compute_type="int8")
            self.microphone = sr.Microphone()
            self.recognizer = sr.Recognizer()

        except ImportError as e:
            error_msg = "faster-whisper library not installed"
            raise ConfigurationError(error_msg, "Please install the faster-whisper library to use Whisper STT.")
        except Exception as e:
            error_msg = f"Failed to initialize Whisper model: {str(e)}"
            raise ConfigurationError(error_msg, "Failed to load Whisper model. Please check your installation.")
    
    def listen(self, timeout: int = 5, phrase_time_limit: int = 10) -> Optional[str]:
        """Listen and transcribe using Whisper"""
        try:
            with self.microphone as source:
                print("🎤 Listening...")
                audio = self.recognizer.listen(
                    source,
                    timeout=timeout,
                    phrase_time_limit=phrase_time_limit
                )
            
            # Convert audio to WAV format for Whisper
            import io
            import wave
            
            wav_data = io.BytesIO()
            with wave.open(wav_data, 'wb') as wav_file:
                wav_file.setnchannels(1)
                wav_file.setsampwidth(2)
                wav_file.setframerate(16000)
                wav_file.writeframes(audio.get_wav_data())
            
            wav_data.seek(0)
            
            # Transcribe with Whisper
            segments, info = self.model.transcribe(wav_data, beam_size=5)
            
            text = " ".join([segment.text for segment in segments])
            
            if text.strip():
                print(f"✅ Recognized: {text}")
                return text.strip()
            
            return None
        
        except sr.WaitTimeoutError:
            print("⏱️ Listening timed out")
            return None
        
        except Exception as e:
            print(f"❌ Error: {e}")
            return None


class WakeWordListener:
    """
    Continuously listens for wake words and triggers callback when detected.
    Wake words: "hello nova", "hi nova", "hey nova", "nova"
    Note: This requires exclusive microphone access and may conflict with other STT instances.
    """
    
    def __init__(self, callback=None, microphone=None):
        self.logger = logging.getLogger('wake_word')
        self.recognizer = sr.Recognizer()
        # Use provided microphone or create new one
        self.microphone = microphone or sr.Microphone()
        self.callback = callback
        self.listening = False
        self.listen_thread = None
        self._owns_microphone = microphone is None
        
        # Wake word patterns (will match variations)
        self.wake_words = [
            'hello nova',
            'hi nova', 
            'hey nova',
            'nova'
        ]
        
        # Note: Ambient noise adjustment is now done in the listen loop
        self.logger.info("Wake word listener initialized and ready")
    
    def _check_wake_word(self, text: str) -> bool:
        """Check if the recognized text contains a wake word"""
        if not text:
            return False
        text_lower = text.lower().strip()
        return any(wake_word in text_lower for wake_word in self.wake_words)
    
    def _listen_loop(self):
        """Main listening loop for wake words"""
        self.logger.info("Wake word listener started - Say 'Hey Nova' or 'Nova' to activate")
        
        # Open microphone once and keep it open for the duration
        try:
            with self.microphone as source:
                # Adjust for ambient noise at start - optimized for speed
                self.recognizer.adjust_for_ambient_noise(source, duration=0.2)
                
                while self.listening:
                    try:
                        # Listen with short timeout for responsive checking
                        audio = self.recognizer.listen(
                            source, 
                            timeout=0.5,  # Reduced timeout to check for stop frequently (was 1)
                            phrase_time_limit=3  # Max 3 seconds for wake word phrase (was 5)
                        )
                        
                        # Try to recognize in a separate try block
                        try:
                            text = self.recognizer.recognize_google(audio)
                            self.logger.debug(f"Wake word listener heard: '{text}'")
                            
                            if self._check_wake_word(text):
                                self.logger.info(f"✅ Wake word detected: '{text}'")
                                
                                # Stop listening FIRST to prevent multiple detections
                                self.listening = False
                                
                                # Run callback in separate thread to avoid blocking
                                if self.callback:
                                    callback_thread = threading.Thread(
                                        target=self.callback, 
                                        args=(text,),
                                        daemon=True
                                    )
                                    callback_thread.start()
                                
                                # Exit the loop immediately
                                break
                                
                        except sr.UnknownValueError:
                            # Didn't understand - that's fine, keep listening
                            pass
                        except sr.RequestError as e:
                            # Network error - log and continue
                            self.logger.warning(f"Network error in wake word detection: {e}")
                            time.sleep(1)
                            
                    except sr.WaitTimeoutError:
                        # Timeout - just continue listening (this is expected)
                        continue
                    except Exception as e:
                        self.logger.error(f"Error in wake word listen cycle: {e}")
                        time.sleep(0.5)
                        
        except Exception as e:
            self.logger.error(f"Fatal error in wake word listen loop: {e}")
        finally:
            self.logger.info("Wake word listen loop ended")
    
    def start(self):
        """Start listening for wake words in background thread"""
        if self.listening:
            self.logger.warning("Wake word listener already running")
            return
        
        self.listening = True
        self.listen_thread = threading.Thread(target=self._listen_loop, daemon=True)
        self.listen_thread.start()
        self.logger.info("Wake word listener started")
    
    def stop(self):
        """Stop the wake word listener"""
        self.listening = False
        if self.listen_thread and self.listen_thread.is_alive():
            # Don't wait too long - the thread may be blocked on microphone access
            self.listen_thread.join(timeout=1)
            if self.listen_thread.is_alive():
                self.logger.warning("Wake word listener thread did not stop gracefully")
        self.logger.info("Wake word listener stopped")
    
    def is_running(self) -> bool:
        """Check if wake word listener is running"""
        return self.listening and self.listen_thread and self.listen_thread.is_alive()


# Test the module
if __name__ == "__main__":
    print("Testing Speech-to-Text...")
    
    # Test Google Speech Recognition
    stt = SpeechToText()
    
    print("\nSay something:")
    text = stt.listen()
    
    if text:
        print(f"\n✅ You said: {text}")
    else:
        print("\n❌ No speech detected")

"""
Text-to-Speech Module
Handles voice output with multiple TTS engines
"""

import pyttsx3
import threading
import queue
import time
import subprocess
import logging
from typing import Optional
from pathlib import Path

# Local exception definitions
class TextToSpeechError(Exception):
    def __init__(self, message, user_message=None):
        super().__init__(message)
        self.user_message = user_message

class ConfigurationError(Exception):
    def __init__(self, message, user_message=None):
        super().__init__(message)
        self.user_message = user_message

class AudioError(Exception):
    def __init__(self, message, user_message=None):
        super().__init__(message)
        self.user_message = user_message

class TextToSpeech:
    """Cross-platform TTS using pyttsx3"""

    def __init__(self, rate: int = 200, volume: float = 0.8):
        """
        Initialize TTS engine with optimized settings for speed

        Args:
            rate: Speech rate (words per minute) - increased for faster response
            volume: Volume level (0.0 to 1.0) - slightly reduced for performance
        """
        self.logger = logging.getLogger('tts')

        # Cache for cleaned text to avoid repeated processing
        self._clean_text_cache = {}
        self._cache_max_size = 100  # Limit cache size to prevent memory issues

        try:
            self.engine = pyttsx3.init()
            self.engine.setProperty('rate', rate)  # Increased from 175 to 200 for faster speech
            self.engine.setProperty('volume', volume)  # Reduced from 0.9 to 0.8 for better performance

            # Get available voices - simplified selection for speed
            voices = self.engine.getProperty('voices')
            if voices:
                # Prefer female voice if available, but don't search all voices
                for voice in voices[:3]:  # Only check first 3 voices for speed
                    if 'female' in voice.name.lower() or 'zira' in voice.name.lower():
                        self.engine.setProperty('voice', voice.id)
                        break

            # Pre-warm TTS engine with a silent test
            try:
                self.engine.say("")
                self.engine.runAndWait()
                self.logger.debug("TTS engine pre-warmed successfully")
            except:
                pass  # Pre-warming might fail on some systems, continue anyway

            # Initialize speech queue and worker thread
            self.speech_queue = queue.Queue(maxsize=3)  # Limit queue size to prevent backlog
            self.running = True
            self.speaking = False  # Track if currently speaking
            self.interrupt_event = threading.Event()  # Event to signal interruption
            self.worker_thread = threading.Thread(target=self._speech_worker, daemon=True)
            self.worker_thread.start()

        except Exception as e:
            error_msg = f"Failed to initialize TTS engine: {str(e)}"
            raise TextToSpeechError(error_msg, "Text-to-speech engine could not be initialized. Please check your system audio setup.")

    def stop(self):
        """Stop the TTS engine and worker thread"""
        try:
            self.running = False
            self.interrupt_event.set()  # Signal interruption
            # Wait for worker thread to finish
            if self.worker_thread.is_alive():
                self.worker_thread.join(timeout=2.0)
            self.engine.stop()
        except Exception as e:
            self.logger.warning(f"Error stopping TTS engine: {str(e)}", exc_info=True)

    def interrupt_speech(self):
        """Interrupt current speech immediately"""
        try:
            self.interrupt_event.set()
            self.engine.stop()
            self.speaking = False
            self.logger.debug("Speech interrupted")
        except Exception as e:
            self.logger.warning(f"Error interrupting speech: {str(e)}", exc_info=True)

    def is_speaking(self) -> bool:
        """Check if TTS is currently speaking"""
        return self.speaking

    def speak(self, text: str):
        """
        Convert text to speech asynchronously

        Args:
            text: Text to speak
        """
        if not text:
            return

        # Enqueue speech request for the worker thread
        self.speech_queue.put(text)
    
    def _speech_worker(self):
        """Worker thread that processes speech requests from the queue"""
        while self.running:
            try:
                # Get next speech request from queue (blocking)
                text = self.speech_queue.get(timeout=1.0)

                if text:
                    self._speak(text)

                # Mark task as done
                self.speech_queue.task_done()

            except queue.Empty:
                # No speech requests, continue waiting
                continue
            except Exception as e:
                error_msg = f"TTS worker error: {str(e)}"
                self.logger.error(error_msg, exc_info=True)

    def _speak(self, text: str):
        """Internal speak method with interruption support"""
        try:
            # Clean text for logging (remove emojis that cause encoding issues)
            clean_text = self._clean_text_for_logging(text)
            self.logger.debug(f"Speaking text: {clean_text[:50]}...")
            self.speaking = True
            self.interrupt_event.clear()  # Reset interrupt event

            # Start speech
            self.engine.say(text)
            self.engine.runAndWait()

            # Check if interrupted after speech completes
            if self.interrupt_event.is_set():
                self.logger.debug("Speech was interrupted")
            else:
                self.logger.debug("Speech completed successfully")

        except Exception as e:
            error_msg = f"TTS speech failed: {str(e)}"
            self.logger.error(error_msg, exc_info=True)
            raise TextToSpeechError(error_msg, "Failed to speak the text. Please check your audio setup.")
        finally:
            self.speaking = False

    def _clean_text_for_logging(self, text: str) -> str:
        """Remove emojis and problematic Unicode characters for logging with caching"""
        # Check cache first
        if text in self._clean_text_cache:
            return self._clean_text_cache[text]

        import re
        # Remove emoji characters (Unicode ranges for emojis)
        emoji_pattern = re.compile(
            "["
            "\U0001F600-\U0001F64F"  # emoticons
            "\U0001F300-\U0001F5FF"  # symbols & pictographs
            "\U0001F680-\U0001F6FF"  # transport & map symbols
            "\U0001F1E0-\U0001F1FF"  # flags (iOS)
            "\U00002700-\U000027BF"  # dingbats
            "\U0001f926-\U0001f937"  # gestures
            "\U00010000-\U0010ffff"  # other unicode
            "\u2640-\u2642"  # gender symbols
            "\u2600-\u2B55"  # misc symbols
            "\u200d"  # zero width joiner
            "\u23cf"  # eject symbol
            "\u23e9"  # fast forward
            "\u231a"  # watch
            "\ufe0f"  # variation selector
            "\u3030"  # wavy dash
            "]+",
            flags=re.UNICODE
        )
        cleaned = emoji_pattern.sub('', text)

        # Cache the result (with size limit)
        if len(self._clean_text_cache) < self._cache_max_size:
            self._clean_text_cache[text] = cleaned
        elif len(self._clean_text_cache) >= self._cache_max_size:
            # Remove oldest entry when cache is full (simple FIFO)
            oldest_key = next(iter(self._clean_text_cache))
            del self._clean_text_cache[oldest_key]
            self._clean_text_cache[text] = cleaned

        return cleaned
    
    def set_voice(self, voice_index: int = 0):
        """Change voice"""
        voices = self.engine.getProperty('voices')
        if 0 <= voice_index < len(voices):
            self.engine.setProperty('voice', voices[voice_index].id)
    
    def list_voices(self):
        """List available voices"""
        voices = self.engine.getProperty('voices')
        print("\n🎙️ Available Voices:")
        for i, voice in enumerate(voices):
            print(f"  {i}: {voice.name} ({voice.languages})")
    
    def set_rate(self, rate: int):
        """Change speech rate"""
        self.engine.setProperty('rate', rate)
    
    def set_volume(self, volume: float):
        """Change volume (0.0 to 1.0)"""
        self.engine.setProperty('volume', max(0.0, min(1.0, volume)))


class PiperTTS:
    """
    High-quality local TTS using Piper
    Requires: piper-tts

    Note: Piper TTS is more complex to set up but provides better quality
    """

    def __init__(self, model_path: Optional[str] = None):
        """
        Initialize Piper TTS

        Args:
            model_path: Path to Piper model file
        """
        self.logger = logging.getLogger('piper_tts')

        try:
            import subprocess
            import shutil

            # Check if piper is installed
            if not shutil.which('piper'):
                error_msg = "Piper TTS not found. Please install: pip install piper-tts"
                self.logger.error(error_msg)
                raise ConfigurationError(error_msg, "Piper TTS is not installed. Please install piper-tts package.")

            self.piper_path = shutil.which('piper')
            self.model_path = model_path

            self.logger.info("Piper TTS initialized successfully")

        except ConfigurationError:
            raise
        except Exception as e:
            error_msg = f"Piper initialization failed: {str(e)}"
            self.logger.error(error_msg)
            raise ConfigurationError(error_msg, "Failed to initialize Piper TTS. Please check your installation.")
    
    def speak(self, text: str, output_file: Optional[str] = None):
        """
        Speak text using Piper

        Args:
            text: Text to speak
            output_file: Optional WAV file to save output
        """
        try:
            import subprocess
            import tempfile

            if output_file is None:
                # Create temporary file
                output_file = tempfile.mktemp(suffix='.wav')

            self.logger.debug(f"Generating speech for text: {text[:50]}...")

            # Run piper
            cmd = [self.piper_path, '--output_file', output_file]
            if self.model_path:
                cmd.extend(['--model', self.model_path])

            process = subprocess.Popen(
                cmd,
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )

            stdout, stderr = process.communicate(input=text)

            if process.returncode != 0:
                error_msg = f"Piper TTS process failed with return code {process.returncode}: {stderr}"
                self.logger.error(error_msg)
                raise TextToSpeechError(error_msg, "Failed to generate speech. Please check Piper TTS installation.")

            # Play the audio file
            self._play_audio(output_file)
            self.logger.debug("Piper TTS speech completed successfully")

        except Exception as e:
            error_msg = f"Piper TTS error: {str(e)}"
            self.logger.error(error_msg, exc_info=True)
            raise TextToSpeechError(error_msg, "Failed to speak text using Piper TTS.")
    
    def _play_audio(self, wav_file: str):
        """Play WAV audio file"""
        try:
            import subprocess
            import platform

            system = platform.system()
            self.logger.debug(f"Playing audio file: {wav_file} on {system}")

            if system == 'Windows':
                import winsound
                winsound.PlaySound(wav_file, winsound.SND_FILENAME)
            elif system == 'Darwin':  # macOS
                result = subprocess.run(['afplay', wav_file], capture_output=True, text=True)
                if result.returncode != 0:
                    raise subprocess.SubprocessError(f"afplay failed: {result.stderr}")
            else:  # Linux
                result = subprocess.run(['aplay', wav_file], capture_output=True, text=True)
                if result.returncode != 0:
                    raise subprocess.SubprocessError(f"aplay failed: {result.stderr}")

            self.logger.debug("Audio playback completed successfully")

        except Exception as e:
            error_msg = f"Audio playback error: {str(e)}"
            self.logger.error(error_msg, exc_info=True)
            raise AudioError(error_msg, "Failed to play the audio file. Please check your audio setup.")


class SmartTTS:
    """
    Smart TTS selector - automatically uses best available TTS with queue support
    """

    def __init__(self):
        """Initialize best available TTS engine"""
        self.logger = logging.getLogger('tts')
        self.engine = None
        self.engine_type = None

        # Try to initialize TTS engines in order of preference
        try:
            self.logger.info("Initializing SmartTTS - trying pyttsx3...")
            self.engine = TextToSpeech()
            self.engine_type = "pyttsx3"
            self.logger.info("SmartTTS initialized successfully with pyttsx3")
        except Exception as e:
            error_msg = f"Failed to initialize any TTS engine: {str(e)}"
            self.logger.warning(error_msg)
            self.logger.info("SmartTTS will fallback to text-only output")

    def speak(self, text: str):
        """Speak text using available engine"""
        if self.engine:
            try:
                self.engine.speak(text)
            except Exception as e:
                error_msg = f"TTS speak failed, falling back to text output: {str(e)}"
                self.logger.error(error_msg, exc_info=True)
                print(f"🔊 [TTS]: {text}")
        else:
            print(f"🔊 [TTS]: {text}")

    def is_available(self) -> bool:
        """Check if TTS is available"""
        return self.engine is not None

    def stop(self):
        """Stop the TTS engine and queue processor"""
        if self.engine and hasattr(self.engine, 'stop'):
            try:
                self.engine.stop()
                self.logger.debug("TTS engine stopped successfully")
            except Exception as e:
                self.logger.error(f"Error stopping TTS engine: {str(e)}", exc_info=True)

    def interrupt_speech(self):
        """Interrupt current speech immediately"""
        if self.engine and hasattr(self.engine, 'interrupt_speech'):
            try:
                self.engine.interrupt_speech()
            except Exception as e:
                self.logger.error(f"Error interrupting speech: {str(e)}", exc_info=True)

    def is_speaking(self) -> bool:
        """Check if TTS is currently speaking"""
        if self.engine and hasattr(self.engine, 'is_speaking'):
            return self.engine.is_speaking()
        return False


# Test the module
if __name__ == "__main__":
    print("Testing Text-to-Speech...")
    
    # Test pyttsx3
    tts = TextToSpeech()
    
    # List available voices
    tts.list_voices()
    
    # Test speech
    print("\n🔊 Testing speech output...")
    tts.speak("Hello! I am your AI voice assistant. How can I help you today?")
    
    # Test async speech
    print("\n🔊 Testing async speech...")
    tts.speak("This is an asynchronous speech test.")
    
    import time
    time.sleep(3)
    
    print("\n✅ TTS test complete!")

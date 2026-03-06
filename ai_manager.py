"""
AI Manager for Voice Assistant
Supports multiple AI models: Gemini and OpenAI GPT
WITH ENHANCED FALLBACK PARSING FOR ALL COMMAND TYPES
"""

import os
import json
import re
import time
import logging
from typing import Dict, Any, Optional, List
from abc import ABC, abstractmethod
import google.generativeai as genai
from dotenv import load_dotenv


def extract_json_safely(text: str) -> Optional[Dict[str, Any]]:
    """
    Extract JSON safely from text using regex.
    Never parse raw response directly - use this function instead.
    
    Args:
        text: The text containing potential JSON data
        
    Returns:
        Parsed JSON dict if found, None otherwise
    """
    try:
        # Try to find JSON object using regex
        match = re.search(r'\{.*\}', text, re.S)
        if match:
            return json.loads(match.group())
    except (json.JSONDecodeError, ValueError):
        pass
    return None

# Load environment variables
load_dotenv()
from enhanced_parser import EnhancedFallbackParser

# Local exception definitions
class APIError(Exception):
    def __init__(self, message, user_message=None):
        super().__init__(message)
        self.user_message = user_message

class ModelError(Exception):
    def __init__(self, message, user_message=None):
        super().__init__(message)
        self.user_message = user_message

class ConfigurationError(Exception):
    def __init__(self, message, user_message=None):
        super().__init__(message)
        self.user_message = user_message

# Simple retry decorator
def retry_on_exception(max_attempts=3, exceptions=(Exception,), delay=1):
    def decorator(func):
        def wrapper(*args, **kwargs):
            for attempt in range(max_attempts):
                try:
                    return func(*args, **kwargs)
                except exceptions as e:
                    if attempt == max_attempts - 1:
                        raise e
                    time.sleep(delay)
            return None
        return wrapper
    return decorator


class AIModel(ABC):
    """Abstract base class for AI models"""

    @abstractmethod
    def parse_command(self, user_input: str, context: Dict[str, Any]) -> Dict[str, Any]:
        """Parse user command and return structured response"""
        pass

    @abstractmethod
    def is_available(self) -> bool:
        """Check if the AI model is available"""
        pass

    @abstractmethod
    def get_model_name(self) -> str:
        """Get the model name"""
        pass


class GeminiModel(AIModel):
    """Gemini AI integration via Google AI Studio"""

    # Capabilities list for help responses
    CAPABILITIES = [
        "💬 Have natural conversations and answer questions",
        "🚀 Open applications like Chrome, Notepad, Calculator, Word, Excel, and more",
        "🔍 Search the web for any information you need",
        "🔊 Control system audio - increase/decrease volume or mute/unmute",
        "💻 Manage system operations - shutdown or restart your computer",
        "📁 Handle file operations - open folders, create directories, access Documents/Downloads",
        "🌤️ Get weather information and forecasts",
        "📋 Remember our conversation history for better context",
        "🎯 Execute commands with voice - just tell me what you want to do"
    ]

    def __init__(self, api_key: str, kb):
        self.logger = logging.getLogger('ai')
        self.api_key = api_key
        self.kb = kb
        self.enhanced_parser = EnhancedFallbackParser(self.kb)
        self.knowledge_base_commands = self.kb.search_commands("")
        self.client = None

        # Rate limit handling
        self.rate_limit_until = 0  # Timestamp when rate limit expires
        self.rate_limit_backoff = 1  # Current backoff multiplier
        self.max_backoff = 60  # Maximum backoff in seconds
        self.consecutive_rate_limits = 0  # Track consecutive rate limit errors

        if api_key:
            try:
                self.logger.debug("Initializing Gemini client...")
                genai.configure(api_key=api_key)
                self.client = genai.GenerativeModel('gemini-2.5-flash')
                self.logger.info("Gemini client initialized successfully")
            except Exception as e:
                error_msg = f"Gemini initialization failed: {str(e)}"
                self.logger.error(error_msg, exc_info=True)
                self.client = None  # Don't raise error, just disable AI

    def is_available(self) -> bool:
        return self.client is not None

    def get_model_name(self) -> str:
        return "gemini-2.5-flash"

    def _is_conversational_input(self, user_input: str) -> bool:
        """
        Determine if input is conversational rather than a command.
        Returns True for: greetings, thanks, questions, casual conversation
        Returns False for: commands (open app, search X, shutdown, etc.)
        
        This ensures fallback parsing is ONLY used for commands,
        and conversational inputs go to the API for better handling.
        """
        user_input_lower = user_input.lower().strip()
        word_count = len(user_input.split())

        # ===== COMMAND CHECKS FIRST (return False = use fallback parser) =====
        
        # Typing/dictation commands - check these FIRST before any conversational checks
        typing_patterns = ['type ', 'write ', 'dictate ', 'enter ', 'input ']
        if any(user_input_lower.startswith(pattern) for pattern in typing_patterns):
            return False
        
        # Other command patterns
        command_patterns = ['open ', 'launch ', 'start ', 'run ', 'search ', 'find ', 'look up ', 
                           'increase ', 'decrease ', 'turn up ', 'turn down ', 'volume ', 'mute ',
                           'set ', 'brightness ', 'screen brightness ', 'display brightness ',
                           'shutdown ', 'restart ', 'create ', 'make ',
                           'close ', 'minimize ', 'minimise ', 'maximize ', 'maximise ', 'restore ']
        if any(pattern in user_input_lower for pattern in command_patterns):
            return False
        
        # Keyboard shortcut patterns - treat as commands, not conversation
        keyboard_patterns = ['press ', 'ctrl ', 'control ', 'alt ', 'shift ', 'enter', 'tab', 'escape', 'esc',
                            'ctrl+', 'control+', 'alt+', 'shift+', 'windows ', 'win ']
        if any(pattern in user_input_lower for pattern in keyboard_patterns):
            return False

        # ===== CONVERSATIONAL INPUTS (use API) =====
        
        # 1. Greetings - always conversational
        greeting_keywords = ['hello', 'hi', 'hey', 'good morning', 'good afternoon', 'good evening', 'greetings', 'sup', "what's up"]
        if any(greeting in user_input_lower.split() or greeting in user_input_lower for greeting in greeting_keywords):
            return True
        
        # 2. Thanks/gratitude - conversational
        thanks_keywords = ['thank', 'thanks', 'thank you', 'appreciate', 'grateful']
        if any(thanks in user_input_lower.split() for thanks in thanks_keywords):
            return True
        
        # 3. Questions (how, what, why, when, where, who, can, could, would, should)
        # These are almost always conversational, not commands
        # Command patterns are checked at the beginning of this method
        
        question_starters = ['how', 'what', 'why', 'when', 'where', 'who', 'can', 'could', 'would', 'should', 'is', 'are', 'do', 'does', "don't", "doesn't"]
        if any(user_input_lower.startswith(q) or user_input_lower.split()[0] in question_starters for q in question_starters):
            return True
        
        # 4. Single word conversational inputs
        single_word_conversational = ['hello', 'hi', 'hey', 'thanks', 'thank', 'okay', 'ok', 'sure', 'yes', 'no', 'maybe', 'nice', 'cool', 'awesome', 'great']
        if word_count == 1 and user_input_lower in single_word_conversational:
            return True
        
        # 5. Help/capabilities queries
        help_keywords = ['what can you do', 'help', 'capabilities', 'what are you', 'what do you', 'features', 'abilities', 'what are your', 'how can you help', 'what can you help with']
        if any(keyword in user_input_lower for keyword in help_keywords):
            return True
        
        # 5.5. Entertainment requests (jokes, stories, fun facts)
        entertainment_keywords = ['joke', 'funny', 'humor', 'laugh', 'story', 'riddle', 'trivia', 'fun fact']
        if any(keyword in user_input_lower for keyword in entertainment_keywords):
            return True

        # 6. Observations/comments - casual conversational statements
        # These indicate the user is making a comment, not giving a command
        # Examples: "oh that's quite a lot you can do", "wow that's cool", "interesting"
        observation_indicators = [
            'oh ', 'oh,', "that's", 'thats', 'quite ', 'a lot', 'wow ', 
            'interesting', 'cool ', 'nice ', 'awesome ', 'amazing ',
            'good to know', 'i see', 'i understand', 'i get it',
            'makes sense', 'i like', 'love ', 'great job',
            'well done', 'good work', 'impressive', 'wow that'
        ]
        if any(indicator in user_input_lower for indicator in observation_indicators):
            return True

        # 7. Responses to questions (short affirmative/negative responses)
        short_responses = ['yes', 'no', 'yeah', 'yep', 'nope', 'maybe', 'sure', 'ok', 'okay', 'alright']
        if word_count <= 2 and user_input_lower in short_responses:
            return True

        # 8. Expressions of confusion or clarification requests
        confusion_indicators = ['what do you mean', 'sorry', 'pardon', 'repeat', 'say again', 'come again']
        if any(indicator in user_input_lower for indicator in confusion_indicators):
            return True

        # 9. Farewells - not commands
        farewell_keywords = ['goodbye', 'bye', 'see you', 'later', 'take care', 'night']
        if any(keyword in user_input_lower for keyword in farewell_keywords):
            return True

        # 10. Self-talk or thinking out loud (not directed at assistant as command)
        # Patterns like "let me think", "hmm", "i wonder"
        thinking_patterns = ['let me think', 'hmm', 'i wonder', 'i guess', 'i suppose']
        if any(pattern in user_input_lower for pattern in thinking_patterns):
            return True

        # ===== COMMANDS (use fallback parser) =====
        # If we get here, it's likely a command - use fallback parser for fast processing
        # Examples: "open chrome", "search for weather", "shutdown computer", "increase volume"
        return False

    def _check_rate_limit(self) -> bool:
        """Check if we're currently rate limited"""
        current_time = time.time()
        if current_time < self.rate_limit_until:
            remaining = int(self.rate_limit_until - current_time)
            print(f"⏳ Rate limited, waiting {remaining} seconds before next API call")
            return True
        return False

    def _handle_rate_limit(self, error_message: str) -> None:
        """Handle rate limit detection and backoff"""
        self.consecutive_rate_limits += 1
        self.logger.warning(f"Rate limit detected (attempt #{self.consecutive_rate_limits}): {error_message}")

        # Calculate backoff time with exponential backoff
        backoff_time = min(self.rate_limit_backoff * (2 ** (self.consecutive_rate_limits - 1)), self.max_backoff)
        self.rate_limit_until = time.time() + backoff_time
        self.rate_limit_backoff = min(self.rate_limit_backoff * 2, self.max_backoff)

        print(f"🚦 Rate limit triggered. Backing off for {backoff_time} seconds. Next attempt at {time.ctime(self.rate_limit_until)}")

    def _reset_rate_limit_state(self) -> None:
        """Reset rate limit state after successful API call"""
        if self.consecutive_rate_limits > 0:
            self.consecutive_rate_limits = 0
            self.rate_limit_backoff = 1  # Reset backoff multiplier
            self.logger.info("Rate limit state reset after successful API call")

    def _handle_conversation(self, user_input: str, context: Dict[str, Any]) -> Dict[str, Any]:
        """Handle conversational input using Gemini API with optimized settings and rate limit handling"""
        print(f"💬 Using Gemini API for conversation: {user_input}")

        # Check if we're currently rate limited
        if self._check_rate_limit():
            print(f"🔄 Rate limited, using fallback parser for conversation: {user_input}")
            return self._fallback_parse(user_input)

        # Check if this is a help/capabilities query
        user_input_lower = user_input.lower().strip()
        help_keywords = ['what can you do', 'help', 'capabilities', 'what are you', 'what do you', 'features', 'abilities', 'what are your', 'how can you help', 'what can you help with']
        is_help_query = any(keyword in user_input_lower for keyword in help_keywords)

        # Build conversation prompt
        if is_help_query:
            system_prompt = f"""You are a friendly, helpful voice assistant. The user is asking about your capabilities.

Here are your capabilities:
{chr(10).join(f"- {cap}" for cap in self.CAPABILITIES)}

Respond naturally and list these capabilities in a friendly, engaging way. Be conversational and warm.
Keep responses informative but concise (2-3 sentences).

Respond with JSON in this format:
{{
    "intent": "conversation",
    "action": "chat",
    "parameters": {{}},
    "command_to_execute": "",
    "response": "Your conversational response here",
    "confidence": 0.9
}}"""
        else:
            system_prompt = """You are a friendly, helpful voice assistant having a natural conversation with a user.

IMPORTANT: You have memory of the past 2 conversations. Use this context to:
- Reference previous topics or questions the user asked
- Maintain continuity in the conversation
- Remember user preferences mentioned earlier
- Build upon what was discussed before

Respond naturally to their message. Be conversational, warm, and engaging.
Keep responses concise (1-2 sentences) but informative.
You can discuss general topics, answer questions, and chat casually.

Respond with JSON in this format:
{
    "intent": "conversation",
    "action": "chat",
    "parameters": {},
    "command_to_execute": "",
    "response": "Your conversational response here",
    "confidence": 0.9
}"""

        # Build conversation history
        conversation_parts = [system_prompt]

        # Limit conversation history to last 4 messages for speed (past 2 prompts context)
        conversation_history = context.get('conversation_history', [])
        for msg in conversation_history[-4:]:
            if msg['role'] == 'user':
                conversation_parts.append(f"User: {msg['content']}")
            elif msg['role'] == 'assistant':
                conversation_parts.append(f"Assistant: {msg['content']}")

        # Add current input
        conversation_parts.append(f"User: {user_input}")

        full_prompt = "\n\n".join(conversation_parts)

        try:
            # Use Google's Gemini API
            response = self.client.generate_content(
                full_prompt,
                generation_config=genai.types.GenerationConfig(
                    temperature=0.1,
                    max_output_tokens=900,
                )
            )

            response_text = response.text.strip()
            
            # Handle truncated/incomplete JSON responses
            if response_text.count("{") > response_text.count("}"):
                # Try to extract response field from incomplete JSON
                response_match = re.search(r'"response"\s*:\s*"([^"]*)"', response_text)
                if response_match:
                    extracted_response = response_match.group(1)
                    return {
                        'success': True,
                        'intent': 'conversation',
                        'action': 'chat',
                        'parameters': {},
                        'command': '',
                        'response': extracted_response,
                        'confidence': 0.7,
                        'model': self.get_model_name()
                    }
                # If can't extract, use the text before the truncation
                truncated_response = response_text[:200] + "..." if len(response_text) > 200 else response_text
                return {
                    'success': True,
                    'intent': 'conversation',
                    'action': 'chat',
                    'parameters': {},
                    'command': '',
                    'response': "I understand. How can I help you today?",
                    'confidence': 0.7,
                    'model': self.get_model_name()
                }
            print(f"🤖 Gemini API response received")

            # Reset rate limit state on successful call
            self._reset_rate_limit_state()

            # Extract JSON more efficiently with better error handling
            try:
                if '```json' in response_text:
                    response_text = response_text.split('```json')[1].split('```')[0].strip()
                elif '```' in response_text:
                    response_text = response_text.split('```')[1].split('```')[0].strip()

                # If response_text is empty after extraction, use fallback
                if not response_text:
                    raise ValueError("Empty response after JSON extraction")

                # Use safe JSON extraction instead of parsing directly
                result = extract_json_safely(response_text)
            except (json.JSONDecodeError, ValueError, IndexError):
                reply_text = response_text.strip()
                confidence = 0.8
                # Try to extract response from plain text
                if '"response"' in response_text:
                    # Look for response field in malformed JSON
                    response_match = re.search(r'"response"\s*:\s*"([^"]*)"', response_text)
                    if response_match:
                        extracted_response = response_match.group(1)
                        return {
                            'success': True,
                            'intent': 'conversation',
                            'action': 'chat',
                            'parameters': {},
                            'command': '',
                            'response': extracted_response,
                            'confidence': 0.7,
                            'model': self.get_model_name()
                        }

                # Final fallback - provide user-friendly message for parsing failures
                return {
                    'success': True,
                    'intent': 'conversation',
                    'action': 'chat',
                    'parameters': {},
                    'command': '',
                    'response': "I am yet not able to understand that, could you try again?",
                    'confidence': 0.5,
                    'model': f"{self.get_model_name()} (Fallback)"
                }

            # Try to extract JSON safely
            result = extract_json_safely(response_text)

            # ✅ If JSON exists → use it
            if result and "response" in result:
                reply_text = result["response"]
                confidence = result.get("confidence", 0.9)

            # ✅ If NO JSON → use plain text response
            else:
                reply_text = response_text.strip()
                confidence = 0.8

            return {
                'success': True,
                'intent': 'conversation',
                'action': 'chat',
                'parameters': {},
                'command': '',
                'response': reply_text,
                'confidence': confidence,
                'model': self.get_model_name()
            }

        except Exception as e:
            error_str = str(e)
            print(f"⚠️ Conversation API error: {error_str}")

            # Check for rate limit errors (429)
            if "429" in error_str or "Too Many Requests" in error_str or "rate limit" in error_str.lower() or "quota" in error_str.lower():
                self._handle_rate_limit(error_str)
                # Fall back to enhanced parser immediately for rate limit scenarios
                print(f"🔄 Rate limit detected, using fallback parser for conversation: {user_input}")
                return self._fallback_parse(user_input)

            # For other errors, fall back to simple response
            return {
                'success': True,
                'intent': 'conversation',
                'action': 'chat',
                'parameters': {},
                'command': '',
                'response': self._generate_conversation_response(user_input),
                'confidence': 0.6,
                'model': f"{self.get_model_name()} (Fallback)"
            }

    def parse_command(self, user_input: str, context: Dict[str, Any]) -> Dict[str, Any]:
        """Parse command using fast fallback parsing with optional API for conversation"""
        if not self.is_available():
            # Even without API, we can still use fallback parsing for commands
            print(f"🔄 Using fallback parsing (API unavailable): {user_input}")
            return self._fallback_parse(user_input)

        # Check if this looks like a conversational query (not a command)
        is_likely_conversation = self._is_conversational_input(user_input)

        # For conversational inputs, use API for natural responses
        if is_likely_conversation:
            return self._handle_conversation(user_input, context)

        # For command inputs, use fast fallback parsing without API calls
        print(f"⚡ Using fast fallback parsing for command: {user_input}")
        result = self._fallback_parse(user_input)
        print(f"DEBUG: Fallback result intent={result.get('intent')}, action={result.get('action')}")
        return result

    def _fallback_parse(self, user_input: str) -> Dict[str, Any]:
        """Enhanced fallback parser using the new EnhancedFallbackParser class"""
        model_name = f"{self.get_model_name()} (Enhanced Fallback)"
        return self.enhanced_parser.parse(user_input, model_name)
    
    def _generate_conversation_response(self, user_input: str) -> str:
        """Generate a conversational response for fallback parser"""
        user_input_lower = user_input.lower()

        # Debug print
        print(f"DEBUG: Processing '{user_input}' -> '{user_input_lower}'")

        # Check for help/capabilities queries first
        help_keywords = ['what can you do', 'help', 'capabilities', 'what are you', 'what do you', 'features', 'abilities', 'what are your', 'how can you help', 'what can you help with']
        if any(keyword in user_input_lower for keyword in help_keywords):
            print("DEBUG: Matched help query")
            capabilities_text = "\n".join(f"• {cap}" for cap in self.CAPABILITIES)
            return f"I'm your voice assistant! Here's what I can help you with:\n\n{capabilities_text}\n\nJust tell me what you'd like to do!"

        if any(greeting in user_input_lower.split() for greeting in ['hello', 'hi', 'hey', 'good morning']):
            print("DEBUG: Matched greeting")
            return "Hello! How can I help you today? (updated)"

        if any(thanks_word in user_input_lower.split() for thanks_word in ['thank', 'thanks']):
            print("DEBUG: Matched thanks")
            return "You're welcome! Is there anything else I can help you with?"

        if (user_input_lower.startswith('how') or user_input_lower.startswith('what') or
            user_input_lower.startswith('why') or user_input_lower.startswith('when') or
            user_input_lower.startswith('where') or user_input_lower.startswith('who')):
            print("DEBUG: Matched question")
            return "That's an interesting question! While I'm primarily here to help with computer tasks, I can try to assist with that."

        print("DEBUG: Default response")
        return "I'm here to help! You can ask me to open applications, search the web, or control system settings. What would you like me to do?"


class AIManager:
    """Manager for multiple AI models"""

    def __init__(self, knowledge_base):
        self.kb = knowledge_base
        self.models = {}
        self.current_model = None
        self._initialize_models()

    def _initialize_models(self):
        """Initialize available AI models"""
        try:
            # Initialize Gemini via OpenRouter
            gemini_key = os.getenv('OPENROUTER_API_KEY') or os.getenv('GOOGLE_API_KEY')
            if gemini_key:
                try:
                    self.models['gemini'] = GeminiModel(gemini_key, self.kb)
                    if not self.current_model:
                        self.current_model = 'gemini'
                    print("✅ AI models initialized successfully")
                except Exception as model_error:
                    print(f"⚠️ Failed to initialize AI model: {str(model_error)}")
                    print("🔄 Continuing with fallback parser only")

            # Log initialization status
            if self.models:
                print(f"✅ AI models available: {list(self.models.keys())}")
            else:
                print("🔄 No AI models available, using fallback parser")
        except Exception as e:
            # Continue with empty models - fallback parser will handle commands
            print(f"⚠️ Error initializing AI models: {str(e)}")
            print("🔄 Continuing with fallback parser only")

    def set_model(self, model_name: str) -> bool:
        """Set the active AI model"""
        if model_name in self.models:
            self.current_model = model_name
            return True
        return False

    def get_available_models(self) -> List[str]:
        """Get list of available models"""
        return list(self.models.keys())

    def get_current_model(self) -> Optional[str]:
        """Get current active model"""
        return self.current_model

    def parse_command(self, user_input: str, context: Dict[str, Any] = None) -> Dict[str, Any]:
        """Parse command using available AI models with conversation context"""
        if context is None:
            context = {}

        # Try all available models in order until one succeeds
        for model_name, model in self.models.items():
            try:
                result = model.parse_command(user_input, context)
                if result.get('success'):
                    return result
            except Exception as e:
                print(f"⚠️ Error with {model_name}: {str(e)}")
                continue

        # Return error if all AI models fail or are unavailable
        return {
            'success': False,
            'intent': 'unknown',
            'action': '',
            'parameters': {},
            'command': '',
            'response': "No AI models are available. Please check your API key configuration.",
            'confidence': 0.0,
            'model': 'None'
        }

    def is_available(self) -> bool:
        """Check if any AI model is available"""
        return len(self.models) > 0

    def get_model_info(self) -> Dict[str, Any]:
        """Get information about available models"""
        info = {}
        for name, model in self.models.items():
            info[name] = {
                'name': model.get_model_name(),
                'available': model.is_available(),
                'active': name == self.current_model
            }
        return info

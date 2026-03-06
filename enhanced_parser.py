"""
Enhanced Fallback Parser Module
Drop-in replacement for the _fallback_parse method in ai_manager.py

Features:
- Fuzzy matching for typos (chorme -> chrome)
- Multiple action keywords (open, launch, start, run, fire up, etc.)
- Flexible word order (chrome open, open chrome)
- Handles filler words (can you please, i want to, etc.)
- Smart command extraction
"""

import re
import os
import logging
from typing import Dict, Any, Optional, List, Tuple
from difflib import SequenceMatcher
import pyautogui

# Local exception definitions
class ParsingError(Exception):
    def __init__(self, message, user_message=None):
        super().__init__(message)
        self.user_message = user_message

class ConfigurationError(Exception):
    def __init__(self, message, user_message=None):
        super().__init__(message)
        self.user_message = user_message


class EnhancedFallbackParser:
    """Advanced fallback parser with fuzzy matching and flexible parsing"""
    
    def __init__(self, knowledge_base):
        self.kb = knowledge_base
        self.logger = logging.getLogger('enhanced_parser')
        
        # Extended action keywords for better matching
        self.action_keywords = {
            'open': ['open', 'launch', 'start', 'run', 'execute', 'fire up', 'bring up', 'load', 'display'],
            'search': ['search','google'],  # Only these two for web search
            'create': ['create', 'make', 'new', 'add', 'build'],
            'close': ['close', 'exit', 'quit', 'shutdown', 'shut down', 'kill', 'terminate'],
            'show': ['show me', 'show'],  # Separate category for 'show' commands
        }
        
        # Common filler words to remove
        self.filler_words = [
            'please', 'can you', 'could you', 'would you', 'i want to', 'i want',
            'i need to', 'i need', 'help me', 'the', 'a', 'an', 'my', 'for me',
            'application', 'app', 'program', 'software', 'go ahead and',
            'how do i', 'how to', 'can i', 'do you', 'will you', 'could i'
        ]
        
        # Words that should NOT be removed when they appear with file/folder operations
        self.contextual_keywords = ['file', 'folder', 'directory']
        
        # Prepositions that should only be removed after action verbs
        # These are handled separately in extract_action_and_target
        self.contextual_prepositions = ['for', 'to', 'about', 'on', 'at', 'with', 'up', 'down']
        
        # Pronouns/objects that are typically filler in commands
        # These should be removed even when standalone
        self.command_filler_pronouns = ['me', 'us', 'them', 'it']
    
    def similarity_score(self, str1: str, str2: str) -> float:
        """Calculate similarity between two strings (0.0 to 1.0)"""
        try:
            return SequenceMatcher(None, str1.lower(), str2.lower()).ratio()
        except Exception as e:
            self.logger.error(f"Error calculating similarity score between '{str1}' and '{str2}': {str(e)}", exc_info=True)
            return 0.0
    
    def fuzzy_match_command(self, query: str, threshold: float = 0.6) -> Optional[Dict]:
        """
        Find closest matching command using fuzzy matching
        Handles typos like: chorme -> chrome, notpad -> notepad
        """
        all_commands = self.kb.search_commands("")
        best_match = None
        best_score = 0.0
        
        # Clean the query
        query_clean = self.clean_text(query)
        
        for cmd in all_commands:
            cmd_name = cmd['name'].lower()
            
            # Try exact match first (highest priority)
            if cmd_name == query_clean:
                return cmd
            
            # Check if command name is contained in query
            if cmd_name in query_clean or query_clean in cmd_name:
                return cmd
            
            # Calculate similarity score
            score = self.similarity_score(query_clean, cmd_name)
            
            # Boost score if words from command name appear in query
            if any(word in query_clean for word in cmd_name.split()):
                score += 0.2
            
            # Boost score for partial matches
            if cmd_name.startswith(query_clean[:3]) or query_clean.startswith(cmd_name[:3]):
                score += 0.15
            
            if score > best_score and score >= threshold:
                best_score = score
                best_match = cmd
        
        return best_match
    
    def clean_text(self, text: str) -> str:
        """Remove filler words and clean text"""
        try:
            text_lower = text.lower()

            # Check if this is a file/folder operation - preserve contextual keywords
            is_file_operation = any(kw in text_lower for kw in self.contextual_keywords)
            
            # Remove multi-word filler phrases FIRST (longer first to avoid partial matches)
            # Only remove complete phrases, not individual words that might be meaningful
            multi_word_fillers = [f for f in self.filler_words if ' ' in f]
            for filler in sorted(multi_word_fillers, key=len, reverse=True):
                text_lower = text_lower.replace(filler, ' ')

            # DON'T remove punctuation yet - keep apostrophes for contractions like "what's", "don't"
            # Just normalize extra spaces
            text_lower = ' '.join(text_lower.split())
            
            # NOW remove single-word fillers, but ONLY as complete words (not letters within words)
            # Use word boundaries to ensure we match whole words only
            # BUT preserve contextual keywords for file operations
            single_word_fillers = [f for f in self.filler_words if ' ' not in f and len(f) > 1]
            for filler in single_word_fillers:
                # Skip removal if this is a file operation and the filler is a contextual keyword
                if is_file_operation and filler in self.contextual_keywords:
                    continue
                # Use word boundary regex to match only complete words
                text_lower = re.sub(r'\b' + re.escape(filler) + r'\b', ' ', text_lower)
            
            # NOW remove punctuation (after removing filler words)
            # This preserves apostrophes in contractions until after word removal
            text_lower = re.sub(r"[^\w\s']", ' ', text_lower)
            
            # Clean up extra spaces again
            text_lower = ' '.join(text_lower.split())
            return text_lower
        except Exception as e:
            self.logger.error(f"Error cleaning text '{text}': {str(e)}", exc_info=True)
            return text.lower().strip()
    
    def extract_action_and_target(self, user_input: str) -> Tuple[Optional[str], str]:
        """
        Extract action (open, search, etc.) and target from user input
        Handles flexible word order: "chrome open" or "open chrome"
        """
        user_input_lower = user_input.lower()
        
        # Special cases: don't let single-letter filler words break app names
        if 'command' in user_input_lower and 'prompt' in user_input_lower:
            return 'open', 'command prompt'
        if 'terminal' in user_input_lower:
            return 'open', 'terminal'
        
        # Try to find action keywords
        found_action = None
        found_action_word = None
        action_position = -1
        
        for action, keywords in self.action_keywords.items():
            for keyword in keywords:
                pos = user_input_lower.find(keyword)
                if pos != -1:
                    found_action = action
                    found_action_word = keyword
                    action_position = pos
                    break
            if found_action:
                break
        
        if not found_action:
            return None, user_input
        
        # Remove the action word/phrase and clean the text
        target = user_input_lower.replace(found_action_word, ' ')
        target = self.clean_text(target)
        
        # Additional cleanup: remove pronouns that appear IMMEDIATELY after action verbs
        # These are typically filler: "find me", "show me", "tell me", "get me"
        # But keep pronouns that appear later in the sentence (they're meaningful there)
        target_words = target.split()
        
        # Remove leading pronouns only (first 1-2 words after action)
        pronouns_to_remove = ['me', 'us', 'them', 'it']
        while target_words and target_words[0] in pronouns_to_remove:
            target_words.pop(0)
            # Only remove up to 2 leading pronouns
            if len([w for w in target.split() if w in pronouns_to_remove]) > 1 and target_words and target_words[0] in pronouns_to_remove:
                target_words.pop(0)
            else:
                break
        
        target = ' '.join(target_words)
        
        # Additional cleanup: remove common prepositions that may remain IMMEDIATELY after action verbs
        # e.g., "search for python" -> removes "for", leaving "python"
        # BUT "search for how long it takes" should keep "for" because it's part of the query
        target_words = target.split()
        
        # Special case: preserve question phrases like "for how long", "about what", etc.
        # Don't remove prepositions if they're part of question structures
        question_phrases = [
            # Two-word question phrases starting with prepositions
            ['for', 'how'], ['for', 'what'], ['for', 'who'], ['for', 'when'], ['for', 'where'], ['for', 'why'],
            ['about', 'what'], ['about', 'how'], ['on', 'what'], ['on', 'how'],
            ['in', 'what'], ['in', 'which'], ['in', 'how'],
            # Question phrases starting directly with question words (no preposition)
            ['how', 'long'], ['how', 'much'], ['how', 'many'], ['how', 'often'], ['how', 'far'],
            ['how', 'to'], ['how', 'do'], ['how', 'does'], ['how', 'did'], ['how', 'can'], ['how', 'will'],
            ['how', 'are'], ['how', 'is'], ['how', 'was'], ['how', 'were'], ['how', 'about'],
            ['what', 'is'], ['what', 'are'], ['what', 'was'], ['what', 'were'], ['what', 'do'], ['what', 'does'],
            ['when', 'is'], ['when', 'are'], ['when', 'was'], ['when', 'were'],
            ['where', 'is'], ['where', 'are'], ['where', 'can'], ['where', 'do'],
            ['why', 'is'], ['why', 'are'], ['why', 'do'], ['why', 'does'],
            ['who', 'is'], ['who', 'are'], ['who', 'was'], ['who', 'were'],
            ['which', 'is'], ['which', 'are'], ['which', 'one']
        ]
        
        # Check if starts with a question phrase
        is_question_phrase = False
        for qphrase in question_phrases:
            if len(target_words) >= len(qphrase):
                if target_words[:len(qphrase)] == qphrase:
                    is_question_phrase = True
                    break
        
        # If it's a question phrase, skip preposition removal entirely
        if is_question_phrase:
            # But first check if there's a LEADING preposition before the question phrase
            # e.g., "for how are you" has "for" before the question phrase ["how", "are"]
            # We should remove that leading preposition
            if target_words and len(target_words) >= 2:
                # Check if first word is preposition and second+ start a question phrase
                if target_words[0] in self.contextual_prepositions:
                    # Check if remaining words form a question phrase
                    for qphrase in question_phrases:
                        if len(target_words[1:]) >= len(qphrase):
                            if target_words[1:len(qphrase)+1] == qphrase:
                                # Remove the leading preposition
                                target_words.pop(0)
                                break
            target = ' '.join(target_words)
        else:
            # Not a question phrase - check for duplicate prepositions and remove leading ones
            # Special handling: if we have "search for for how long", the second "for" is part of "for how long"
            # So we should NOT remove duplicates if what follows is a question phrase
            
            # First, check if removing the duplicate would create a question phrase
            has_duplicate_preposition = False
            if found_action_word and target_words:
                action_words = found_action_word.split()
                if action_words and action_words[-1] in self.contextual_prepositions:
                    last_action_word = action_words[-1]
                    # Check if first target word is the same preposition
                    if target_words[0] == last_action_word:
                        has_duplicate_preposition = True
                        # Check if after removing this duplicate, we'd get a question phrase
                        remaining_words = target_words[1:]  # Words after removing duplicate
                        for qphrase in question_phrases:
                            if len(remaining_words) >= len(qphrase):
                                if remaining_words[:len(qphrase)] == qphrase:
                                    # Don't remove the duplicate - it's part of a question phrase!
                                    has_duplicate_preposition = False
                                    break
            
            # Remove duplicate preposition only if it won't break a question phrase
            if has_duplicate_preposition:
                target_words.pop(0)
            
            # Now remove leading prepositions (at most 2 words) only if not a question phrase
            while target_words and target_words[0] in self.contextual_prepositions:
                target_words.pop(0)
                # Check if next word is also a preposition
                if target_words and target_words[0] in self.contextual_prepositions:
                    target_words.pop(0)
                else:
                    break
            
            target = ' '.join(target_words)
        
        return found_action, target
    
    def parse(self, user_input: str, model_name: str = "Enhanced Fallback") -> Dict[str, Any]:
        """
        Main parsing method with enhanced capabilities:
        - Handles multiple action verb variations
        - Flexible word order
        - Fuzzy matching for typos
        - Smart command extraction
        """
        try:
            user_input_lower = user_input.lower()
            print(f"🔄 Enhanced fallback parsing: {user_input}")
        except Exception as e:
            self.logger.error(f"Error initializing parse for input '{user_input}': {str(e)}", exc_info=True)
            return self._create_error_response(f"Failed to process input: {str(e)}")
        
        # ===== 0. JOKES/ENTERTAINMENT =====
        joke_keywords = ['joke', 'tell me a joke', 'funny', 'make me laugh', 'humor me']
        if any(keyword in user_input_lower for keyword in joke_keywords):
            jokes = [
                "Why don't scientists trust atoms? Because they make up everything!",
                "Why did the scarecrow win an award? He was outstanding in his field!",
                "Why don't eggs tell jokes? They'd crack each other up!",
                "What do you call a fake noodle? An impasta!",
                "Why did the math book look sad? Because it had too many problems!",
                "What do you call a bear with no teeth? A gummy bear!",
                "Why couldn't the bicycle stand up by itself? It was two tired!",
                "What do you call cheese that isn't yours? Nacho cheese!",
                "Why did the golfer bring two pairs of pants? In case he got a hole in one!",
                "What do you call a sleeping dinosaur? A dino-snore!"
            ]
            import random
            joke = random.choice(jokes)
            return self._create_response(
                intent='conversation',
                action='chat',
                parameters={'type': 'joke'},
                command='',
                response=joke,
                confidence=0.9,
                model=model_name
            )

        # ===== 0.5. TEXT TYPING/DICTATION =====
        # Handle voice typing commands: "type hello world", "write hello", "dictate hello"
        # Also handles inline case: "type this is python in uppercase"
        # Only direct keywords - no phrases like "type this", "write this", etc.
        type_keywords = ['type ', 'write ', 'dictate ', 'enter ', 'input ']
        
        is_type_command = any(user_input_lower.startswith(keyword) for keyword in type_keywords)
        
        if is_type_command:
            # Check for inline case command at the end (e.g., "... in uppercase")
            case_type = None
            case_patterns = [
                (r'\s+in\s+(uppercase|upper case|all caps|caps)$', 'uppercase'),
                (r'\s+in\s+(lowercase|lower case|small case|small)$', 'lowercase'),
                (r'\s+in\s+(title case|titlecase)$', 'title'),
                (r'\s+in\s+(sentence case|sentencecase)$', 'sentence'),
                (r'\s+in\s+(camel case|camelcase)$', 'camel'),
                (r'\s+in\s+(pascal case|pascalcase)$', 'pascal'),
                (r'\s+in\s+(snake case|snakecase|underscore)$', 'snake'),
                (r'\s+in\s+(kebab case|kebabcase|hyphen case)$', 'kebab'),
            ]
            
            text_to_process_lower = user_input_lower
            text_to_process_original = user_input
            for pattern, case_name in case_patterns:
                match = re.search(pattern, user_input_lower)
                if match:
                    case_type = case_name
                    # Remove the case command from both lowercase and original text
                    text_to_process_lower = user_input_lower[:match.start()]
                    text_to_process_original = user_input[:match.start()]
                    break
            
            # Extract the text to type
            text_to_type = self._extract_text_to_type_with_case(text_to_process_lower, text_to_process_original)
            
            # Handle "type space" command - insert a single space
            if not text_to_type and re.match(r'^(?:type|write|dictate|enter|input)\s+space\b', text_to_process_lower):
                text_to_type = ' '
            
            if text_to_type:
                params = {'text': text_to_type}
                if case_type:
                    params['case_type'] = case_type
                    response_text = f"Typing in {case_type}: {text_to_type[:50]}{'...' if len(text_to_type) > 50 else ''}"
                else:
                    response_text = f"Typing: {text_to_type[:50]}{'...' if len(text_to_type) > 50 else ''}"
                
                return self._create_response(
                    intent='type_text',
                    action='type',
                    parameters=params,
                    command=f'type:{text_to_type}',
                    response=response_text,
                    confidence=0.95,
                    model=model_name
                )

        # ===== 0.54. CLEAR TEXT =====
        # Handle "clear", "clear all", "clear text" commands
        clear_patterns = [
            (r'^(?:clear|clear all|delete all|clear text|clear everything|clear textbox|clear input)$', 'clear_all', 'Clearing all text'),
        ]
        
        for pattern, action, response_text in clear_patterns:
            if re.match(pattern, user_input_lower):
                return self._create_response(
                    intent='text_edit',
                    action=action,
                    parameters={},
                    command=f'text_edit:{action}',
                    response=response_text,
                    confidence=0.95,
                    model=model_name
                )

        # ===== 0.545. MULTIPLE DELETE/BACKSPACE =====
        # Handle "delete 2 words", "backspace 3 letters", etc.
        delete_patterns = [
            (r'^(?:delete|backspace|del)\s+(\d+|one|two|three|four|five|six|seven|eight|nine|ten)\s*(?:words?|word)$', 'delete_words', 'Deleting words'),
            (r'^(?:delete|backspace|del)\s+(\d+|one|two|three|four|five|six|seven|eight|nine|ten)\s*(?:letters?|letter|chars?|char|characters?|character)$', 'delete_chars', 'Deleting characters'),
            (r'^(?:delete|backspace|del)\s+(\d+|one|two|three|four|five|six|seven|eight|nine|ten)$', 'delete_chars', 'Deleting characters'),
        ]
        
        for pattern, action, response_text in delete_patterns:
            match = re.match(pattern, user_input_lower)
            if match:
                count_str = match.group(1)
                # Convert word numbers to digits
                word_to_num = {'one': 1, 'two': 2, 'three': 3, 'four': 4, 'five': 5, 
                               'six': 6, 'seven': 7, 'eight': 8, 'nine': 9, 'ten': 10}
                count = word_to_num.get(count_str, int(count_str) if count_str.isdigit() else 1)
                return self._create_response(
                    intent='text_edit',
                    action=action,
                    parameters={'count': count},
                    command=f'text_edit:{action}:{count}',
                    response=f'{response_text} ({count})',
                    confidence=0.95,
                    model=model_name
                )

        # ===== 0.55. KEYBOARD SHORTCUTS =====
        # Handle special keys and combinations: "press enter", "press tab", "alt tab", "ctrl c", etc.
        keyboard_shortcuts = [
            # Single keys
            (r'^(?:press\s+)?(enter|return)$', ['enter'], 'Pressing Enter'),
            (r'^(?:press\s+)?(tab)$', ['tab'], 'Pressing Tab'),
            (r'^(?:press\s+)?(shift)$', ['shift'], 'Pressing Shift'),
            (r'^(?:press\s+)?(ctrl|control)$', ['ctrl'], 'Pressing Ctrl'),
            (r'^(?:press\s+)?(alt)$', ['alt'], 'Pressing Alt'),
            (r'^(?:press\s+)?(escape|esc)$', ['esc'], 'Pressing Escape'),
            (r'^(?:press\s+)?(backspace|delete|del)$', ['backspace'], 'Pressing Backspace'),
            (r'^(?:press\s+)?(home)$', ['home'], 'Pressing Home'),
            (r'^(?:press\s+)?(end)$', ['end'], 'Pressing End'),
            (r'^(?:press\s+)?(page up|pageup)$', ['pageup'], 'Pressing Page Up'),
            (r'^(?:press\s+)?(page down|pagedown)$', ['pagedown'], 'Pressing Page Down'),
            (r'^(?:press\s+)?(up|up arrow)$', ['up'], 'Pressing Up Arrow'),
            (r'^(?:press\s+)?(down|down arrow)$', ['down'], 'Pressing Down Arrow'),
            (r'^(?:press\s+)?(left|left arrow)$', ['left'], 'Pressing Left Arrow'),
            (r'^(?:press\s+)?(right|right arrow)$', ['right'], 'Pressing Right Arrow'),
            (r'^(?:press\s+)?(insert|ins)$', ['insert'], 'Pressing Insert'),
            (r'^(?:press\s+)?(print screen|printscreen|prtsc|prt sc|snapshot)$', ['printscreen'], 'Pressing Print Screen'),
            (r'^(?:press\s+)?(f1|f2|f3|f4|f5|f6|f7|f8|f9|f10|f11|f12)$', None, 'Pressing Function Key'),  # Special handling
            
            # Two-key combinations
            (r'^(?:press\s+)?(ctrl\s*\+\s*c|control\s+c|ctrl\s+c)$', ['ctrl', 'c'], 'Copying (Ctrl+C)'),
            (r'^(?:press\s+)?(ctrl\s*\+\s*v|control\s+v|ctrl\s+v)$', ['ctrl', 'v'], 'Pasting (Ctrl+V)'),
            (r'^(?:press\s+)?(ctrl\s*\+\s*x|control\s+x|ctrl\s+x)$', ['ctrl', 'x'], 'Cutting (Ctrl+X)'),
            (r'^(?:press\s+)?(ctrl\s*\+\s*z|control\s+z|ctrl\s+z)$', ['ctrl', 'z'], 'Undoing (Ctrl+Z)'),
            (r'^(?:press\s+)?(ctrl\s*\+\s*y|control\s+y|ctrl\s+y)$', ['ctrl', 'y'], 'Redoing (Ctrl+Y)'),
            (r'^(?:press\s+)?(ctrl\s*\+\s*a|control\s+a|ctrl\s+a)$', ['ctrl', 'a'], 'Selecting All (Ctrl+A)'),
            (r'^(?:press\s+)?(ctrl\s*\+\s*s|control\s+s|ctrl\s+s)$', ['ctrl', 's'], 'Saving (Ctrl+S)'),
            (r'^(?:press\s+)?(ctrl\s*\+\s*f|control\s+f|ctrl\s+f)$', ['ctrl', 'f'], 'Finding (Ctrl+F)'),
            (r'^(?:press\s+)?(ctrl\s*\+\s*p|control\s+p|ctrl\s+p)$', ['ctrl', 'p'], 'Printing (Ctrl+P)'),
            (r'^(?:press\s+)?(ctrl\s*\+\s*n|control\s+n|ctrl\s+n)$', ['ctrl', 'n'], 'New (Ctrl+N)'),
            (r'^(?:press\s+)?(ctrl\s*\+\s*o|control\s+o|ctrl\s+o)$', ['ctrl', 'o'], 'Opening (Ctrl+O)'),
            (r'^(?:press\s+)?(ctrl\s*\+\s*t|control\s+t|ctrl\s+t|control\s+tea|control\s+ti|ctrl\s+tea|ctrl\s+ti)$', ['ctrl', 't'], 'New Tab (Ctrl+T)'),
            (r'^(?:press\s+)?(ctrl\s*\+\s*w|control\s+w|ctrl\s+w)$', ['ctrl', 'w'], 'Closing Tab (Ctrl+W)'),
            (r'^(?:press\s+)?(alt\s*\+\s*tab|alt\s+tab)$', ['alt', 'tab'], 'Switching Window (Alt+Tab)'),
            (r'^(?:press\s+)?(alt\s*\+\s*f4|alt\s+f4)$', ['alt', 'f4'], 'Closing Window (Alt+F4)'),
            (r'^(?:press\s+)?(shift\s*\+\s*enter|shift\s+enter)$', ['shift', 'enter'], 'Shift+Enter'),
            (r'^(?:press\s+)?(ctrl\s*\+\s*enter|control\s+enter|ctrl\s+enter)$', ['ctrl', 'enter'], 'Ctrl+Enter'),
            (r'^(?:press\s+)?(shift\s*\+\s*tab|shift\s+tab)$', ['shift', 'tab'], 'Shift+Tab'),
            (r'^(?:press\s+)?(ctrl\s*\+\s*shift\s*\+\s*t|control\s+shift\s+t|ctrl\s+shift\s+t)$', ['ctrl', 'shift', 't'], 'Reopening Closed Tab'),
            (r'^(?:press\s+)?(windows|win|command|cmd)$', ['win'], 'Pressing Windows Key'),
            (r'^(?:press\s+)?(windows\s*\+\s*d|win\s*\+\s*d|windows\s+d|win\s+d)$', ['win', 'd'], 'Show Desktop'),
            (r'^(?:press\s+)?(windows\s*\+\s*e|win\s*\+\s*e|windows\s+e|win\s+e)$', ['win', 'e'], 'Opening File Explorer'),
            (r'^(?:press\s+)?(windows\s*\+\s*r|win\s*\+\s*r|windows\s+r|win\s+r)$', ['win', 'r'], 'Opening Run Dialog'),
            (r'^(?:press\s+)?(windows\s*\+\s*l|win\s*\+\s*l|windows\s+l|win\s+l)$', ['win', 'l'], 'Locking Computer'),
                        
            # Three-key combinations
            (r'^(?:press\s+)?(ctrl\s*\+\s*shift\s*\+\s*esc|control\s+shift\s+escape|ctrl\s+shift\s+escape|ctrl\s+shift\s+esc)$', ['ctrl', 'shift', 'esc'], 'Opening Task Manager'),
            (r'^(?:press\s+)?(ctrl\s*\+\s*alt\s*\+\s*del|control\s+alt\s+delete|ctrl\s+alt\s+delete)$', ['ctrl', 'alt', 'delete'], 'Ctrl+Alt+Delete'),
            (r'^(?:press\s+)?(ctrl\s*\+\s*shift\s*\+\s*n|control\s+shift\s+n|ctrl\s+shift\s+n)$', ['ctrl', 'shift', 'n'], 'New Folder/Incognito'),
            (r'^(?:press\s+)?(ctrl\s*\+\s*alt\s*\+\s*t|control\s+alt\s+t|ctrl\s+alt\s+t)$', ['ctrl', 'alt', 't'], 'Ctrl+Alt+T'),
            (r'^(?:press\s+)?(alt\s*\+\s*shift\s*\+\s*tab|alt\s+shift\s+tab)$', ['alt', 'shift', 'tab'], 'Switch Windows Backward'),
            (r'^(?:press\s+)?(ctrl\s*\+\s*shift\s*\+\s*p|control\s+shift\s+p|ctrl\s+shift\s+p)$', ['ctrl', 'shift', 'p'], 'Private/Incognito Window'),
            (r'^(?:press\s+)?(ctrl\s*\+\s*shift\s*\+\s*s|control\s+shift\s+s|ctrl\s+shift\s+s)$', ['ctrl', 'shift', 's'], 'Save As'),
            (r'^(?:press\s+)?(ctrl\s*\+\s*shift\s*\+\s*c|control\s+shift\s+c|ctrl\s+shift\s+c)$', ['ctrl', 'shift', 'c'], 'Inspect Element'),
            (r'^(?:press\s+)?(ctrl\s*\+\s*shift\s*\+\s*v|control\s+shift\s+v|ctrl\s+shift\s+v)$', ['ctrl', 'shift', 'v'], 'Paste as Plain Text'),
            (r'^(?:press\s+)?(ctrl\s*\+\s*shift\s*\+\s*b|control\s+shift\s+b|ctrl\s+shift\s+b)$', ['ctrl', 'shift', 'b'], 'Toggle Bookmarks Bar'),
            (r'^(?:press\s+)?(ctrl\s*\+\s*shift\s*\+\s*e|control\s+shift\s+e|ctrl\s+shift\s+e)$', ['ctrl', 'shift', 'e'], 'Open Explorer'),
            (r'^(?:press\s+)?(ctrl\s*\+\s*shift\s*\+\s*o|control\s+shift\s+o|ctrl\s+shift\s+o)$', ['ctrl', 'shift', 'o'], 'Open Bookmarks Manager'),
        ]
        
        for pattern, keys, response_text in keyboard_shortcuts:
            match = re.match(pattern, user_input_lower)
            if match:
                # Special handling for function keys (F1-F12)
                if keys is None:
                    keys = [match.group(1)]
                return self._create_response(
                    intent='keyboard_shortcut',
                    action='press_keys',
                    parameters={'keys': keys},
                    command=f'keys:{"+".join(keys)}',
                    response=response_text,
                    confidence=0.95,
                    model=model_name
                )

        # ===== 0.56. CAPS LOCK CONTROL =====
        # Handle caps lock commands: "caps on", "caps off", "caps lock on", "caps lock off"
        # Also handles common misrecognitions like "caps of" instead of "caps off"
        caps_patterns = [
            (r'^(?:caps lock on|caps on|capitals on|all caps on)$', 'on', 'Caps Lock ON'),
            (r'^(?:caps lock off|caps lock of|caps off|caps of|capitals off|all caps off)$', 'off', 'Caps Lock OFF'),
            (r'^(?:caps lock|toggle caps|toggle caps lock)$', 'toggle', 'Toggling Caps Lock'),
        ]
        
        for pattern, action, response_text in caps_patterns:
            if re.match(pattern, user_input_lower):
                return self._create_response(
                    intent='caps_lock',
                    action=action,
                    parameters={},
                    command=f'caps_lock:{action}',
                    response=response_text,
                    confidence=0.95,
                    model=model_name
                )

        # ===== 0.6. WINDOW CONTROL BY APPLICATION NAME =====
        # Handle commands like "maximize chrome", "minimize notepad", "close calculator"
        # Check this FIRST before generic window control
        window_action_patterns = [
            (r'^(?:minimize|minimise)\s+(.+)$', 'minimize', 'Minimizing'),
            (r'^(?:maximize|maximise)\s+(.+)$', 'maximize', 'Maximizing'),
            (r'^(?:restore)\s+(.+)$', 'maximize', 'Restoring'),
            (r'^(?:close)\s+(.+)$', 'close', 'Closing'),
        ]
        
        for pattern, action, verb in window_action_patterns:
            match = re.match(pattern, user_input_lower)
            if match:
                app_name = match.group(1).strip()
                print(f"DEBUG: Window control matched - action={action}, app={app_name}")
                # Don't match if it's just "close window" or "close this" (handled below)
                if app_name not in ['window', 'this', 'this window', 'application', 'app', 'program']:
                    return self._create_response(
                        intent='window_control_app',
                        action=action,
                        parameters={'app_name': app_name},
                        command=f'{action}_app:{app_name}',
                        response=f'{verb} {app_name.title()}',
                        confidence=0.9,
                        model=model_name
                    )

        # ===== 0.6.5 WINDOW CONTROL (ACTIVE WINDOW) =====
        # Handle window management commands for active window (minimize, maximize, close)
        # Only if no specific app name was mentioned above
        minimize_keywords = ['minimize', 'minimise', 'minimize window', 'minimise window', 'minimize this', 'minimise this']
        maximize_keywords = ['maximize', 'maximise', 'maximize window', 'maximise window', 'maximize this', 'maximise this', 'restore window', 'restore this']
        close_window_keywords = ['close','close window', 'close this window', 'close this', 'close application', 'close app', 'close program']
        
        if any(keyword == user_input_lower or user_input_lower.startswith(keyword + ' ') for keyword in minimize_keywords):
            return self._create_response(
                intent='window_control',
                action='minimize',
                parameters={},
                command='minimize_window',
                response='Minimizing window',
                confidence=0.95,
                model=model_name
            )
        
        if any(keyword == user_input_lower or user_input_lower.startswith(keyword + ' ') for keyword in maximize_keywords):
            return self._create_response(
                intent='window_control',
                action='maximize',
                parameters={},
                command='maximize_window',
                response='Maximizing window',
                confidence=0.95,
                model=model_name
            )
        
        if any(keyword == user_input_lower or user_input_lower.startswith(keyword + ' ') for keyword in close_window_keywords):
            return self._create_response(
                intent='window_control',
                action='close',
                parameters={},
                command='close_window',
                response='Closing window',
                confidence=0.95,
                model=model_name
            )

        # ===== 0.7. HELP/CAPABILITIES =====
        help_keywords = ['what can you do', 'help', 'capabilities', 'what are you', 'what do you', 'features', 'abilities', 'what are your', 'how can you help', 'what can you help with']
        if any(keyword in user_input_lower for keyword in help_keywords):
            capabilities = [
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
            full_response = "Here's what I can help you with:\n\n" + "\n".join(f"• {cap}" for cap in capabilities) + "\n\nJust tell me what you'd like me to do!"
            return self._create_response(
                intent='information',
                action='list_capabilities',
                parameters={},
                command='help',
                response=full_response,
                confidence=0.95,
                model=model_name
            )

        # ===== 1. WEATHER INFORMATION =====
        weather_keywords = ['weather', 'temperature', 'forecast', 'rain', 'sunny', 'cold', 'hot', 'climate']
        # Use word boundary matching to avoid false positives (e.g., "hot" in "photos")
        matched_weather = [kw for kw in weather_keywords if re.search(r'\b' + re.escape(kw) + r'\b', user_input_lower)]
        if matched_weather:
            return self._create_response(
                intent='information',
                action='weather',
                parameters={'query': 'weather'},
                command='weather_search',
                response="Let me search for weather information",
                confidence=0.9,
                model=model_name
            )



        # ===== 2. VOLUME CONTROLS =====
        volume_patterns = ['volume', 'sound', 'audio']
        if any(word in user_input_lower for word in volume_patterns):
            # Check for setting volume to a specific level FIRST: "set volume to 50", "volume 33", "increase volume to 70", "decrease volume to 0"
            # This regex handles: "volume to 50", "volume 50", "set volume to 50", "increase volume to 70", "decrease volume to 0"
            volume_set_match = re.search(r'(?:set|increase|decrease|up|down)?\s*(?:volume|sound|audio)\s+(?:to\s+)?(\d+|zero|one|two|three|four|five|six|seven|eight|nine|ten)', user_input_lower)
            if volume_set_match:
                level_str = volume_set_match.group(1)
                # Convert word numbers to digits
                word_to_num = {'zero': 0, 'one': 1, 'two': 2, 'three': 3, 'four': 4, 'five': 5, 
                               'six': 6, 'seven': 7, 'eight': 8, 'nine': 9, 'ten': 10}
                level = word_to_num.get(level_str, int(level_str) if level_str.isdigit() else 50)
                # Clamp level between 0 and 100
                level = max(0, min(100, level))
                return self._create_response(
                    intent='system_command',
                    action='volume_set',
                    parameters={'level': level},
                    command=f'volume_set:{level}',
                    response=f"Setting volume to {level}%",
                    confidence=0.95,
                    model=model_name
                )
            # Then check for increase/decrease without specific numbers
            elif any(word in user_input_lower for word in ['up', 'increase', 'raise', 'louder', 'higher', 'more']):
                return self._create_response(
                    intent='system_command',
                    action='volume_up',
                    command='volume_up',
                    response="Increasing volume",
                    confidence=0.95,
                    model=model_name
                )
            elif any(word in user_input_lower for word in ['down', 'decrease', 'lower', 'quieter', 'less']):
                return self._create_response(
                    intent='system_command',
                    action='volume_down',
                    command='volume_down',
                    response="Decreasing volume",
                    confidence=0.95,
                    model=model_name
                )
        
        # ===== 2. BRIGHTNESS CONTROLS =====
        brightness_patterns = ['brightness', 'screen brightness', 'display brightness']
        if any(word in user_input_lower for word in brightness_patterns):
            # Check for setting brightness to a specific level FIRST: "set brightness to 50", "brightness 33", "increase brightness to 70", "decrease brightness to 0"
            # This regex handles: "brightness to 50", "brightness 50", "set brightness to 50", "increase brightness to 70", "decrease brightness to 0"
            brightness_set_match = re.search(r'(?:set|increase|decrease|up|down)?\s*(?:brightness|screen brightness|display brightness)\s+(?:to\s+)?(\d+|zero|one|two|three|four|five|six|seven|eight|nine|ten)', user_input_lower)
            if brightness_set_match:
                level_str = brightness_set_match.group(1)
                # Convert word numbers to digits
                word_to_num = {'zero': 0, 'one': 1, 'two': 2, 'three': 3, 'four': 4, 'five': 5, 
                               'six': 6, 'seven': 7, 'eight': 8, 'nine': 9, 'ten': 10}
                level = word_to_num.get(level_str, int(level_str) if level_str.isdigit() else 50)
                # Clamp level between 0 and 100
                level = max(0, min(100, level))
                return self._create_response(
                    intent='system_command',
                    action='brightness_set',
                    parameters={'level': level},
                    command=f'brightness_set:{level}',
                    response=f"Setting brightness to {level}%",
                    confidence=0.95,
                    model=model_name
                )
            # Then check for increase/decrease without specific numbers
            elif any(word in user_input_lower for word in ['up', 'increase', 'raise', 'higher', 'more']):
                return self._create_response(
                    intent='system_command',
                    action='brightness_up',
                    command='brightness_up',
                    response="Increasing brightness",
                    confidence=0.95,
                    model=model_name
                )
            elif any(word in user_input_lower for word in ['down', 'decrease', 'lower', 'less']):
                return self._create_response(
                    intent='system_command',
                    action='brightness_down',
                    command='brightness_down',
                    response="Decreasing brightness",
                    confidence=0.95,
                    model=model_name
                )
        
        # ===== 2. MUTE/UNMUTE =====
        if any(word in user_input_lower for word in ['mute', 'unmute', 'silence', 'silent']):
            return self._create_response(
                intent='system_command',
                action='mute',
                command='mute',
                response="Toggling mute",
                confidence=0.95,
                model=model_name
            )
        


        # ===== 4. FILE/FOLDER OPERATIONS =====
        # Documents folder
        if any(word in user_input_lower for word in ['document', 'documents']):
            if any(word in user_input_lower for word in ['folder', 'open', 'show', 'go to']):
                return self._create_response(
                    intent='file_operation',
                    action='open_folder',
                    parameters={'folder': str(os.path.expanduser('~/Documents'))},
                    command='explorer Documents',
                    response="Opening Documents folder",
                    confidence=0.95,
                    model=model_name
                )
        
        # Downloads folder
        if any(word in user_input_lower for word in ['download', 'downloads']):
            if any(word in user_input_lower for word in ['folder', 'open', 'show', 'go to']):
                return self._create_response(
                    intent='file_operation',
                    action='open_folder',
                    parameters={'folder': str(os.path.expanduser('~/Downloads'))},
                    command='explorer Downloads',
                    response="Opening Downloads folder",
                    confidence=0.95,
                    model=model_name
                )
        
        # ===== 4.5 OPEN SPECIFIC FILE/FOLDER =====
        # Handle "open ... folder" or "open ... file" commands
        # This should be checked before general application opening
        if 'open' in user_input_lower:
            # Check for "open ... folder" pattern - captures everything between "open" and "folder/directory"
            # Handles: "open java script workshop folder", "open boot camp folder", etc.
            folder_match = re.search(r'open\s+(.+?)(?:\s+folder|\s+directory)$', user_input_lower)
            if folder_match:
                folder_name = folder_match.group(1).strip()
                return self._create_response(
                    intent='file_operation',
                    action='open_folder_by_name',
                    parameters={'folder_name': folder_name},
                    command=f'open_folder:{folder_name}',
                    response=f"Opening folder: {folder_name}",
                    confidence=0.9,
                    model=model_name
                )
            
            # Check for "open filename.ext file" pattern - handles "open style.js file", "open document.pdf file"
            # This must come BEFORE the general "open ... file" pattern to properly extract filenames with extensions
            # Pattern: "open" + space + filename with extension + space + "file" at the end
            file_with_ext_and_file_match = re.search(r'open\s+([\w\-\.]+\.[a-z]{2,4})\s+file$', user_input_lower)
            if file_with_ext_and_file_match:
                file_name = file_with_ext_and_file_match.group(1).strip()
                return self._create_response(
                    intent='file_operation',
                    action='open_file_by_name',
                    parameters={'file_name': file_name},
                    command=f'open_file:{file_name}',
                    response=f"Opening file: {file_name}",
                    confidence=0.9,
                    model=model_name
                )
            
            # Check for "open ... file" pattern - captures everything between "open" and "file"
            # This handles cases without file extension like "open my document file"
            file_match = re.search(r'open\s+(.+?)(?:\s+file)$', user_input_lower)
            if file_match:
                file_name = file_match.group(1).strip()
                return self._create_response(
                    intent='file_operation',
                    action='open_file_by_name',
                    parameters={'file_name': file_name},
                    command=f'open_file:{file_name}',
                    response=f"Opening file: {file_name}",
                    confidence=0.9,
                    model=model_name
                )
            
            # Check for "open filename.ext" pattern (files with extensions, no "file" keyword)
            # Matches: "open style.js", "open document.pdf", "open script.py"
            # Also matches: "open style.js file", "open document.pdf please"
            # IMPORTANT: Must have a valid extension (2-4 chars after dot) to avoid matching app names
            open_file_ext_match = re.search(r'open\s+([\w\-\.]+\.[a-z]{2,4})(?:\s+(?:file|please|folder|directory))?(?:\s*$)', user_input_lower)
            if open_file_ext_match:
                file_name = open_file_ext_match.group(1).strip()
                # Double-check it actually has an extension (contains a dot)
                if '.' in file_name:
                    return self._create_response(
                        intent='file_operation',
                        action='open_file_by_name',
                        parameters={'file_name': file_name},
                        command=f'open_file:{file_name}',
                        response=f"Opening file: {file_name}",
                        confidence=0.85,
                        model=model_name
                    )
        
        # File Explorer
        explorer_patterns = ['file explorer', 'files', 'file manager', 'my files', 'explorer']
        if any(pattern in user_input_lower for pattern in explorer_patterns):
            return self._create_response(
                intent='open_application',
                action='open',
                parameters={'app': 'explorer'},
                command='explorer',
                response="Opening File Explorer",
                confidence=0.95,
                model=model_name
            )
        
        # Create folder
        if 'create' in user_input_lower or 'new' in user_input_lower:
            if 'folder' in user_input_lower or 'directory' in user_input_lower:
                folder_name = self._extract_name(user_input_lower, ['folder', 'directory'])
                return self._create_response(
                    intent='file_operation',
                    action='create_folder',
                    parameters={'folder': folder_name},
                    command=f'mkdir {folder_name}',
                    response=f"Creating folder: {folder_name}",
                    confidence=0.85,
                    model=model_name
                )
        
        # ===== 5. WEB SEARCH =====
        # Check for search intent
        action, target = self.extract_action_and_target(user_input)

        if action == 'search' and target:
            return self._create_response(
                intent='web_search',
                action='search',
                parameters={'query': target},
                command=f'search:{target}',
                response=f"Searching for: {target}",
                confidence=0.9,
                model=model_name
            )
        
        # Handle 'show' commands - treat as search if target is instructional/query-like
        if action == 'show' and target:
            # If target contains question words or multiple words, treat as search
            question_words = ['how', 'what', 'when', 'where', 'why', 'who', 'which']
            target_words = target.split()
            is_instructional = any(qw in target_words for qw in question_words) or len(target_words) >= 3
            
            if is_instructional:
                return self._create_response(
                    intent='web_search',
                    action='search',
                    parameters={'query': target},
                    command=f'search:{target}',
                    response=f"Searching for: {target}",
                    confidence=0.85,
                    model=model_name
                )

        # ===== 6. OPEN APPLICATIONS =====
        # This now handles all variations: open, launch, start, run, fire up, etc.
        if action == 'open':
            # Special handling for "open google" - should be web search, not application
            target_lower = target.lower().strip()
            if target_lower == 'google':
                return self._create_response(
                    intent='web_search',
                    action='search',
                    parameters={'query': ''},
                    command='search:',
                    response="Opening Google",
                    confidence=0.95,
                    model=model_name
                )

            # Special handling for common application aliases

            # Handle "command prompt" variations - more robust matching
            if ('command' in target_lower and 'prompt' in target_lower) or 'cmd' in target_lower or 'terminal' in target_lower:
                cmd_info = self.kb.get_command('cmd')
                if cmd_info:
                    return self._create_response(
                        intent='open_application',
                        action='open',
                        parameters={'app': 'cmd'},
                        command=cmd_info['path'],
                        response="Opening Command Prompt",
                        confidence=0.95,
                        model=model_name
                    )

            # Handle "file explorer" variations
            if 'file' in target_lower and ('explorer' in target_lower or target_lower == 'file'):
                cmd_info = self.kb.get_command('explorer')
                if cmd_info:
                    return self._create_response(
                        intent='open_application',
                        action='open',
                        parameters={'app': 'explorer'},
                        command=cmd_info['path'],
                        response="Opening File Explorer",
                        confidence=0.95,
                        model=model_name
                    )

            # Handle other common aliases
            app_aliases = {
                'notepad': ['notepad', 'text editor', 'editor'],
                'calculator': ['calc', 'calculator'],
                'paint': ['paint', 'mspaint'],
                'word': ['word', 'microsoft word', 'ms word'],
                'excel': ['excel', 'microsoft excel', 'ms excel'],
                'chrome': ['chrome', 'google chrome', 'browser'],
                'vscode': ['vscode', 'visual studio code', 'code'],
                'taskmgr': ['task manager', 'taskmgr'],
                'control': ['control panel', 'control'],
                'settings': ['settings', 'windows settings']
            }

            for app_name, aliases in app_aliases.items():
                if any(alias in target_lower for alias in aliases):
                    cmd_info = self.kb.get_command(app_name)
                    if cmd_info:
                        try:
                            cmd_path = cmd_info['path']
                            # Use app_name for response, not description (which may contain "Open")
                            return self._create_response(
                                intent='open_application',
                                action='open',
                                parameters={'app': app_name},
                                command=cmd_path,
                                response=f"Opening {app_name.title()}",
                                confidence=0.9,
                                model=model_name
                            )
                        except (KeyError, TypeError):
                            # Handle Mock objects or missing keys
                            return self._create_response(
                                intent='open_application',
                                action='open',
                                parameters={'app': app_name},
                                command=app_name,
                                response=f"Opening {app_name}",
                                confidence=0.5,
                                model=model_name
                            )

            # Try exact match first
            cmd_info = self.kb.get_command(target)

            # If not found, try fuzzy matching
            if not cmd_info:
                fuzzy_match = self.fuzzy_match_command(target, threshold=0.6)
                if fuzzy_match:
                    cmd_info = fuzzy_match
                    print(f"✨ Fuzzy matched: '{target}' → '{fuzzy_match['name']}'")

            if cmd_info:
                try:
                    cmd_name = cmd_info['name']
                    cmd_path = cmd_info['path']
                    cmd_desc = cmd_info['description']
                    # Remove "Open " prefix from description if present to avoid "Opening Open Calculator"
                    clean_desc = cmd_desc
                    if clean_desc.lower().startswith('open '):
                        clean_desc = clean_desc[5:]
                    return self._create_response(
                        intent='open_application',
                        action='open',
                        parameters={'app': cmd_name},
                        command=cmd_path,
                        response=f"Opening {clean_desc}",
                        confidence=0.9,
                        model=model_name
                    )
                except (KeyError, TypeError):
                    # Handle Mock objects or missing keys
                    return self._create_response(
                        intent='open_application',
                        action='open',
                        parameters={'app': target},
                        command=target,
                        response=f"Attempting to open: {target}",
                        confidence=0.5,
                        model=model_name
                    )
            else:
                # Try to open anyway (might be a valid executable name)
                return self._create_response(
                    intent='open_application',
                    action='open',
                    parameters={'app': target},
                    command=target,
                    response=f"Attempting to open: {target}",
                    confidence=0.5,
                    model=model_name
                )
        
        # ===== 7. SYSTEM COMMANDS =====
        # Check for abort/shutdown cancel FIRST (higher priority)
        abort_keywords = ['abort', 'cancel', 'stop', 'nevermind', "never mind"]
        if any(keyword in user_input_lower for keyword in abort_keywords):
            if 'shutdown' in user_input_lower or 'restart' in user_input_lower or 'shut down' in user_input_lower:
                return self._create_response(
                    intent='system_command',
                    action='abort_shutdown',
                    command='abort_shutdown',
                    response="Cancelling shutdown/restart",
                    confidence=0.95,
                    model=model_name
                )
        
        shutdown_keywords = ['shutdown', 'shut down', 'power off', 'turn off']
        if any(keyword in user_input_lower for keyword in shutdown_keywords):
            return self._create_response(
                intent='system_command',
                action='shutdown',
                command='shutdown',
                response="Initiating shutdown",
                confidence=0.95,
                model=model_name
            )
        
        restart_keywords = ['restart', 'reboot', 'reset']
        if any(keyword in user_input_lower for keyword in restart_keywords):
            return self._create_response(
                intent='system_command',
                action='restart',
                command='restart',
                response="Initiating restart",
                confidence=0.95,
                model=model_name
            )
        
        # ===== 8. CONVERSATION/GREETINGS =====
        # Handle conversational inputs and greetings
        conversational_phrases = [
            'hello', 'hi', 'hey', 'good morning', 'good afternoon', 'good evening',
            'thanks', 'thank you', 'thank', 'how are you', 'how do you do',
            'nice to meet you', 'pleased to meet you', 'what\'s up', 'sup',
            'how\'s it going', 'hru', 'how r u', 'greetings'
        ]

        # Check for exact matches with conversational phrases
        if user_input_lower in conversational_phrases:
            return self._create_response(
                intent='conversation',
                action='chat',
                parameters={'input': user_input},
                command='',
                response=self._generate_greeting_response(user_input_lower),
                confidence=0.9,
                model=model_name
            )

        # Handle single word conversational inputs
        if len(user_input.split()) == 1:
            single_word_conversational = ['hello', 'hi', 'hey', 'thanks', 'thank']
            if user_input_lower in single_word_conversational:
                return self._create_response(
                    intent='conversation',
                    action='chat',
                    parameters={'input': user_input},
                    command='',
                    response=self._generate_greeting_response(user_input_lower),
                    confidence=0.8,
                    model=model_name
                )

        # ===== 9. FAILED TO PARSE =====
        return {
            'success': False,
            'intent': 'unknown',
            'action': '',
            'parameters': {},
            'command': '',
            'response': self._generate_failure_message(user_input),
            'confidence': 0.0,
            'model': model_name
        }
    
    def _extract_name(self, text: str, keywords: List[str]) -> str:
        """Extract name from text after keywords like 'called', 'named'"""
        for indicator in ['called', 'named', 'name']:
            if indicator in text:
                parts = text.split(indicator, 1)
                if len(parts) > 1:
                    name = parts[1].strip()
                    # Remove keyword if present
                    for keyword in keywords:
                        name = name.replace(keyword, '').strip()
                    return name if name else 'New Folder'
        return 'New Folder'
    
    def _extract_text_to_type(self, user_input: str) -> str:
        """Extract text to type from voice command"""
        return self._extract_text_to_type_with_case(user_input.lower(), user_input)

    def _extract_text_to_type_with_case(self, user_input_lower: str, original_input: str) -> str:
        """Extract text to type from voice command with pre-processed lowercase version"""
        import re
        
        # Pattern to match typing commands: "type X", "write X", "dictate X", etc.
        pattern = r'^(?:type|write|dictate|enter|input)\s+(.+)$'
        
        match = re.match(pattern, user_input_lower)
        if match:
            text = match.group(1)
            # Get the original case text from the end of original_input
            start_pos = user_input_lower.find(text)
            if start_pos != -1:
                text = original_input[start_pos:]
            
            # Clean up common filler words at the start
            filler_starts = ['saying ', 'that says ', 'quote ', 'text ', 'the text ']
            text_lower = text.lower()
            for filler in filler_starts:
                if text_lower.startswith(filler):
                    text = text[len(filler):].strip()
                    text_lower = text.lower()
            
            # Replace "space" keyword with actual space character
            # Handle patterns like "hello space world" -> "hello world"
            text = self._replace_space_keywords(text)
            
            return text.strip()
        
        # Fallback: return everything after first word
        words = original_input.split(None, 1)
        if len(words) > 1:
            text = words[1].strip()
            # Replace "space" keyword with actual space character
            text = self._replace_space_keywords(text)
            return text
        return ''
    
    def _replace_space_keywords(self, text: str) -> str:
        """Replace 'space' keywords with actual space characters"""
        import re
        
        # Replace standalone 'space' (as a word) with actual space
        # Use word boundaries to avoid replacing 'space' within other words like 'spaceman'
        result = re.sub(r'\bspace\b', ' ', text, flags=re.IGNORECASE)
        
        # Normalize multiple spaces to single space
        result = re.sub(r'\s+', ' ', result)
        
        return result
    
    def _generate_greeting_response(self, greeting: str) -> str:
        """Generate appropriate response for greetings"""
        greeting_responses = {
            'hello': "Hello! How can I help you today?",
            'hi': "Hi there! What can I do for you?",
            'hey': "Hey! How can I assist you?",
            'good morning': "Good morning! How can I help you start your day?",
            'good afternoon': "Good afternoon! How can I help you?",
            'good evening': "Good evening! How can I assist you?",
            'thanks': "You're welcome! Is there anything else I can help with?",
            'thank you': "You're welcome! Happy to help.",
            'thank': "You're welcome! Is there anything else I can do for you?",
            'how are you': "I'm doing well, thank you! How can I help you today?",
            'how do you do': "I'm doing well! How can I assist you?",
            'nice to meet you': "Nice to meet you too! How can I help?",
            'pleased to meet you': "The pleasure is mine! How can I assist you?",
            'what\'s up': "Not much! How can I help you today?",
            'sup': "Hey! What's up? How can I help?",
            'how\'s it going': "Going well! How can I help you today?",
            'hru': "I'm good! How about you? How can I help?",
            'how r u': "I'm doing great! How can I assist you?",
            'greetings': "Greetings! How can I help you today?"
        }

        return greeting_responses.get(greeting, "Hello! How can I help you?")

    def _generate_failure_message(self, user_input: str) -> str:
        """Generate helpful failure message with suggestions"""
        # Try to suggest similar commands
        suggestions = []
        all_commands = self.kb.search_commands("")[:5]  # Top 5 commands

        for cmd in all_commands:
            suggestions.append(cmd['name'])

        suggestion_text = ", ".join(suggestions) if suggestions else "chrome, notepad, calculator"

        return f"I didn't understand '{user_input}'. Try commands like: open {suggestion_text}"
    
    def _create_response(self, intent: str, action: str, command: str = '', 
                        parameters: Dict = None, response: str = '', 
                        confidence: float = 0.9, model: str = 'Enhanced Fallback') -> Dict[str, Any]:
        """Helper to create standardized response"""
        return {
            'success': True,
            'intent': intent,
            'action': action,
            'parameters': parameters or {},
            'command': command,
            'response': response,
            'confidence': confidence,
            'model': model
        }
    
    def _create_error_response(self, error_message: str, model: str = 'Enhanced Fallback') -> Dict[str, Any]:
        """Helper to create standardized error response"""
        return {
            'success': False,
            'intent': 'error',
            'action': '',
            'parameters': {},
            'command': '',
            'response': f"I encountered an error: {error_message}",
            'confidence': 0.0,
            'model': model
        }

"""
Voice Assistant with Multi-Model AI Integration
Complete implementation with STT, TTS, Knowledge Base, and System Control
Currently uses Gemini Flash 2.0 via OpenRouter API
"""

import os
import json
import sqlite3
import subprocess
import webbrowser
import logging
import time
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, Optional
from dotenv import load_dotenv

# Import our AI manager
from ai_manager import AIManager

# Local exception definitions
class CommandExecutionError(Exception):
    def __init__(self, message: str, user_message: str = None):
        super().__init__(message)
        self.user_message = user_message or message

class DatabaseError(Exception):
    def __init__(self, message: str, user_message: str = None):
        super().__init__(message)
        self.user_message = user_message or message

class ConfigurationError(Exception):
    def __init__(self, message: str, user_message: str = None):
        super().__init__(message)
        self.user_message = user_message or message

class InitializationError(Exception):
    def __init__(self, message: str, user_message: str = None):
        super().__init__(message)
        self.user_message = user_message or message

# Simple retry decorator (replacing utils.py)
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

# Load environment variables
load_dotenv()

class KnowledgeBase:
    """SQLite-based knowledge base for commands and system information"""
    
    def __init__(self, db_path: str):
        self.db_path = db_path
        self.logger = logging.getLogger('knowledge_base')
        self.conn = sqlite3.connect(db_path, check_same_thread=False)
        self.cursor = self.conn.cursor()
        self._initialize_database()
    
    def _initialize_database(self):
        """Create tables for knowledge base"""
        try:
            # Commands table
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS commands (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    command_name TEXT UNIQUE,
                    command_type TEXT,
                    execution_path TEXT,
                    parameters TEXT,
                    description TEXT,
                    usage_count INTEGER DEFAULT 0,
                    last_used TIMESTAMP
                )
            ''')

            # Command history
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    user_input TEXT,
                    intent TEXT,
                    executed_command TEXT,
                    success BOOLEAN,
                    response TEXT
                )
            ''')

            # User preferences
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS preferences (
                    key TEXT PRIMARY KEY,
                    value TEXT,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')

            # Conversation history for context memory
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS conversation_context (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    session_id TEXT DEFAULT 'default',
                    role TEXT NOT NULL,
                    content TEXT NOT NULL,
                    timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')

            self.conn.commit()
            self.logger.debug("Database tables initialized successfully")
            self._populate_default_commands()
        except sqlite3.Error as e:
            error_msg = f"Failed to initialize database tables: {str(e)}"
            self.logger.error(error_msg, exc_info=True)
            raise DatabaseError(error_msg, "Unable to create database tables. Please check database integrity.")
    
    def _populate_default_commands(self):
        """Add default system commands"""
        default_commands = [
            # Web Browsers
            ('chrome', 'application', 'chrome', '', 'Open Google Chrome', 0, None),
            ('firefox', 'application', 'firefox', '', 'Open Mozilla Firefox', 0, None),
            ('edge', 'application', 'msedge', '', 'Open Microsoft Edge', 0, None),
            ('brave', 'application', 'brave', '', 'Open Brave Browser', 0, None),
            ('opera', 'application', 'opera', '', 'Open Opera Browser', 0, None),
            
            # System Applications
            ('notepad', 'application', 'notepad', '', 'Open Notepad', 0, None),
            ('notepad_plus_plus', 'application', 'notepad++', '', 'Open Notepad++', 0, None),
            ('wordpad', 'application', 'write', '', 'Open WordPad', 0, None),
            ('explorer', 'application', 'explorer', '', 'Open File Explorer', 0, None),
            ('calculator', 'application', 'calc', '', 'Open Calculator', 0, None),
            ('paint', 'application', 'mspaint', '', 'Open Paint', 0, None),
            ('paint3d', 'application', 'paint3d', '', 'Open Paint 3D', 0, None),
            ('snipping_tool', 'application', 'snippingtool', '', 'Open Snipping Tool', 0, None),
            ('photos', 'application', 'start ms-photos:', '', 'Open Photos App', 0, None),
            ('camera', 'application', 'start microsoft.windows.camera:', '', 'Open Camera', 0, None),
            
            # Development Tools
            ('vscode', 'application', 'code', '', 'Open Visual Studio Code', 0, None),
            ('visual_studio', 'application', 'devenv', '', 'Open Visual Studio', 0, None),
            ('pycharm', 'application', 'pycharm', '', 'Open PyCharm', 0, None),
            ('intellij', 'application', 'idea', '', 'Open IntelliJ IDEA', 0, None),
            ('eclipse', 'application', 'eclipse', '', 'Open Eclipse IDE', 0, None),
            ('sublime_text', 'application', 'sublime_text', '', 'Open Sublime Text', 0, None),
            ('atom', 'application', 'atom', '', 'Open Atom Editor', 0, None),
            ('git_bash', 'application', 'sh', '--login -i', 'Open Git Bash', 0, None),
            ('terminal', 'application', 'wt', '', 'Open Windows Terminal', 0, None),
            ('powershell', 'application', 'powershell', '', 'Open PowerShell', 0, None),
            ('cmd', 'application', 'cmd', '', 'Open Command Prompt', 0, None),
            ('azure_data_studio', 'application', 'azuredatastudio', '', 'Open Azure Data Studio', 0, None),
            
            # Microsoft Office
            ('word', 'application', 'winword', '', 'Open Microsoft Word', 0, None),
            ('excel', 'application', 'excel', '', 'Open Microsoft Excel', 0, None),
            ('powerpoint', 'application', 'powerpnt', '', 'Open Microsoft PowerPoint', 0, None),
            ('outlook', 'application', 'outlook', '', 'Open Microsoft Outlook', 0, None),
            ('onenote', 'application', 'onenote', '', 'Open Microsoft OneNote', 0, None),
            ('access', 'application', 'msaccess', '', 'Open Microsoft Access', 0, None),
            ('publisher', 'application', 'mspub', '', 'Open Microsoft Publisher', 0, None),
            ('teams', 'application', 'teams', '', 'Open Microsoft Teams', 0, None),
            ('skype', 'application', 'skype', '', 'Open Skype', 0, None),
            
            # Media Players
            ('vlc', 'application', 'vlc', '', 'Open VLC Media Player', 0, None),
            ('windows_media_player', 'application', 'wmplayer', '', 'Open Windows Media Player', 0, None),
            ('spotify', 'application', 'spotify', '', 'Open Spotify', 0, None),
            ('itunes', 'application', 'itunes', '', 'Open iTunes', 0, None),
            ('groove_music', 'application', 'start ms-windows-store://pdp/?ProductId=9wzdncrfj3qw', '', 'Open Groove Music', 0, None),
            ('movies_tv', 'application', 'start microsoft.zunevideo:', '', 'Open Movies & TV', 0, None),
            
            # Communication Apps
            ('discord', 'application', 'discord', '', 'Open Discord', 0, None),
            ('slack', 'application', 'slack', '', 'Open Slack', 0, None),
            ('zoom', 'application', 'zoom', '', 'Open Zoom', 0, None),
            ('whatsapp', 'application', 'whatsapp', '', 'Open WhatsApp', 0, None),
            ('telegram', 'application', 'telegram', '', 'Open Telegram', 0, None),
            ('signal', 'application', 'signal', '', 'Open Signal', 0, None),
            
            # Cloud Storage
            ('onedrive', 'application', 'onedrive', '', 'Open OneDrive', 0, None),
            ('dropbox', 'application', 'dropbox', '', 'Open Dropbox', 0, None),
            ('google_drive', 'application', 'start https://drive.google.com', '', 'Open Google Drive in Browser', 0, None),
            
            # System Utilities
            ('task_manager', 'application', 'taskmgr', '', 'Open Task Manager', 0, None),
            ('control_panel', 'application', 'control', '', 'Open Control Panel', 0, None),
            ('settings', 'application', 'ms-settings:', '', 'Open Windows Settings', 0, None),
            ('device_manager', 'application', 'devmgmt.msc', '', 'Open Device Manager', 0, None),
            ('disk_management', 'application', 'diskmgmt.msc', '', 'Open Disk Management', 0, None),
            ('event_viewer', 'application', 'eventvwr.msc', '', 'Open Event Viewer', 0, None),
            ('services', 'application', 'services.msc', '', 'Open Services', 0, None),
            ('registry_editor', 'application', 'regedit', '', 'Open Registry Editor', 0, None),
            ('system_config', 'application', 'msconfig', '', 'Open System Configuration', 0, None),
            ('directx_diagnostic', 'application', 'dxdiag', '', 'Open DirectX Diagnostic Tool', 0, None),
            ('character_map', 'application', 'charmap', '', 'Open Character Map', 0, None),
            ('resource_monitor', 'application', 'perfmon /res', '', 'Open Resource Monitor', 0, None),
            ('performance_monitor', 'application', 'perfmon', '', 'Open Performance Monitor', 0, None),
            
            # Games & Entertainment
            ('xbox', 'application', 'start xbox:', '', 'Open Xbox App', 0, None),
            ('steam', 'application', 'steam', '', 'Open Steam', 0, None),
            ('epic_games', 'application', 'EpicGamesLauncher', '', 'Open Epic Games Launcher', 0, None),
            ('minecraft', 'application', 'minecraftlauncher', '', 'Open Minecraft Launcher', 0, None),
            
            # PDF Readers
            ('adobe_reader', 'application', 'acrord32', '', 'Open Adobe Reader', 0, None),
            ('adobe_acrobat', 'application', 'acrobat', '', 'Open Adobe Acrobat', 0, None),
            ('foxit_reader', 'application', 'foxitreader', '', 'Open Foxit Reader', 0, None),
            
            # Archive Tools
            ('winrar', 'application', 'winrar', '', 'Open WinRAR', 0, None),
            ('7zip', 'application', '7zfm', '', 'Open 7-Zip', 0, None),
            
            # Security
            ('windows_security', 'application', 'start windowsdefender:', '', 'Open Windows Security', 0, None),
            ('bitlocker', 'application', 'manage-bde.wsh', '', 'Open BitLocker Management', 0, None),
        ]
        
        for cmd in default_commands:
            try:
                self.cursor.execute('''
                    INSERT OR IGNORE INTO commands 
                    (command_name, command_type, execution_path, parameters, description, usage_count, last_used)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', cmd)
            except sqlite3.IntegrityError:
                pass
        
        self.conn.commit()
    
    def get_command(self, command_name: str) -> Optional[Dict]:
        """Retrieve command information. Prefer exact match so 'notepad' doesn't match 'notepad_plus_plus'."""
        try:
            # Exact match first (case-insensitive)
            self.cursor.execute('''
                SELECT command_name, command_type, execution_path, parameters, description
                FROM commands WHERE LOWER(command_name) = LOWER(?)
                LIMIT 1
            ''', (command_name.strip(),))
            result = self.cursor.fetchone()

            # Fallback: partial match, prefer shortest name (e.g. "word" over "wordpad")
            if not result:
                self.cursor.execute('''
                    SELECT command_name, command_type, execution_path, parameters, description
                    FROM commands WHERE command_name LIKE ?
                    ORDER BY LENGTH(command_name) ASC, usage_count DESC
                    LIMIT 1
                ''', (f'%{command_name.strip()}%',))
                result = self.cursor.fetchone()

            if result:
                return {
                    'name': result[0],
                    'type': result[1],
                    'path': result[2],
                    'params': result[3],
                    'description': result[4]
                }
            return None
        except sqlite3.Error as e:
            error_msg = f"Failed to retrieve command '{command_name}': {str(e)}"
            self.logger.error(error_msg, exc_info=True)
            raise DatabaseError(error_msg, "Unable to retrieve command information from database.")
    
    def update_usage(self, command_name: str):
        """Update command usage statistics"""
        self.cursor.execute('''
            UPDATE commands 
            SET usage_count = usage_count + 1, last_used = CURRENT_TIMESTAMP
            WHERE command_name = ?
        ''', (command_name,))
        self.conn.commit()
    
    def add_to_history(self, user_input: str, intent: str, command: str, success: bool, response: str):
        """Add command to history"""
        try:
            self.cursor.execute('''
                INSERT INTO history (user_input, intent, executed_command, success, response)
                VALUES (?, ?, ?, ?, ?)
            ''', (user_input, intent, command, success, response))
            self.conn.commit()
            self.logger.debug(f"Added command to history: {intent} - {command}")
        except sqlite3.Error as e:
            error_msg = f"Failed to add command to history: {str(e)}"
            self.logger.error(error_msg, exc_info=True)
            raise DatabaseError(error_msg, "Unable to save command history.")
    
    def get_recent_history(self, limit: int = 10) -> list:
        """Get recent command history"""
        self.cursor.execute('''
            SELECT timestamp, user_input, intent, executed_command, success, response
            FROM history ORDER BY timestamp DESC LIMIT ?
        ''', (limit,))
        
        return [
            {
                'timestamp': row[0],
                'input': row[1],
                'intent': row[2],
                'command': row[3],
                'success': bool(row[4]),
                'response': row[5]
            }
            for row in self.cursor.fetchall()
        ]
    
    def search_commands(self, query: str) -> list:
        """Search for commands by name or description"""
        try:
            self.cursor.execute('''
                SELECT command_name, description, execution_path
                FROM commands
                WHERE command_name LIKE ? OR description LIKE ?
                ORDER BY usage_count DESC
            ''', (f'%{query}%', f'%{query}%'))

            return [
                {'name': row[0], 'description': row[1], 'path': row[2]}
                for row in self.cursor.fetchall()
            ]
        except sqlite3.Error as e:
            error_msg = f"Failed to search commands with query '{query}': {str(e)}"
            self.logger.error(error_msg, exc_info=True)
            raise DatabaseError(error_msg, "Unable to search commands in database.")
    
    def close(self):
        """Close database connection"""
        self.conn.close()

    def save_conversation_message(self, role: str, content: str, session_id: str = 'default'):
        """Save a conversation message to persistent storage"""
        try:
            self.cursor.execute('''
                INSERT INTO conversation_context (session_id, role, content, timestamp)
                VALUES (?, ?, ?, CURRENT_TIMESTAMP)
            ''', (session_id, role, content))
            self.conn.commit()
            self.logger.debug(f"Saved conversation message: {role} - {content[:50]}...")
        except sqlite3.Error as e:
            self.logger.warning(f"Failed to save conversation message: {str(e)}")

    def get_conversation_history(self, limit: int = 10, session_id: str = 'default') -> list:
        """Get recent conversation history from database"""
        try:
            self.cursor.execute('''
                SELECT role, content, timestamp
                FROM conversation_context
                WHERE session_id = ?
                ORDER BY timestamp ASC
                LIMIT ?
            ''', (session_id, limit))
            
            return [
                {
                    'role': row[0],
                    'content': row[1],
                    'timestamp': row[2]
                }
                for row in self.cursor.fetchall()
            ]
        except sqlite3.Error as e:
            self.logger.warning(f"Failed to get conversation history: {str(e)}")
            return []

    def clear_conversation_history(self, session_id: str = 'default'):
        """Clear conversation history for a session"""
        try:
            self.cursor.execute('''
                DELETE FROM conversation_context WHERE session_id = ?
            ''', (session_id,))
            self.conn.commit()
            self.logger.info(f"Cleared conversation history for session: {session_id}")
        except sqlite3.Error as e:
            self.logger.warning(f"Failed to clear conversation history: {str(e)}")





class CommandExecutor:
    """Execute system commands based on parsed intent"""

    def __init__(self, knowledge_base: KnowledgeBase):
        self.kb = knowledge_base
        self.logger = logging.getLogger('command_executor')
    
    def execute(self, parsed_command: Dict[str, Any]) -> Dict[str, Any]:
        """Execute the parsed command"""
        intent = parsed_command.get('intent')
        action = parsed_command.get('action')
        parameters = parsed_command.get('parameters', {})
        command = parsed_command.get('command', '')

        try:
            if intent == 'open_application':
                return self._open_application(parameters, command)

            elif intent == 'web_search':
                return self._web_search(parameters)

            elif intent == 'system_command':
                return self._system_command(action, parameters)

            elif intent == 'file_operation':
                return self._file_operation(action, parameters)

            elif intent == 'type_text':
                return self._type_text(parameters)

            elif intent == 'change_text_case':
                return self._change_text_case(parameters)

            elif intent == 'caps_lock':
                return self._toggle_caps_lock(action)

            elif intent == 'window_control':
                return self._window_control(action)

            elif intent == 'window_control_app':
                return self._window_control_app(action, parameters)

            elif intent == 'keyboard_shortcut':
                return self._press_keyboard_shortcut(parameters)

            elif intent == 'text_edit':
                return self._text_edit(action, parameters)

            elif intent == 'information':
                return self._handle_information(action, parameters, parsed_command)

            elif intent == 'conversation':
                # For pure conversation, just return the response without executing anything
                return {
                    'success': True,
                    'message': parsed_command.get('response', 'I understand.')
                }

            elif intent == 'affirmative':
                # Handle acknowledgments, thanks, appreciation
                return {
                    'success': True,
                    'message': parsed_command.get('response', "You're welcome! Happy to help.")
                }

            elif intent == 'greeting':
                # Handle greetings
                return {
                    'success': True,
                    'message': parsed_command.get('response', 'Hello! How can I assist you?')
                }

            else:
                return {
                    'success': False,
                    'message': f"Unknown intent: {intent}"
                }

        except Exception as e:
            return {
                'success': False,
                'message': f"Execution error: {str(e)}"
            }
    
    @retry_on_exception(max_attempts=2, exceptions=(OSError, FileNotFoundError))
    def _open_application(self, params: Dict, command: str = '') -> Dict:
        """Open an application"""
        # Use command field if available, otherwise fall back to params
        app_path = command or params.get('app', '')

        if not app_path:
            error_msg = "No application specified for opening"
            self.logger.warning(error_msg)
            raise CommandExecutionError(error_msg, "I need to know which application you'd like me to open. Try saying 'open chrome' or 'open notepad'.")

        try:
            self.logger.debug(f"Attempting to open application: {app_path}")
            
            # Handle UWP apps and protocols (ms-photos:, steam://, etc.)
            if app_path.startswith(('ms-', 'microsoft.', 'steam://', 'com.epicgames.')):
                # Use shell execute for protocol handlers
                subprocess.run(['cmd', '/c', 'start', '', app_path], shell=False)
            elif os.path.isfile(app_path) or (os.path.sep in app_path and os.path.exists(app_path)):
                # Full path to an existing file - use startfile
                os.startfile(app_path)
            else:
                # App name (e.g. notepad, calc, chrome) - must use 'start' to launch from PATH
                # os.startfile('notepad') fails because it looks for a file named "notepad"
                subprocess.run(['cmd', '/c', 'start', '', app_path], shell=False)

            # Update usage in knowledge base
            if 'app' in params:
                try:
                    self.kb.update_usage(params.get('app'))
                except DatabaseError as e:
                    self.logger.warning(f"Failed to update usage for {params.get('app')}: {str(e)}")

            # Wait a moment for the application to start, then bring it to foreground
            import time
            import win32gui
            import win32con
            
            app_name = params.get('app', app_path).lower()
            
            # Try multiple times with reduced delays for faster response (some apps take longer to initialize)
            hwnd = None
            for attempt, delay in enumerate([0.15, 0.25, 0.35, 0.5], 1):
                time.sleep(delay)
                hwnd = self._find_window_by_app_name(app_name)
                if hwnd and win32gui.IsWindow(hwnd):
                    self.logger.info(f"Found {app_path} window on attempt {attempt}")
                    break
            
            if hwnd and win32gui.IsWindow(hwnd):
                try:
                    # Get the current foreground window
                    foreground_hwnd = win32gui.GetForegroundWindow()
                    
                    # If we're not already the foreground, we need to use some tricks
                    if hwnd != foreground_hwnd:
                        # Show the window (restore if minimized)
                        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                        time.sleep(0.05)
                        
                        # Attach to the foreground window's thread (required for SetForegroundWindow)
                        current_thread = win32gui.GetCurrentThreadId()
                        foreground_thread = win32gui.GetWindowThreadProcessId(foreground_hwnd)[0]
                        
                        if foreground_thread != current_thread:
                            win32gui.AttachThreadInput(foreground_thread, current_thread, True)
                        
                        try:
                            # Bring window to foreground
                            win32gui.SetForegroundWindow(hwnd)
                            # Also try SetActiveWindow
                            win32gui.SetActiveWindow(hwnd)
                            self.logger.info(f"Brought {app_path} to foreground")
                        finally:
                            # Detach threads
                            if foreground_thread != current_thread:
                                win32gui.AttachThreadInput(foreground_thread, current_thread, False)
                    
                except Exception as e:
                    self.logger.warning(f"Could not bring {app_path} to foreground: {e}")

            self.logger.info(f"Successfully opened application: {app_path}")
            return {
                'success': True,
                'message': f"Opened {app_path}"
            }

        except FileNotFoundError as e:
            error_msg = f"Application not found: {app_path}"
            self.logger.error(error_msg, exc_info=True)
            raise CommandExecutionError(error_msg, f"I couldn't find the application '{app_path}'. Please make sure it's installed and try again.")
        except OSError as e:
            error_msg = f"OS error opening application {app_path}: {str(e)}"
            self.logger.error(error_msg, exc_info=True)
            raise CommandExecutionError(error_msg, f"There was a problem opening '{app_path}'. It might not be installed or there could be a permissions issue.")
        except Exception as e:
            error_msg = f"Unexpected error opening application {app_path}: {str(e)}"
            self.logger.error(error_msg, exc_info=True)
            raise CommandExecutionError(error_msg, f"Something unexpected happened while trying to open '{app_path}'. Please try again.")
    
    def _web_search(self, params: Dict) -> Dict:
        """Perform web search"""
        query = params.get('query', '')
        
        try:
            url = f"https://www.google.com/search?q={query.replace(' ', '+')}"
            webbrowser.open(url)
            
            return {
                'success': True,
                'message': f"Searching for: {query}"
            }
        
        except Exception as e:
            return {
                'success': False,
                'message': f"Search failed: {str(e)}"
            }
    
    @retry_on_exception(max_attempts=2, exceptions=(subprocess.CalledProcessError, OSError))
    def _system_command(self, action: str, params: Dict) -> Dict:
        """Execute system command"""
        if not action:
            error_msg = "No system action specified"
            self.logger.warning(error_msg)
            raise CommandExecutionError(error_msg, 'I need to know what system action you want. Try saying "increase volume" or "mute".')

        try:
            self.logger.debug(f"Executing system command: {action}")

            if action == 'volume_up':
                # Windows volume up
                subprocess.run(['powershell', '-Command',
                               '(New-Object -ComObject WScript.Shell).SendKeys([char]175)'],
                               capture_output=True, check=True)
                self.logger.info("Volume increased successfully")
                return {'success': True, 'message': 'Volume increased'}

            elif action == 'volume_down':
                # Windows volume down
                subprocess.run(['powershell', '-Command',
                               '(New-Object -ComObject WScript.Shell).SendKeys([char]174)'],
                               capture_output=True, check=True)
                self.logger.info("Volume decreased successfully")
                return {'success': True, 'message': 'Volume decreased'}

            elif action == 'mute':
                # Windows mute
                subprocess.run(['powershell', '-Command',
                               '(New-Object -ComObject WScript.Shell).SendKeys([char]173)'],
                               capture_output=True, check=True)
                self.logger.info("Mute toggled successfully")
                return {'success': True, 'message': 'Mute toggled'}

            elif action == 'volume_set':
                # Set volume to specific level
                level = params.get('level', 50)
                try:
                    # Use PowerShell with Windows CoreAudio - simpler approach
                    ps_script = f'''
$level = {level} / 100.0
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
[Guid("5CDF2C82-841E-4546-9722-0CF74078229A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IAudioEndpointVolume {{
    int f1(); int f2(); int f3(); int f4();
    int SetMasterVolumeLevelScalar(float fLevel, IntPtr pguidEventContext);
}}
[Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDevice {{
    int Activate(ref Guid id, int clsCtx, int activationParams, out IAudioEndpointVolume aev);
}}
[Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDeviceEnumerator {{
    int f1(); int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice endpoint);
}}
[ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
class MMDeviceEnumerator {{ }}
public class Volume {{
    public static void SetVolume(float level) {{
        var enumerator = new MMDeviceEnumerator() as IMMDeviceEnumerator;
        IMMDevice dev;
        enumerator.GetDefaultAudioEndpoint(0, 1, out dev);
        IAudioEndpointVolume epv;
        var epvid = typeof(IAudioEndpointVolume).GUID;
        dev.Activate(ref epvid, 0, 0, out epv);
        epv.SetMasterVolumeLevelScalar(level, IntPtr.Zero);
    }}
}}
"@
[Volume]::SetVolume($level)
'''
                    result = subprocess.run(['powershell', '-Command', ps_script], capture_output=True, text=True)
                    if result.returncode == 0:
                        self.logger.info(f"Volume set to {level}%")
                        return {'success': True, 'message': f'Volume set to {level}%'}
                    else:
                        self.logger.error(f"PowerShell error: {result.stderr}")
                        # Fallback to key presses
                        return self._adjust_volume_by_keys(level)
                except Exception as e:
                    self.logger.error(f"Failed to set volume: {e}")
                    # Fallback to key presses
                    return self._adjust_volume_by_keys(level)

            elif action == 'brightness_set':
                # Set brightness to specific level
                level = params.get('level', 50)
                return self._set_brightness(level)

            elif action == 'brightness_up':
                return self._adjust_brightness('up')

            elif action == 'brightness_down':
                return self._adjust_brightness('down')

            elif action == 'shutdown':
                subprocess.run(['shutdown', '/s', '/t', '10'], capture_output=True, check=True)
                self.logger.info("System shutdown initiated")
                return {'success': True, 'message': 'Shutting down in 10 seconds'}

            elif action == 'restart':
                subprocess.run(['shutdown', '/r', '/t', '10'], capture_output=True, check=True)
                self.logger.info("System restart initiated")
                return {'success': True, 'message': 'Restarting in 10 seconds'}

            elif action == 'abort_shutdown':
                result = subprocess.run(['shutdown', '/a'], capture_output=True, text=True)
                if result.returncode == 0:
                    self.logger.info("Shutdown/restart aborted successfully")
                    return {'success': True, 'message': 'Shutdown/restart cancelled!'}
                else:
                    # No shutdown was scheduled
                    self.logger.info("No pending shutdown found")
                    return {'success': True, 'message': 'No pending shutdown to cancel'}

            else:
                error_msg = f"Unknown system action: {action}"
                self.logger.warning(error_msg)
                raise CommandExecutionError(error_msg, f"I don't know how to handle '{action}'. Try volume control commands.")

        except subprocess.CalledProcessError as e:
            error_msg = f"System command '{action}' failed with exit code {e.returncode}"
            self.logger.error(error_msg, exc_info=True)
            raise CommandExecutionError(error_msg, f"The system command '{action}' failed. There might be a permissions issue.")
        except FileNotFoundError as e:
            error_msg = f"Required system tools not found for action '{action}'"
            self.logger.error(error_msg, exc_info=True)
            raise CommandExecutionError(error_msg, f"I couldn't find the required system tools to perform '{action}'. This might be a system configuration issue.")
        except Exception as e:
            error_msg = f"Unexpected error in system command '{action}': {str(e)}"
            self.logger.error(error_msg, exc_info=True)
            raise CommandExecutionError(error_msg, f"Something went wrong with the '{action}' command. Please try again.")

    def _adjust_volume_by_keys(self, target_level: int) -> Dict:
        """Fallback: adjust volume using keyboard shortcuts"""
        try:
            import pyautogui
            # Mute first to get to 0, then unmute and adjust
            # This is approximate - press volume up multiple times
            presses = int(target_level / 2)  # Each press is roughly 2%
            for _ in range(presses):
                pyautogui.keyDown('volumeup')
                pyautogui.keyUp('volumeup')
            return {'success': True, 'message': f'Volume adjusted to approximately {target_level}%'}
        except Exception as e:
            return {'success': False, 'message': f'Could not set volume to {target_level}%'}

    def _set_brightness(self, level: int) -> Dict:
        """Set brightness to a specific level using WMI or keyboard fallback"""
        try:
            # Use PowerShell to set brightness via WMI
            ps_script = f'''
$brightness = {level}
try {{
    $monitor = Get-WmiObject -Namespace root/wmi -Class WmiMonitorBrightnessMethods -ErrorAction Stop
    if ($monitor) {{
        $monitor.WmiSetBrightness(1, $brightness)
        Write-Output "SUCCESS"
    }} else {{
        Write-Output "NO_MONITOR"
    }}
}} catch {{
    Write-Output "ERROR: $_"
}}
'''
            result = subprocess.run(['powershell', '-Command', ps_script], capture_output=True, text=True)
            output = result.stdout.strip()
            
            if "SUCCESS" in output:
                self.logger.info(f"Brightness set to {level}%")
                return {'success': True, 'message': f'Brightness set to {level}%'}
            elif "NO_MONITOR" in output or "ERROR" in output:
                # Fallback to keyboard brightness keys
                return self._adjust_brightness_by_keys(level)
            else:
                return self._adjust_brightness_by_keys(level)
        except Exception as e:
            self.logger.error(f"Failed to set brightness: {e}")
            return self._adjust_brightness_by_keys(level)

    def _adjust_brightness(self, direction: str) -> Dict:
        """Adjust brightness up or down by 10%"""
        try:
            ps_script = f'''
try {{
    $monitor = Get-WmiObject -Namespace root/wmi -Class WmiMonitorBrightnessMethods -ErrorAction Stop
    $current = (Get-WmiObject -Namespace root/wmi -Class WmiMonitorBrightness -ErrorAction Stop).CurrentBrightness
    if ("{direction}" -eq "up") {{
        $new = [Math]::Min(100, $current + 10)
    }} else {{
        $new = [Math]::Max(0, $current - 10)
    }}
    $monitor.WmiSetBrightness(1, $new)
    Write-Output "SUCCESS"
}} catch {{
    Write-Output "ERROR"
}}
'''
            result = subprocess.run(['powershell', '-Command', ps_script], capture_output=True, text=True)
            if "SUCCESS" in result.stdout:
                action_text = "increased" if direction == 'up' else "decreased"
                self.logger.info(f"Brightness {action_text}")
                return {'success': True, 'message': f'Brightness {action_text}'}
            else:
                # Fallback to keyboard
                return self._adjust_brightness_by_keys(10 if direction == 'up' else -10)
        except Exception as e:
            self.logger.error(f"Failed to adjust brightness: {e}")
            return self._adjust_brightness_by_keys(10 if direction == 'up' else -10)

    def _adjust_brightness_by_keys(self, change: int) -> Dict:
        """Fallback: adjust brightness using keyboard brightness keys"""
        try:
            import pyautogui
            # Use brightness keys (Fn + F5/F6 or dedicated brightness keys)
            presses = abs(int(change / 10))  # Each press is roughly 10%
            key = 'brightnessup' if change > 0 else 'brightnessdown'
            
            for _ in range(max(1, presses)):
                pyautogui.keyDown(key)
                pyautogui.keyUp(key)
            
            action_text = "increased" if change > 0 else "decreased"
            return {'success': True, 'message': f'Brightness {action_text} (using keyboard keys)'}
        except Exception as e:
            return {'success': False, 'message': 'Brightness control not available on this device. This feature requires a laptop with built-in display or a monitor that supports DDC/CI.'}
    
    def _handle_information(self, action: str, params: Dict, parsed_command: Dict) -> Dict:
        """Handle information requests - including weather"""
        try:
            response = parsed_command.get('response', '')
            
            # Handle weather specifically with web search
            if action == 'weather':
                query = params.get('query', 'weather')
                url = f"https://www.google.com/search?q={query.replace(' ', '+')}"
                webbrowser.open(url)
                return {
                    'success': True,
                    'message': response or 'Opening weather information in your browser'
                }
            
            # Handle capabilities
            elif action == 'list_capabilities':
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
                response = response or "Here's what I can help you with:\n\n" + "\n".join(f"• {cap}" for cap in capabilities) + "\n\nJust tell me what you'd like to do!"
            
            return {
                'success': True,
                'message': response
            }
        
        except Exception as e:
            return {
                'success': False,
                'message': f'Information request failed: {str(e)}'
            }
    
    def _file_operation(self, action: str, params: Dict) -> Dict:
        """Perform file operations"""
        try:
            if action == 'open_folder':
                folder = params.get('folder', str(Path.home()))
                os.startfile(folder)
                return {'success': True, 'message': f'Opened {folder}'}
            
            elif action == 'create_folder':
                folder = params.get('folder', '')
                Path(folder).mkdir(parents=True, exist_ok=True)
                return {'success': True, 'message': f'Created folder: {folder}'}
            
            elif action == 'open_folder_by_name':
                folder_name = params.get('folder_name', '')
                return self._open_folder_by_name(folder_name)
            
            elif action == 'open_file_by_name':
                file_name = params.get('file_name', '')
                return self._open_file_by_name(file_name)
            
            else:
                return {'success': False, 'message': f'Unknown file operation: {action}'}
        
        except Exception as e:
            return {'success': False, 'message': f'File operation failed: {str(e)}'}

    def _open_folder_by_name(self, folder_name: str) -> Dict:
        """Open a folder by searching for it - prioritizes current folder, then falls back to common locations"""
        try:
            import os
            from pathlib import Path
            import re
            
            # Clean up folder name
            folder_name = folder_name.strip().lower()
            
            # Get current working directory (represents current opened folder)
            current_folder = Path.cwd()
            
            # Common locations to search
            search_paths = [
                current_folder,  # Prioritize current working directory first
                Path.home(),  # User home directory
                Path.home() / 'Desktop',
                Path.home() / 'Documents',
                Path.home() / 'Downloads',
                Path.home() / 'Pictures',
                Path.home() / 'Music',
                Path.home() / 'Videos',
                Path('C:\\'),
            ]
            
            # Add D: drive if it exists
            if os.path.exists('D:\\'):
                search_paths.append(Path('D:\\'))
            
            # Create variations of the folder name for matching
            # e.g., "java script" -> "javascript", "boot camp" -> "bootcamp"
            folder_name_no_spaces = folder_name.replace(' ', '')
            folder_name_with_spaces = ' '.join(folder_name.split())  # normalize spaces
            
            # Search for the folder
            for base_path in search_paths:
                if base_path is None or not base_path.exists():
                    continue
                    
                try:
                    folders = [item for item in base_path.iterdir() if item.is_dir()]
                    
                    # Determine if this is the current folder for better messaging
                    is_current_folder = base_path == current_folder
                    location_text = "from current folder" if is_current_folder else f"from {base_path.name}"
                    
                    # Try exact match first
                    for item in folders:
                        item_name_lower = item.name.lower()
                        if item_name_lower == folder_name:
                            os.startfile(str(item))
                            return {'success': True, 'message': f'Opened folder: {item.name} ({location_text})'}
                    
                    # Try match without spaces (e.g., "java script" matches "javascript")
                    for item in folders:
                        item_name_lower = item.name.lower()
                        item_name_no_spaces = item_name_lower.replace(' ', '')
                        if item_name_no_spaces == folder_name_no_spaces:
                            os.startfile(str(item))
                            return {'success': True, 'message': f'Opened folder: {item.name} ({location_text})'}
                    
                    # Try partial match - folder name contains search term or vice versa
                    for item in folders:
                        item_name_lower = item.name.lower()
                        item_name_no_spaces = item_name_lower.replace(' ', '')
                        
                        # Check if search term is in folder name
                        if folder_name in item_name_lower or folder_name_no_spaces in item_name_no_spaces:
                            os.startfile(str(item))
                            return {'success': True, 'message': f'Opened folder: {item.name} ({location_text})'}
                        
                        # Check if folder name is in search term
                        if item_name_lower in folder_name or item_name_no_spaces in folder_name_no_spaces:
                            os.startfile(str(item))
                            return {'success': True, 'message': f'Opened folder: {item.name} ({location_text})'}
                        
                        # Check word-by-word match (for multi-word folders)
                        search_words = set(folder_name.split())
                        item_words = set(item_name_lower.split())
                        if search_words and item_words and len(search_words & item_words) >= min(len(search_words), len(item_words)) * 0.5:
                            os.startfile(str(item))
                            return {'success': True, 'message': f'Opened folder: {item.name} ({location_text})'}
                            
                except (PermissionError, OSError):
                    continue
            
            # If not found, try to open as absolute path
            if os.path.exists(folder_name):
                os.startfile(folder_name)
                return {'success': True, 'message': f'Opened: {folder_name}'}
            
            return {'success': False, 'message': f'Could not find folder: {folder_name}'}
            
        except Exception as e:
            return {'success': False, 'message': f'Failed to open folder: {str(e)}'}

    def _open_file_by_name(self, file_name: str) -> Dict:
        """Open a file by searching for it - prioritizes current folder, then falls back to common locations"""
        try:
            import os
            from pathlib import Path
            
            # Clean up file name
            file_name = file_name.strip().lower()
            file_name_no_spaces = file_name.replace(' ', '')
            
            # Get current working directory (represents current opened folder)
            current_folder = Path.cwd()
            
            # Common locations to search (same as folder opening)
            search_paths = [
                current_folder,  # Prioritize current working directory first
                current_folder.parent,  # Also check parent directory
                Path.home(),  # User home directory
                Path.home() / 'Desktop',
                Path.home() / 'Documents',
                Path.home() / 'Downloads',
                Path.home() / 'Pictures',
                Path.home() / 'Music',
                Path.home() / 'Videos',
                Path('C:\\'),
            ]
            
            # Add D: drive if it exists
            if os.path.exists('D:\\'):
                search_paths.append(Path('D:\\'))
            
            # Search for the file
            for base_path in search_paths:
                if not base_path.exists():
                    continue
                    
                try:
                    files = [item for item in base_path.iterdir() if item.is_file()]
                    
                    # Check if file name has an extension (contains a dot)
                    has_extension = '.' in file_name
                    
                    # Determine if this is the current folder for better messaging
                    is_current_folder = base_path == current_folder
                    location_text = "from current folder" if is_current_folder else f"from {base_path.name}"
                    
                    # Try exact match first (filename without extension)
                    for item in files:
                        item_stem_lower = item.stem.lower()
                        item_name_lower = item.name.lower()
                        
                        # Match by stem (without extension)
                        if item_stem_lower == file_name:
                            os.startfile(str(item))
                            return {'success': True, 'message': f'Opened file: {item.name} ({location_text})'}
                        
                        # Match by full filename (with extension) if search term has extension
                        if has_extension and item_name_lower == file_name:
                            os.startfile(str(item))
                            return {'success': True, 'message': f'Opened file: {item.name} ({location_text})'}
                    
                    # Try match without spaces (e.g., "my file" matches "myfile")
                    for item in files:
                        item_stem_lower = item.stem.lower()
                        item_stem_no_spaces = item_stem_lower.replace(' ', '')
                        if item_stem_no_spaces == file_name_no_spaces:
                            os.startfile(str(item))
                            return {'success': True, 'message': f'Opened file: {item.name} ({location_text})'}
                    
                    # Try partial match - file name contains search term or vice versa
                    for item in files:
                        item_name_lower = item.name.lower()
                        item_stem_lower = item.stem.lower()
                        item_stem_no_spaces = item_stem_lower.replace(' ', '')
                        
                        # Check if search term is in file name (both stem and full name)
                        if file_name in item_name_lower or file_name in item_stem_lower or file_name_no_spaces in item_stem_no_spaces:
                            os.startfile(str(item))
                            return {'success': True, 'message': f'Opened file: {item.name} ({location_text})'}
                        
                        # Check if file name is in search term
                        if item_stem_lower in file_name or item_stem_no_spaces in file_name_no_spaces or item_name_lower in file_name:
                            os.startfile(str(item))
                            return {'success': True, 'message': f'Opened file: {item.name} ({location_text})'}
                        
                        # Check word-by-word match (for multi-word files)
                        search_words = set(file_name.split())
                        item_words = set(item_stem_lower.split())
                        if search_words and item_words and len(search_words & item_words) >= min(len(search_words), len(item_words)) * 0.5:
                            os.startfile(str(item))
                            return {'success': True, 'message': f'Opened file: {item.name} ({location_text})'}
                            
                except (PermissionError, OSError):
                    continue
            
            # If not found, try to open as absolute path
            if os.path.exists(file_name):
                os.startfile(file_name)
                return {'success': True, 'message': f'Opened: {file_name}'}
            
            return {'success': False, 'message': f'Could not find file: {file_name}'}
            
        except Exception as e:
            return {'success': False, 'message': f'Failed to open file: {str(e)}'}
    
    def _type_text(self, params: Dict) -> Dict:
        """Type text into the active window using pyautogui"""
        try:
            import pyautogui
            
            text = params.get('text', '')
            case_type = params.get('case_type', '')
            
            if not text:
                return {'success': False, 'message': 'No text provided to type'}
            
            # Apply case transformation if specified
            if case_type:
                text = self._apply_case_transform(text, case_type)
            
            # Safety: ensure pyautogui failsafe is enabled (move mouse to corner to abort)
            pyautogui.FAILSAFE = True
            
            # Small delay to let user move cursor if needed
            import time
            time.sleep(0.5)
            
            # Type the text
            pyautogui.typewrite(text, interval=0.01)
            
            self.logger.info(f"Successfully typed text: {text[:50]}{'...' if len(text) > 50 else ''}")
            return {
                'success': True, 
                'message': f'Typed: {text[:50]}{"..." if len(text) > 50 else ""}'
            }
            
        except ImportError:
            self.logger.error("pyautogui not installed")
            return {'success': False, 'message': 'Typing functionality not available - pyautogui not installed'}
        except Exception as e:
            self.logger.error(f"Error typing text: {str(e)}")
            return {'success': False, 'message': f'Failed to type text: {str(e)}'}

    def _apply_case_transform(self, text: str, case_type: str) -> str:
        """Apply case transformation to text"""
        case_type = case_type.lower()
        
        if case_type in ['uppercase', 'upper', 'caps', 'all caps']:
            return text.upper()
        elif case_type in ['lowercase', 'lower', 'small']:
            return text.lower()
        elif case_type in ['title', 'titlecase']:
            return text.title()
        elif case_type in ['sentence', 'sentencecase']:
            return text.capitalize()
        elif case_type in ['camel', 'camelcase']:
            words = text.split()
            if words:
                return words[0].lower() + ''.join(word.capitalize() for word in words[1:])
            return text.lower()
        elif case_type in ['pascal', 'pascalcase']:
            words = text.split()
            return ''.join(word.capitalize() for word in words)
        elif case_type in ['snake', 'snakecase']:
            return text.lower().replace(' ', '_')
        elif case_type in ['kebab', 'kebabcase']:
            return text.lower().replace(' ', '-')
        else:
            return text

    def _toggle_caps_lock(self, action: str) -> Dict:
        """Toggle caps lock on/off"""
        try:
            import pyautogui
            import win32api
            import win32con
            
            # Safety: ensure pyautogui failsafe is enabled
            pyautogui.FAILSAFE = True
            
            # Check current caps lock state
            current_state = win32api.GetKeyState(win32con.VK_CAPITAL)
            is_caps_on = current_state & 0x0001 != 0
            
            if action == 'on':
                if not is_caps_on:
                    # Turn caps lock on
                    pyautogui.keyDown('capslock')
                    pyautogui.keyUp('capslock')
                    self.logger.info("Caps Lock turned ON")
                    return {'success': True, 'message': 'Caps Lock is now ON'}
                else:
                    return {'success': True, 'message': 'Caps Lock is already ON'}
                    
            elif action == 'off':
                if is_caps_on:
                    # Turn caps lock off
                    pyautogui.keyDown('capslock')
                    pyautogui.keyUp('capslock')
                    self.logger.info("Caps Lock turned OFF")
                    return {'success': True, 'message': 'Caps Lock is now OFF'}
                else:
                    return {'success': True, 'message': 'Caps Lock is already OFF'}
            
            elif action == 'toggle':
                # Just toggle regardless of current state
                pyautogui.keyDown('capslock')
                pyautogui.keyUp('capslock')
                new_state = "ON" if not is_caps_on else "OFF"
                self.logger.info(f"Caps Lock toggled to {new_state}")
                return {'success': True, 'message': f'Caps Lock is now {new_state}'}
            
            else:
                return {'success': False, 'message': f'Unknown caps lock action: {action}'}
                
        except ImportError:
            self.logger.error("pyautogui or win32api not installed")
            return {'success': False, 'message': 'Caps Lock control not available - required modules not installed'}
        except Exception as e:
            self.logger.error(f"Error toggling caps lock: {str(e)}")
            return {'success': False, 'message': f'Failed to toggle caps lock: {str(e)}'}

    def _change_text_case(self, params: Dict) -> Dict:
        """Change the case of selected text or provided text"""
        try:
            import pyautogui
            import time
            
            # Safety: ensure pyautogui failsafe is enabled
            pyautogui.FAILSAFE = True
            
            case_type = params.get('case_type', 'lowercase').lower()
            text = params.get('text', '')
            
            # If no text provided, try to get selected text from clipboard
            if not text:
                # Save current clipboard content
                try:
                    original_clipboard = pyautogui.clipboard.paste()
                except:
                    original_clipboard = ""
                
                # Copy selected text to clipboard
                pyautogui.keyDown('ctrl')
                pyautogui.keyDown('c')
                pyautogui.keyUp('c')
                pyautogui.keyUp('ctrl')
                time.sleep(0.1)
                
                try:
                    text = pyautogui.clipboard.paste()
                    # Restore original clipboard if nothing was selected
                    if not text or text == original_clipboard:
                        return {'success': False, 'message': 'No text selected. Please select text first or provide text to convert.'}
                except:
                    return {'success': False, 'message': 'Could not access clipboard. Please provide text to convert.'}
            
            # Apply case transformation
            if case_type == 'uppercase' or case_type == 'upper':
                converted_text = text.upper()
                case_name = 'UPPERCASE'
            elif case_type == 'lowercase' or case_type == 'lower':
                converted_text = text.lower()
                case_name = 'lowercase'
            elif case_type == 'title' or case_type == 'titlecase':
                converted_text = text.title()
                case_name = 'Title Case'
            elif case_type == 'sentence' or case_type == 'sentencecase':
                converted_text = text.capitalize()
                case_name = 'Sentence case'
            elif case_type == 'camel' or case_type == 'camelcase':
                # Convert to camelCase (first word lowercase, rest title case, no spaces)
                words = text.split()
                if words:
                    converted_text = words[0].lower() + ''.join(word.capitalize() for word in words[1:])
                else:
                    converted_text = text.lower()
                case_name = 'camelCase'
            elif case_type == 'pascal' or case_type == 'pascalcase':
                # Convert to PascalCase (all words title case, no spaces)
                words = text.split()
                converted_text = ''.join(word.capitalize() for word in words)
                case_name = 'PascalCase'
            elif case_type == 'snake' or case_type == 'snakecase':
                # Convert to snake_case (lowercase with underscores)
                converted_text = text.lower().replace(' ', '_')
                case_name = 'snake_case'
            elif case_type == 'kebab' or case_type == 'kebabcase':
                # Convert to kebab-case (lowercase with hyphens)
                converted_text = text.lower().replace(' ', '-')
                case_name = 'kebab-case'
            elif case_type == 'default' or case_type == 'normal':
                # Return text as-is (default/original)
                converted_text = text
                case_name = 'default'
            else:
                # Default to lowercase if unknown case type
                converted_text = text.lower()
                case_name = 'lowercase'
            
            # If we got text from selection, paste the converted text back
            if not params.get('text'):
                # Copy converted text to clipboard
                pyautogui.clipboard.copy(converted_text)
                time.sleep(0.05)
                
                # Paste to replace selected text
                pyautogui.keyDown('ctrl')
                pyautogui.keyDown('v')
                pyautogui.keyUp('v')
                pyautogui.keyUp('ctrl')
                time.sleep(0.05)
                
                # Restore original clipboard content
                if original_clipboard:
                    pyautogui.clipboard.copy(original_clipboard)
                
                self.logger.info(f"Converted selected text to {case_name}")
                return {
                    'success': True,
                    'message': f'Converted to {case_name}: {converted_text[:50]}{"..." if len(converted_text) > 50 else ""}'
                }
            else:
                # Just return the converted text without pasting
                self.logger.info(f"Converted text to {case_name}")
                return {
                    'success': True,
                    'message': f'{case_name}: {converted_text}',
                    'converted_text': converted_text
                }
                
        except ImportError:
            self.logger.error("pyautogui not installed")
            return {'success': False, 'message': 'Text case control not available - pyautogui not installed'}
        except Exception as e:
            self.logger.error(f"Error changing text case: {str(e)}")
            return {'success': False, 'message': f'Failed to change text case: {str(e)}'}
    
    def _window_control(self, action: str) -> Dict:
        """Control the active window (minimize, maximize, close) using keyboard shortcuts"""
        try:
            import pyautogui
            import time
            
            # Safety: ensure pyautogui failsafe is enabled
            pyautogui.FAILSAFE = True
            
            # Small delay before action - reduced for faster response
            time.sleep(0.1)
            
            if action == 'minimize':
                # Use the hotkey function which is more reliable than keyDown/keyUp
                pyautogui.hotkey('win', 'down')
                time.sleep(0.05)  # Brief delay after (reduced)
                self.logger.info("Minimized active window")
                return {'success': True, 'message': 'Minimized window'}
                
            elif action == 'maximize':
                # Use the hotkey function which is more reliable
                pyautogui.hotkey('win', 'up')
                time.sleep(0.05)  # Brief delay after (reduced)
                self.logger.info("Maximized active window")
                return {'success': True, 'message': 'Maximized window'}
                
            elif action == 'close':
                # Use hotkey for Alt+F4
                pyautogui.hotkey('alt', 'f4')
                time.sleep(0.05)  # Brief delay after (reduced)
                self.logger.info("Closed active window")
                return {'success': True, 'message': 'Closed window'}
                
            else:
                return {'success': False, 'message': f'Unknown window action: {action}'}
                
        except ImportError:
            self.logger.error("pyautogui not installed")
            return {'success': False, 'message': 'Window control not available - pyautogui not installed'}
        except Exception as e:
            self.logger.error(f"Error controlling window: {str(e)}")
            return {'success': False, 'message': f'Failed to control window: {str(e)}'}
    
    def _window_control_app(self, action: str, params: Dict) -> Dict:
        """Control a specific application's window (minimize, maximize, close) by app name"""
        try:
            import pyautogui
            import time
            import win32gui
            import win32con
            import re
            
            app_name = params.get('app_name', '').lower()
            if not app_name:
                return {'success': False, 'message': 'No application name provided'}
            
            # Find window by application name
            hwnd = self._find_window_by_app_name(app_name)
            
            if not hwnd:
                # If closing an app that's not open, offer to open it first
                if action == 'close':
                    return {
                        'success': False, 
                        'message': f'{app_name.title()} is not currently open. Would you like me to open it instead?',
                        'suggestion': f'Say "open {app_name}" to launch it.'
                    }
                elif action == 'maximize':
                    return {
                        'success': False, 
                        'message': f'{app_name.title()} is not currently open. Would you like me to open and maximize it?',
                        'suggestion': f'Say "open {app_name}" to launch it.'
                    }
                elif action == 'minimize':
                    return {
                        'success': False, 
                        'message': f'{app_name.title()} is not currently open. I can only minimize apps that are already running.',
                        'suggestion': f'Say "open {app_name}" to launch it first.'
                    }
                return {'success': False, 'message': f'Could not find window for {app_name}'}
            
            # Validate window is still valid before trying to control it
            if not win32gui.IsWindow(hwnd):
                return {'success': False, 'message': f'{app_name} window is no longer valid'}
            
            # Bring window to foreground first
            try:
                win32gui.SetForegroundWindow(hwnd)
            except Exception as e:
                self.logger.warning(f"Could not set foreground window: {e}")
            time.sleep(0.1)  # Reduced delay to ensure window is active (was 0.2)
            
            if action == 'minimize':
                # Minimize the window
                try:
                    win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                    self.logger.info(f"Minimized {app_name} window")
                    return {'success': True, 'message': f'Minimized {app_name}'}
                except Exception as e:
                    return {'success': False, 'message': f'Could not minimize {app_name}: {str(e)}'}
                
            elif action == 'maximize':
                # Maximize the window
                try:
                    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                    self.logger.info(f"Maximized {app_name} window")
                    return {'success': True, 'message': f'Maximized {app_name}'}
                except Exception as e:
                    return {'success': False, 'message': f'Could not maximize {app_name}: {str(e)}'}
                
            elif action == 'close':
                # Close the window by sending WM_CLOSE message
                try:
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                    self.logger.info(f"Closed {app_name} window")
                    return {'success': True, 'message': f'Closed {app_name}'}
                except Exception as e:
                    return {'success': False, 'message': f'Could not close {app_name}: {str(e)}'}
                
            else:
                return {'success': False, 'message': f'Unknown window action: {action}'}
                
        except ImportError as e:
            self.logger.error(f"Required module not installed: {str(e)}")
            return {'success': False, 'message': 'Window control not available - required modules not installed'}
        except Exception as e:
            self.logger.error(f"Error controlling {app_name} window: {str(e)}")
            return {'success': False, 'message': f'Failed to control {app_name} window: {str(e)}'}

    def _press_keyboard_shortcut(self, params: Dict) -> Dict:
        """Press keyboard shortcut keys using pyautogui"""
        try:
            import pyautogui
            import time
            import subprocess
            
            keys = params.get('keys', [])
            if not keys:
                return {'success': False, 'message': 'No keys specified'}
            
            # Special handling for Ctrl+Shift+Esc (Task Manager)
            # Windows blocks programmatic sending of this combo for security
            if keys == ['ctrl', 'shift', 'esc']:
                try:
                    # Use os.startfile which is more reliable for Windows apps
                    import os
                    os.startfile('taskmgr')
                    self.logger.info("Opened Task Manager directly")
                    return {
                        'success': True,
                        'message': 'Opening Task Manager'
                    }
                except Exception as e:
                    self.logger.error(f"Failed to open Task Manager: {e}")
                    # Fallback to subprocess
                    try:
                        subprocess.run(['taskmgr'], capture_output=True, shell=True)
                        return {
                            'success': True,
                            'message': 'Opening Task Manager'
                        }
                    except Exception as e2:
                        self.logger.error(f"Fallback also failed: {e2}")
                        return {'success': False, 'message': 'Could not open Task Manager'}
            
            # Special handling for Ctrl+Alt+Delete - cannot be simulated
            if keys == ['ctrl', 'alt', 'delete']:
                return {
                    'success': False,
                    'message': 'Ctrl+Alt+Delete cannot be automated for security reasons. Please press it manually.'
                }
            
            # Safety: ensure pyautogui failsafe is enabled
            pyautogui.FAILSAFE = True
            
            # Small delay before pressing keys
            time.sleep(0.2)
            
            # Press keys using hotkey (handles multiple keys properly)
            if len(keys) == 1:
                pyautogui.press(keys[0])
                self.logger.info(f"Pressed key: {keys[0]}")
            else:
                pyautogui.hotkey(*keys)
                self.logger.info(f"Pressed hotkey: {'+'.join(keys)}")
            
            return {
                'success': True,
                'message': f"Pressed {'+'.join(keys)}"
            }
            
        except ImportError:
            self.logger.error("pyautogui not installed")
            return {'success': False, 'message': 'Keyboard shortcuts not available - pyautogui not installed'}
        except Exception as e:
            self.logger.error(f"Error pressing keyboard shortcut: {str(e)}")
            return {'success': False, 'message': f'Failed to press keys: {str(e)}'}

    def _text_edit(self, action: str, params: Dict) -> Dict:
        """Handle text editing commands like clear, delete words/chars"""
        try:
            import pyautogui
            import time
            
            # Safety: ensure pyautogui failsafe is enabled
            pyautogui.FAILSAFE = True
            
            if action == 'clear_all':
                # Select all and delete
                pyautogui.keyDown('ctrl')
                pyautogui.keyDown('a')
                pyautogui.keyUp('a')
                pyautogui.keyUp('ctrl')
                time.sleep(0.1)
                pyautogui.keyDown('delete')
                pyautogui.keyUp('delete')
                self.logger.info("Cleared all text")
                return {'success': True, 'message': 'Cleared all text'}
                
            elif action == 'delete_words':
                count = params.get('count', 1)
                # Ctrl+Shift+Left selects word by word, then delete
                for _ in range(count):
                    pyautogui.keyDown('ctrl')
                    pyautogui.keyDown('shift')
                    pyautogui.keyDown('left')
                    pyautogui.keyUp('left')
                    pyautogui.keyUp('shift')
                    pyautogui.keyUp('ctrl')
                    time.sleep(0.02)
                    pyautogui.keyDown('delete')
                    pyautogui.keyUp('delete')
                    time.sleep(0.02)
                self.logger.info(f"Deleted {count} words")
                return {'success': True, 'message': f'Deleted {count} words'}
                
            elif action == 'delete_chars':
                count = params.get('count', 1)
                # Press backspace multiple times
                for _ in range(count):
                    pyautogui.keyDown('backspace')
                    pyautogui.keyUp('backspace')
                    time.sleep(0.01)
                self.logger.info(f"Deleted {count} characters")
                return {'success': True, 'message': f'Deleted {count} characters'}
                
            else:
                return {'success': False, 'message': f'Unknown text edit action: {action}'}
                
        except ImportError:
            self.logger.error("pyautogui not installed")
            return {'success': False, 'message': 'Text editing not available - pyautogui not installed'}
        except Exception as e:
            self.logger.error(f"Error in text edit: {str(e)}")
            return {'success': False, 'message': f'Failed to edit text: {str(e)}'}
    
    def _find_window_by_app_name(self, app_name: str):
        """Find window handle (hwnd) by application name"""
        import win32gui
        import re
        
        # Common app name mappings
        app_mappings = {
            'chrome': ['chrome', 'google chrome'],
            'notepad': ['notepad', 'text editor'],
            'calculator': ['calculator', 'calc'],
            'word': ['word', 'microsoft word', 'winword'],
            'excel': ['excel', 'microsoft excel'],
            'vscode': ['visual studio code', 'vscode', 'code'],
            'explorer': ['file explorer', 'explorer'],
            'cmd': ['command prompt', 'cmd', 'consolewindowclass'],
            'paint': ['paint', 'mspaint'],
            'edge': ['microsoft edge', 'edge'],
            'firefox': ['firefox', 'mozilla firefox'],
            'terminal': ['windows terminal', 'terminal'],
        }
        
        # Reverse mapping: if user says "command prompt", map to "cmd"
        reverse_mappings = {}
        for key, values in app_mappings.items():
            for value in values:
                reverse_mappings[value] = key
        
        # Normalize app name - check if it's a reverse mapping first
        normalized_name = reverse_mappings.get(app_name, app_name)
        
        # Get possible names for the app
        search_names = [normalized_name, app_name]
        if normalized_name in app_mappings:
            search_names.extend(app_mappings[normalized_name])
        
        def callback(hwnd, extra):
            try:
                # Check if window is valid and visible
                if win32gui.IsWindow(hwnd) and win32gui.IsWindowVisible(hwnd):
                    window_text = win32gui.GetWindowText(hwnd).lower()
                    class_name = win32gui.GetClassName(hwnd).lower()
                    
                    # Check if any search name matches window title or class name
                    for name in search_names:
                        if name in window_text or name in class_name:
                            extra.append(hwnd)
                            return False  # Stop enumeration
            except Exception:
                # Window might be invalid/destroyed, skip it
                pass
            return True
        
        matching_windows = []
        try:
            win32gui.EnumWindows(callback, matching_windows)
        except Exception as e:
            self.logger.warning(f"Error enumerating windows: {e}")
        
        return matching_windows[0] if matching_windows else None


class VoiceAssistant:
    """Main voice assistant orchestrator"""

    def __init__(self, data_dir: str = None):
        self.logger = logging.getLogger('voice_assistant')

        if data_dir is None:
            # Store database in project folder for easy access
            # Simple folder name: 'db'
            data_dir = str(Path(__file__).parent / 'db')

        self.data_dir = Path(data_dir)
        try:
            self.data_dir.mkdir(exist_ok=True)
            self.logger.debug(f"Data directory created/verified: {self.data_dir}")
        except OSError as e:
            error_msg = f"Failed to create data directory {self.data_dir}: {str(e)}"
            self.logger.error(error_msg, exc_info=True)
            raise ConfigurationError(error_msg, "Unable to create data directory for voice assistant.")

        # Initialize components
        db_path = str(self.data_dir / 'knowledge.db')
        try:
            self.kb = KnowledgeBase(db_path)
            self.logger.debug("Knowledge base initialized successfully")
        except DatabaseError as e:
            self.logger.error(f"Failed to initialize knowledge base: {str(e)}", exc_info=True)
            raise

        # Initialize AI manager (supports multiple models)
        try:
            self.ai_manager = AIManager(self.kb)
            self.logger.debug("AI manager initialized successfully")
        except Exception as e:
            error_msg = f"Failed to initialize AI manager: {str(e)}"
            self.logger.error(error_msg, exc_info=True)
            raise InitializationError(error_msg, "Unable to initialize AI components.")

        try:
            self.executor = CommandExecutor(self.kb)
            self.logger.debug("Command executor initialized successfully")
        except Exception as e:
            error_msg = f"Failed to initialize command executor: {str(e)}"
            self.logger.error(error_msg, exc_info=True)
            raise InitializationError(error_msg, "Unable to initialize command execution system.")

        self.logger.info("Voice assistant initialized successfully")

        # Conversation context and memory
        self.max_history_length = 20
        self.session_id = 'default'
        # Load conversation history from database (last 2 conversations = 4 messages)
        self.conversation_history = self._load_conversation_history()
        self.user_context = {
            'name': None,
            'preferences': {},
            'last_topic': None,
            'mood': 'neutral',
            'personality': 'helpful'
        }

    def _load_conversation_history(self) -> list:
        """Load conversation history from database"""
        try:
            # Load last 4 messages (2 conversations: user->assistant, user->assistant)
            history = self.kb.get_conversation_history(limit=4, session_id=self.session_id)
            self.logger.info(f"Loaded {len(history)} messages from conversation history")
            return history
        except Exception as e:
            self.logger.warning(f"Failed to load conversation history: {str(e)}")
            return []

    def _build_context(self) -> Dict[str, Any]:
        """Build conversation context for AI parsing"""
        return {
            'conversation_history': self.conversation_history[-4:],  # Last 2 conversations (4 messages)
            'user_context': self.user_context,
            'recent_commands': self.kb.get_recent_history(5),  # Last 5 commands
            'available_commands': self.kb.search_commands("")[:10]  # Top 10 commands
        }

    def _add_to_conversation_history(self, role: str, content: str):
        """Add message to conversation history and persist to database"""
        message = {
            'role': role,
            'content': content,
            'timestamp': datetime.now().isoformat()
        }
        self.conversation_history.append(message)

        # Persist to database for context memory across sessions
        try:
            self.kb.save_conversation_message(role, content, self.session_id)
        except Exception as e:
            self.logger.warning(f"Failed to persist conversation message: {str(e)}")

        # Keep history within limits
        if len(self.conversation_history) > self.max_history_length:
            self.conversation_history = self.conversation_history[-self.max_history_length:]

    def _update_user_context(self, parsed_command: Dict, execution_result: Dict):
        """Update user context based on interaction"""
        intent = parsed_command.get('intent', '')

        # Update last topic based on intent
        if intent == 'web_search':
            self.user_context['last_topic'] = 'search'
        elif intent == 'open_application':
            self.user_context['last_topic'] = 'application'
        elif intent == 'conversation':
            self.user_context['last_topic'] = 'chat'
        elif intent == 'system_command':
            self.user_context['last_topic'] = 'system_control'

        # Update mood based on success/failure
        if execution_result.get('success'):
            if self.user_context['mood'] == 'frustrated':
                self.user_context['mood'] = 'neutral'
        else:
            if self.user_context['mood'] == 'neutral':
                self.user_context['mood'] = 'frustrated'
    
    def process_command(self, user_input: str) -> Dict[str, Any]:
        """Process a voice command with conversation context"""
        self.logger.info(f"Processing command: {user_input[:50]}...")

        try:
            # Add to conversation history
            self._add_to_conversation_history('user', user_input)

            # Build context for AI parsing
            context = self._build_context()

            # Parse command using AI manager with context
            parsed = self.ai_manager.parse_command(user_input, context)

            if not parsed.get('success'):
                # Get available commands for suggestions
                try:
                    available_commands = self.kb.search_commands("")
                    command_names = [cmd['name'] for cmd in available_commands[:5]]  # Top 5 commands
                except DatabaseError:
                    command_names = ['chrome', 'notepad', 'calculator']  # Fallback suggestions

                # Create a more conversational response
                suggestions = ", ".join(command_names)
                response_text = f"I'm not sure what you mean by that. You could try saying 'open {command_names[0]}' or 'search for something'. Some things I can help with include: {suggestions}."

                self.logger.warning(f"Failed to parse command: {user_input}")
                self.logger.debug(f"Suggestions provided: {suggestions}")

                # Add AI response to conversation history
                self._add_to_conversation_history('assistant', response_text)

                return {
                    'success': False,
                    'message': "Could not understand command",
                    'response': response_text,
                    'suggestions': command_names
                }

            self.logger.debug(f"Parsed intent: {parsed['intent']}, action: {parsed.get('action', 'N/A')}")

            # Handle conversation intents differently
            if parsed['intent'] == 'conversation':
                # For pure conversation, don't execute commands
                result = {'success': True, 'message': parsed.get('response', 'I understand.')}
                response_text = parsed.get('response', 'I understand.')
            else:
                # Execute command for other intents
                try:
                    result = self.executor.execute(parsed)
                    response_text = parsed.get('response', result['message'])
                except CommandExecutionError as e:
                    self.logger.error(f"Command execution failed: {str(e)}", exc_info=True)
                    result = {'success': False, 'message': str(e.user_message)}
                    response_text = str(e.user_message)

            # Save to history
            try:
                self.kb.add_to_history(
                    user_input=user_input,
                    intent=parsed['intent'],
                    command=parsed['command'],
                    success=result['success'],
                    response=result['message']
                )
            except DatabaseError as e:
                self.logger.warning(f"Failed to save command to history: {str(e)}")

            # Add AI response to conversation history
            self._add_to_conversation_history('assistant', response_text)

            # Update user context based on the interaction
            self._update_user_context(parsed, result)

            return {
                'success': result['success'],
                'message': result['message'],
                'response': response_text
            }

        except Exception as e:
            error_msg = f"Unexpected error processing command '{user_input}': {str(e)}"
            self.logger.error(error_msg, exc_info=True)

            # Try to provide a helpful response even on error
            response_text = "I'm having trouble processing that right now. Please try again."
            self._add_to_conversation_history('assistant', response_text)

            return {
                'success': False,
                'message': "Processing error occurred",
                'response': response_text
            }
    
    def get_history(self, limit: int = 10):
        """Get command history"""
        return self.kb.get_recent_history(limit)

    def get_available_models(self) -> list:
        """Get list of available AI models"""
        try:
            return self.ai_manager.get_available_models()
        except Exception as e:
            self.logger.error(f"Failed to get available models: {str(e)}", exc_info=True)
            return []  # Return empty list on error

    def get_current_model(self) -> str:
        """Get the currently active AI model"""
        try:
            return self.ai_manager.get_current_model()
        except Exception as e:
            self.logger.error(f"Failed to get current model: {str(e)}", exc_info=True)
            return "unknown"  # Return default on error

    def set_model(self, model_name: str) -> bool:
        """Switch to a different AI model"""
        try:
            result = self.ai_manager.set_model(model_name)
            if result:
                self.logger.info(f"Successfully switched to model: {model_name}")
            else:
                self.logger.warning(f"Failed to switch to model: {model_name}")
            return result
        except Exception as e:
            self.logger.error(f"Error switching to model {model_name}: {str(e)}", exc_info=True)
            return False

    def get_model_info(self) -> dict:
        """Get information about all AI models"""
        try:
            return self.ai_manager.get_model_info()
        except Exception as e:
            self.logger.error(f"Failed to get model info: {str(e)}", exc_info=True)
            return {}  # Return empty dict on error

    def clear_conversation_memory(self):
        """Clear conversation history from memory and database"""
        self.conversation_history = []
        try:
            self.kb.clear_conversation_history(self.session_id)
            self.logger.info("Conversation memory cleared")
        except Exception as e:
            self.logger.warning(f"Failed to clear conversation memory: {str(e)}")

    def cleanup(self):
        """Cleanup resources"""
        try:
            self.kb.close()
            self.logger.info("Voice assistant resources cleaned up successfully")
        except Exception as e:
            self.logger.error(f"Error during cleanup: {str(e)}", exc_info=True)


if __name__ == "__main__":
    # Test the assistant
    assistant = VoiceAssistant()
    
    # Test commands
    test_commands = [
        "open chrome",
        "search for python tutorials",
        "open notepad",
        "increase volume"
    ]
    
    for cmd in test_commands:
        result = assistant.process_command(cmd)
        print(f"✅ Result: {result['message']}\n")
    
    assistant.cleanup()

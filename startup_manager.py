"""
Startup Manager for Nova Voice Assistant
Manages Windows startup registration
"""
import os
import sys
import winreg
from pathlib import Path

class StartupManager:
    """Manage Windows startup registration for Nova"""
    
    def __init__(self):
        self.app_name = "Nova Voice Assistant"
        self.registry_key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        
    def get_python_script_path(self):
        """Get the path to main_app.py"""
        # Get the directory where this script is located
        script_dir = Path(__file__).parent
        main_app = script_dir / 'main_app.py'
        return str(main_app)
    
    def is_startup_enabled(self):
        """Check if Nova is registered to start with Windows"""
        try:
            # Check if batch file exists in startup folder
            startup_folder = Path(os.getenv('APPDATA')) / r'Microsoft\Windows\Start Menu\Programs\Startup'
            batch_file = startup_folder / 'NovaVoiceAssistant.bat'
            
            if batch_file.exists():
                return True, str(batch_file)
            else:
                return False, None
        except Exception as e:
            return False, str(e)
    
    def enable_startup(self):
        """Add Nova to Windows startup with auto-browser launch"""
        try:
            # Get the Python executable path
            python_exe = sys.executable
            script_path = self.get_python_script_path()
            
            # Create batch file for startup that launches web server AND opens browser
            batch_content = f'''@echo off
timeout /t 5 /nobreak >nul
start "" "{python_exe}" "{script_path}"
timeout /t 3 /nobreak >nul
start http://localhost:5000
exit
'''
            
            # Save batch file in startup folder
            startup_folder = Path(os.getenv('APPDATA')) / r'Microsoft\Windows\Start Menu\Programs\Startup'
            startup_folder.mkdir(parents=True, exist_ok=True)
            batch_file = startup_folder / 'NovaVoiceAssistant.bat'
            
            with open(batch_file, 'w') as f:
                f.write(batch_content)
            
            return True, "Successfully added to startup (auto-launches web page)"
        except Exception as e:
            return False, f"Error: {str(e)}"
    
    def disable_startup(self):
        """Remove Nova from Windows startup"""
        try:
            # Remove batch file from startup folder
            startup_folder = Path(os.getenv('APPDATA')) / r'Microsoft\Windows\Start Menu\Programs\Startup'
            batch_file = startup_folder / 'NovaVoiceAssistant.bat'
            
            if batch_file.exists():
                batch_file.unlink()
                return True, "Successfully removed from startup"
            else:
                return True, "Not in startup"
        except Exception as e:
            return False, f"Error: {str(e)}"
    
    def toggle_startup(self):
        """Toggle startup registration"""
        enabled, _ = self.is_startup_enabled()
        if enabled:
            return self.disable_startup()
        else:
            return self.enable_startup()

if __name__ == '__main__':
    manager = StartupManager()
    
    if len(sys.argv) > 1:
        action = sys.argv[1]
        if action == 'check':
            enabled, value = manager.is_startup_enabled()
            print(f"Startup Status: {'Enabled' if enabled else 'Disabled'}")
            if value:
                print(f"Command: {value}")
        elif action == 'enable':
            success, msg = manager.enable_startup()
            print(f"{'✅' if success else '❌'} {msg}")
        elif action == 'disable':
            success, msg = manager.disable_startup()
            print(f"{'✅' if success else '❌'} {msg}")
        elif action == 'toggle':
            success, msg = manager.toggle_startup()
            print(f"{'✅' if success else '❌'} {msg}")
    else:
        print("Usage: python startup_manager.py [check|enable|disable|toggle]")

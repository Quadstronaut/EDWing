"""
Elite Dangerous Command Relay - Fixed Looping Version
Key changes:
1. Longer delays between keys (Elite needs time to process)
2. Longer focus delay after SetForegroundWindow
3. File-based input to avoid console focus stealing
4. Continuous loop with command file monitoring

Requirements:
- pip install pywin32
"""

import time
import logging
from typing import List, Tuple, Optional
from pathlib import Path
import os

# Windows API imports
import win32api
import win32con
import win32gui
import win32process

# Configuration
CONFIG = {
    "window_title_contains": "Elite - Dangerous (CLIENT)",
    "process_name": "elitedangerous64",
    "commanders": ["Bistronaut", "Tristronaut", "Quadstronaut"],
    "primary_commander": "Duvrazh",
    "key_send_delay": 0.15,      # 150ms between keys
    "focus_delay": 0.5,          # 500ms after focusing window
    "inter_window_delay": 1.0,   # 1 second between windows
    "command_file": "elite_commands.txt",  # File to read commands from
    "check_interval": 0.5,       # How often to check for new commands
}

# Logging setup
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.StreamHandler(), logging.FileHandler("elite_command_relay.log")],
)
logger = logging.getLogger(__name__)


class CommandRelay:
    def __init__(self):
        self.all_commanders = CONFIG["commanders"] + [CONFIG["primary_commander"]]
        self.running = True
        self.command_file = Path(CONFIG["command_file"])
        self.last_modified = 0
        
        # Create command file if it doesn't exist
        if not self.command_file.exists():
            self.command_file.write_text("")
            self.last_modified = self.command_file.stat().st_mtime
        
        print("=" * 70)
        print("Elite Dangerous Command Relay - FIXED LOOPING VERSION")
        print("=" * 70)
        print(f"Looking for process: '{CONFIG['process_name']}.exe'")
        print(f"Window title must contain: '{CONFIG['window_title_contains']}'")
        print(f"Named commanders: {', '.join(CONFIG['commanders'])}")
        print(f"Primary commander: {CONFIG['primary_commander']}")
        print("")
        print("USAGE:")
        print(f"1. Write commands to: {self.command_file.absolute()}")
        print("2. Commands are sent automatically when file is saved")
        print("3. File is cleared after each broadcast")
        print("")
        print("EXAMPLES:")
        print("  w w s s")
        print("  1 q q d")
        print("  w a s d")
        print("")
        print("Press Ctrl+C to exit")
        print("-" * 70)

    def find_elite_window(self, target_commander: str = None) -> Optional[int]:
        """Find Elite Dangerous window handle - EXACT copy from autohonk.py"""
        def enum_windows_callback(hwnd, windows):
            try:
                if win32gui.IsWindowVisible(hwnd):
                    title = win32gui.GetWindowText(hwnd)
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    process_handle = win32api.OpenProcess(
                        win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, 
                        False, 
                        pid
                    )
                    process_name = win32process.GetModuleFileNameEx(process_handle, 0).lower()
                    win32api.CloseHandle(process_handle)
                    
                    if 'elitedangerous64' in process_name and CONFIG['window_title_contains'].lower() in title.lower():
                        if target_commander:
                            if target_commander in CONFIG["commanders"]:
                                if target_commander.lower() in title.lower():
                                    windows.append((hwnd, title, target_commander))
                            elif target_commander == CONFIG["primary_commander"]:
                                has_other_commander = any(
                                    cmd.lower() in title.lower() 
                                    for cmd in CONFIG["commanders"]
                                )
                                if not has_other_commander:
                                    windows.append((hwnd, title, target_commander))
                        else:
                            windows.append((hwnd, title, "Unknown"))
                        
            except Exception:
                pass
            return True
        
        try:
            windows = []
            win32gui.EnumWindows(enum_windows_callback, windows)
            
            if windows:
                hwnd, title, commander = windows[0]
                return hwnd
            else:
                return None
                
        except Exception as e:
            logger.error(f"Error finding Elite window: {e}")
            return None

    def find_all_elite_windows(self) -> List[Tuple[int, str, str]]:
        """Find all Elite Dangerous windows."""
        all_windows = []
        
        for commander in CONFIG["commanders"]:
            hwnd = self.find_elite_window(commander)
            if hwnd:
                title = win32gui.GetWindowText(hwnd)
                all_windows.append((hwnd, title, commander))
        
        hwnd = self.find_elite_window(CONFIG["primary_commander"])
        if hwnd:
            title = win32gui.GetWindowText(hwnd)
            all_windows.append((hwnd, title, CONFIG["primary_commander"]))
        
        return all_windows

    def get_virtual_key_code(self, key: str) -> Optional[int]:
        """Get Windows virtual key code - EXACT copy from autohonk.py"""
        special_keys = {
            ' ': win32con.VK_SPACE,
            '\n': win32con.VK_RETURN,
            '\r': win32con.VK_RETURN,
            '\t': win32con.VK_TAB,
        }
        
        if key.lower() in special_keys:
            return special_keys[key.lower()]
        elif len(key) == 1:
            return ord(key.upper())
        else:
            return None

    def send_keys_to_window(self, hwnd: int, command: str, commander: str) -> bool:
        """Send entire command to a window with proper timing."""
        try:
            print(f"🎯 Focusing {commander}...")
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(CONFIG["focus_delay"])
            
            print(f"📤 Sending '{command}' to {commander}...")
            
            for i, char in enumerate(command):
                vk_code = self.get_virtual_key_code(char)
                if vk_code is None:
                    print(f"⚠️  Unknown key: {char}")
                    continue
                
                # Key down
                win32api.keybd_event(vk_code, 0, 0, 0)
                time.sleep(0.05)
                
                # Key up
                win32api.keybd_event(vk_code, 0, win32con.KEYEVENTF_KEYUP, 0)
                
                if i < len(command) - 1:
                    time.sleep(CONFIG["key_send_delay"])
            
            print(f"✅ Sent {len(command)} keys to {commander}")
            return True
            
        except Exception as e:
            print(f"❌ Error sending to {commander}: {e}")
            logger.error(f"Error sending keys to {commander}: {e}")
            return False

    def send_command_to_all_windows(self, command: str):
        """Send command sequence to all Elite Dangerous windows."""
        if not command.strip():
            return
            
        print(f"\n{'=' * 70}")
        print(f"🚀 Broadcasting: '{command}'")
        print(f"{'=' * 70}")
        
        windows = self.find_all_elite_windows()
        
        if not windows:
            print("❌ No Elite Dangerous windows found!")
            return
        
        print(f"\n📡 Found {len(windows)} Elite window(s):")
        for _, title, commander in windows:
            print(f"   • {commander}")
        
        print(f"\n🎮 Sending commands...\n")
        
        success_count = 0
        for i, (hwnd, title, commander) in enumerate(windows):
            if self.send_keys_to_window(hwnd, command, commander):
                success_count += 1
            
            if i < len(windows) - 1:
                print(f"⏳ Waiting {CONFIG['inter_window_delay']}s before next window...\n")
                time.sleep(CONFIG["inter_window_delay"])
        
        print(f"\n{'=' * 70}")
        print(f"🎉 Results: {success_count}/{len(windows)} windows")
        print(f"{'=' * 70}\n")

    def check_for_commands(self):
        """Check if command file has been modified and process commands."""
        try:
            if not self.command_file.exists():
                return
            
            current_modified = self.command_file.stat().st_mtime
            
            if current_modified > self.last_modified:
                self.last_modified = current_modified
                
                # Read command
                command = self.command_file.read_text().strip()
                
                if command:
                    # Send command
                    self.send_command_to_all_windows(command)
                    
                    # Clear file
                    self.command_file.write_text("")
                    self.last_modified = self.command_file.stat().st_mtime
                    
                    print("✅ Ready for next command...\n")
                    
        except Exception as e:
            logger.error(f"Error checking commands: {e}")

    def run(self):
        """Main execution loop."""
        try:
            # Test window detection on startup
            print("🔍 Testing window detection...")
            windows = self.find_all_elite_windows()
            if windows:
                print(f"✅ Found {len(windows)} Elite window(s):")
                for _, title, commander in windows:
                    print(f"   • {commander}")
            else:
                print("⚠️  No Elite windows found - make sure Elite is running!")
            
            print(f"\n🎮 Monitoring {self.command_file.absolute()}")
            print("✅ Ready! Write commands to the file and save.\n")
            
            # Main loop
            while self.running:
                self.check_for_commands()
                time.sleep(CONFIG["check_interval"])
                
        except KeyboardInterrupt:
            print("\n🛑 Shutting down...")
            self.running = False
        
        print("\n👋 Command Relay stopped!")


def main():
    """Main function."""
    print("Starting Elite Dangerous Command Relay...\n")
    
    relay = CommandRelay()
    relay.run()


if __name__ == "__main__":
    main()
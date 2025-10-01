"""
Elite Dangerous Command Relay - Multi-Window Input Broadcasting
Uses PostMessage (WM_KEYDOWN/WM_KEYUP) - THE WORKING METHOD from your window library

Requirements:
- pip install pywin32
"""

import time
import threading
import logging
from typing import List, Tuple, Optional
import sys
import msvcrt
import ctypes

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
    "typing_timeout": 2.0,      # 2 seconds as requested
    "key_press_duration": 0.1,  # Duration to hold key (like your library)
    "key_send_delay": 0.05,     # Delay between keys
    "window_delay": 0.2,        # Delay between windows
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
        self.command_buffer = ""
        self.last_keypress_time = 0
        self.running = True
        self.input_thread = None
        self.timer_thread = None
        self.buffer_lock = threading.Lock()
        self.console_hwnd = None
        
        # Get our console window handle
        self.console_hwnd = self.get_console_window()
        
        print("=" * 70)
        print("Elite Dangerous Command Relay - PostMessage Method")
        print("Using WM_KEYDOWN/WM_KEYUP like your working window library!")
        print("=" * 70)
        print(f"Looking for process: '{CONFIG['process_name']}.exe'")
        print(f"Window title must contain: '{CONFIG['window_title_contains']}'")
        print(f"Named commanders: {', '.join(CONFIG['commanders'])}")
        print(f"Primary commander: {CONFIG['primary_commander']}")
        print("")
        print("INSTRUCTIONS:")
        print("1. Focus this console window")
        print("2. Type your command (e.g., '1qq' or 'swsw')")
        print(f"3. Wait {CONFIG['typing_timeout']} seconds - command broadcasts to ALL Elite windows")
        print("4. Press Ctrl+C to exit")
        print("-" * 70)

    def get_console_window(self) -> Optional[int]:
        """Get the console window handle using kernel32."""
        try:
            kernel32 = ctypes.windll.kernel32
            hwnd = kernel32.GetConsoleWindow()
            return hwnd if hwnd else None
        except Exception as e:
            logger.error(f"Error getting console window handle: {e}")
            return None

    def find_elite_window(self, target_commander: str = None) -> Optional[int]:
        """Find Elite Dangerous window handle."""
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
                logger.info(f"Found Elite window for {target_commander or 'testing'}: '{title}'")
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
        """Get Windows virtual key code."""
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

    def press_key(self, hwnd: int, key_code: int, duration: float = None):
        """
        Press a key using PostMessage - EXACT method from your working library!
        def press(self, key, duration=.1):
            self.key_down(key)
            time.sleep(duration)
            self.key_up(key)
        """
        if duration is None:
            duration = CONFIG["key_press_duration"]
        
        # Key down - PostMessage with WM_KEYDOWN
        win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, key_code, 0)
        time.sleep(duration)
        # Key up - PostMessage with WM_KEYUP
        win32api.PostMessage(hwnd, win32con.WM_KEYUP, key_code, 0)

    def send_keys_to_window(self, hwnd: int, command: str, commander: str) -> bool:
        """Send entire command to a window using PostMessage."""
        try:
            print(f"🎯 Sending '{command}' to {commander} using PostMessage...")
            
            # Send each character
            for char in command:
                vk_code = self.get_virtual_key_code(char)
                if vk_code is None:
                    print(f"⚠️ Unknown key: {char}")
                    continue
                
                # Use PostMessage method - EXACT copy from your library
                self.press_key(hwnd, vk_code)
                
                # Delay between keys
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
            
        print(f"\n🚀 Broadcasting command: '{command}' (length: {len(command)})")
        
        # Find all Elite windows
        windows = self.find_all_elite_windows()
        
        if not windows:
            print("⚠️  No Elite Dangerous windows found!")
            return
        
        print(f"📡 Found {len(windows)} Elite window(s):")
        for _, title, commander in windows:
            print(f"   • {commander}: {title}")
        
        print("\n🎮 Sending commands with PostMessage...")
        
        # Send to each window
        success_count = 0
        for hwnd, title, commander in windows:
            if self.send_keys_to_window(hwnd, command, commander):
                success_count += 1
            time.sleep(CONFIG["window_delay"])
        
        print(f"\n🎉 Successfully sent to {success_count}/{len(windows)} windows")
        
        # Focus back to console
        if self.console_hwnd:
            try:
                win32gui.SetForegroundWindow(self.console_hwnd)
                time.sleep(0.1)
                print("🔄 Console refocused - ready for next command")
            except:
                pass
        
        print("-" * 70)

    def input_monitor(self):
        """Monitor for keyboard input in the console."""
        print("🎧 Input monitor started. Type your commands...")
        
        while self.running:
            try:
                if msvcrt.kbhit():
                    char = msvcrt.getch().decode('utf-8', errors='ignore')
                    
                    if ord(char) == 3:  # Ctrl+C
                        print("\n🛑 Ctrl+C detected - shutting down...")
                        self.running = False
                        break
                    elif ord(char) == 8:  # Backspace
                        with self.buffer_lock:
                            if self.command_buffer:
                                self.command_buffer = self.command_buffer[:-1]
                                print(f"\rCommand: '{self.command_buffer}'", end=" " * 10, flush=True)
                                self.last_keypress_time = time.time()
                        continue
                    elif ord(char) == 13:  # Enter
                        char = '\n'
                    
                    with self.buffer_lock:
                        self.command_buffer += char
                        self.last_keypress_time = time.time()
                        print(f"\rCommand: '{self.command_buffer}'", end="", flush=True)
                
                time.sleep(0.01)
                
            except Exception as e:
                logger.error(f"Error in input monitor: {e}")
                time.sleep(0.1)

    def timer_monitor(self):
        """Monitor for typing timeout and send commands when ready."""
        while self.running:
            try:
                with self.buffer_lock:
                    if (self.command_buffer and 
                        self.last_keypress_time > 0 and 
                        time.time() - self.last_keypress_time >= CONFIG["typing_timeout"]):
                        
                        command_to_send = self.command_buffer
                        self.command_buffer = ""
                        self.last_keypress_time = 0
                        
                        print()  # New line
                        self.send_command_to_all_windows(command_to_send)
                
                time.sleep(0.1)
                
            except Exception as e:
                logger.error(f"Error in timer monitor: {e}")
                time.sleep(0.1)

    def run(self):
        """Main execution logic."""
        try:
            # Test window detection
            print("🔍 Testing window detection...")
            windows = self.find_all_elite_windows()
            if windows:
                print(f"✅ Found {len(windows)} Elite window(s):")
                for _, title, commander in windows:
                    print(f"   • {commander}: {title}")
            else:
                print("⚠️  No Elite windows found - make sure Elite is running!")
            
            print("\n🎮 Ready for input! Type commands and wait 2 seconds...")
            
            # Start monitoring threads
            self.input_thread = threading.Thread(target=self.input_monitor, daemon=True)
            self.input_thread.start()
            
            self.timer_thread = threading.Thread(target=self.timer_monitor, daemon=True)
            self.timer_thread.start()
            
            # Main loop
            while self.running:
                time.sleep(0.1)
                
        except KeyboardInterrupt:
            print("\n🛑 Shutting down...")
            self.running = False
        
        print("\n👋 Command Relay stopped!")


def main():
    """Main function."""
    print("Starting Elite Dangerous Command Relay...")
    print("Using PostMessage (WM_KEYDOWN/WM_KEYUP) method\n")
    
    relay = CommandRelay()
    relay.run()


if __name__ == "__main__":
    main()
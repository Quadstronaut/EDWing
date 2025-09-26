"""
Elite Dangerous Command Relay - Multi-Window Input Broadcasting
Captures keyboard input and relays commands to all Elite Dangerous windows after typing stops.

Requirements:
- pip install pywin32
"""

import time
import threading
import logging
from typing import List, Tuple, Dict, Optional
import queue
import sys
import msvcrt
import ctypes
from ctypes import wintypes

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
    "typing_timeout": 1.0,  # Wait 1 second after last keypress before sending
    "key_send_delay": 0.05,  # Delay between each key send
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
        print("Elite Dangerous Command Relay - Multi-Window Broadcasting")
        print("=" * 70)
        print(f"Looking for process: '{CONFIG['process_name']}.exe'")
        print(f"Window title must contain: '{CONFIG['window_title_contains']}'")
        print(f"Named commanders: {', '.join(CONFIG['commanders'])}")
        print(f"Primary commander (no name in title): {CONFIG['primary_commander']}")
        print(f"Typing timeout: {CONFIG['typing_timeout']} seconds")
        print("")
        print("INSTRUCTIONS:")
        print("1. Focus this console window")
        print("2. Type your command sequence (e.g., '1 q q d [space]')")
        print("3. Wait 1 second - command will be sent to all Elite windows")
        print("4. Press Ctrl+C to exit")
        print("-" * 70)
        print("Ready for input...")
        print("")

    def get_console_window(self) -> Optional[int]:
        """Get the console window handle using kernel32."""
        try:
            # Use kernel32.GetConsoleWindow() instead of win32gui
            kernel32 = ctypes.windll.kernel32
            hwnd = kernel32.GetConsoleWindow()
            if hwnd:
                logger.debug(f"Console window handle: {hwnd}")
                return hwnd
            else:
                logger.warning("Could not get console window handle")
                return None
        except Exception as e:
            logger.error(f"Error getting console window handle: {e}")
            return None

    def find_all_elite_windows(self) -> List[Tuple[int, str, str]]:
        """Find all Elite Dangerous windows."""
        
        def enum_windows_callback(hwnd, windows):
            try:
                if win32gui.IsWindowVisible(hwnd):
                    title = win32gui.GetWindowText(hwnd)
                    
                    # Get the process info for this window
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    try:
                        process_handle = win32api.OpenProcess(
                            win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, 
                            False, 
                            pid
                        )
                        process_name = win32process.GetModuleFileNameEx(process_handle, 0).lower()
                        win32api.CloseHandle(process_handle)
                        
                        # Check if this is the Elite Dangerous process
                        if CONFIG["process_name"].lower() in process_name:
                            # Check if title contains our base window title
                            if title.strip() and CONFIG["window_title_contains"].lower() in title.lower():
                                
                                # First, check for named commanders
                                for commander in CONFIG["commanders"]:
                                    if commander.lower() in title.lower():
                                        windows.append((hwnd, title, commander))
                                        logger.debug(f"Found Elite window for {commander}: '{title}'")
                                        return True
                                
                                # If no named commanders found, check if this could be the primary commander
                                primary_commander = CONFIG["primary_commander"]
                                # Check if title contains any OTHER commander names - if not, it's likely the primary
                                has_other_commander = any(cmd.lower() in title.lower() for cmd in CONFIG["commanders"])
                                if not has_other_commander:
                                    # This window doesn't contain any named commander, assume it's primary
                                    windows.append((hwnd, title, primary_commander))
                                    logger.debug(f"Found Elite window for {primary_commander} (no name in title): '{title}'")
                                        
                    except Exception as e:
                        logger.debug(f"Could not get process info for PID {pid}: {e}")
                        
            except Exception as e:
                logger.debug(f"Error processing window {hwnd}: {e}")
                pass
            return True

        try:
            windows = []
            win32gui.EnumWindows(enum_windows_callback, windows)
            return windows
        except Exception as e:
            logger.error("Error finding Elite windows: %s", e)
            return []

    def get_virtual_key_code(self, char: str) -> Optional[int]:
        """Get Windows virtual key code for a character."""
        if char == ' ':
            return win32con.VK_SPACE
        elif char == '\n' or char == '\r':
            return win32con.VK_RETURN
        elif char == '\t':
            return win32con.VK_TAB
        elif len(char) == 1 and char.isalnum():
            return ord(char.upper())
        else:
            # For other characters, try to get the virtual key
            try:
                return ord(char.upper())
            except:
                return None

    def send_key_to_window(self, hwnd: int, char: str, commander: str) -> bool:
        """Send a single key to a specific window."""
        try:
            # Handle special characters
            if char == ' ':
                # Send space key
                win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_SPACE, 0)
                time.sleep(0.01)
                win32gui.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_SPACE, 0)
                logger.debug(f"Sent SPACE to {commander}")
                return True
            elif char == '\n' or char == '\r':
                # Send enter key
                win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
                time.sleep(0.01)
                win32gui.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
                logger.debug(f"Sent ENTER to {commander}")
                return True
            elif len(char) == 1:
                # Send character using WM_CHAR (most reliable for text input)
                win32gui.SendMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
                logger.debug(f"Sent '{char}' to {commander}")
                return True
            else:
                logger.warning(f"Unsupported character: '{char}'")
                return False
                
        except Exception as e:
            logger.error(f"Error sending key '{char}' to {commander}: {e}")
            return False

    def send_command_to_windows(self, command: str):
        """Send command sequence to all Elite Dangerous windows."""
        if not command.strip():
            return
            
        print(f"🎯 Broadcasting command: '{command}'")
        
        # Find all Elite windows
        windows = self.find_all_elite_windows()
        
        if not windows:
            print("⚠️  No Elite Dangerous windows found!")
            return
        
        print(f"📡 Found {len(windows)} Elite Dangerous window(s)")
        for _, title, commander in windows:
            print(f"   • {commander}: {title}")
        
        # Send command to each window
        success_count = 0
        for hwnd, title, commander in windows:
            try:
                print(f"📤 Sending to {commander}...", end=" ")
                
                # Try to focus window, but don't fail if it doesn't work
                try:
                    win32gui.SetForegroundWindow(hwnd)
                    win32gui.SetActiveWindow(hwnd)
                    time.sleep(0.05)  # Brief pause
                except Exception as focus_error:
                    logger.debug(f"Could not focus {commander}: {focus_error}")
                    # Continue anyway - SendMessage should still work
                
                # Send each character in the command
                char_success = 0
                for char in command:
                    if self.send_key_to_window(hwnd, char, commander):
                        char_success += 1
                    time.sleep(CONFIG["key_send_delay"])
                
                print(f"✅ ({char_success}/{len(command)} chars)")
                success_count += 1
                
                # Brief pause between windows
                time.sleep(0.05)
                
            except Exception as e:
                print(f"❌ Error: {e}")
                logger.error(f"Error sending command to {commander}: {e}")
        
        print(f"🎉 Command sent to {success_count}/{len(windows)} windows")
        
        # Focus back to console window
        if self.console_hwnd:
            try:
                win32gui.SetForegroundWindow(self.console_hwnd)
                win32gui.SetActiveWindow(self.console_hwnd)
                print("🔄 Console window refocused")
            except Exception as e:
                logger.debug(f"Could not refocus console window: {e}")
        
        print("-" * 50)
        print("Ready for next command...")

    def input_monitor(self):
        """Monitor for keyboard input in the console."""
        print("🎧 Input monitor started. Type your commands...")
        
        while self.running:
            try:
                if msvcrt.kbhit():
                    char = msvcrt.getch().decode('utf-8', errors='ignore')
                    
                    # Handle special keys
                    if ord(char) == 3:  # Ctrl+C
                        print("\n🛑 Ctrl+C detected - shutting down...")
                        self.running = False
                        break
                    elif ord(char) == 8:  # Backspace
                        with self.buffer_lock:
                            if self.command_buffer:
                                self.command_buffer = self.command_buffer[:-1]
                                print(f"\rCurrent command: '{self.command_buffer}'", end=" " * 10, flush=True)
                                self.last_keypress_time = time.time()
                        continue
                    elif ord(char) == 13:  # Enter - treat as part of command
                        char = '\n'
                    
                    # Add character to buffer
                    with self.buffer_lock:
                        self.command_buffer += char
                        self.last_keypress_time = time.time()
                        print(f"\rCurrent command: '{self.command_buffer}'", end="", flush=True)
                
                time.sleep(0.01)  # Small delay to prevent excessive CPU usage
                
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
                        
                        # Time to send the command
                        command_to_send = self.command_buffer
                        self.command_buffer = ""  # Clear buffer
                        self.last_keypress_time = 0
                        
                        print()  # New line after current command display
                        self.send_command_to_windows(command_to_send)
                
                time.sleep(0.1)  # Check every 100ms
                
            except Exception as e:
                logger.error(f"Error in timer monitor: {e}")
                time.sleep(0.1)

    def run(self):
        """Main execution logic."""
        try:
            # Start input monitoring thread
            self.input_thread = threading.Thread(target=self.input_monitor, daemon=True)
            self.input_thread.start()
            
            # Start timer monitoring thread
            self.timer_thread = threading.Thread(target=self.timer_monitor, daemon=True)
            self.timer_thread.start()
            
            # Main loop - just wait for shutdown
            while self.running:
                time.sleep(0.1)
                
        except KeyboardInterrupt:
            print("\n🛑 Keyboard interrupt - shutting down...")
            self.running = False
        
        # Wait for threads to finish
        if self.input_thread:
            self.input_thread.join(timeout=1.0)
        if self.timer_thread:
            self.timer_thread.join(timeout=1.0)
        
        print("\n👋 Command Relay stopped. Goodbye!")


def main():
    """Main function to start the Command Relay."""
    print("Starting Elite Dangerous Command Relay...")
    
    relay = CommandRelay()
    relay.run()
    
    logger.info("Command Relay process finished")


if __name__ == "__main__":
    main()
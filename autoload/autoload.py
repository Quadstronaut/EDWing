"""
Elite Dangerous AutoLoad - Multi-Commander Script
Double clicks in the center of each client window to skip the opening cutscene.

Requirements:
- pip install pywin32 pycaw
"""

import time
import logging
from typing import Optional, Set, List, Tuple
import ctypes
from ctypes import wintypes

# Third party API imports for window handling
import win32api
import win32con
import win32gui
import win32process

# Audio detection imports
from pycaw.pycaw import AudioUtilities

# Configuration
CONFIG = {
    "window_title_contains": "Elite - Dangerous (CLIENT)",
    "process_name": "elitedangerous64",
    # For multibox setups define your alts here
    "commanders": ["Bistronaut", "Tristronaut", "Quadstronaut"],
    "primary_commander": "Duvrazh",  # Primary commander default without name in window title
}

# Logging setup
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.StreamHandler(), logging.FileHandler("elite_autoload.log")],
)
logger = logging.getLogger(__name__)


class MultiCommanderAutoLoad:
    def __init__(self):
        self.processed_commanders: Set[str] = set()
        self.all_commanders = CONFIG["commanders"] + [CONFIG["primary_commander"]]
        self.total_commanders = len(self.all_commanders)
        
        print("=" * 60)
        print("Elite Dangerous Multi-Commander AutoLoad - Double Click Skip")
        print("=" * 60)
        print(f"Looking for process: '{CONFIG['process_name']}.exe'")
        print(f"Window title must contain: '{CONFIG['window_title_contains']}'")
        print(f"Named commanders: {', '.join(CONFIG['commanders'])}")
        print(f"Primary commander (no name in title): {CONFIG['primary_commander']}")
        print("Will wait for window title AND audio output before double-clicking...")
        print("-" * 60)

    def find_unprocessed_elite_windows(self) -> List[Tuple[int, str, str]]:
        """Find Elite Dangerous windows for commanders that haven't been processed yet."""
        
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
                                    if commander not in self.processed_commanders and commander.lower() in title.lower():
                                        windows.append((hwnd, title, commander))
                                        logger.info(f"Found Elite window for {commander}: '{title}' from process: {process_name}")
                                        return True  # Found a named commander, continue
                                
                                # If no named commanders found and primary commander not processed,
                                # check if this could be the primary commander's window
                                primary_commander = CONFIG["primary_commander"]
                                if primary_commander not in self.processed_commanders:
                                    # Check if title contains any OTHER commander names - if not, it's likely the primary
                                    has_other_commander = any(cmd.lower() in title.lower() for cmd in CONFIG["commanders"])
                                    if not has_other_commander:
                                        # This window doesn't contain any named commander, assume it's primary
                                        windows.append((hwnd, title, primary_commander))
                                        logger.info(f"Found Elite window for {primary_commander} (no name in title): '{title}' from process: {process_name}")
                                        
                            elif title.strip():
                                logger.debug(f"Elite process found but title doesn't match pattern: '{title}' from process: {process_name}")
                            else:
                                logger.debug(f"Elite process found but title is blank from process: {process_name}")
                                
                    except Exception as e:
                        logger.debug(f"Could not get process info for PID {pid}: {e}")
                        
            except Exception as e:
                logger.debug(f"Error processing window {hwnd}: {e}")
                pass  # Ignore windows we can't access
            return True

        try:
            windows = []
            win32gui.EnumWindows(enum_windows_callback, windows)
            return windows
        except Exception as e:
            logger.error("Error finding Elite windows: %s", e)
            return []

    def is_elite_producing_audio(self) -> bool:
        """Check if any Elite Dangerous process is currently producing audio."""
        try:
            # Get all audio sessions
            sessions = AudioUtilities.GetAllSessions()
            
            for session in sessions:
                if session.Process:
                    process_name = session.Process.name().lower()
                    if CONFIG["process_name"].lower() in process_name:
                        # Check if the session is active and has volume > 0
                        volume = session.SimpleAudioVolume
                        if volume and not volume.GetMute():
                            # Session exists and is not muted - Elite is producing audio
                            logger.debug(f"Elite audio session found: {process_name}, muted: {volume.GetMute()}")
                            return True
                        elif volume:
                            logger.debug(f"Elite audio session found but muted: {process_name}")
                        else:
                            logger.debug(f"Elite process found but no audio volume interface: {process_name}")
            
            return False
            
        except Exception as e:
            logger.debug(f"Error checking audio sessions: {e}")
            return False

    def wait_for_next_commander_window(self) -> Optional[Tuple[int, str, str]]:
        """Wait for the next unprocessed commander's window to appear."""
        remaining_commanders = [cmd for cmd in self.all_commanders if cmd not in self.processed_commanders]
        
        if not remaining_commanders:
            return None
            
        print(f"\nWaiting for windows from remaining commanders: {', '.join(remaining_commanders)}")
        
        check_count = 0
        while remaining_commanders:
            windows = self.find_unprocessed_elite_windows()
            if windows:
                # Prioritize named commanders over primary commander
                named_windows = [(hwnd, title, cmd) for hwnd, title, cmd in windows if cmd in CONFIG["commanders"]]
                primary_windows = [(hwnd, title, cmd) for hwnd, title, cmd in windows if cmd == CONFIG["primary_commander"]]
                
                # Process named commanders first, then primary
                if named_windows:
                    hwnd, title, commander = named_windows[0]
                elif primary_windows and len(self.processed_commanders) == len(CONFIG["commanders"]):
                    # Only process primary commander after all named commanders are done
                    hwnd, title, commander = primary_windows[0]
                else:
                    # Wait for named commanders first
                    check_count += 1
                    if check_count % 10 == 0:  # Every 5 seconds
                        print(f"Still waiting for named commander windows... (checked {check_count} times)")
                    time.sleep(0.5)
                    continue
                
                print(f"Found window for commander {commander}: '{title}'")
                logger.info(f"Elite Dangerous window detected for {commander}")
                return hwnd, title, commander
            
            check_count += 1
            if check_count % 10 == 0:  # Every 5 seconds
                print(f"Still waiting for commander windows... (checked {check_count} times)")
                
            time.sleep(0.5)  # Check every 500ms
            
            # Update remaining commanders list in case some were processed
            remaining_commanders = [cmd for cmd in self.all_commanders if cmd not in self.processed_commanders]
            
        return None

    def wait_for_audio(self, commander: str) -> None:
        """Wait until Elite Dangerous starts producing audio."""
        print(f"Window found for {commander}! Now waiting for Elite Dangerous to start producing audio...")
        
        audio_check_count = 0
        while True:
            if self.is_elite_producing_audio():
                print(f"Elite Dangerous is now producing audio for {commander}!")
                logger.info(f"Elite Dangerous audio detected for {commander}")
                return
            
            audio_check_count += 1
            if audio_check_count % 10 == 0:  # Every 5 seconds
                print(f"Still waiting for audio for {commander}... (checked {audio_check_count} times)")
                
            time.sleep(0.5)  # Check every 500ms

    def get_window_center(self, hwnd: int) -> Tuple[int, int]:
        """Get the center coordinates of a window."""
        try:
            rect = win32gui.GetWindowRect(hwnd)
            left, top, right, bottom = rect
            center_x = (left + right) // 2
            center_y = (top + bottom) // 2
            
            logger.debug(f"Window rect: {rect}, center: ({center_x}, {center_y})")
            return center_x, center_y
        except Exception as e:
            logger.error(f"Error getting window center: {e}")
            raise

    def double_click_window_center(self, elite_hwnd: int, commander: str):
        """Double click in the center of the Elite Dangerous window."""
        try:
            print(f"Focusing window and double-clicking center for {commander}...")
            
            # Get window dimensions and center point
            center_x, center_y = self.get_window_center(elite_hwnd)
            
            # Get window client area coordinates (relative to window)
            rect = win32gui.GetWindowRect(elite_hwnd)
            client_rect = win32gui.GetClientRect(elite_hwnd)
            
            # Convert to client coordinates
            client_point = win32gui.ScreenToClient(elite_hwnd, (center_x, center_y))
            client_x, client_y = client_point
            
            print(f"Window center - Screen coords: ({center_x}, {center_y}), Client coords: ({client_x}, {client_y})")
            
            # Ensure window is properly focused and brought to front
            win32gui.SetForegroundWindow(elite_hwnd)
            win32gui.SetActiveWindow(elite_hwnd)
            win32gui.BringWindowToTop(elite_hwnd)
            time.sleep(0.3)  # Give time for focus to establish
            
            # Verify the window is focused
            focused_hwnd = win32gui.GetForegroundWindow()
            if focused_hwnd != elite_hwnd:
                logger.warning(f"Window focus verification failed for {commander}")
                win32gui.ShowWindow(elite_hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(elite_hwnd)
                time.sleep(0.3)
            
            # Method 1: Send click messages directly to the window (most reliable)
            def send_click_via_messages():
                # Convert client coordinates to lParam format
                lparam = win32api.MAKELONG(client_x, client_y)
                
                # Send double-click sequence
                print(f"Sending double-click to client coordinates ({client_x}, {client_y})...")
                
                # First click
                win32gui.SendMessage(elite_hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
                time.sleep(0.01)  # Small delay
                win32gui.SendMessage(elite_hwnd, win32con.WM_LBUTTONUP, 0, lparam)
                
                time.sleep(0.05)  # Short delay between clicks
                
                # Second click (double-click)
                win32gui.SendMessage(elite_hwnd, win32con.WM_LBUTTONDBLCLK, win32con.MK_LBUTTON, lparam)
                time.sleep(0.01)
                win32gui.SendMessage(elite_hwnd, win32con.WM_LBUTTONUP, 0, lparam)
                
                logger.info(f"Double-click messages sent to {commander} at ({client_x}, {client_y})")
                return True
            
            # Method 2: Use SetCursorPos + mouse_event (fallback)
            def send_click_via_mouse_event():
                # Save current cursor position
                current_pos = win32gui.GetCursorPos()
                
                # Move cursor to window center
                win32api.SetCursorPos((center_x, center_y))
                time.sleep(0.1)
                
                # Perform double-click using mouse_event
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                time.sleep(0.05)
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                
                # Restore cursor position
                win32api.SetCursorPos(current_pos)
                
                logger.info(f"Double-click via mouse_event sent to {commander} at ({center_x}, {center_y})")
                return True
            
            # Try Method 1 first (window messages)
            try:
                if send_click_via_messages():
                    print(f"✅ Double-click sent successfully to {commander} via window messages")
                    logger.info(f"Double-click completed for {commander} using SendMessage")
                    return
            except Exception as e:
                logger.warning(f"SendMessage method failed for {commander}: {e}")
            
            # Fallback to Method 2 (mouse_event)
            try:
                print(f"Trying fallback method (mouse_event) for {commander}...")
                if send_click_via_mouse_event():
                    print(f"✅ Double-click sent successfully to {commander} via mouse_event")
                    logger.info(f"Double-click completed for {commander} using mouse_event")
                    return
            except Exception as e:
                logger.warning(f"mouse_event method failed for {commander}: {e}")
            
            # If both methods failed
            print(f"❌ Failed to send double-click to {commander} - all methods failed")
            logger.error(f"All double-click methods failed for {commander}")
                
        except Exception as e:
            logger.error(f"Error in double_click_window_center for {commander}: {e}")
            print(f"Critical error sending double-click to {commander}: {e}")

    def process_commander(self, hwnd: int, title: str, commander: str):
        """Process a single commander's window."""
        print(f"\n{'='*60}")
        print(f"Processing Commander: {commander}")
        print(f"Window Title: {title}")
        print(f"{'='*60}")
        
        # Wait for Elite to start producing audio
        self.wait_for_audio(commander)

        # Brief pause to ensure everything is ready
        print(f"Audio detected for {commander}! Waiting 1 second before double-clicking...")
        time.sleep(1)
        
        # Double-click the center of the window to skip cutscene
        self.double_click_window_center(hwnd, commander)
        
        # Mark this commander as processed
        self.processed_commanders.add(commander)
        remaining = self.total_commanders - len(self.processed_commanders)
        
        print(f"✅ Commander {commander} processed successfully!")
        print(f"Remaining commanders: {remaining}")
        logger.info(f"Commander {commander} processed. {remaining} remaining.")

    def run(self):
        """Main execution logic."""
        print("Starting Multi-Commander AutoLoad process...")
        
        while len(self.processed_commanders) < self.total_commanders:
            # Wait for next commander's window
            result = self.wait_for_next_commander_window()
            
            if result is None:
                print("No more commanders to process.")
                break
                
            hwnd, title, commander = result
            
            # Process this commander
            self.process_commander(hwnd, title, commander)
            
            if len(self.processed_commanders) < self.total_commanders:
                print(f"\n🔄 Looking for next commander...")
                time.sleep(1)  # Brief pause before looking for next commander
        
        print(f"\n{'='*60}")
        print("🎉 All commanders processed!")
        print(f"Processed: {', '.join(sorted(self.processed_commanders))}")
        print(f"{'='*60}")


def main():
    """Main function to start the Multi-Commander AutoLoad process."""
    print("Starting Elite Dangerous Multi-Commander AutoLoad...")
    
    autoload = MultiCommanderAutoLoad()
    autoload.run()
    
    print("Multi-Commander AutoLoad complete. Exiting.")
    logger.info("Multi-Commander AutoLoad process finished")


if __name__ == "__main__":
    main()
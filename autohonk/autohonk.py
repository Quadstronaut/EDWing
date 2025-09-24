"""
Elite Dangerous AutoHonk - Standalone Script
Monitors Elite Dangerous journal files and auto-presses your Primary Fire key when jumping to new systems.
Modified to hold key until FSSDiscoveryScan event instead of fixed duration.

Requirements:
- pip install pywin32 watchdog
"""

import os
import json
import time
import threading
from pathlib import Path
from typing import Optional
import xml.etree.ElementTree as ET
import glob
import logging
from datetime import datetime

# Windows API imports
import win32api
import win32con
import win32gui
import win32process

# File monitoring
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# Configuration
CONFIG = {
    'window_title_contains': 'Elite - Dangerous (CLIENT)',  # Part of Elite window title to look for
    'delay_after_jump': 2.0,  # Wait 2 seconds after jump before honking
    'max_honk_duration': 7.0,  # Maximum time to honk (safety fallback)
    'key_press_interval': 0.1,  # How often to send key presses (for continuous hold)
    'auto_detect_primary_fire': True,  # Auto-detect from bindings
    'manual_key_override': None,  # Set to specific key if needed (e.g., 'numpad_add')
    'journal_folder': Path.home() / 'Saved Games' / 'Frontier Developments' / 'Elite Dangerous'
}

# Logging setup
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('elite_autohonk.log')
    ]
)
logger = logging.getLogger(__name__)

class AutoHonk:
    def __init__(self):
        self.current_system = None
        self.primary_fire_key = None
        self.elite_hwnd = None
        self.running = True
        self.honking_active = False
        self.honk_thread = None
        self.honk_lock = threading.Lock()
        
        # Detect primary fire key on startup
        self.detect_primary_fire_key()
        
        print("=" * 60)
        print("Elite Dangerous AutoHonk - Standalone (FSS Discovery Mode)")
        print("=" * 60)
        print(f"Monitoring journal folder: {CONFIG['journal_folder']}")
        print(f"Looking for window containing: '{CONFIG['window_title_contains']}'")
        print(f"Detected primary fire key: {self.primary_fire_key or 'Not detected'}")
        print(f"Max honk duration (safety): {CONFIG['max_honk_duration']} seconds")
        print("Will honk until FSSDiscoveryScan event is detected...")
        print("Waiting for FSD jumps...")
        print("-" * 60)
    
    def detect_primary_fire_key(self):
        """Detect Primary Fire key from Elite Dangerous bindings."""
        try:
            bindings_dir = Path(os.environ.get('LOCALAPPDATA')) / "Frontier Developments" / "Elite Dangerous" / "Options" / "Bindings"
            
            if not bindings_dir.exists():
                logger.warning(f"Bindings directory not found: {bindings_dir}")
                return
            
            binds_files = list(bindings_dir.glob("*.binds"))
            if not binds_files:
                logger.warning("No .binds files found")
                return
            
            # Use the most recent bindings file
            latest_bindings = max(binds_files, key=lambda x: x.stat().st_mtime)
            logger.info(f"Reading bindings from: {latest_bindings}")
            
            tree = ET.parse(latest_bindings)
            root = tree.getroot()
            
            primary_fire = root.find(".//PrimaryFire")
            if primary_fire is not None:
                primary = primary_fire.find("Primary")
                if primary is not None:
                    device = primary.get("Device")
                    key_attr = primary.get("Key")
                    
                    if device == "Keyboard" and key_attr:
                        self.primary_fire_key = self.convert_elite_key_name(key_attr)
                        logger.info(f"Detected Primary Fire key: {key_attr} -> {self.primary_fire_key}")
                        return
            
            logger.warning("PrimaryFire binding not found in bindings file")
            
        except Exception as e:
            logger.error(f"Error detecting primary fire key: {e}")
    
    def convert_elite_key_name(self, elite_key: str) -> str:
        """Convert Elite Dangerous key name to Windows virtual key format."""
        if elite_key.startswith("Key_"):
            elite_key = elite_key[4:]
        
        key_mapping = {
            'Numpad_Add': 'numpad_add',
            'Numpad_Subtract': 'numpad_subtract',
            'Numpad_Multiply': 'numpad_multiply',
            'Numpad_Divide': 'numpad_divide',
            'Space': 'space',
            'Enter': 'enter',
            'Tab': 'tab',
            'F1': 'f1', 'F2': 'f2', 'F3': 'f3', 'F4': 'f4',
            'F5': 'f5', 'F6': 'f6', 'F7': 'f7', 'F8': 'f8',
            'F9': 'f9', 'F10': 'f10', 'F11': 'f11', 'F12': 'f12',
        }
        
        return key_mapping.get(elite_key, elite_key.lower())
    
    def find_elite_window(self) -> Optional[int]:
        """Find Elite Dangerous window handle by process name and window title."""
        def enum_windows_callback(hwnd, windows):
            try:
                if win32gui.IsWindowVisible(hwnd):
                    # Get window title
                    title = win32gui.GetWindowText(hwnd)
                    
                    # Get process ID and name
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    process_handle = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
                    process_name = win32process.GetModuleFileNameEx(process_handle, 0).lower()
                    win32api.CloseHandle(process_handle)
                    
                    # Check if it's Elite Dangerous process with matching window title
                    if 'elitedangerous64' in process_name and CONFIG['window_title_contains'].lower() in title.lower():
                        windows.append((hwnd, title, process_name))
                        
            except Exception:
                # Skip windows we can't access
                pass
            return True
        
        try:
            windows = []
            win32gui.EnumWindows(enum_windows_callback, windows)
            
            if windows:
                hwnd, title, process = windows[0]
                logger.info(f"Found Elite window: '{title}' (PID: {process})")
                return hwnd
            else:
                logger.warning(f"Elite Dangerous window not found (looking for process EliteDangerous64 with title containing '{CONFIG['window_title_contains']}')")
                return None
                
        except Exception as e:
            logger.error(f"Error finding Elite window: {e}")
            return None
    
    def get_virtual_key_code(self, key: str) -> Optional[int]:
        """Get Windows virtual key code for the key."""
        special_keys = {
            'numpad_add': win32con.VK_ADD,
            'numpad_subtract': win32con.VK_SUBTRACT,
            'numpad_multiply': win32con.VK_MULTIPLY,
            'numpad_divide': win32con.VK_DIVIDE,
            'space': win32con.VK_SPACE,
            'enter': win32con.VK_RETURN,
            'tab': win32con.VK_TAB,
            'f1': win32con.VK_F1, 'f2': win32con.VK_F2, 'f3': win32con.VK_F3, 'f4': win32con.VK_F4,
            'f5': win32con.VK_F5, 'f6': win32con.VK_F6, 'f7': win32con.VK_F7, 'f8': win32con.VK_F8,
            'f9': win32con.VK_F9, 'f10': win32con.VK_F10, 'f11': win32con.VK_F11, 'f12': win32con.VK_F12,
        }
        
        if key.lower() in special_keys:
            return special_keys[key.lower()]
        elif len(key) == 1:
            return ord(key.upper())
        else:
            return None
    
    def continuous_keypress(self, key: str):
        """Send continuous keypress until stopped."""
        try:
            # Find Elite window
            elite_hwnd = self.find_elite_window()
            if not elite_hwnd:
                print("❌ Elite Dangerous window not found - cannot send keypress")
                return
            
            # Get virtual key code
            vk_code = self.get_virtual_key_code(key)
            if vk_code is None:
                print(f"❌ Unknown key: {key}")
                return
            
            print(f"🎯 Starting continuous keypress '{key}' to Elite Dangerous...")
            print("   Will continue until FSSDiscoveryScan event or timeout...")
            
            # Bring window to foreground
            win32gui.SetForegroundWindow(elite_hwnd)
            time.sleep(0.2)  # Brief delay to ensure focus
            
            start_time = time.time()
            key_down = False
            
            while self.honking_active and self.running:
                # Send key down if not already down
                if not key_down:
                    win32api.keybd_event(vk_code, 0, 0, 0)
                    print(f"⬇️ Key DOWN: {key}")
                    key_down = True
                
                # Check for timeout (safety mechanism)
                elapsed = time.time() - start_time
                if elapsed >= CONFIG['max_honk_duration']:
                    print(f"⏰ Timeout reached ({CONFIG['max_honk_duration']}s) - stopping honk")
                    break
                
                # Short sleep to prevent excessive CPU usage
                time.sleep(CONFIG['key_press_interval'])
            
            # Always send key up when done
            if key_down:
                win32api.keybd_event(vk_code, 0, win32con.KEYEVENTF_KEYUP, 0)
                print(f"⬆️ Key UP: {key}")
            
            elapsed = time.time() - start_time
            print(f"✅ Honking complete! Duration: {elapsed:.1f} seconds")
            
        except Exception as e:
            print(f"❌ Error during continuous keypress: {e}")
            logger.error(f"Continuous keypress error: {e}")
    
    def start_honking(self, key: str):
        """Start the honking process in a separate thread."""
        with self.honk_lock:
            if self.honking_active:
                print("⚠️ Already honking - ignoring duplicate request")
                return
            
            self.honking_active = True
            self.honk_thread = threading.Thread(target=self.continuous_keypress, args=(key,), daemon=True)
            self.honk_thread.start()
    
    def stop_honking(self):
        """Stop the honking process."""
        with self.honk_lock:
            if not self.honking_active:
                return
            
            print("🛑 FSSDiscoveryScan detected - stopping honk")
            self.honking_active = False
            
            # Wait for thread to finish
            if self.honk_thread and self.honk_thread.is_alive():
                self.honk_thread.join(timeout=1.0)
    
    def process_journal_entry(self, entry: dict):
        """Process a journal entry and trigger honk if needed."""
        try:
            event_type = entry.get('event')
            timestamp = entry.get('timestamp', 'Unknown')
            
            if event_type == 'FSDJump':
                new_system = entry.get('StarSystem')
                if new_system and new_system != self.current_system:
                    print(f"\n🚀 FSD JUMP DETECTED!")
                    print(f"   Time: {timestamp}")
                    print(f"   From: {self.current_system or 'Unknown'}")
                    print(f"   To: {new_system}")
                    
                    self.current_system = new_system
                    
                    # Stop any existing honking first
                    self.stop_honking()
                    
                    # Determine which key to use
                    key_to_use = None
                    if CONFIG['manual_key_override']:
                        key_to_use = CONFIG['manual_key_override']
                        print(f"   Using manual key override: {key_to_use}")
                    elif CONFIG['auto_detect_primary_fire'] and self.primary_fire_key:
                        key_to_use = self.primary_fire_key
                        print(f"   Using detected Primary Fire key: {key_to_use}")
                    else:
                        key_to_use = '1'  # Default fallback
                        print(f"   Using fallback key: {key_to_use}")
                    
                    # Schedule the honk
                    print(f"   Waiting {CONFIG['delay_after_jump']} seconds before honking...")
                    
                    def delayed_honk():
                        time.sleep(CONFIG['delay_after_jump'])
                        self.start_honking(key_to_use)
                    
                    # Run in separate thread so it doesn't block file monitoring
                    threading.Thread(target=delayed_honk, daemon=True).start()
            
            elif event_type == 'FSSDiscoveryScan':
                # This is the event that tells us the discovery scan is complete
                bodies_count = entry.get('BodyCount', 'Unknown')
                non_bodies_count = entry.get('NonBodyCount', 'Unknown')
                print(f"\n📡 FSS DISCOVERY SCAN COMPLETE!")
                print(f"   Time: {timestamp}")
                print(f"   Bodies found: {bodies_count}")
                print(f"   Non-body signals: {non_bodies_count}")
                
                # Stop honking
                self.stop_honking()
                print("-" * 60)
                    
            elif event_type in ['Location', 'LoadGame', 'StartUp']:
                # Track current system from these events too
                system = entry.get('StarSystem')
                if system and system != self.current_system:
                    self.current_system = system
                    print(f"📍 Current system: {system}")
                    
        except Exception as e:
            logger.error(f"Error processing journal entry: {e}")

class JournalMonitor(FileSystemEventHandler):
    def __init__(self, autohonk: AutoHonk):
        self.autohonk = autohonk
        self.current_file = None
        self.file_position = 0
        
        # Find the latest journal file
        self.find_latest_journal()
    
    def find_latest_journal(self):
        """Find and start monitoring the latest journal file."""
        try:
            journal_files = list(CONFIG['journal_folder'].glob('Journal.*.log'))
            if journal_files:
                latest_journal = max(journal_files, key=lambda x: x.stat().st_mtime)
                self.current_file = latest_journal
                self.file_position = latest_journal.stat().st_size  # Start at end of file
                logger.info(f"Monitoring journal file: {latest_journal}")
                print(f"📖 Monitoring: {latest_journal.name}")
            else:
                logger.warning("No journal files found")
                print("⚠️ No journal files found")
        except Exception as e:
            logger.error(f"Error finding journal files: {e}")
    
    def on_modified(self, event):
        """Handle file modification events."""
        if event.is_directory:
            return
            
        file_path = Path(event.src_path)
        
        # Check if it's a journal file
        if file_path.name.startswith('Journal.') and file_path.name.endswith('.log'):
            self.read_new_lines(file_path)
    
    def on_created(self, event):
        """Handle new file creation (new journal files)."""
        if event.is_directory:
            return
            
        file_path = Path(event.src_path)
        
        if file_path.name.startswith('Journal.') and file_path.name.endswith('.log'):
            print(f"\n📖 New journal file detected: {file_path.name}")
            self.current_file = file_path
            self.file_position = 0
    
    def read_new_lines(self, file_path: Path):
        """Read new lines from the journal file."""
        try:
            if file_path != self.current_file:
                return
                
            with open(file_path, 'r', encoding='utf-8') as f:
                f.seek(self.file_position)
                new_lines = f.readlines()
                self.file_position = f.tell()
                
                for line in new_lines:
                    line = line.strip()
                    if line:
                        try:
                            entry = json.loads(line)
                            self.autohonk.process_journal_entry(entry)
                        except json.JSONDecodeError:
                            pass  # Skip invalid JSON lines
                            
        except Exception as e:
            logger.error(f"Error reading journal file: {e}")

def main():
    """Main function to start the AutoHonk monitor."""
    print("Starting Elite Dangerous AutoHonk (FSS Discovery Mode)...")
    
    # Check if journal folder exists
    if not CONFIG['journal_folder'].exists():
        print(f"❌ Journal folder not found: {CONFIG['journal_folder']}")
        print("Make sure Elite Dangerous has been run at least once.")
        input("Press Enter to exit...")
        return
    
    # Initialize AutoHonk
    autohonk = AutoHonk()
    
    # Set up file monitoring
    event_handler = JournalMonitor(autohonk)
    observer = Observer()
    observer.schedule(event_handler, str(CONFIG['journal_folder']), recursive=False)
    
    # Start monitoring
    observer.start()
    
    try:
        print("\n✅ AutoHonk is running! Press Ctrl+C to stop.")
        while autohonk.running:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n\n🛑 Stopping AutoHonk...")
        autohonk.running = False
        autohonk.stop_honking()  # Make sure to stop any active honking
        observer.stop()
    
    observer.join()
    print("👋 AutoHonk stopped. Goodbye!")

if __name__ == "__main__":
    main()
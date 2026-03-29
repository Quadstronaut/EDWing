"""
Elite Dangerous AutoHonk - Sandboxie Multi-Instance Support

Monitors Elite Dangerous journal files and auto-presses Primary Fire
when jumping to a new system. Holds until FSSDiscoveryScan event.

Supports running one instance per sandbox by passing --sandbox <BoxName>.
When running outside a sandbox, monitors the default journal folder.

Requirements: pip install -r requirements.txt
"""

import argparse
import json
import logging
import os
import sys
import threading
import time
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Optional

import win32api
import win32con
import win32gui
import win32process
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

logger = logging.getLogger("autohonk")

# Key name mapping from Elite Dangerous bindings XML to Windows VK names
ELITE_KEY_MAP = {
    "Numpad_Add": win32con.VK_ADD,
    "Numpad_Subtract": win32con.VK_SUBTRACT,
    "Numpad_Multiply": win32con.VK_MULTIPLY,
    "Numpad_Divide": win32con.VK_DIVIDE,
    "Space": win32con.VK_SPACE,
    "Enter": win32con.VK_RETURN,
    "Tab": win32con.VK_TAB,
    **{f"F{n}": getattr(win32con, f"VK_F{n}") for n in range(1, 13)},
}


def resolve_journal_folder(sandbox: Optional[str] = None) -> Path:
    """Return the journal folder path, accounting for Sandboxie virtualisation."""
    default = Path.home() / "Saved Games" / "Frontier Developments" / "Elite Dangerous"

    if not sandbox:
        return default

    # Sandboxie stores virtualised user files under:
    # C:\Sandbox\<User>\<BoxName>\user\current\Saved Games\...
    # The exact root depends on Sandboxie config; check common locations.
    username = os.environ.get("USERNAME", "")
    candidates = [
        Path(f"C:/Sandbox/{username}/{sandbox}/user/current/Saved Games/Frontier Developments/Elite Dangerous"),
        Path(f"C:/Sandbox/{username}/{sandbox}/drive/C/Users/{username}/Saved Games/Frontier Developments/Elite Dangerous"),
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate

    logger.warning("Sandboxie journal folder not found for box '%s', falling back to default", sandbox)
    return default


def resolve_bindings_folder(sandbox: Optional[str] = None) -> Optional[Path]:
    """Return the Elite key bindings folder."""
    local_app_data = os.environ.get("LOCALAPPDATA", "")
    default = Path(local_app_data) / "Frontier Developments" / "Elite Dangerous" / "Options" / "Bindings"

    if not sandbox:
        return default if default.exists() else None

    username = os.environ.get("USERNAME", "")
    candidates = [
        Path(f"C:/Sandbox/{username}/{sandbox}/user/current/AppData/Local/Frontier Developments/Elite Dangerous/Options/Bindings"),
        default,  # bindings are often shared, not virtualised
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


def detect_primary_fire_key(bindings_dir: Optional[Path]) -> Optional[int]:
    """Read the Primary Fire key from the most recent .binds file. Returns VK code."""
    if not bindings_dir or not bindings_dir.exists():
        return None

    binds_files = list(bindings_dir.glob("*.binds"))
    if not binds_files:
        return None

    latest = max(binds_files, key=lambda p: p.stat().st_mtime)
    logger.info("Reading bindings from %s", latest)

    try:
        tree = ET.parse(latest)
        primary_fire = tree.getroot().find(".//PrimaryFire")
        if primary_fire is None:
            return None
        primary = primary_fire.find("Primary")
        if primary is None or primary.get("Device") != "Keyboard":
            return None

        key_name = primary.get("Key", "")
        if key_name.startswith("Key_"):
            key_name = key_name[4:]

        if key_name in ELITE_KEY_MAP:
            return ELITE_KEY_MAP[key_name]
        if len(key_name) == 1:
            return ord(key_name.upper())
    except Exception:
        logger.exception("Failed to parse bindings file")

    return None


class AutoHonk:
    def __init__(self, sandbox: Optional[str], window_filter: Optional[str],
                 delay: float, max_duration: float, manual_vk: Optional[int]):
        self.sandbox = sandbox
        self.window_filter = window_filter  # substring to match in window title
        self.delay_after_jump = delay
        self.max_honk_duration = max_duration
        self.manual_vk = manual_vk

        self.current_system: Optional[str] = None
        self.running = True
        self.honking_active = False
        self.honk_thread: Optional[threading.Thread] = None
        self.honk_lock = threading.Lock()

        # Detect primary fire key
        bindings_dir = resolve_bindings_folder(sandbox)
        self.fire_vk = manual_vk or detect_primary_fire_key(bindings_dir)
        if not self.fire_vk:
            logger.warning("Could not detect Primary Fire key; defaulting to '1'")
            self.fire_vk = ord("1")

    def find_elite_hwnd(self) -> Optional[int]:
        """Find the Elite Dangerous window matching our filter."""
        results = []

        def callback(hwnd, _):
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                title = win32gui.GetWindowText(hwnd)
                if "Elite - Dangerous" not in title:
                    return True
                if self.window_filter and self.window_filter.lower() not in title.lower():
                    return True
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                handle = win32api.OpenProcess(
                    win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid
                )
                try:
                    exe = win32process.GetModuleFileNameEx(handle, 0).lower()
                finally:
                    win32api.CloseHandle(handle)
                if "elitedangerous64" in exe:
                    results.append(hwnd)
            except Exception:
                pass
            return True

        win32gui.EnumWindows(callback, None)
        return results[0] if results else None

    def _do_honk(self):
        """Hold the fire key until stopped or timeout."""
        hwnd = self.find_elite_hwnd()
        if not hwnd:
            logger.warning("Elite window not found - skipping honk")
            return

        try:
            win32gui.SetForegroundWindow(hwnd)
        except Exception:
            logger.warning("Could not focus Elite window")
            return

        time.sleep(0.2)
        vk = self.fire_vk
        win32api.keybd_event(vk, 0, 0, 0)
        start = time.time()

        try:
            while self.honking_active and self.running:
                if time.time() - start >= self.max_honk_duration:
                    logger.info("Honk timeout (%.1fs)", self.max_honk_duration)
                    break
                time.sleep(0.05)
        finally:
            win32api.keybd_event(vk, 0, win32con.KEYEVENTF_KEYUP, 0)
            logger.info("Honk finished after %.1fs", time.time() - start)

    def start_honking(self):
        with self.honk_lock:
            if self.honking_active:
                return
            self.honking_active = True
            self.honk_thread = threading.Thread(target=self._do_honk, daemon=True)
            self.honk_thread.start()

    def stop_honking(self):
        with self.honk_lock:
            if not self.honking_active:
                return
            self.honking_active = False
            if self.honk_thread and self.honk_thread.is_alive():
                self.honk_thread.join(timeout=2.0)

    def process_entry(self, entry: dict):
        event = entry.get("event")

        if event == "FSDJump":
            system = entry.get("StarSystem")
            if system and system != self.current_system:
                logger.info("FSD Jump: %s -> %s", self.current_system or "?", system)
                self.current_system = system
                self.stop_honking()

                def delayed():
                    time.sleep(self.delay_after_jump)
                    self.start_honking()

                threading.Thread(target=delayed, daemon=True).start()

        elif event == "FSSDiscoveryScan":
            bodies = entry.get("BodyCount", "?")
            logger.info("FSS scan complete: %s bodies", bodies)
            self.stop_honking()

        elif event in ("Location", "LoadGame", "StartUp"):
            system = entry.get("StarSystem")
            if system:
                self.current_system = system
                logger.info("Current system: %s", system)


class JournalWatcher(FileSystemEventHandler):
    def __init__(self, honker: AutoHonk):
        self.honker = honker
        self.current_file: Optional[Path] = None
        self.file_position = 0
        self._find_latest()

    def _find_latest(self):
        journal_dir = resolve_journal_folder(self.honker.sandbox)
        journals = sorted(journal_dir.glob("Journal.*.log"), key=lambda p: p.stat().st_mtime)
        if journals:
            self.current_file = journals[-1]
            self.file_position = self.current_file.stat().st_size
            logger.info("Tailing %s", self.current_file.name)

    def on_modified(self, event):
        if event.is_directory:
            return
        path = Path(event.src_path)
        if path.name.startswith("Journal.") and path.name.endswith(".log") and path == self.current_file:
            self._read_new(path)

    def on_created(self, event):
        if event.is_directory:
            return
        path = Path(event.src_path)
        if path.name.startswith("Journal.") and path.name.endswith(".log"):
            logger.info("New journal: %s", path.name)
            self.current_file = path
            self.file_position = 0

    def _read_new(self, path: Path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                f.seek(self.file_position)
                for line in f:
                    line = line.strip()
                    if line:
                        try:
                            self.honker.process_entry(json.loads(line))
                        except json.JSONDecodeError:
                            pass
                self.file_position = f.tell()
        except Exception:
            logger.exception("Error reading journal")


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Elite Dangerous AutoHonk")
    p.add_argument("--sandbox", "-s", help="Sandboxie box name (e.g. CMDRBistronaut)")
    p.add_argument("--window-filter", "-w",
                    help="Substring to match in Elite window title (default: sandbox name or 'Elite - Dangerous')")
    p.add_argument("--delay", type=float, default=2.0, help="Seconds after jump before honking (default: 2)")
    p.add_argument("--max-duration", type=float, default=7.0, help="Max honk duration in seconds (default: 7)")
    p.add_argument("--key", help="Manual key override (e.g. '1', 'space', 'numpad_add')")
    p.add_argument("--verbose", "-v", action="store_true")
    return p


def main():
    args = build_parser().parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[logging.StreamHandler()],
    )

    # Resolve manual key override
    manual_vk = None
    if args.key:
        key_upper = args.key.strip().replace(" ", "_").title()
        if key_upper in ELITE_KEY_MAP:
            manual_vk = ELITE_KEY_MAP[key_upper]
        elif len(args.key) == 1:
            manual_vk = ord(args.key.upper())
        else:
            logger.error("Unknown key: %s", args.key)
            sys.exit(1)

    window_filter = args.window_filter or args.sandbox
    journal_folder = resolve_journal_folder(args.sandbox)

    if not journal_folder.exists():
        logger.error("Journal folder not found: %s", journal_folder)
        logger.error("Make sure Elite Dangerous has been run at least once%s.",
                      f" in sandbox '{args.sandbox}'" if args.sandbox else "")
        sys.exit(1)

    honker = AutoHonk(
        sandbox=args.sandbox,
        window_filter=window_filter,
        delay=args.delay,
        max_duration=args.max_duration,
        manual_vk=manual_vk,
    )

    watcher = JournalWatcher(honker)
    observer = Observer()
    observer.schedule(watcher, str(journal_folder), recursive=False)
    observer.start()

    label = f" (sandbox: {args.sandbox})" if args.sandbox else ""
    logger.info("AutoHonk running%s - monitoring %s", label, journal_folder)
    logger.info("Primary fire VK code: 0x%02X", honker.fire_vk)
    logger.info("Press Ctrl+C to stop")

    try:
        while honker.running:
            time.sleep(1)
    except KeyboardInterrupt:
        logger.info("Shutting down...")
        honker.running = False
        honker.stop_honking()
        observer.stop()

    observer.join()


if __name__ == "__main__":
    main()

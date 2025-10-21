#!/usr/bin/env python
# -*- coding: utf-8 -*-
 
import os
import sys
import time
import ctypes
import random
import subprocess
from dataclasses import dataclass
from typing import List, Optional, Dict, Any
import psutil
import win32gui
import win32con
import win32process
import win32api
 
# ==============================
# Configuration
# ==============================
 
@dataclass
class Config:
    launch_elite_dangerous: bool = True
    skip_intro: bool = True
    pg_entry: bool = True
    launch_edeb: bool = False
    launch_edmc: bool = True
    python_path: str = r'C:\Users\Quadstronaut\scoop\apps\python\current\python.exe'
    window_poll_interval: float = 0.333  # seconds
    process_wait_interval: float = 0.5   # seconds
    window_move_retry_interval: float = 0.5
    max_retries: int = 3
 
config = Config()
 
# Commander names
CMDR_NAMES = [
    "CMDRDuvrazh",
    "CMDRBistronaut",
    "CMDRTristronaut",
    "CMDRQuadstronaut"
]
 
# Alt accounts (skip first)
ELITE_DANGEROUS_CMDRS = CMDR_NAMES[1:]
 
# Paths
SANDBOXIE_START = r'C:\Users\Quadstronaut\scoop\apps\sandboxie-plus-np\current\Start.exe'
MIN_ED_LAUNCHER = r'G:\SteamLibrary\steamapps\common\Elite Dangerous\MinEdLauncher.exe'
EDEB_LAUNCHER = r'G:\EliteApps\EDEB\Elite Dangerous Exploration Buddy.exe'
 
# Window configurations
window_configs: List[Dict[str, Any]] = []
 
 
# ==============================
# Windows API via ctypes
# ==============================
 
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
 
SW_MAXIMIZE = 3
SWP_NOZORDER = 0x0004
 
 
def set_window_pos(hwnd: int, x: int, y: int, width: int, height: int, maximize: bool = False) -> bool:
    """Move and resize a window using Windows API."""
    if maximize:
        return bool(user32.ShowWindow(hwnd, SW_MAXIMIZE))
    else:
        return bool(user32.SetWindowPos(
            hwnd, 0, x, y, width, height,
            SWP_NOZORDER
        ))
 
 
def find_window(process_name: str, title_contains: str) -> Optional[int]:
    """Find window handle by process name and partial title match."""
    def enum_windows_callback(hwnd, windows):
        if win32gui.IsWindowVisible(hwnd):
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            try:
                proc = psutil.Process(pid)
                if proc.name().lower().startswith(process_name.lower()):
                    title = win32gui.GetWindowText(hwnd)
                    if title_contains in title:
                        windows.append((hwnd, title))
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
        return True
 
    windows = []
    win32gui.EnumWindows(enum_windows_callback, windows)
    return windows[0][0] if windows else None
 
 
# ==============================
# Window Positioning Function
# ==============================
 
def set_window_position(
    process_name: str,
    window_title: str,
    x: int,
    y: int,
    width: int,
    height: int,
    maximize: bool = False
) -> bool:
    hwnd = find_window(process_name, window_title)
    if not hwnd:
        print(f"[WARN] Window not found: {process_name} | Title: {window_title}")
        return False
 
    time.sleep(0.1)  # Small delay for stability
 
    success = set_window_pos(hwnd, x, y, width, height, maximize)
    if success:
        action = "Maximized" if maximize else f"Moved to ({x}, {y}), {width}x{height}"
        print(f"[INFO] {action} window: '{window_title}' ({process_name})")
    else:
        print(f"[ERROR] Failed to position window: '{window_title}'")
    return success
 
 
# ==============================
# Build Window Configurations
# ==============================
 
if config.launch_elite_dangerous:
    elite_windows = [
        {"name": ELITE_DANGEROUS_CMDRS[0], "x": -1080, "y": -387, "w": 800, "h": 600, "moved": False, "retry": 0},
        {"name": ELITE_DANGEROUS_CMDRS[1], "x": -1080, "y": 213,  "w": 800, "h": 600, "moved": False, "retry": 0},
        {"name": ELITE_DANGEROUS_CMDRS[2], "x": -1080, "y": 813,  "w": 800, "h": 600, "moved": False, "retry": 0},
    ]
    for w in elite_windows:
        w["process"] = "EliteDangerous64"
        w["maximize"] = False
    window_configs.extend(elite_windows)
 
if config.launch_edeb:
    edeb_windows = [
        {"process": "Elite Dangerous Exploration Buddy", "name": CMDR_NAMES[0], "maximize": True, "moved": False, "retry": 0}
    ]
    window_configs.extend(edeb_windows)
    # Launch EDEB
    if os.path.exists(EDEB_LAUNCHER):
        subprocess.Popen([EDEB_LAUNCHER])
    else:
        print(f"[ERROR] EDEB launcher not found: {EDEB_LAUNCHER}")
 
if config.launch_edmc:
    edmc_windows = [
        {"name": CMDR_NAMES[0], "x": 100,  "y": 100,  "w": 300, "h": 600, "moved": False, "retry": 0},
        {"name": CMDR_NAMES[1], "x": -280, "y": -387, "w": 300, "h": 600, "moved": False, "retry": 0},
        {"name": CMDR_NAMES[2], "x": -280, "y": 213,  "w": 300, "h": 600, "moved": False, "retry": 0},
        {"name": CMDR_NAMES[3], "x": -280, "y": 813,  "w": 300, "h": 600, "moved": False, "retry": 0},
    ]
    for w in edmc_windows:
        w["process"] = "EDMarketConnector"
        w["maximize"] = False
    window_configs.extend(edmc_windows)
 
 
# ==============================
# Validation
# ==============================
 
def validate_paths() -> bool:
    sbs_exists = os.path.exists(SANDBOXIE_START)
    edl_exists = os.path.exists(MIN_ED_LAUNCHER)
   
    if not sbs_exists:
        print(f"[ERROR] Sandboxie not found: {SANDBOXIE_START}")
    if not edl_exists:
        print(f"[ERROR] MinEDLauncher not found: {MIN_ED_LAUNCHER}")
   
    return sbs_exists and edl_exists
 
 
# ==============================
# Main Execution
# ==============================
 
def main():
    if not validate_paths():
        sys.exit(1)
 
    print("Starting Elite Dangerous multibox...")
 
    # Launch all Elite instances via Sandboxie
    if config.launch_elite_dangerous:
        for i, cmdr in enumerate(CMDR_NAMES):
            args = [
                SANDBOXIE_START,
                f"/box:{cmdr}",
                MIN_ED_LAUNCHER,
                f"/frontier", f"Account{i+1}", "/edo", "/autorun", "/autoquit", "/skipInstallPrompt"
            ]
            subprocess.Popen(args)
            print(f"Launched {cmdr} in sandbox")
 
    if not window_configs:
        print("No windows to manage.")
        return
 
    # --- WINDOW DETECTION PHASE ---
    print("\nWaiting for application windows to load...")
    previous_count = -1
    all_found = False
 
    while not all_found:
        found = 0
        for wc in window_configs:
            if find_window(wc["process"], wc["name"]):
                found += 1
 
        if found != previous_count:
            print(f"Found {found} of {len(window_configs)}")
            previous_count = found
 
        if found == len(window_configs):
            all_found = True
        else:
            time.sleep(config.window_poll_interval)
 
    # --- WINDOW POSITIONING PHASE ---
    print("\nAll windows detected. Beginning positioning...")
    first_elite_wait = True
 
    while any(not wc["moved"] for wc in window_configs):
        for wc in window_configs:
            if wc["moved"]:
                continue
 
            # Wait for correct window title
            hwnd = None
            while not hwnd:
                hwnd = find_window(wc["process"], wc["name"])
                if not hwnd:
                    time.sleep(config.process_wait_interval)
 
            # Special delay for Elite Dangerous (first instance only)
            if wc["process"] == "EliteDangerous64" and first_elite_wait:
                delay = random.randint(7, 11)
                print(f"Waiting {delay} seconds for Elite Dangerous to stabilize...")
                time.sleep(delay)
                first_elite_wait = False
 
            # Try to position
            if wc["retry"] < config.max_retries:
                success = set_window_position(
                    process_name=wc["process"],
                    window_title=wc["name"],
                    x=wc.get("x", 0),
                    y=wc.get("y", 0),
                    width=wc.get("w", 0),
                    height=wc.get("h", 0),
                    maximize=wc.get("maximize", False)
                )
                if success:
                    wc["moved"] = True
                else:
                    wc["retry"] += 1
                    print(f"Retry {wc['retry']}/{config.max_retries} for {wc['name']}")
            else:
                print(f"[WARN] Max retries exceeded for {wc['name']}. Skipping.")
                wc["moved"] = True
 
        time.sleep(config.window_move_retry_interval)
 
    print("\nWindow positioning complete!")
 
    # --- Post-launch scripts ---
    script_dir = os.path.dirname(os.path.abspath(__file__))
    intro_skipped = False
 
    if config.skip_intro:
        intro_script = os.path.join(script_dir, 'clicker_scripts', 'cutscene.ps1')
        if os.path.exists(intro_script):
            subprocess.Popen(['powershell', '-ExecutionPolicy', 'Bypass', '-File', intro_script])
            intro_skipped = True
            print("Launched intro skip script.")
        else:
            print(f"[WARN] Intro skip script not found: {intro_script}")
 
    if config.pg_entry and intro_skipped:
        time.sleep(20)  # Tuned delay
        pg_script = os.path.join(script_dir, 'clicker_scripts', 'continue-pg.ps1')
        if os.path.exists(pg_script):
            subprocess.Popen(['powershell', '-ExecutionPolicy', 'Bypass', '-File', pg_script])
            print("Launched PG entry script.")
        else:
            print(f"[WARN] PG entry script not found: {pg_script}")
 
    # Optional: Run Python input broadcaster
    # input_script = os.path.join(script_dir, 'input_broadcast.py')
    # if os.path.exists(input_script):
    #     subprocess.Popen([config.python_path, input_script])
 
    # Optional: Run PowerShell input broadcaster
    # input_ps = os.path.join(script_dir, 'input_broadcast.ps1')
    # if os.path.exists(input_ps):
    #     subprocess.Popen(['powershell', '-File', input_ps])
 
 
if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n[INFO] Stopped by user.")
    except Exception as e:
        print(f"[FATAL] {e}")
        sys.exit(1)

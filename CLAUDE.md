# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Is

EDWing is a personal automation suite for multiboxing Elite Dangerous with 4 commander accounts via Sandboxie-Plus sandboxes. It launches all instances, positions their windows across monitors, skips intro cutscenes, auto-selects Private Group, and provides in-game automation (autohonk on system jump, input broadcasting).

## How It Runs

There is no build step, test suite, or linter. This is a collection of PowerShell and Python scripts run directly.

```powershell
# Main launcher - starts 4 sandboxed Elite instances + companion apps, positions windows
.\Get-Wing.ps1

# AutoHonk - run one per sandbox (or unsandboxed for primary CMDR)
pip install -r requirements.txt
python autohonk\autohonk.py --sandbox CMDRBistronaut
python autohonk\autohonk.py  # primary commander, no sandbox

# Input broadcast (experimental, nonfunctional at time of writing)
.\input_broadcast.ps1
python input_broadcast.py
```

## Architecture

**Get-Wing.ps1** is the orchestrator. Flow:
1. Loads defaults, then dot-sources `wing.conf.ps1` (gitignored) for local overrides
2. Builds a `$windowConfigurations` array based on feature flags (`launchEliteDangerous`, `launchEDMC`, `launchEDEB`)
3. Validates that required executables exist (only checks paths for enabled features)
4. Launches each CMDR in a Sandboxie sandbox via `Start.exe /box:<name>`
5. Polls for windows by process name + title match until all appear
6. Positions windows using Win32 `SetWindowPos`/`ShowWindowAsync` with retry logic
7. Runs clicker scripts for cutscene skip and PG entry

**Sandboxie model:** Each CMDR name (e.g. `CMDRBistronaut`) is both the sandbox box name and the window title substring used for matching. The first CMDR in `$cmdrNames` plays on the primary monitor; the rest go to alt monitors at negative X coordinates.

**clicker_scripts/** use absolute screen coordinates to automate mouse clicks. `MouseUtil.ps1` is the shared module — both scripts dot-source it. Coordinates are tuned to a specific multi-monitor layout and resolution.

**autohonk/autohonk.py** monitors Elite journal files via watchdog, detects FSDJump events, and holds the Primary Fire key until FSSDiscoveryScan completes. Supports `--sandbox <BoxName>` to resolve Sandboxie's virtualised journal folder paths.

**input_broadcast.ps1 / input_broadcast.py** are experimental and nonfunctional. They attempt to relay keypresses to all Elite windows via PostMessage or keybd_event.

## Key Constraints

- All Win32 `Add-Type` calls must check `[System.Management.Automation.PSTypeName]` before loading to avoid duplicate type errors on re-runs.
- Window configs in `$windowConfigurations` must always include all keys: `Name`, `ProcessName`, `X`, `Y`, `Width`, `Height`, `Maximize`, `Moved`, `RetryCount`. Missing keys cause parameter binding failures in `Set-WindowPosition`.
- Sandboxie journal paths vary by installation. `autohonk.py` checks `C:/Sandbox/<user>/<box>/user/current/...` and `C:/Sandbox/<user>/<box>/drive/C/...` as candidates.
- Clicker script coordinates are brittle — they're pixel positions for a specific monitor layout. Changing resolution or monitor arrangement breaks them.
- Elite Dangerous windows need a settle delay after detection before positioning works reliably (renderer not ready immediately).

## Git Workflow

Feature branches with PR, squash merge to `master`. Commit style: imperative present tense, ≤72 char subject. Branch prefixes: `feat/`, `fix/`, `docs/`, `chore/`.

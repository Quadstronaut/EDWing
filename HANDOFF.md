# EDWing — work-in-progress handoff (resume here)

**Branch:** `feat/mode2-launcher` (all work committed + pushed to origin; `master` is unrelated/behind).
**Goal:** make the multibox launcher functional again — best way, not the old way. **Mode 2 first**
(all 4 clients borderless full-size on the primary monitor; a bot/you bring the wanted CMDR
forward, act, cycle). Mode 1 (old tiled layout) is secondary. Eventually ED-AFK drives all 4.

Last worked: 2026-07-13. Session paused for a reboot; the only thing left is a **keyboard session**
(logins + live validation) that I can't do remotely.

---

## What's built (all committed on this branch)

| File | Purpose | State |
|------|---------|-------|
| `WingLib.ps1` | Window identity (Sandboxie box-tag → HWND), `AttachThreadInput` focus, live-monitor rects, `ShowWindowAsync`, borderless stack | ✅ built, read-only validated |
| `WingCreds.ps1` | Identity-anchored native-first credential + `settings.json` seeding, integrity-checked | ✅ built; **ran it → all 4 boxes credentialed** |
| `Get-Wing.ps1` | Mode-2 launcher: concurrent bounded wait, re-auth-safe skip, stack-on-primary, per-click focus re-verify. `-Mode stacked|tiled`, `-NoCreds`, `-WhatIf` | ✅ built, parses, `-WhatIf` clean |
| `tools/Get-CursorPos.ps1` | Capture per-CMDR menu click coords | ✅ |
| `tools/Show-Cmdr.ps1` | Bring one CMDR window to foreground (standalone) | ✅ |
| `example_configs/wing.conf.ps1.example` | Updated to the new schema | ✅ |
| removed `min-ed-launcher.ini` | Was fictional (min-ed uses settings.json + interactive-login `.cred`) | ✅ |

Committed machine defaults let `.\Get-Wing.ps1` run **without** a wing.conf.

---

## Credential state
Seeding was actually run: **all 4 boxes now have a token + settings.json** (Duvrazh seeded from the
native Account1 token; Bistronaut/Tristronaut/Quadstronaut already had tokens; native tokens backed up
to gitignored `cred_backup/`). Identity map is pinned: Duvrazh=Account1, Bistronaut=Account2,
Tristronaut=Account3, Quadstronaut=Account4.

**UNVERIFIED — the whole point of the keyboard session:** whether those seeded DPAPI tokens actually
*decrypt inside the sandbox* or prompt for a fresh login.

---

## RESUME RUNBOOK (run at the keyboard, from the repo dir)

```powershell
# 0. Close the native/unsandboxed ED first (cleaner; the launcher ignores it regardless)

# 1. Smoke-test ONE box — does the seeded token auto-login, or ask for login/2FA?
& 'C:\Users\Quadstronaut\scoop\apps\sandboxie-plus-np\current\Start.exe' /box:CMDRBistronaut `
  'G:\SteamLibrary\steamapps\common\Elite Dangerous\MinEdLauncher.exe' `
  /frontier Account2 /edo /autorun /autoquit /skipInstallPrompt

# 2. Once its window is up — capture the EXACT box-tag title format + confirm detection:
. .\WingLib.ps1; Get-CmdrWindows -BoxNames @('CMDRBistronaut') | Format-List Box,Title,Hwnd

# 3. Focus across the sandbox boundary?
.\tools\Show-Cmdr.ps1 CMDRBistronaut

# 4. If 1–3 are good: the full run (all 4, stacked on primary)
.\Get-Wing.ps1

# 5. One-time: set each client to Borderless display mode in ED graphics (for clean stacking)
# 6. Capture menu coords: .\tools\Get-CursorPos.ps1  → paste into wing.conf.ps1 CmdrClickSets
```

**Report back to Claude:** (a) did tokens auto-login or prompt? (b) paste the `Title` from step 2
(the real tag format) (c) did focus + stacking behave? → then finalize the click flow.

---

## Deferred TODOs (documented, not yet done)
- Replace the flat `PgReadyDelaySeconds` sleep with a **per-CMDR journal readiness signal** (poll for
  Fileheader/LoadGame, like autohonk) before firing menu clicks — avoids clicking into a re-auth prompt.
- **Already-running detection** before re-launching a box (avoid a second instance / stale-window pick).
- **Re-poll timed-out boxes** for late arrivals so they get placed instead of intruding on the layout.
- Update **README.md** (still describes the OLD launcher) after keyboard validation confirms behavior.
- Optional hardening: move machine-specific paths from `Get-Wing.ps1` into gitignored `wing.conf.ps1`.

## Review items intentionally NOT done (don't re-litigate)
- "Restore WS_CAPTION on non-borderless SetRect" — **rejected**: would add a title bar to a game
  legitimately in Borderless mode. Cosmetic cross-mode-rerun caveat instead.
- Get-BoxMinEdDir second path convention / DPAPI-in-sandbox — acknowledged assumptions; the keyboard
  test is the empirical resolution.

## Environment quick-facts
- Monitors: DISPLAY1 primary 1920×1080 @ (0,0); DISPLAY2 vertical 1080×1920 @ (-1080,-406).
- Sandboxie boxes have `BoxNameTitle=y` (title carries the box name), no `AutoDelete` (persist),
  `OpenWinClass=EliteDangerous64.exe/IgnoreUIPI` (launcher can manipulate the sandboxed windows),
  companions auto-run via `RunCommand`. min-ed-launcher v0.12.2.
- AutoHotkey is NOT installed (native AttachThreadInput used instead; AHK is the fallback if it fails).

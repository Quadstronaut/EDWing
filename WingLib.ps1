# WingLib.ps1 - Core window identity / focus / positioning primitives for EDWing.
# Dot-source from the launcher:  . "$PSScriptRoot\WingLib.ps1"
#
# Why this exists: the old launcher matched Elite windows by the game's title, which
# is the CMDR-less "Elite - Dangerous (CLIENT)" and can match several processes at
# once (wrong-window bug). Here we key off the Sandboxie box-name tag that each box
# stamps into its window title (BoxNameTitle=y), which uniquely identifies the CMDR,
# and we resolve exactly one HWND per box. This module also owns the foreground
# primitive that mode 2 (bot brings a client forward) and future ED-AFK depend on.

#Requires -Version 5.1

# --- Win32 surface (loaded once; guarded per repo convention) ---
if (-not ([System.Management.Automation.PSTypeName]'WingWin32').Type) {
    Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;

public class WingWin32 {
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr p);
    delegate bool EnumProc(IntPtr h, IntPtr p);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] static extern bool BringWindowToTop(IntPtr h);
    // ShowWindowAsync POSTS the show request (non-blocking) instead of SendMessage-ing it,
    // so a target ED thread that isn't pumping its queue (loading screen, re-auth modal,
    // fullscreen transition) can't block our single launcher thread. See review rank 2.
    [DllImport("user32.dll")] static extern bool ShowWindowAsync(IntPtr h, int n);
    [DllImport("user32.dll")] static extern bool AttachThreadInput(uint a, uint b, bool attach);
    [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
    [DllImport("user32.dll")] static extern int GetWindowLong(IntPtr h, int idx);
    [DllImport("user32.dll")] static extern int SetWindowLong(IntPtr h, int idx, int val);
    [DllImport("user32.dll", EntryPoint="GetWindowLongPtrW")] static extern IntPtr GetWindowLongPtr64(IntPtr h, int idx);
    [DllImport("user32.dll", EntryPoint="SetWindowLongPtrW")] static extern IntPtr SetWindowLongPtr64(IntPtr h, int idx, IntPtr val);

    const int GWL_STYLE = -16;
    const long WS_CAPTION     = 0x00C00000;
    const long WS_THICKFRAME  = 0x00040000;
    const long WS_SYSMENU     = 0x00080000;
    const long WS_MINIMIZEBOX = 0x00020000;
    const long WS_MAXIMIZEBOX = 0x00010000;
    static readonly IntPtr HWND_TOP = IntPtr.Zero;
    const uint SWP_FRAMECHANGED = 0x0020;
    const uint SWP_SHOWWINDOW   = 0x0040;
    const uint SWP_NOZORDER     = 0x0004;
    const uint SWP_NOACTIVATE   = 0x0010;   // position without stealing focus (review rank 7)
    const int SW_RESTORE = 9;

    public struct WinInfo { public long Hwnd; public uint Pid; public string Title; public bool Visible; }

    public static long ForegroundHwnd() { return GetForegroundWindow().ToInt64(); }

    // Every top-level window owned by a PID (there can be more than one; caller filters).
    public static WinInfo[] WindowsForPid(uint pid) {
        var list = new List<WinInfo>();
        EnumWindows((h, p) => {
            uint wpid; GetWindowThreadProcessId(h, out wpid);
            if (wpid == pid) {
                var sb = new StringBuilder(512); GetWindowText(h, sb, 512);
                var wi = new WinInfo();
                wi.Hwnd = h.ToInt64(); wi.Pid = wpid; wi.Title = sb.ToString(); wi.Visible = IsWindowVisible(h);
                list.Add(wi);
            }
            return true;
        }, IntPtr.Zero);
        return list.ToArray();
    }

    static long GetStyle(IntPtr h) {
        return IntPtr.Size == 8 ? GetWindowLongPtr64(h, GWL_STYLE).ToInt64() : (long)GetWindowLong(h, GWL_STYLE);
    }
    static void SetStyle(IntPtr h, long val) {
        if (IntPtr.Size == 8) SetWindowLongPtr64(h, GWL_STYLE, new IntPtr(val));
        else SetWindowLong(h, GWL_STYLE, (int)val);
    }

    // Raise a specific window to the foreground, defeating Windows' foreground-lock by
    // briefly attaching our input queue to the current foreground thread (the trick AHK
    // uses internally). Returns true only if the window actually became foreground.
    public static bool ForceForeground(long hwnd) {
        IntPtr h = new IntPtr(hwnd);
        ShowWindowAsync(h, SW_RESTORE);
        IntPtr fore = GetForegroundWindow();
        uint foreThread; GetWindowThreadProcessId(fore, out foreThread);
        uint thisThread = GetCurrentThreadId();
        bool attached = false;
        if (foreThread != 0 && foreThread != thisThread) attached = AttachThreadInput(thisThread, foreThread, true);
        BringWindowToTop(h);
        SetForegroundWindow(h);
        if (attached) AttachThreadInput(thisThread, foreThread, false);
        return GetForegroundWindow() == h;
    }

    // Position/size a window WITHOUT activating it (mode-2 stacking must not fight for
    // focus). Borderless strips the caption+frame first (a no-op if the game is already
    // in Borderless display mode). Note: we deliberately never re-add caption bits - a
    // game legitimately in Borderless mode must not sprout a title bar (review rank 12).
    public static bool SetRect(long hwnd, int x, int y, int w, int h, bool borderless) {
        IntPtr hh = new IntPtr(hwnd);
        ShowWindowAsync(hh, SW_RESTORE);
        if (borderless) {
            long style = GetStyle(hh);
            style &= ~(WS_CAPTION | WS_THICKFRAME | WS_SYSMENU | WS_MINIMIZEBOX | WS_MAXIMIZEBOX);
            SetStyle(hh, style);
            return SetWindowPos(hh, HWND_TOP, x, y, w, h, SWP_FRAMECHANGED | SWP_SHOWWINDOW | SWP_NOACTIVATE);
        }
        return SetWindowPos(hh, HWND_TOP, x, y, w, h, SWP_NOZORDER | SWP_SHOWWINDOW | SWP_NOACTIVATE);
    }
}
'@
}

# --- Identity ---

# Return one object per visible Elite window: Box (derived from the Sandboxie title tag),
# Pid, Hwnd, Title. Untagged windows (a native/unsandboxed client) come back with Box=$null.
# When several box names match a title (a future name that is a substring of another), the
# LONGEST match wins so a prefix name can't steal a longer box's window (review rank 13).
function Get-CmdrWindows {
    [CmdletBinding()]
    param(
        [string]$ProcessName = 'EliteDangerous64',
        [string[]]$BoxNames  = @()
    )
    $out = @()
    foreach ($p in (Get-Process -Name $ProcessName -ErrorAction SilentlyContinue)) {
        foreach ($w in [WingWin32]::WindowsForPid([uint32]$p.Id)) {
            if (-not $w.Visible) { continue }                       # skip IME/helper windows
            if ([string]::IsNullOrWhiteSpace($w.Title)) { continue }
            $hits = @($BoxNames | Where-Object { $w.Title -like "*$_*" })
            $box  = if ($hits) { $hits | Sort-Object { $_.Length } -Descending | Select-Object -First 1 } else { $null }
            $out += [pscustomobject]@{
                Box = $box; Pid = [int]$w.Pid; Hwnd = [long]$w.Hwnd; Title = $w.Title
            }
        }
    }
    return $out
}

# Bounded wait for ONE box's window. Returns the window object, or $null on timeout.
function Wait-CmdrWindow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$BoxName,
        [int]$TimeoutSec = 120,
        [int]$PollMs     = 2000,
        [string]$ProcessName = 'EliteDangerous64'
    )
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    do {
        $win = Get-CmdrWindows -ProcessName $ProcessName -BoxNames @($BoxName) |
            Where-Object { $_.Box -eq $BoxName } | Select-Object -First 1
        if ($win) { return $win }
        Start-Sleep -Milliseconds $PollMs
    } while ((Get-Date) -lt $deadline)
    return $null
}

# Bounded wait for MANY boxes CONCURRENTLY: one shared poll retires each box as it appears,
# so total wait is bounded by the single largest timeout, not the sum (review rank 9).
# Returns @{ Ready = <window objects>; NotReady = <box names that timed out> }.
function Wait-AllCmdrWindows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$BoxNames,
        [int]$TimeoutSec = 120,
        [int]$PollMs     = 2000,
        [string]$ProcessName = 'EliteDangerous64'
    )
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    $ready = @{}
    do {
        $wins = Get-CmdrWindows -ProcessName $ProcessName -BoxNames $BoxNames
        foreach ($b in $BoxNames) {
            if ($ready.ContainsKey($b)) { continue }
            $w = $wins | Where-Object { $_.Box -eq $b } | Select-Object -First 1
            if ($w) { $ready[$b] = $w; Write-Host ("  {0} ready (hwnd 0x{1:X})" -f $b, $w.Hwnd) -ForegroundColor Green }
        }
        if ($ready.Count -lt $BoxNames.Count) { Start-Sleep -Milliseconds $PollMs }
    } while ($ready.Count -lt $BoxNames.Count -and (Get-Date) -lt $deadline)

    $readyList = @(); $notReady = @()
    foreach ($b in $BoxNames) {
        if ($ready.ContainsKey($b)) { $readyList += $ready[$b] } else { $notReady += $b }
    }
    return [pscustomobject]@{ Ready = $readyList; NotReady = $notReady }
}

# --- Monitors ---

# Live monitor list so placement isn't baked to one machine's pixel constants.
function Get-WingMonitorLayout {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Screen]::AllScreens | ForEach-Object {
        [pscustomobject]@{
            Name        = $_.DeviceName
            Primary     = $_.Primary
            X           = $_.Bounds.X
            Y           = $_.Bounds.Y
            Width       = $_.Bounds.Width
            Height      = $_.Bounds.Height
            Orientation = $(if ($_.Bounds.Height -gt $_.Bounds.Width) { 'Vertical' } else { 'Landscape' })
        }
    }
}

function Get-WingPrimaryRect {
    Get-WingMonitorLayout | Where-Object Primary | Select-Object -First 1
}

# --- Actions (thin PowerShell over the Win32 class) ---

function Set-CmdrForeground {
    [CmdletBinding()]
    param([Parameter(Mandatory)][long]$Hwnd)
    return [WingWin32]::ForceForeground($Hwnd)
}

function Set-CmdrWindowRect {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][long]$Hwnd,
        [Parameter(Mandatory)][int]$X,
        [Parameter(Mandatory)][int]$Y,
        [Parameter(Mandatory)][int]$Width,
        [Parameter(Mandatory)][int]$Height,
        [switch]$Borderless
    )
    return [WingWin32]::SetRect($Hwnd, $X, $Y, $Width, $Height, [bool]$Borderless)
}

# Currently-foreground HWND, for verifying a focus attempt actually took (review rank 5).
function Get-ForegroundHwnd {
    return [long][WingWin32]::ForegroundHwnd()
}

# --- Self-test ---

# Read-only by default: enumerates Elite windows + monitors and prints what it sees.
# -TestFocus / -TestMove exercise the disruptive primitives - and are HARD-GUARDED to
# only ever touch a Box-TAGGED (sandboxed) window, never the live/native ED (review rank 1).
function Invoke-WingSelfTest {
    [CmdletBinding()]
    param(
        [string[]]$BoxNames = @('CMDRDuvrazh','CMDRBistronaut','CMDRTristronaut','CMDRQuadstronaut'),
        [switch]$TestFocus,
        [switch]$TestMove
    )
    Write-Host "=== Monitors ===" -ForegroundColor Cyan
    Get-WingMonitorLayout | Format-Table -AutoSize | Out-String | Write-Host

    Write-Host "=== Elite windows ===" -ForegroundColor Cyan
    $wins = Get-CmdrWindows -BoxNames $BoxNames
    if (-not $wins) { Write-Host "  (none running)" }
    else { $wins | Format-Table Box, Pid, @{n='Hwnd';e={'0x{0:X}' -f $_.Hwnd}}, Title -AutoSize | Out-String | Write-Host }

    if ($TestFocus -or $TestMove) {
        # NEVER select an untagged (native/unsandboxed) window as a test target.
        $t = $wins | Where-Object { $_.Box } | Select-Object -First 1
        if (-not $t) {
            Write-Warning "No sandboxed (box-tagged) CMDR window found - refusing to focus/move an untagged/native window."
            return $wins
        }
        if ($TestFocus) {
            Write-Host "Focus test -> $($t.Box)" -ForegroundColor Yellow
            Write-Host ("  ForceForeground returned: {0}" -f (Set-CmdrForeground -Hwnd $t.Hwnd))
        }
        if ($TestMove) {
            $r = Get-WingPrimaryRect
            Write-Host "Move test -> $($t.Box) to primary $($r.Width)x$($r.Height)" -ForegroundColor Yellow
            Write-Host ("  SetRect returned: {0}" -f (Set-CmdrWindowRect -Hwnd $t.Hwnd -X $r.X -Y $r.Y -Width $r.Width -Height $r.Height))
        }
    }
    return $wins
}

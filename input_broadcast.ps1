# Elite Dangerous Input Broadcaster - PowerShell Version
# Uses SendInput API like the working cutscene script
# Try PostMessage first (no focus)
# .\input_broadcast.ps1 -UseDirect
# If that doesn't work, use keybd_event (focus mode, like autohonk)
# .\input_broadcast.ps1

# Configuration
$config = @{
    WindowTitleContains = "Elite - Dangerous (CLIENT)"
    ProcessName         = "EliteDangerous64"
    Commanders          = @("Bistronaut", "Tristronaut", "Quadstronaut")
    PrimaryCommander    = "Duvrazh"
    TypingTimeout       = 1.0    # Seconds to wait after typing stops
    KeySendDelay        = 50     # Milliseconds between each key
    FocusDelay          = 200    # Milliseconds to wait after focusing (like autohonk)
}

# Add Windows Forms for keyboard handling
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Runtime.InteropServices

# Define Windows API functions
Add-Type @"
using System;
using System.Runtime.InteropServices;

public class WinAPI {
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
    
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
    
    [DllImport("user32.dll")]
    public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    
    [DllImport("user32.dll")]
    public static extern uint MapVirtualKey(uint uCode, uint uMapType);
    
    public const uint WM_KEYDOWN = 0x0100;
    public const uint WM_KEYUP = 0x0101;
    public const uint KEYEVENTF_KEYUP = 0x0002;
}
"@

function Get-EliteWindows {
    $windows = @()
    
    # Get all Elite Dangerous processes
    $processes = Get-Process -Name $config.ProcessName -ErrorAction SilentlyContinue
    
    foreach ($proc in $processes) {
        if ($proc.MainWindowHandle -ne 0) {
            $title = $proc.MainWindowTitle
            
            if ($title -like "*$($config.WindowTitleContains)*") {
                # Check for named commanders
                $commanderFound = $false
                foreach ($commander in $config.Commanders) {
                    if ($title -like "*$commander*") {
                        $windows += [PSCustomObject]@{
                            Handle    = $proc.MainWindowHandle
                            Title     = $title
                            Commander = $commander
                        }
                        $commanderFound = $true
                        break
                    }
                }
                
                # If no named commander found, check if it's primary
                if (-not $commanderFound) {
                    $hasOtherCommander = $false
                    foreach ($commander in $config.Commanders) {
                        if ($title -like "*$commander*") {
                            $hasOtherCommander = $true
                            break
                        }
                    }
                    
                    if (-not $hasOtherCommander) {
                        $windows += [PSCustomObject]@{
                            Handle    = $proc.MainWindowHandle
                            Title     = $title
                            Commander = $config.PrimaryCommander
                        }
                    }
                }
            }
        }
    }
    
    return $windows
}

function Get-VirtualKeyCode {
    param([char]$Char)
    
    # Special keys
    $specialKeys = @{
        ' '  = 0x20  # VK_SPACE
        "`n" = 0x0D  # VK_RETURN
        "`r" = 0x0D  # VK_RETURN
        "`t" = 0x09  # VK_TAB
    }
    
    if ($specialKeys.ContainsKey($Char)) {
        return $specialKeys[$Char]
    }
    elseif ($Char -match '[a-zA-Z0-9]') {
        return [byte][char]$Char.ToString().ToUpper()
    }
    else {
        # Try to convert the character
        try {
            return [byte][char]$Char.ToString().ToUpper()
        }
        catch {
            return $null
        }
    }
}

function Send-KeyToWindow {
    param(
        [IntPtr]$WindowHandle,
        [char]$Char,
        [string]$Commander,
        [switch]$UseDirect
    )
    
    $vkCode = Get-VirtualKeyCode -Char $Char
    
    if ($null -eq $vkCode) {
        Write-Host "  ⚠️  Unsupported character: '$Char'" -ForegroundColor Yellow
        return $false
    }
    
    try {
        if ($UseDirect) {
            # Use PostMessage - doesn't require focus
            $scanCode = [WinAPI]::MapVirtualKey($vkCode, 0)
            $lparamDown = 1 -bor ($scanCode -shl 16)
            $lparamUp = 1 -bor ($scanCode -shl 16) -bor (1 -shl 30) -bor (1 -shl 31)
            
            [WinAPI]::PostMessage($WindowHandle, [WinAPI]::WM_KEYDOWN, [IntPtr]$vkCode, [IntPtr]$lparamDown) | Out-Null
            Start-Sleep -Milliseconds 20
            [WinAPI]::PostMessage($WindowHandle, [WinAPI]::WM_KEYUP, [IntPtr]$vkCode, [IntPtr]$lparamUp) | Out-Null
        }
        else {
            # Use keybd_event - requires focus (like autohonk)
            [WinAPI]::SetForegroundWindow($WindowHandle) | Out-Null
            Start-Sleep -Milliseconds $config.FocusDelay
            
            [WinAPI]::keybd_event($vkCode, 0, 0, [UIntPtr]::Zero)
            Start-Sleep -Milliseconds 10
            [WinAPI]::keybd_event($vkCode, 0, [WinAPI]::KEYEVENTF_KEYUP, [UIntPtr]::Zero)
        }
        
        return $true
    }
    catch {
        Write-Host "  ❌ Error sending '$Char' to $Commander`: $_" -ForegroundColor Red
        return $false
    }
}

function Send-CommandToWindows {
    param(
        [string]$Command,
        [switch]$UseDirect
    )
    
    if ([string]::IsNullOrWhiteSpace($Command)) {
        return
    }
    
    Write-Host "`n🎯 Broadcasting: '$Command'" -ForegroundColor Cyan
    
    # Find Elite windows
    $windows = Get-EliteWindows
    
    if ($windows.Count -eq 0) {
        Write-Host "⚠️  No Elite Dangerous windows found!" -ForegroundColor Yellow
        return
    }
    
    Write-Host "📡 Found $($windows.Count) window(s)" -ForegroundColor Green
    foreach ($win in $windows) {
        Write-Host "   • $($win.Commander)" -ForegroundColor Gray
    }
    
    $method = if ($UseDirect) { "PostMessage" } else { "keybd_event" }
    Write-Host "📤 Method: $method" -ForegroundColor Cyan
    
    # Send to each window
    $successCount = 0
    foreach ($win in $windows) {
        Write-Host "   $($win.Commander)..." -NoNewline
        
        $charSuccess = 0
        foreach ($char in $Command.ToCharArray()) {
            if (Send-KeyToWindow -WindowHandle $win.Handle -Char $char -Commander $win.Commander -UseDirect:$UseDirect) {
                $charSuccess++
            }
            
            Start-Sleep -Milliseconds $config.KeySendDelay
        }
        
        Write-Host " ✅ ($charSuccess/$($Command.Length))" -ForegroundColor Green
        $successCount++
        
        Start-Sleep -Milliseconds 100
    }
    
    Write-Host "🎉 Sent to $successCount/$($windows.Count) windows" -ForegroundColor Green
    Write-Host ("-" * 50) -ForegroundColor DarkGray
    Write-Host "Ready for next command..." -ForegroundColor Cyan
}

# Main script
Write-Host "=" * 70 -ForegroundColor Cyan
Write-Host "Elite Dangerous Input Broadcaster - PowerShell Edition" -ForegroundColor Cyan
Write-Host "=" * 70 -ForegroundColor Cyan
Write-Host ""
Write-Host "INSTRUCTIONS:" -ForegroundColor Yellow
Write-Host "1. Keep this window focused"
Write-Host "2. Type your command (e.g., '1 q q d ')"
Write-Host "3. Wait $($config.TypingTimeout) second(s) - command broadcasts automatically"
Write-Host "4. Press Ctrl+C to exit"
Write-Host ""
Write-Host "MODES:" -ForegroundColor Yellow
Write-Host "  -UseDirect    : Use PostMessage (no focus required)"
Write-Host "  [no param]    : Use keybd_event (focus each window, like autohonk)"
Write-Host ""

# Check for -UseDirect parameter
$useDirect = $args -contains "-UseDirect"
if ($useDirect) {
    Write-Host "🔧 Using PostMessage (direct)" -ForegroundColor Green
}
else {
    Write-Host "🔧 Using keybd_event (focus mode, like autohonk)" -ForegroundColor Green
}

Write-Host ("-" * 70) -ForegroundColor DarkGray

# Test Elite window detection
Write-Host "🔍 Scanning for Elite windows..." -ForegroundColor Cyan
$windows = Get-EliteWindows
if ($windows.Count -gt 0) {
    Write-Host "✅ Found $($windows.Count) window(s):" -ForegroundColor Green
    foreach ($win in $windows) {
        Write-Host "   • $($win.Commander)" -ForegroundColor Gray
    }
}
else {
    Write-Host "⚠️  No Elite windows found" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Ready for input..." -ForegroundColor Green
Write-Host ""

# Input loop
$commandBuffer = ""
$lastKeypressTime = [DateTime]::MinValue

try {
    while ($true) {
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            
            # Handle Ctrl+C
            if ($key.Key -eq 'C' -and $key.Modifiers -eq 'Control') {
                Write-Host "`n🛑 Exiting..." -ForegroundColor Yellow
                break
            }
            
            # Handle Backspace
            if ($key.Key -eq 'Backspace') {
                if ($commandBuffer.Length -gt 0) {
                    $commandBuffer = $commandBuffer.Substring(0, $commandBuffer.Length - 1)
                    Write-Host "`rCommand: '$commandBuffer'$((' ' * 20))" -NoNewline
                }
                $lastKeypressTime = Get-Date
                continue
            }
            
            # Handle Enter
            if ($key.Key -eq 'Enter') {
                $commandBuffer += "`n"
            }
            else {
                $commandBuffer += $key.KeyChar
            }
            
            Write-Host "`rCommand: '$commandBuffer'" -NoNewline
            $lastKeypressTime = Get-Date
        }
        
        # Check if we should send the command
        if ($commandBuffer.Length -gt 0) {
            $elapsed = (Get-Date) - $lastKeypressTime
            if ($elapsed.TotalSeconds -ge $config.TypingTimeout) {
                $toSend = $commandBuffer
                $commandBuffer = ""
                $lastKeypressTime = [DateTime]::MinValue
                
                Write-Host ""
                Send-CommandToWindows -Command $toSend -UseDirect:$useDirect
            }
        }
        
        Start-Sleep -Milliseconds 100
    }
}
catch {
    Write-Host "`n❌ Error: $_" -ForegroundColor Red
}
finally {
    Write-Host "`n👋 Goodbye!" -ForegroundColor Cyan
}
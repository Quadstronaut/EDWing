#AutoIt3Wrapper_Icon=E:\Elite Dangerous.ico
#AutoIt3Wrapper_Compression=9

# ====================================================================================================================
# SCRIPT CONFIGURATION
#
# Enter your four commander names here. These names must match the window titles of your Elite Dangerous clients.
# AutoIt will search for windows with titles containing these names.
# ====================================================================================================================
Local $aCmdrNames[4] = ["CMDR Alice", "CMDR Bob", "CMDR Charlie", "CMDR Dave"]

; The process name for the Elite Dangerous window.
Local $sProcessName = "EliteDangerous64.exe"

; The key that, when pressed, will broadcast to all windows.
; We'll use the 'Numpad 1' key as the default broadcast key.
; You can change this to any key you want, for example: "{F1}", "a", etc.
Local $sBroadcastKey = "{Numpad1}"

; ====================================================================================================================
; SCRIPT LOGIC
; ====================================================================================================================

; Keep track of the window handles for our four Elite Dangerous clients.
Local $aWindowHandles[4]
Local $iFoundWindows = 0

; Loop until all four windows are found.
While $iFoundWindows < 4
    For $i = 0 To UBound($aCmdrNames) - 1
        ; Find the window by process name and partial title (the commander's name).
        ; We're using a partial match because the full window title might change.
        Local $hWnd = WinGetHandle($sProcessName, $aCmdrNames[$i])
        
        If $hWnd <> "" Then
            ; Check if we've already found this window.
            If Not IsDeclared("$aWindowHandles[" & $i & "]") Or $aWindowHandles[$i] <> $hWnd Then
                $aWindowHandles[$i] = $hWnd
                $iFoundWindows += 1
                ConsoleWrite("Found window for " & $aCmdrNames[$i] & " with handle " & $hWnd & @CRLF)
            EndIf
        EndIf
    Next
    Sleep(500) ; Wait 500ms before checking again.
WEnd

ConsoleWrite("All four Elite Dangerous windows found. Starting key broadcast..." & @CRLF)

; Main loop to monitor for key presses and broadcast them.
While True
    ; Check if the broadcast key is pressed.
    ; This script will only function when the specified key is pressed down.
    If _IsPressed($sBroadcastKey) Then
        ; Get the currently active window's handle.
        Local $hActiveWindow = WinGetHandle("[ACTIVE]")

        ; If the active window is one of our Elite Dangerous clients, proceed with the broadcast.
        If IsInArray($hActiveWindow, $aWindowHandles) Then
            ; Get the key that was pressed.
            Local $sKey = _GetPressedKey()

            ; Broadcast the key to all four windows.
            For $i = 0 To UBound($aWindowHandles) - 1
                Local $hWnd = $aWindowHandles[$i]
                
                ; Send the key only if the window is not the active window.
                ; This prevents double input on the main window.
                If $hWnd <> $hActiveWindow Then
                    ControlSend($hWnd, "", "", $sKey)
                    ConsoleWrite("Broadcasting '" & $sKey & "' to window " & $hWnd & @CRLF)
                EndIf
            Next
        EndIf
    EndIf
    
    Sleep(10) ; Small delay to prevent high CPU usage.
WEnd

; ====================================================================================================================
; HELPER FUNCTIONS
; ====================================================================================================================

; A helper function to check if a key is currently pressed.
Func _IsPressed($sKey)
    Local $iKey = Asc(StringLeft($sKey, 1))
    Return BitAND(DllStructGetData(DllCall("user32.dll", "int", "GetKeyState", "int", $iKey), 1), 0x8000)
EndFunc

; A helper function to get the key code of the key that was pressed.
Func _GetPressedKey()
    Local $aKeys = "1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ!@#$%^&*()_+=-`~,.<>/?;:'""[{]}\|"
    For $i = 1 To StringLen($aKeys)
        Local $sKey = StringMid($aKeys, $i, 1)
        If _IsPressed($sKey) Then
            Return $sKey
        EndIf
    Next
    Return ""
EndFunc

; A helper function to check if a value exists in an array.
Func IsInArray($sValue, $aArray)
    For $i = 0 To UBound($aArray) - 1
        If $aArray[$i] = $sValue Then
            Return True
        EndIf
    Next
    Return False
EndFunc
Attribute VB_Name = "modKeyStates"
'Author:  David Zimmer <dzzie@yahoo.com>
'AI:      Claude.ai
'Site:    http://sandsprite.com
'License: MIT
'---------------------------------------------------


'=========================================================================
' Keyboard and Win32 Helper Functions - Refactored
' Professional naming and cleaner implementation
'=========================================================================
Option Explicit

' Win32 API
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer


' Win32 API for UTF-8 conversion
Declare Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal codePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cchMultiByte As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long) As Long

Declare Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal codePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cchMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long) As Long

Global Const CP_UTF8 As Long = 65001

' Virtual key codes
Private Const VK_SHIFT = &H10
Private Const VK_CONTROL = &H11
Private Const VK_MENU = &H12      ' ALT key

' Modifier key flags (matches Windows convention)
Public Enum ModifierKeys
    MOD_NONE = 0
    MOD_SHIFT = 1
    MOD_ALT = 2
    MOD_CONTROL = 4
    MOD_SHIFT_CTRL = 5         ' Shift + Ctrl
    MOD_SHIFT_ALT = 3          ' Shift + Alt
    MOD_CTRL_ALT = 6           ' Ctrl + Alt
    MOD_ALL = 7                ' Shift + Ctrl + Alt
End Enum

'=========================================================================
' UTF-8 Conversion Helpers
'=========================================================================

' Convert UTF-8 byte array to VB6 Unicode string
Function UTF8BytesToString(ByRef bytes() As Byte, ByVal byteCount As Long) As String
    Dim wideChars As Long
    Dim result As String
    
    If byteCount = 0 Then
        UTF8BytesToString = ""
        Exit Function
    End If
    
    ' Get required buffer size for wide chars
    wideChars = MultiByteToWideChar(CP_UTF8, 0, VarPtr(bytes(0)), byteCount, 0, 0)
    
    If wideChars = 0 Then
        ' Conversion failed - return empty string
        UTF8BytesToString = ""
        Exit Function
    End If
    
    ' Allocate string buffer
    result = String$(wideChars, 0)
    
    ' Convert UTF-8 to Unicode
    wideChars = MultiByteToWideChar(CP_UTF8, 0, VarPtr(bytes(0)), byteCount, StrPtr(result), wideChars)
    
    UTF8BytesToString = result
End Function

' Convert VB6 Unicode string to UTF-8 byte array
Function StringToUTF8Bytes(ByVal str As String) As Byte()
    Dim utf8Len As Long
    Dim result() As Byte
    
    If Len(str) = 0 Then
        ReDim result(0 To 0)
        result(0) = 0  ' Null terminator
        StringToUTF8Bytes = result
        Exit Function
    End If
    
    ' Get required buffer size for UTF-8
    utf8Len = WideCharToMultiByte(CP_UTF8, 0, StrPtr(str), Len(str), 0, 0, 0, 0)
    
    If utf8Len = 0 Then
        ReDim result(0 To 0)
        result(0) = 0
        StringToUTF8Bytes = result
        Exit Function
    End If
    
    ' Allocate byte buffer (with null terminator)
    ReDim result(0 To utf8Len) As Byte
    
    ' Convert to UTF-8
    utf8Len = WideCharToMultiByte(CP_UTF8, 0, StrPtr(str), Len(str), VarPtr(result(0)), utf8Len, 0, 0)
    result(utf8Len) = 0  ' Add null terminator
    
    StringToUTF8Bytes = result
End Function


'=========================================================================
' Keyboard State Functions
'=========================================================================

' Get current modifier key state as a bitmask
' Returns: Combination of MOD_SHIFT (1), MOD_ALT (2), MOD_CONTROL (4)
Public Function GetModifierState() As Long
    Dim modifiers As Long
    
    modifiers = 0
    
    If IsKeyPressed(VK_SHIFT) Then
        modifiers = modifiers Or MOD_SHIFT
    End If
    
    If IsKeyPressed(VK_MENU) Then   ' ALT key
        modifiers = modifiers Or MOD_ALT
    End If
    
    If IsKeyPressed(VK_CONTROL) Then
        modifiers = modifiers Or MOD_CONTROL
    End If
    
    GetModifierState = modifiers
End Function


' Check if a specific key is currently pressed
' Uses GetAsyncKeyState to check real-time key state
Public Function IsKeyPressed(ByVal virtualKeyCode As Long) As Boolean
    Dim keyState As Integer
    
    keyState = GetAsyncKeyState(virtualKeyCode)
    
    ' High-order bit set means key is currently pressed
    IsKeyPressed = (keyState And &H8000&) <> 0
End Function

' Check if ONLY a specific modifier is pressed (no others)
Public Function IsModifierOnly(ByVal modifier As ModifierKeys) As Boolean
    IsModifierOnly = (GetModifierState() = modifier)
End Function

' Check if a specific modifier is pressed (possibly with others)
Public Function IsModifierPressed(ByVal modifier As ModifierKeys) As Boolean
    IsModifierPressed = (GetModifierState() And modifier) <> 0
End Function

'=========================================================================
' Convenience Functions for Common Key Combinations
'=========================================================================

Public Function IsCtrlPressed() As Boolean
    IsCtrlPressed = IsKeyPressed(VK_CONTROL)
End Function

Public Function IsShiftPressed() As Boolean
    IsShiftPressed = IsKeyPressed(VK_SHIFT)
End Function

Public Function IsAltPressed() As Boolean
    IsAltPressed = IsKeyPressed(VK_MENU)
End Function

' Check for specific combinations
Public Function IsCtrlShiftPressed() As Boolean
    IsCtrlShiftPressed = IsModifierOnly(MOD_SHIFT_CTRL)
End Function

Public Function IsCtrlOnly() As Boolean
    IsCtrlOnly = IsModifierOnly(MOD_CONTROL)
End Function

Public Function IsShiftOnly() As Boolean
    IsShiftOnly = IsModifierOnly(MOD_SHIFT)
End Function

'=========================================================================
' Win32 LPARAM/WPARAM Helper Functions
'=========================================================================

' Extract high and low words from a Long value
' Used for parsing Win32 message parameters
Public Sub SplitLong(ByVal value As Long, ByRef hiWord As Long, ByRef loWord As Long)
    ' Extract high word (upper 16 bits)
    hiWord = (value And &HFFFF0000) \ &H10000
    
    ' Extract low word (lower 16 bits)
    loWord = value And &HFFFF&
    
    ' Handle sign extension for negative low word
    If loWord And &H8000& Then
        loWord = loWord Or &HFFFF0000
    End If
End Sub

' Alternative: Return as separate values
Public Function GetHiWord(ByVal value As Long) As Long
    GetHiWord = (value And &HFFFF0000) \ &H10000
End Function

Public Function GetLoWord(ByVal value As Long) As Long
    GetLoWord = value And &HFFFF&
    If GetLoWord And &H8000& Then
        GetLoWord = GetLoWord Or &HFFFF0000
    End If
End Function

' Create a Long from two Words (useful for constructing lParam/wParam)
Public Function MakeLong(ByVal loWord As Integer, ByVal hiWord As Integer) As Long
    MakeLong = (CLng(hiWord) * &H10000) Or (loWord And &HFFFF&)
End Function

'=========================================================================
' Usage Examples
'=========================================================================

' Example 1: Check for Ctrl+Space in WM_CHAR handler
' Old way:
'   If wParam = 32 And piGetShiftState = 4 Then
'
' New way:
'   If wParam = 32 And IsCtrlOnly() Then

' Example 2: Check for Ctrl+Shift+S
' Old way:
'   If piGetShiftState = 5 Then ' 1 (shift) + 4 (ctrl)
'
' New way:
'   If IsModifierOnly(MOD_SHIFT_CTRL) Then
'   Or: If IsCtrlShiftPressed() Then

' Example 3: Check if Ctrl is pressed (regardless of other modifiers)
' Old way:
'   If (piGetShiftState And 4) Then
'
' New way:
'   If IsCtrlPressed() Then
'   Or: If IsModifierPressed(MOD_CONTROL) Then

' Example 4: Parse mouse coordinates from lParam
' Old way:
'   pGetHiWordLoWord lParam, y, x
'
' New way:
'   SplitLong lParam, y, x
'   Or: x = GetLoWord(lParam): y = GetHiWord(lParam)

' Example 5: In your Scintilla subclass handler
'   Case WM_CHAR
'       If wParam = 32 And IsCtrlOnly() Then
'           ' Ctrl+Space pressed
'           bHandled = True
'           lReturn = 0
'           RaiseEvent AutoCompleteEvent(Helper.WordAtCaret())
'       End If
'
'   Case WM_KEYDOWN
'       Select Case wParam
'           Case vbKeyS
'               If IsCtrlOnly() Then
'                   ' Ctrl+S - Save
'               ElseIf IsCtrlShiftPressed() Then
'                   ' Ctrl+Shift+S - Save As
'               End If
'       End Select

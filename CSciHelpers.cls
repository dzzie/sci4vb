VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSciHelpers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Author:  David Zimmer <dzzie@yahoo.com>
'AI:      Claude.ai
'Site:    http://sandsprite.com
'License: MIT
'---------------------------------------------------



'=========================================================================
' CSciHelpers - Editor utility and helper methods
'=========================================================================
Option Explicit

Private m_Sci As SciWrapper

' Scintilla messages
Private Const SCI_WORDSTARTPOSITION = 2266
Private Const SCI_WORDENDPOSITION = 2267
Private Const SCI_POSITIONFROMPOINT = 2022
Private Const SCI_GETLINEINDENTATION = 2127
Private Const SCI_SETLINEINDENTATION = 2126
Private Const SCI_GETLINEINDENTPOSITION = 2128
Private Const SCI_BRACEHIGHLIGHT = 2351
Private Const SCI_BRACEBADLIGHT = 2352

' Win32 API for mouse and window operations
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Friend Sub Init(sciCtrl As SciWrapper)
    Set m_Sci = sciCtrl
End Sub


'=========================================================================
' Word Operations
'=========================================================================

' Get the word under the mouse cursor at a specific position
Public Function WordUnderMouse(ByVal pos As Long, Optional ByVal ignoreWhiteSpace As Boolean = False) As String
    Dim wordStart As Long
    Dim wordEnd As Long
    
    On Error Resume Next
    
    ' If ignoring whitespace, adjust position back to last non-space character
    If ignoreWhiteSpace Then
        pos = pos - GetSpaceCountBack(pos)
    End If
    
    ' Get word boundaries
    wordStart = m_Sci.SciMsg(SCI_WORDSTARTPOSITION, pos, 1)  ' onlyWordCharacters = True
    wordEnd = m_Sci.SciMsg(SCI_WORDENDPOSITION, pos, 1)
    
    ' Extract word
    If wordEnd > wordStart Then
        Dim text As String
        text = m_Sci.doc.text
        WordUnderMouse = Mid$(text, wordStart + 1, wordEnd - wordStart)
    Else
        WordUnderMouse = ""
    End If
End Function

' Get the word at the current caret position
Public Function WordAtCaret(Optional ByVal ignoreWhiteSpace As Boolean = False) As String
    WordAtCaret = WordUnderMouse(m_Sci.sel.currentPos, ignoreWhiteSpace)
End Function

' Get the previous word before the caret

' Get word at mouse coordinates (screen coordinates)
Public Function WordAtMouseCoords(ByVal screenX As Long, ByVal screenY As Long) As String
    Dim pos As Long
    pos = m_Sci.SciMsg(SCI_POSITIONFROMPOINT, screenX, screenY)
    WordAtMouseCoords = WordUnderMouse(pos)
End Function

'=========================================================================
' Indentation Helpers
'=========================================================================

' Maintain indentation when pressing Enter - matches previous line's indentation
Public Sub MaintainIndent()
    Dim currentLine As Long
    Dim previousLine As Long
    Dim indentAmount As Long
    Dim lineText As String
    
    On Error Resume Next
    
    currentLine = m_Sci.lines.LineFromPosition(m_Sci.sel.currentPos)
    
    If currentLine = 0 Then Exit Sub
    
    ' Find previous non-empty line
    previousLine = currentLine - 1
    Do While previousLine >= 0
        lineText = m_Sci.lines.GetLine(previousLine)
        
        ' Skip empty lines
        If Len(Trim$(lineText)) > 0 Then
            ' Get indentation of previous non-empty line
            indentAmount = m_Sci.lines.GetLineIndentation(previousLine)
            
            ' Apply to current line
            If indentAmount > 0 Then
                m_Sci.lines.SetLineIndentation currentLine, indentAmount
                
                ' Move caret to end of indentation
                Dim indentPos As Long
                indentPos = m_Sci.lines.GetLineIndentPosition(currentLine)
                m_Sci.sel.currentPos = indentPos
                m_Sci.sel.anchor = indentPos
            End If
            Exit Sub
        End If
        
        previousLine = previousLine - 1
    Loop
End Sub

' Smart indent - increases indent if previous line suggests it
Public Sub SmartIndent()
    Dim currentLine As Long
    Dim previousLine As Long
    Dim indentAmount As Long
    Dim lineText As String
    Dim trimmedText As String
    
    On Error Resume Next
    
    currentLine = m_Sci.lines.LineFromPosition(m_Sci.sel.currentPos)
    
    If currentLine = 0 Then Exit Sub
    
    previousLine = currentLine - 1
    lineText = m_Sci.lines.GetLine(previousLine)
    trimmedText = Trim$(lineText)
    indentAmount = m_Sci.lines.GetLineIndentation(previousLine)
    
    ' Check if previous line should increase indent
    Dim shouldIncreaseIndent As Boolean
    shouldIncreaseIndent = False
    
    ' Check for common indent triggers
    If Len(trimmedText) > 0 Then
        Dim lastChar As String
        lastChar = Right$(trimmedText, 1)
        
        ' VB-style indent triggers
        If Right$(trimmedText, 4) = "Then" Or _
           InStr(1, trimmedText, "Sub ", vbTextCompare) > 0 Or _
           InStr(1, trimmedText, "Function ", vbTextCompare) > 0 Or _
           InStr(1, trimmedText, "Property ", vbTextCompare) > 0 Or _
           InStr(1, trimmedText, "If ", vbTextCompare) > 0 Or _
           InStr(1, trimmedText, "For ", vbTextCompare) > 0 Or _
           InStr(1, trimmedText, "While ", vbTextCompare) > 0 Or _
           InStr(1, trimmedText, "Do", vbTextCompare) > 0 Or _
           InStr(1, trimmedText, "Select Case", vbTextCompare) > 0 Then
            shouldIncreaseIndent = True
        End If
        
        ' C-style indent triggers
        If lastChar = "{" Or lastChar = ":" Then
            shouldIncreaseIndent = True
        End If
    End If
    
    ' Apply indentation
    If shouldIncreaseIndent Then
        indentAmount = indentAmount + m_Sci.edit.TabWidth
    End If
    
    If indentAmount > 0 Then
        m_Sci.lines.SetLineIndentation currentLine, indentAmount
        
        ' Move caret to end of indentation
        Dim indentPos As Long
        indentPos = m_Sci.lines.GetLineIndentPosition(currentLine)
        m_Sci.sel.currentPos = indentPos
        m_Sci.sel.anchor = indentPos
    End If
End Sub

'=========================================================================
' Bracket/Brace Highlighting
'=========================================================================

' Highlight matching bracket/brace at current position
Public Sub HighlightMatchingBrace()
    Dim pos As Long
    Dim ch As String
    Dim matchPos As Long
    
    pos = m_Sci.sel.currentPos
    
    ' Check character before caret
    If pos > 0 Then
        ch = Mid$(m_Sci.doc.text, pos, 1)
        
        If InStr("()[]{}""", ch) > 0 Then
            matchPos = FindMatchingBrace(pos - 1)
            
            If matchPos >= 0 Then
                m_Sci.SciMsg SCI_BRACEHIGHLIGHT, pos - 1, matchPos
            Else
                m_Sci.SciMsg SCI_BRACEBADLIGHT, pos - 1
            End If
            Exit Sub
        End If
    End If
    
    ' Check character at caret
    If pos < m_Sci.doc.TextLength Then
        ch = Mid$(m_Sci.doc.text, pos + 1, 1)
        
        If InStr("()[]{}""", ch) > 0 Then
            matchPos = FindMatchingBrace(pos)
            
            If matchPos >= 0 Then
                m_Sci.SciMsg SCI_BRACEHIGHLIGHT, pos, matchPos
            Else
                m_Sci.SciMsg SCI_BRACEBADLIGHT, pos
            End If
        End If
    End If
End Sub

' Highlight a specific word throughout the document
Public Sub HighlightWord(ByVal word As String, ByVal indicatorNum As Long)
    Dim pos As Long
    Dim wordLen As Long
    
    If Len(word) = 0 Then Exit Sub
    
    ' Clear previous highlights
    m_Sci.Indic.Current = indicatorNum
    m_Sci.Indic.ClearRange 0, m_Sci.doc.TextLength
    
    ' Search for all occurrences using the fixed Search.Find
    wordLen = Len(word)
    pos = 0
    
    Do
        pos = m_Sci.Search.Find(word, pos, False, True, False, False)  ' wholeWord = True
        
        If pos >= 0 Then
            ' Highlight this occurrence
            m_Sci.Indic.Current = indicatorNum
            m_Sci.Indic.FillRange pos, wordLen
            pos = pos + wordLen
        End If
    Loop While pos >= 0 And pos < m_Sci.doc.TextLength
End Sub

' Clear word highlighting
Public Sub ClearWordHighlight(ByVal indicatorNum As Long)
    m_Sci.Indic.Current = indicatorNum
    m_Sci.Indic.ClearRange 0, m_Sci.doc.TextLength
End Sub

'=========================================================================
' Call Tip Detection
'=========================================================================

' Check if mouse is currently over a call tip window
Public Function IsMouseOverCallTip() As Boolean
    Dim pt As POINTAPI
    Dim hwnd As Long
    Dim className As String
    Dim classLen As Long
    
    On Error Resume Next
    
    ' Get cursor position
    GetCursorPos pt
    
    ' Get window under cursor
    hwnd = WindowFromPoint(pt.X, pt.Y)
    
    If hwnd <> 0 Then
        ' Get window class name
        className = String$(100, 0)
        classLen = GetClassName(hwnd, className, 100)
        
        If classLen > 0 Then
            className = Left$(className, classLen)
            IsMouseOverCallTip = (className = "CallTip")
        End If
    End If
End Function

'=========================================================================
' Private Helper Functions
'=========================================================================

' Find matching bracket/brace for the character at position
Private Function FindMatchingBrace(ByVal pos As Long) As Long
    Dim text As String
    Dim ch As String
    Dim matchChar As String
    Dim direction As Long
    Dim depth As Long
    Dim i As Long
    Dim currentChar As String
    
    text = m_Sci.doc.text
    If pos < 0 Or pos >= Len(text) Then
        FindMatchingBrace = -1
        Exit Function
    End If
    
    ch = Mid$(text, pos + 1, 1)
    
    ' Determine direction and matching character
    Select Case ch
        Case "(": matchChar = ")": direction = 1
        Case ")": matchChar = "(": direction = -1
        Case "[": matchChar = "]": direction = 1
        Case "]": matchChar = "[": direction = -1
        Case "{": matchChar = "}": direction = 1
        Case "}": matchChar = "{": direction = -1
        Case Else
            FindMatchingBrace = -1
            Exit Function
    End Select
    
    ' Search for matching brace
    depth = 1
    i = pos + direction
    
    Do While i >= 0 And i < Len(text)
        currentChar = Mid$(text, i + 1, 1)
        
        If currentChar = ch Then
            depth = depth + 1
        ElseIf currentChar = matchChar Then
            depth = depth - 1
            If depth = 0 Then
                FindMatchingBrace = i
                Exit Function
            End If
        End If
        
        i = i + direction
    Loop
    
    FindMatchingBrace = -1
End Function

' Count spaces backwards from position
Private Function GetSpaceCountBack(ByVal pos As Long) As Long
    Dim line As Long
    Dim lineStart As Long
    Dim lineText As String
    Dim curCol As Long
    Dim i As Long
    Dim Count As Long
    Dim ch As String
    
    On Error Resume Next
    
    line = m_Sci.lines.LineFromPosition(pos)
    lineStart = m_Sci.lines.PositionFromLine(line)
    lineText = m_Sci.lines.GetLine(line)
    curCol = pos - lineStart
    
    If curCol <= 0 Then Exit Function
    
    ' Count spaces backwards
    For i = curCol To 1 Step -1
        ch = Mid$(lineText, i, 1)
        If ch = " " Or ch = vbTab Then
            Count = Count + 1
        Else
            Exit For
        End If
    Next i
    
    GetSpaceCountBack = Count
End Function

' Check if character is whitespace
Private Function IsWhiteSpace(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then
        IsWhiteSpace = False
    Else
        IsWhiteSpace = (ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf)
    End If
End Function

'=========================================================================
' Professional-grade Word Functions for CSciHelpers
' Uses Scintilla API properly - no manual string parsing
'=========================================================================

'-------------------------------------------------------------------------
' CurrentWord - Get word at caret with exact boundary positions
' This is what you should use instead of CurrentWordInternal
'-------------------------------------------------------------------------
Public Function currentWord(Optional ByRef iStart As Long, Optional ByRef iEnd As Long) As String
    Dim currentPos As Long
    Dim text As String
    
    On Error Resume Next
    
    currentPos = m_Sci.sel.currentPos
    
    ' Use Scintilla's word boundary detection
    ' onlyWordCharacters = True (1) means stop at any non-word character
    iStart = m_Sci.SciMsg(SCI_WORDSTARTPOSITION, currentPos, 1)
    iEnd = m_Sci.SciMsg(SCI_WORDENDPOSITION, currentPos, 1)
    
    ' Handle special case: cursor immediately after a dot (member access)
    ' Check if we're right after a '.' and before that is a valid word
    If iEnd = iStart And currentPos > 0 Then
        text = m_Sci.doc.text
        If Mid$(text, currentPos, 1) = "." Then
            ' We're at a dot - get the word before it
            iStart = m_Sci.SciMsg(SCI_WORDSTARTPOSITION, currentPos - 1, 1)
            iEnd = m_Sci.SciMsg(SCI_WORDENDPOSITION, currentPos - 1, 1)
        End If
    End If
    
    ' Extract the word text
    If iEnd > iStart Then
        text = m_Sci.doc.text
        currentWord = Mid$(text, iStart + 1, iEnd - iStart)
    Else
        currentWord = ""
    End If
End Function

'-------------------------------------------------------------------------
' PreviousWord - Improved version with better dot handling
' Gets the word immediately before the current word/position
'-------------------------------------------------------------------------
Public Function PreviousWord() As String
    Dim currentPos As Long
    Dim wordStart As Long
    Dim prevWordStart As Long
    Dim prevWordEnd As Long
    Dim text As String
    Dim ch As String
    
    On Error Resume Next
    
    currentPos = m_Sci.sel.currentPos
    
    ' Get start of current word (or position if not in a word)
    wordStart = m_Sci.SciMsg(SCI_WORDSTARTPOSITION, currentPos, 1)
    
    ' If we're in the middle of a word, wordStart < currentPos
    ' Move back to find the previous word
    If wordStart > 0 Then
        prevWordEnd = wordStart - 1
        
        ' Get the full text once for efficiency
        text = m_Sci.doc.text
        
        ' Skip backwards over whitespace and dots
        Do While prevWordEnd > 0
            ch = Mid$(text, prevWordEnd + 1, 1)
            If Not IsWhiteSpace(ch) And ch <> "." Then
                Exit Do
            End If
            prevWordEnd = prevWordEnd - 1
        Loop
        
        ' Now we should be at the end of the previous word
        If prevWordEnd > 0 Then
            ' Get the previous word boundaries using Scintilla
            prevWordStart = m_Sci.SciMsg(SCI_WORDSTARTPOSITION, prevWordEnd, 1)
            prevWordEnd = m_Sci.SciMsg(SCI_WORDENDPOSITION, prevWordEnd, 1)
            
            If prevWordEnd > prevWordStart Then
                PreviousWord = Mid$(text, prevWordStart + 1, prevWordEnd - prevWordStart)
            End If
        End If
    End If
End Function

'-------------------------------------------------------------------------
' GetWordBeforeDot - Specifically for autocomplete after typing '.'
' This is optimized for the "obj." autocomplete scenario
'-------------------------------------------------------------------------
Public Function GetWordBeforeDot() As String
    Dim currentPos As Long
    Dim checkPos As Long
    Dim text As String
    Dim ch As String
    Dim wordStart As Long
    Dim wordEnd As Long
    
    On Error Resume Next
    
    currentPos = m_Sci.sel.currentPos
    
    ' We need to check if there's a dot immediately before the cursor
    If currentPos < 1 Then
        GetWordBeforeDot = ""
        Exit Function
    End If
    
    text = m_Sci.doc.text
    
    ' Check the character right before cursor
    ch = Mid$(text, currentPos, 1)
    
    If ch = "." Then
        ' Found a dot! Get the word immediately before it
        checkPos = currentPos - 2  ' Position before the dot
        
        If checkPos >= 0 Then
            ' Get word boundaries for the position before the dot
            wordStart = m_Sci.SciMsg(SCI_WORDSTARTPOSITION, checkPos, 1)
            wordEnd = m_Sci.SciMsg(SCI_WORDENDPOSITION, checkPos, 1)
            
            ' Make sure the word actually touches the dot
            If wordEnd = currentPos - 1 And wordEnd > wordStart Then
                GetWordBeforeDot = Mid$(text, wordStart + 1, wordEnd - wordStart)
            End If
        End If
    End If
End Function

'-------------------------------------------------------------------------
' GetCurrentWordOrDotPrefix - For autocomplete triggers
' Returns either the current partial word or the word before a dot
' This handles both "obj." and "obj.mem" scenarios
'-------------------------------------------------------------------------
Public Function GetCurrentWordOrDotPrefix(Optional ByRef isDotCompletion As Boolean) As String
    Dim currentPos As Long
    Dim text As String
    Dim checkPos As Long
    Dim ch As String
    Dim wordStart As Long
    Dim wordEnd As Long
    Dim currentWord As String
    Dim i As Long
    
    On Error Resume Next
    
    currentPos = m_Sci.sel.currentPos
    text = m_Sci.doc.text
    isDotCompletion = False
    
    ' Get current word at cursor
    wordStart = m_Sci.SciMsg(SCI_WORDSTARTPOSITION, currentPos, 1)
    wordEnd = m_Sci.SciMsg(SCI_WORDENDPOSITION, currentPos, 1)
    
    If wordEnd > wordStart Then
        currentWord = Mid$(text, wordStart + 1, wordEnd - wordStart)
    End If
    
    ' Now check if there's a dot before the current word
    If wordStart > 0 Then
        ' Check character immediately before current word
        checkPos = wordStart - 1
        
        ' Skip whitespace backwards
        Do While checkPos > 0
            ch = Mid$(text, checkPos, 1)
            If Not IsWhiteSpace(ch) Then Exit Do
            checkPos = checkPos - 1
        Loop
        
        ' If we found a dot, this is a member access completion
        If checkPos > 0 And Mid$(text, checkPos, 1) = "." Then
            ' Get the word before the dot
            checkPos = checkPos - 1
            
            If checkPos >= 0 Then
                Dim beforeDotStart As Long
                Dim beforeDotEnd As Long
                
                beforeDotStart = m_Sci.SciMsg(SCI_WORDSTARTPOSITION, checkPos, 1)
                beforeDotEnd = m_Sci.SciMsg(SCI_WORDENDPOSITION, checkPos, 1)
                
                If beforeDotEnd > beforeDotStart Then
                    isDotCompletion = True
                    GetCurrentWordOrDotPrefix = Mid$(text, beforeDotStart + 1, beforeDotEnd - beforeDotStart)
                    Exit Function
                End If
            End If
        End If
    End If
    
    ' Not a dot completion - return current word
    GetCurrentWordOrDotPrefix = currentWord
End Function

'-------------------------------------------------------------------------
' Example usage in autocomplete event handler:
'-------------------------------------------------------------------------
' Case WM_KEYUP
'     If wParam = 190 Then ' period key
'         Dim className As String
'         className = Helper.GetWordBeforeDot()
'         If Len(className) > 0 Then
'             RaiseEvent AutoCompleteEvent(className)
'         End If
'     End If
'
' Or for CharAdded event (when they're typing):
'     Dim isDot As Boolean
'     Dim prefix As String
'     prefix = Helper.GetCurrentWordOrDotPrefix(isDot)
'     If isDot Then
'         ' Show members of 'prefix' object
'         RaiseEvent AutoCompleteEvent(prefix)
'     ElseIf Len(prefix) >= 2 Then
'         ' Show general autocomplete for partial word
'         RaiseEvent AutoCompleteEvent(prefix)
'     End If


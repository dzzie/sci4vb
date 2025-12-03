VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSciSelection"
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
' CSciSelection - Selection and caret management
'=========================================================================
Option Explicit

Private m_Sci As SciWrapper

' Scintilla messages
Private Const SCI_REPLACESEL = 2170
Private Const SCI_REPLACETARGET = 2194
Private Const SCI_SETTARGETSTART = 2190
Private Const SCI_SETTARGETEND = 2192
Private Const SCI_WORDSTARTPOSITION = 2266
Private Const SCI_WORDENDPOSITION = 2267

Private Const SCI_SETSEL = 2160
Private Const SCI_GETSELTEXT = 2161
Private Const SCI_GETSELECTIONSTART = 2143
Private Const SCI_SETSELECTIONSTART = 2142
Private Const SCI_GETSELECTIONEND = 2145
Private Const SCI_SETSELECTIONEND = 2144
Private Const SCI_SELECTALL = 2013
Private Const SCI_SETCURRENTPOS = 2141
Private Const SCI_GETCURRENTPOS = 2008
Private Const SCI_SETANCHOR = 2026
Private Const SCI_GETANCHOR = 2009
Private Const SCI_SETSELECTIONMODE = 2422
Private Const SCI_GETSELECTIONMODE = 2423
Private Const SCI_GETSELECTIONS = 2570
Private Const SCI_CLEARSELECTIONS = 2571
Private Const SCI_SETSELECTION = 2572
Private Const SCI_ADDSELECTION = 2573
Private Const SCI_SETMAINSELECTION = 2574
Private Const SCI_GETMAINSELECTION = 2575
Private Const SCI_SETSELECTIONNSTART = 2579
Private Const SCI_GETSELECTIONNSTART = 2585
Private Const SCI_SETSELECTIONNEND = 2580
Private Const SCI_GETSELECTIONNEND = 2586
Private Const SCI_GOTOPOS = 2025
Private Const SCI_GOTOLINE = 2024
Private Const SCI_SETXCARETPOLICY = 2402
Private Const SCI_SETYCARETPOLICY = 2403
Private Const SCI_SETEMPTYSELECTION = 2556

' Selection modes
Public Enum eSelectionMode
    selStream = 0
    selRectangle = 1
    selLines = 2
    selThin = 3
End Enum

' Caret policy flags
Public Enum eCaretPolicy
    caretSlop = &H1
    caretStrict = &H4
    caretJumps = &H10
    caretEven = &H8
End Enum

Friend Sub Init(sciCtrl As SciWrapper)
    Set m_Sci = sciCtrl
End Sub

' ========== Traditional VB-style Selection Properties ==========

Public Property Get SelStart() As Long
    ' Returns position of selection start (or caret if no selection)
    SelStart = m_Sci.SciMsg(SCI_GETSELECTIONSTART, 0, 0)
End Property

Public Property Let SelStart(ByVal pos As Long)
    ' Set selection start (anchor), keeping current position
    m_Sci.SciMsg SCI_SETSELECTIONSTART, pos, 0
End Property

Public Property Get SelLength() As Long
    ' Returns length of selection (0 if no selection)
    Dim startPos As Long, endPos As Long
    startPos = m_Sci.SciMsg(SCI_GETSELECTIONSTART, 0, 0)
    endPos = m_Sci.SciMsg(SCI_GETSELECTIONEND, 0, 0)
    SelLength = endPos - startPos
End Property

Public Property Let SelLength(ByVal length As Long)
    ' Set selection length from current anchor
    Dim startPos As Long
    startPos = m_Sci.SciMsg(SCI_GETSELECTIONSTART, 0, 0)
    m_Sci.SciMsg SCI_SETSELECTIONEND, startPos + length, 0
End Property

Public Property Get selText() As String
    ' Returns selected text (empty if no selection)
    Dim length As Long
    length = Me.SelLength
    
    If length = 0 Then
        selText = ""
        Exit Property
    End If
    
    Dim buffer As String
    buffer = String$(length + 1, vbNullChar)
    m_Sci.SciMsgStr SCI_GETSELTEXT, 0, buffer
    selText = Left$(buffer, length)
End Property

Public Property Let selText(ByVal text As String)
    ' Replace selection with text
    m_Sci.SciMsgStr SCI_REPLACESEL, 0, text
End Property

' ========== Convenience Methods ==========

Public Sub SelectRange(ByVal startPos As Long, ByVal length As Long)
    ' Select text from startPos for length characters
    m_Sci.SciMsg SCI_SETSEL, startPos, startPos + length
End Sub

Public Sub ClearSelection()
    ' Remove selection (collapse to current position)
    Dim pos As Long
    pos = m_Sci.SciMsg(SCI_GETCURRENTPOS, 0, 0)
    m_Sci.SciMsg SCI_SETSEL, pos, pos
End Sub

'=========================================================================
' Selection Range
'=========================================================================
Public Sub SetSelection(ByVal anchor As Long, ByVal caret As Long)
    m_Sci.SciMsg SCI_SETSEL, anchor, caret
End Sub

Public Property Get SelectionStart() As Long
    SelectionStart = m_Sci.SciMsg(SCI_GETSELECTIONSTART)
End Property

Public Property Let SelectionStart(ByVal position As Long)
    m_Sci.SciMsg SCI_SETSELECTIONSTART, position
End Property

Public Property Get SelectionEnd() As Long
    SelectionEnd = m_Sci.SciMsg(SCI_GETSELECTIONEND)
End Property

Public Property Let SelectionEnd(ByVal position As Long)
    m_Sci.SciMsg SCI_SETSELECTIONEND, position
End Property

Public Sub SelectAll()
    m_Sci.SciMsg SCI_SELECTALL
End Sub

Public Function GetSelectedText() As String
    Dim length As Long
    Dim buffer() As Byte
    Dim wideBuffer As String
    Dim wideLen As Long
    
    length = SelectionEnd - SelectionStart
    If length <= 0 Then
        GetSelectedText = ""
        Exit Function
    End If
    
    ' Allocate byte array for UTF-8 data from Scintilla
    ReDim buffer(0 To length) As Byte
    m_Sci.SciMsgPtr SCI_GETSELTEXT, 0, VarPtr(buffer(0))
    
    ' Calculate required buffer size for wide string
    wideLen = MultiByteToWideChar(CP_UTF8, 0, VarPtr(buffer(0)), length, 0, 0)
    
    If wideLen > 0 Then
        ' Allocate wide string buffer
        wideBuffer = String$(wideLen, 0)
        
        ' Convert UTF-8 to Unicode
        MultiByteToWideChar CP_UTF8, 0, VarPtr(buffer(0)), length, StrPtr(wideBuffer), wideLen
        
        GetSelectedText = wideBuffer
    Else
        GetSelectedText = ""
    End If
End Function

'=========================================================================
' Text Replacement
'=========================================================================

' Replace the current selection with new text
Public Sub ReplaceSelection(ByVal text As String)
    Dim utf8Bytes() As Byte
    If Len(text) = 0 Then
        m_Sci.SciMsg SCI_REPLACESEL, 0, 0
    Else
        utf8Bytes = StrConv(text, vbFromUnicode)
        m_Sci.SciMsgPtr SCI_REPLACESEL, 0, VarPtr(utf8Bytes(0))
    End If
End Sub

' Replace text in a specific range (more efficient)
Public Sub ReplaceRange(ByVal startPos As Long, ByVal endPos As Long, ByVal text As String)
    Dim utf8Bytes() As Byte
    
    m_Sci.SciMsg SCI_SETTARGETSTART, startPos
    m_Sci.SciMsg SCI_SETTARGETEND, endPos
    
    If Len(text) = 0 Then
        m_Sci.SciMsg SCI_REPLACETARGET, 0, 0
    Else
        utf8Bytes = StrConv(text, vbFromUnicode)
        m_Sci.SciMsgPtr SCI_REPLACETARGET, UBound(utf8Bytes) + 1, VarPtr(utf8Bytes(0))
    End If
End Sub

' Convenience property for selected text (read/write)
Public Property Get text() As String
    text = GetSelectedText()
End Property

Public Property Let text(ByVal newText As String)
    ReplaceSelection newText
End Property

' Get length of selection
Public Property Get length() As Long
    length = SelectionEnd - SelectionStart
End Property

' Check if there's a selection
Public Property Get HasSelection() As Boolean
    HasSelection = (SelectionEnd > SelectionStart)
End Property

' Select word at caret
Public Sub SelectWord()
    Dim pos As Long, wordStart As Long, wordEnd As Long
    pos = currentPos
    wordStart = m_Sci.SciMsg(SCI_WORDSTARTPOSITION, pos, 1)
    wordEnd = m_Sci.SciMsg(SCI_WORDENDPOSITION, pos, 1)
    If wordEnd > wordStart Then SetSelection wordStart, wordEnd
End Sub

Public Function SelectLine() As Long
  Dim curLine As Long
  curLine = currentLine
  m_Sci.sel.SetSelection m_Sci.lines.PositionFromLine(curLine), m_Sci.lines.GetLineEndPosition(curLine)
End Function

'=========================================================================
' Caret Position
'=========================================================================
Public Property Get currentPos() As Long
    currentPos = m_Sci.SciMsg(SCI_GETCURRENTPOS)
End Property

Public Property Let currentPos(ByVal position As Long)
    m_Sci.SciMsg SCI_SETCURRENTPOS, position
End Property

Public Property Get anchor() As Long
    anchor = m_Sci.SciMsg(SCI_GETANCHOR)
End Property

Public Property Let anchor(ByVal position As Long)
    m_Sci.SciMsg SCI_SETANCHOR, position
End Property

Public Sub GotoPos(ByVal position As Long)
    m_Sci.SciMsg SCI_GOTOPOS, position
End Sub

Public Sub GotoLine(ByVal line As Long)
    m_Sci.SciMsg SCI_GOTOLINE, line
End Sub

Public Sub SetEmptySelection(ByVal caret As Long)
    m_Sci.SciMsg SCI_SETEMPTYSELECTION, caret
End Sub

'=========================================================================
' Convenience Properties
'=========================================================================

' Get/set the current line number (0-based) - no more confusion!
Public Property Get currentLine() As Long
    currentLine = m_Sci.lines.LineFromPosition(currentPos)
End Property

Public Property Let currentLine(ByVal line As Long)
    GotoLine line
End Property

'=========================================================================
' Selection Mode
'=========================================================================
Public Property Get SelectionMode() As eSelectionMode
    SelectionMode = m_Sci.SciMsg(SCI_GETSELECTIONMODE)
End Property

Public Property Let SelectionMode(ByVal value As eSelectionMode)
    m_Sci.SciMsg SCI_SETSELECTIONMODE, value
End Property

'=========================================================================
' Multiple Selections
'=========================================================================
Public Function GetSelectionCount() As Long
    GetSelectionCount = m_Sci.SciMsg(SCI_GETSELECTIONS)
End Function

Public Sub ClearSelections()
    m_Sci.SciMsg SCI_CLEARSELECTIONS
End Sub

Public Sub SetSelectionN(ByVal n As Long, ByVal caret As Long, ByVal anchor As Long)
    m_Sci.SciMsg SCI_SETSELECTION, n
    m_Sci.SciMsg SCI_SETSELECTIONNSTART, n, anchor
    m_Sci.SciMsg SCI_SETSELECTIONNEND, n, caret
End Sub

Public Sub AddSelection(ByVal caret As Long, ByVal anchor As Long)
    m_Sci.SciMsg SCI_ADDSELECTION, caret, anchor
End Sub

Public Property Get MainSelection() As Long
    MainSelection = m_Sci.SciMsg(SCI_GETMAINSELECTION)
End Property

Public Property Let MainSelection(ByVal selection As Long)
    m_Sci.SciMsg SCI_SETMAINSELECTION, selection
End Property

Public Function GetSelectionNStart(ByVal selection As Long) As Long
    GetSelectionNStart = m_Sci.SciMsg(SCI_GETSELECTIONNSTART, selection)
End Function

Public Function GetSelectionNEnd(ByVal selection As Long) As Long
    GetSelectionNEnd = m_Sci.SciMsg(SCI_GETSELECTIONNEND, selection)
End Function

'=========================================================================
' Caret Policy
'=========================================================================
Public Sub SetXCaretPolicy(ByVal caretPolicy As eCaretPolicy, ByVal caretSlop As Long)
    m_Sci.SciMsg SCI_SETXCARETPOLICY, caretPolicy, caretSlop
End Sub

Public Sub SetYCaretPolicy(ByVal caretPolicy As eCaretPolicy, ByVal caretSlop As Long)
    m_Sci.SciMsg SCI_SETYCARETPOLICY, caretPolicy, caretSlop
End Sub


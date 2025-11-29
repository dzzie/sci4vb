VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSciLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Author:  David Zimmer <dzzie@yahoo.com> + Claude.ai
'Site:    http://sandsprite.com
'License: MIT
'---------------------------------------------------


'=========================================================================
' CSciLines - Line-based operations
'=========================================================================
Option Explicit

Private m_Sci As SciWrapper

' Scintilla messages
Private Const SCI_GETLINECOUNT = 2154
Private Const SCI_LINEFROMPOSITION = 2166
Private Const SCI_POSITIONFROMLINE = 2167
Private Const SCI_GETLINEENDPOSITION = 2136
Private Const SCI_LINELENGTH = 2350
Private Const SCI_GETLINE = 2153
Private Const SCI_GETCURLINE = 2027
Private Const SCI_GETLINEINDENTATION = 2127
Private Const SCI_SETLINEINDENTATION = 2126
Private Const SCI_GETLINEINDENTPOSITION = 2128
Private Const SCI_GETCOLUMN = 2129
Private Const SCI_FINDCOLUMN = 2456
Private Const SCI_LINETRANSPOSE = 2339
Private Const SCI_LINEREVERSE = 2354
Private Const SCI_LINEDELETE = 2338
Private Const SCI_LINECUT = 2337
Private Const SCI_LINECOPY = 2455
Private Const SCI_LINEDUPLICATE = 2404
Private Const SCI_GETFIRSTVISIBLELINE = 2152
Private Const SCI_SETFIRSTVISIBLELINE = 2613
Private Const SCI_LINESONSCREEN = 2370
Private Const SCI_VISIBLEFROMDOCLINE = 2220
Private Const SCI_DOCLINEFROMVISIBLE = 2221
Private Const SCI_ENSUREVISIBLE = 2232
Private Const SCI_ENSUREVISIBLEENFORCEPOLICY = 2234
Private Const SCI_LINESCROLL = 2168
Private Const SCI_SCROLLCARET = 2169
Private Const SCI_GOTOLINE = 2024

Friend Sub Init(sciCtrl As SciWrapper)
    Set m_Sci = sciCtrl
End Sub

'=========================================================================
' Line Count and Position
'=========================================================================
Public Property Get Count() As Long
    Count = m_Sci.SciMsg(SCI_GETLINECOUNT)
End Property

Public Function LineFromPosition(ByVal pos As Long) As Long
    LineFromPosition = m_Sci.SciMsg(SCI_LINEFROMPOSITION, pos)
End Function

Public Function PositionFromLine(ByVal line As Long) As Long
    PositionFromLine = m_Sci.SciMsg(SCI_POSITIONFROMLINE, line)
End Function

Public Function GetLineEndPosition(ByVal line As Long) As Long
    GetLineEndPosition = m_Sci.SciMsg(SCI_GETLINEENDPOSITION, line)
End Function

Public Function GetLineLength(ByVal line As Long) As Long
    GetLineLength = m_Sci.SciMsg(SCI_LINELENGTH, line)
End Function

'=========================================================================
' Line Content
'=========================================================================
Public Function GetLine(ByVal line As Long) As String
    Dim length As Long
    Dim buffer() As Byte
    Dim wideBuffer As String
    Dim wideLen As Long
    
    length = GetLineLength(line)
    If length = 0 Then
        GetLine = ""
        Exit Function
    End If
    
    ' Allocate byte array for UTF-8 data
    ReDim buffer(0 To length) As Byte
    m_Sci.SciMsgPtr SCI_GETLINE, line, VarPtr(buffer(0))
    
    ' Convert UTF-8 to Unicode
    wideLen = MultiByteToWideChar(CP_UTF8, 0, VarPtr(buffer(0)), length, 0, 0)
    If wideLen > 0 Then
        wideBuffer = String$(wideLen, 0)
        MultiByteToWideChar CP_UTF8, 0, VarPtr(buffer(0)), length, StrPtr(wideBuffer), wideLen
        GetLine = wideBuffer
    Else
        GetLine = ""
    End If
End Function

Public Function GetCurrentLine() As String
    Dim length As Long
    Dim buffer() As Byte
    Dim wideBuffer As String
    Dim wideLen As Long
    
    length = m_Sci.SciMsg(SCI_GETCURLINE, 0, 0)
    If length = 0 Then
        GetCurrentLine = ""
        Exit Function
    End If
    
    ' Allocate byte array for UTF-8 data
    ReDim buffer(0 To length) As Byte
    m_Sci.SciMsgPtr SCI_GETCURLINE, length + 1, VarPtr(buffer(0))
    
    ' Convert UTF-8 to Unicode
    wideLen = MultiByteToWideChar(CP_UTF8, 0, VarPtr(buffer(0)), length, 0, 0)
    If wideLen > 0 Then
        wideBuffer = String$(wideLen, 0)
        MultiByteToWideChar CP_UTF8, 0, VarPtr(buffer(0)), length, StrPtr(wideBuffer), wideLen
        GetCurrentLine = wideBuffer
    Else
        GetCurrentLine = ""
    End If
End Function

'=========================================================================
' Line Indentation
'=========================================================================
Public Function GetLineIndentation(ByVal line As Long) As Long
    GetLineIndentation = m_Sci.SciMsg(SCI_GETLINEINDENTATION, line)
End Function

Public Sub SetLineIndentation(ByVal line As Long, ByVal indentSize As Long)
    m_Sci.SciMsg SCI_SETLINEINDENTATION, line, indentSize
End Sub

Public Function GetLineIndentPosition(ByVal line As Long) As Long
    GetLineIndentPosition = m_Sci.SciMsg(SCI_GETLINEINDENTPOSITION, line)
End Function

'=========================================================================
' Column Operations
'=========================================================================
Public Function GetColumn(ByVal pos As Long) As Long
    GetColumn = m_Sci.SciMsg(SCI_GETCOLUMN, pos)
End Function

Public Function FindColumn(ByVal line As Long, ByVal Column As Long) As Long
    FindColumn = m_Sci.SciMsg(SCI_FINDCOLUMN, line, Column)
End Function

'=========================================================================
' Line Manipulation
'=========================================================================
Public Sub Transpose()
    m_Sci.SciMsg SCI_LINETRANSPOSE
End Sub

Public Sub Reverse()
    m_Sci.SciMsg SCI_LINEREVERSE
End Sub

Public Sub Delete(ByVal line As Long)
    Dim pos As Long
    pos = PositionFromLine(line)
    m_Sci.sel.GotoPos pos
    m_Sci.SciMsg SCI_LINEDELETE
End Sub

Public Sub Cut()
    m_Sci.SciMsg SCI_LINECUT
End Sub

Public Sub Copy()
    m_Sci.SciMsg SCI_LINECOPY
End Sub

Public Sub Duplicate()
    m_Sci.SciMsg SCI_LINEDUPLICATE
End Sub

Sub GotoLineCentered(ByVal line As Long, Optional selected As Boolean = True)
    Dim mline As Long
    line = line - 1
    mline = line - CInt(LinesOnScreen / 2)
    If mline > 0 Then FirstVisibleLine = mline
    GotoLine line
    If selected Then m_Sci.sel.SelectLine
End Sub

Property Get FirstVisibleLine() As Long
    'returns the displayed line index, not absolute. if word wrap is on, it will be wrong..that was hard to find!
    
    Dim X As Long
    X = GetFirstVisibleLine
    
    If m_Sci.doc.WordWrap Or m_Sci.doc.Folding Then
        X = m_Sci.lines.DocLineFromVisible(X)
    End If
    
    FirstVisibleLine = X
    
End Property

Public Sub GotoLine(ByVal line As Long)
    m_Sci.SciMsg SCI_GOTOLINE, line
End Sub

Property Let FirstVisibleLine(topLine As Long)

    GotoLine topLine + LinesOnScreen + 5 'go past it
    GotoLine topLine   'now go to it and it will be topmost line..

End Property

'=========================================================================
' Line Visibility
'=========================================================================
'Public Property Get FirstVisibleLine() As Long
'    FirstVisibleLine = m_Sci.SciMsg(SCI_GETFIRSTVISIBLELINE)
'End Property
'
'Public Property Let FirstVisibleLine(ByVal line As Long)
'    m_Sci.SciMsg SCI_SETFIRSTVISIBLELINE, line
'End Property

' Convenience methods (match old API style)
Public Function GetFirstVisibleLine() As Long
    GetFirstVisibleLine = FirstVisibleLine
End Function

Public Sub SetFirstVisibleLine(ByVal line As Long)
    FirstVisibleLine = line
End Sub

Public Property Get LinesOnScreen() As Long
    LinesOnScreen = m_Sci.SciMsg(SCI_LINESONSCREEN)
End Property

Public Function VisibleFromDocLine(ByVal docLine As Long) As Long
    VisibleFromDocLine = m_Sci.SciMsg(SCI_VISIBLEFROMDOCLINE, docLine)
End Function

Public Function DocLineFromVisible(ByVal displayLine As Long) As Long
    DocLineFromVisible = m_Sci.SciMsg(SCI_DOCLINEFROMVISIBLE, displayLine)
End Function

Public Sub EnsureVisible(ByVal line As Long)
    m_Sci.SciMsg SCI_ENSUREVISIBLE, line
End Sub

Public Sub EnsureVisibleEnforcePolicy(ByVal line As Long)
    m_Sci.SciMsg SCI_ENSUREVISIBLEENFORCEPOLICY, line
End Sub

'=========================================================================
' Scrolling
'=========================================================================
Public Sub Scroll(ByVal columns As Long, ByVal lines As Long)
    m_Sci.SciMsg SCI_LINESCROLL, columns, lines
End Sub

Public Sub ScrollCaret()
    m_Sci.SciMsg SCI_SCROLLCARET
End Sub


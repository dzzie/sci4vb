VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSciSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Author:  David Zimmer <dzzie@yahoo.com> + Claude.ai
'Site:    http://sandsprite.com
'License: MIT
'---------------------------------------------------


'=========================================================================
' CSciSearch - Search and replace operations
'=========================================================================
Option Explicit

Private m_Sci As SciWrapper

' Search messages
Private Const SCI_FINDTEXT = 2150
Private Const SCI_SEARCHINTARGET = 2197
Private Const SCI_SETTARGETSTART = 2190
Private Const SCI_GETTARGETSTART = 2191
Private Const SCI_SETTARGETEND = 2192
Private Const SCI_GETTARGETEND = 2193
Private Const SCI_SETSEARCHFLAGS = 2198
Private Const SCI_GETSEARCHFLAGS = 2199
Private Const SCI_REPLACETARGET = 2194
Private Const SCI_REPLACETARGETRE = 2195
Private Const SCI_GETTAG = 2616

' Search flags
Public Enum eSearchFlags
    findNone = 0
    findMatchCase = &H4
    findWholeWord = &H2
    findWordStart = &H100000
    findRegexp = &H200000
    findPosix = &H400000
End Enum

' Find text structure for API
Private Type TextToFind
    chrg_cpMin As Long
    chrg_cpMax As Long
    lpstrText As Long
    chrgText_cpMin As Long
    chrgText_cpMax As Long
End Type

Friend Sub Init(sciCtrl As SciWrapper)
    Set m_Sci = sciCtrl
End Sub

'=========================================================================
' Search Flags
'=========================================================================
Public Property Get searchFlags() As eSearchFlags
    searchFlags = m_Sci.SciMsg(SCI_GETSEARCHFLAGS)
End Property

Public Property Let searchFlags(ByVal flags As eSearchFlags)
    m_Sci.SciMsg SCI_SETSEARCHFLAGS, flags
End Property

'=========================================================================
' Target Range
'=========================================================================
Public Property Get TargetStart() As Long
    TargetStart = m_Sci.SciMsg(SCI_GETTARGETSTART)
End Property

Public Property Let TargetStart(ByVal pos As Long)
    m_Sci.SciMsg SCI_SETTARGETSTART, pos
End Property

Public Property Get TargetEnd() As Long
    TargetEnd = m_Sci.SciMsg(SCI_GETTARGETEND)
End Property

Public Property Let TargetEnd(ByVal pos As Long)
    m_Sci.SciMsg SCI_SETTARGETEND, pos
End Property

Public Sub SetTargetRange(ByVal startPos As Long, ByVal endPos As Long)
    m_Sci.SciMsg SCI_SETTARGETSTART, startPos
    m_Sci.SciMsg SCI_SETTARGETEND, endPos
End Sub

'=========================================================================
' Search Operations
'=========================================================================
Public Function findText(ByVal startPos As Long, ByVal endPos As Long, ByVal searchText As String, ByVal flags As eSearchFlags) As Long
    Dim ttf As TextToFind
    Dim oldFlags As Long
    Dim utf8Bytes() As Byte
    
    ' Save current flags and set new ones
    oldFlags = searchFlags
    searchFlags = flags
    
    ' Convert search text to UTF-8
    utf8Bytes = StrConv(searchText, vbFromUnicode)
    
    ' Setup search structure
    ttf.chrg_cpMin = startPos
    ttf.chrg_cpMax = endPos
    ttf.lpstrText = VarPtr(utf8Bytes(0))
    
    ' Perform search
    findText = m_Sci.SciMsgPtr(SCI_FINDTEXT, flags, VarPtr(ttf))
    
    ' Restore original flags
    searchFlags = oldFlags
End Function

Public Function SearchInTarget(ByVal searchText As String) As Long
    Dim utf8Bytes() As Byte
    utf8Bytes = StrConv(searchText, vbFromUnicode)
    SearchInTarget = m_Sci.SciMsgPtr(SCI_SEARCHINTARGET, UBound(utf8Bytes) + 1, VarPtr(utf8Bytes(0)))
End Function

'=========================================================================
' Replace Operations
'=========================================================================
Public Function ReplaceTarget(ByVal replacementText As String) As Long
    Dim utf8Bytes() As Byte
    If Len(replacementText) = 0 Then
        ReplaceTarget = m_Sci.SciMsg(SCI_REPLACETARGET, 0, 0)
    Else
        utf8Bytes = StrConv(replacementText, vbFromUnicode)
        ReplaceTarget = m_Sci.SciMsgPtr(SCI_REPLACETARGET, UBound(utf8Bytes) + 1, VarPtr(utf8Bytes(0)))
    End If
End Function

Public Function ReplaceTargetRE(ByVal replacementText As String) As Long
    Dim utf8Bytes() As Byte
    If Len(replacementText) = 0 Then
        ReplaceTargetRE = m_Sci.SciMsg(SCI_REPLACETARGETRE, 0, 0)
    Else
        utf8Bytes = StrConv(replacementText, vbFromUnicode)
        ReplaceTargetRE = m_Sci.SciMsgPtr(SCI_REPLACETARGETRE, UBound(utf8Bytes) + 1, VarPtr(utf8Bytes(0)))
    End If
End Function

Public Function GetTag(ByVal tagNumber As Long) As String
    Dim buffer() As Byte
    Dim length As Long
    Dim wideBuffer As String
    Dim wideLen As Long
    
    ' Allocate byte array
    ReDim buffer(0 To 255) As Byte
    length = m_Sci.SciMsgPtr(SCI_GETTAG, tagNumber, VarPtr(buffer(0)))
    
    If length > 0 Then
        ' Convert UTF-8 to Unicode
        wideLen = MultiByteToWideChar(CP_UTF8, 0, VarPtr(buffer(0)), length, 0, 0)
        If wideLen > 0 Then
            wideBuffer = String$(wideLen, 0)
            MultiByteToWideChar CP_UTF8, 0, VarPtr(buffer(0)), length, StrPtr(wideBuffer), wideLen
            GetTag = wideBuffer
        Else
            GetTag = ""
        End If
    Else
        GetTag = ""
    End If
End Function

'=========================================================================
' High-Level Search Methods
'=========================================================================
Public Function Find(ByVal searchText As String, Optional ByVal startPos As Long = 0, Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = False, Optional ByVal wordStart As Boolean = False, Optional ByVal useRegex As Boolean = False) As Long
    Dim flags As eSearchFlags
    Dim endPos As Long
    
    ' Build search flags
    flags = findNone
    If matchCase Then flags = flags Or findMatchCase
    If wholeWord Then flags = flags Or findWholeWord
    If wordStart Then flags = flags Or findWordStart
    If useRegex Then flags = flags Or findRegexp
    
    ' Search from startPos to end of document
    endPos = m_Sci.doc.TextLength
    Find = findText(startPos, endPos, searchText, flags)
End Function

Public Function FindNext(ByVal searchText As String, Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = False, Optional ByVal wordStart As Boolean = False, Optional ByVal useRegex As Boolean = False) As Long
    Dim currentPos As Long
    
    currentPos = m_Sci.sel.currentPos
    FindNext = Find(searchText, currentPos + 1, matchCase, wholeWord, wordStart, useRegex)
End Function

Public Function FindPrevious(ByVal searchText As String, Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = False, Optional ByVal wordStart As Boolean = False, Optional ByVal useRegex As Boolean = False) As Long
    Dim flags As eSearchFlags
    Dim currentPos As Long
    
    ' Build search flags
    flags = findNone
    If matchCase Then flags = flags Or findMatchCase
    If wholeWord Then flags = flags Or findWholeWord
    If wordStart Then flags = flags Or findWordStart
    If useRegex Then flags = flags Or findRegexp
    
    ' Search backwards from current position to start
    currentPos = m_Sci.sel.SelectionStart
    If currentPos > 0 Then
        FindPrevious = findText(0, currentPos - 1, searchText, flags)
    Else
        FindPrevious = -1
    End If
End Function

Public Function ReplaceAll(ByVal searchText As String, ByVal replaceText As String, Optional ByVal matchCase As Boolean = False, Optional ByVal wholeWord As Boolean = False, Optional ByVal wordStart As Boolean = False, Optional ByVal useRegex As Boolean = False) As Long
    Dim Count As Long
    Dim pos As Long
    Dim flags As eSearchFlags
    
    ' Build search flags
    flags = findNone
    If matchCase Then flags = flags Or findMatchCase
    If wholeWord Then flags = flags Or findWholeWord
    If wordStart Then flags = flags Or findWordStart
    If useRegex Then flags = flags Or findRegexp
    
    ' Begin undo action
    m_Sci.doc.BeginUndoAction
    
    Count = 0
    pos = 0
    
    Do
        pos = findText(pos, m_Sci.doc.TextLength, searchText, flags)
        If pos >= 0 Then
            SetTargetRange pos, pos + Len(searchText)
            If useRegex Then
                ReplaceTargetRE replaceText
            Else
                ReplaceTarget replaceText
            End If
            Count = Count + 1
            pos = pos + Len(replaceText)
        End If
    Loop While pos >= 0
    
    ' End undo action
    m_Sci.doc.EndUndoAction
    
    ReplaceAll = Count
End Function


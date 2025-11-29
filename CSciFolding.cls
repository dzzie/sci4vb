VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSciFolding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Author:  David Zimmer <dzzie@yahoo.com> + Claude.ai
'Site:    http://sandsprite.com
'License: MIT
'---------------------------------------------------


'=========================================================================
' CSciFolding - Code folding operations
'=========================================================================
Option Explicit

Private m_Sci As SciWrapper

' Scintilla messages
Private Const SCI_VISIBLEFROMDOCLINE = 2220
Private Const SCI_DOCLINEFROMVISIBLE = 2221
Private Const SCI_SHOWLINES = 2226
Private Const SCI_HIDELINES = 2227
Private Const SCI_GETLINEVISIBLE = 2228
Private Const SCI_SETFOLDLEVEL = 2222
Private Const SCI_GETFOLDLEVEL = 2223
Private Const SCI_SETFOLDFLAGS = 2233
Private Const SCI_GETLASTCHILD = 2224
Private Const SCI_GETFOLDPARENT = 2225
Private Const SCI_SETFOLDEXPANDED = 2229
Private Const SCI_GETFOLDEXPANDED = 2230
Private Const SCI_TOGGLEFOLD = 2231
Private Const SCI_FOLDLINE = 2237
Private Const SCI_FOLDCHILDREN = 2238
Private Const SCI_FOLDALL = 2662
Private Const SCI_EXPANDCHILDREN = 2239
Private Const SCI_CONTRACTEDFOLDNEXT = 2618

' Fold level constants
Private Const SC_FOLDLEVELBASE = &H400
Private Const SC_FOLDLEVELWHITEFLAG = &H1000
Private Const SC_FOLDLEVELHEADERFLAG = &H2000
Private Const SC_FOLDLEVELNUMBERMASK = &HFFF

' Fold flags
Public Enum eFoldFlags
    foldLevelNumbers = &H40
    foldLineBeforeExpanded = &H2
    foldLineBeforeContracted = &H4
    foldLineAfterExpanded = &H8
    foldLineAfterContracted = &H10
End Enum

' Fold actions
Public Enum eFoldAction
    foldContract = 0
    foldExpand = 1
    foldToggle = 2
End Enum

Friend Sub Init(sciCtrl As SciWrapper)
    Set m_Sci = sciCtrl
    SetFoldFlags 16   ' Enable fold line editor display
End Sub


'=========================================================================
' Fold Level Management
'=========================================================================
Public Function GetFoldLevel(ByVal line As Long) As Long
    GetFoldLevel = m_Sci.SciMsg(SCI_GETFOLDLEVEL, line)
End Function

Public Sub SetFoldLevel(ByVal line As Long, ByVal level As Long)
    m_Sci.SciMsg SCI_SETFOLDLEVEL, line, level
End Sub

Public Function GetFoldLevelNumber(ByVal line As Long) As Long
    GetFoldLevelNumber = GetFoldLevel(line) And SC_FOLDLEVELNUMBERMASK
End Function

Public Function IsFoldHeader(ByVal line As Long) As Boolean
    IsFoldHeader = ((GetFoldLevel(line) And SC_FOLDLEVELHEADERFLAG) <> 0)
End Function

Public Function GetLastChild(ByVal line As Long, Optional ByVal level As Long = -1) As Long
    GetLastChild = m_Sci.SciMsg(SCI_GETLASTCHILD, line, level)
End Function

Public Function GetFoldParent(ByVal line As Long) As Long
    GetFoldParent = m_Sci.SciMsg(SCI_GETFOLDPARENT, line)
End Function

'=========================================================================
' Fold Expansion
'=========================================================================
Public Function GetFoldExpanded(ByVal line As Long) As Boolean
    GetFoldExpanded = (m_Sci.SciMsg(SCI_GETFOLDEXPANDED, line) <> 0)
End Function

Public Sub SetFoldExpanded(ByVal line As Long, ByVal expanded As Boolean)
    m_Sci.SciMsg SCI_SETFOLDEXPANDED, line, IIf(expanded, 1, 0)
End Sub

Public Sub ToggleFold(ByVal line As Long)
    m_Sci.SciMsg SCI_TOGGLEFOLD, line
End Sub

Public Sub FoldLine(ByVal line As Long, ByVal action As eFoldAction)
    m_Sci.SciMsg SCI_FOLDLINE, line, action
End Sub

Public Sub FoldChildren(ByVal line As Long, ByVal action As eFoldAction)
    m_Sci.SciMsg SCI_FOLDCHILDREN, line, action
End Sub

Public Sub FoldAll(ByVal action As eFoldAction)
    m_Sci.SciMsg SCI_FOLDALL, action
End Sub

Public Sub ExpandChildren(ByVal line As Long, ByVal level As Long)
    m_Sci.SciMsg SCI_EXPANDCHILDREN, line, level
End Sub

'=========================================================================
' Line Visibility
'=========================================================================
Public Function GetLineVisible(ByVal line As Long) As Boolean
    GetLineVisible = (m_Sci.SciMsg(SCI_GETLINEVISIBLE, line) <> 0)
End Function

Public Sub ShowLines(ByVal lineStart As Long, ByVal lineEnd As Long)
    m_Sci.SciMsg SCI_SHOWLINES, lineStart, lineEnd
End Sub

Public Sub HideLines(ByVal lineStart As Long, ByVal lineEnd As Long)
    m_Sci.SciMsg SCI_HIDELINES, lineStart, lineEnd
End Sub

Public Function ContractedFoldNext(ByVal lineStart As Long) As Long
    ContractedFoldNext = m_Sci.SciMsg(SCI_CONTRACTEDFOLDNEXT, lineStart)
End Function

'=========================================================================
' Fold Flags
'=========================================================================
Public Sub SetFoldFlags(ByVal flags As eFoldFlags)
    m_Sci.SciMsg SCI_SETFOLDFLAGS, flags
End Sub
 

'=========================================================================
' Helper Functions
'=========================================================================
Public Sub CollapseAll()
    Dim lineCount As Long
    Dim i As Long
    
    lineCount = m_Sci.lines.Count
    For i = 0 To lineCount - 1
        If IsFoldHeader(i) Then
            If GetFoldExpanded(i) Then
                ToggleFold i
            End If
        End If
    Next i
End Sub

Public Sub ExpandAll()
    Dim lineCount As Long
    Dim i As Long
    
    lineCount = m_Sci.lines.Count
    For i = 0 To lineCount - 1
        If IsFoldHeader(i) Then
            If Not GetFoldExpanded(i) Then
                ToggleFold i
            End If
        End If
    Next i
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSciMarkers"
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
' CSciMarkers - Marker management
'=========================================================================
Option Explicit

Private m_Sci As SciWrapper

Private Const SCI_MARKERDEFINE = 2040
Private Const SCI_MARKERSETFORE = 2041
Private Const SCI_MARKERSETBACK = 2042
Private Const SCI_MARKERADD = 2043
Private Const SCI_MARKERDELETE = 2044
Private Const SCI_MARKERDELETEALL = 2045
Private Const SCI_MARKERGET = 2046
Private Const SCI_MARKERNEXT = 2047
Private Const SCI_MARKERPREVIOUS = 2048
Private Const SCI_MARKERSETALPHA = 2476
Private Const SCI_MARKERSETBACKSELECTED = 2292
Private Const SCI_MARKERENABLEHIGHLIGHT = 2293

' Fold marker numbers (fixed by Scintilla)
Const SC_MARKNUM_FOLDEREND = 25
Const SC_MARKNUM_FOLDEROPENMID = 26
Const SC_MARKNUM_FOLDERMIDTAIL = 27
Const SC_MARKNUM_FOLDERTAIL = 28
Const SC_MARKNUM_FOLDERSUB = 29
Const SC_MARKNUM_FOLDER = 30
Const SC_MARKNUM_FOLDEROPEN = 31

' Marker symbols - CORRECTED to match Scintilla.h
Const SC_MARK_BOXPLUS = 9
Const SC_MARK_BOXMINUS = 10
Const SC_MARK_BOXPLUSCONNECTED = 11
Const SC_MARK_BOXMINUSCONNECTED = 12
Const SC_MARK_TCORNER = 13
Const SC_MARK_LCORNER = 14
Const SC_MARK_VLINE = 15
    
    
Public Enum eMarkerSymbol
    markerCircle = 0
    markerRoundRect = 1
    markerArrow = 2
    markerSmallRect = 3
    markerShortArrow = 4
    markerEmpty = 5
    markerArrowDown = 6
    markerminus = 7
    markerPlus = 8
    ' FIXED: Box-tree fold markers now match Scintilla constants
    markerBoxPlus = 9                ' SC_MARK_BOXPLUS
    markerBoxMinus = 10              ' SC_MARK_BOXMINUS
    markerBoxPlusConnected = 11      ' SC_MARK_BOXPLUSCONNECTED
    markerBoxMinusConnected = 12     ' SC_MARK_BOXMINUSCONNECTED
    markerTCorner = 13               ' SC_MARK_TCORNER
    markerLCorner = 14               ' SC_MARK_LCORNER
    markerVLine = 15                 ' SC_MARK_VLINE
    markerLCornerCurve = 16
    markerTCornerCurve = 17
    markerCirclePlus = 18
    markerCirclePlusConnected = 19
    markerCircleMinus = 20
    markerCircleMinusConnected = 21
    markerBackground = 22
    markerDotDotDot = 23
    markerArrows = 24
    markerPixmap = 25
    markerFullRect = 26
    markerLeftRect = 27
    markerBookmark = 31
End Enum

Friend Sub Init(sciCtrl As SciWrapper)
    Set m_Sci = sciCtrl
    
    Define SC_MARKNUM_FOLDEROPEN, markerminus 'markerBoxMinus           ' 31 -> 14
    Define SC_MARKNUM_FOLDER, markerPlus 'markerBoxPlus                ' 30 -> 12
    Define SC_MARKNUM_FOLDERSUB, markerEmpty               ' 29 -> 5 (EMPTY!)
    Define SC_MARKNUM_FOLDERTAIL, markerEmpty              ' 28 -> 5 (EMPTY!)
    Define SC_MARKNUM_FOLDEREND, markerEmpty               ' 25 -> 5 (EMPTY!)
    Define SC_MARKNUM_FOLDEROPENMID, markerEmpty           ' 26 -> 5 (EMPTY!)
    Define SC_MARKNUM_FOLDERMIDTAIL, markerEmpty           ' 27 -> 5 (EMPTY!)
    
    ' Colors only for the visible markers
    SetFore SC_MARKNUM_FOLDEROPEN, &HFFFFFF
    SetBack SC_MARKNUM_FOLDEROPEN, &H808080
    SetFore SC_MARKNUM_FOLDER, &HFFFFFF
    SetBack SC_MARKNUM_FOLDER, &H808080
    

    'sciCtrl.SciMsg 2233, 0, 0    ' Clear all flags first
'    Define 31, markerminus  ' FOLDEROPEN = BOXMINUS (shows on expanded headers)
'    Define 30, markerPlus  ' FOLDER = BOXPLUS (shows on collapsed headers)
'    Define 29, 5   ' FOLDERSUB = EMPTY (no symbol, just connects)
'    Define 28, 5   ' FOLDERTAIL = EMPTY
'    Define 27, 5   ' FOLDERMIDTAIL = EMPTY
'    Define 26, 5   ' FOLDEROPENMID = EMPTY
'    Define 25, 5   ' FOLDEREND = EMPTY
 
    
End Sub

 

Public Sub Define(ByVal marker As Long, ByVal symbol As eMarkerSymbol)
    m_Sci.SciMsg SCI_MARKERDEFINE, marker, symbol
End Sub

Public Sub SetFore(ByVal marker As Long, ByVal color As Long)
    m_Sci.SciMsg SCI_MARKERSETFORE, marker, color
End Sub

Public Sub SetBack(ByVal marker As Long, ByVal color As Long)
    m_Sci.SciMsg SCI_MARKERSETBACK, marker, color
End Sub

Public Function add(ByVal line As Long, ByVal marker As Long) As Long
    add = m_Sci.SciMsg(SCI_MARKERADD, line, marker)
End Function

Public Sub Delete(ByVal line As Long, ByVal marker As Long)
    m_Sci.SciMsg SCI_MARKERDELETE, line, marker
End Sub

Public Sub DeleteAll(ByVal marker As Long)
    m_Sci.SciMsg SCI_MARKERDELETEALL, marker
End Sub

Public Function GetMarkers(ByVal line As Long) As Long
    GetMarkers = m_Sci.SciMsg(SCI_MARKERGET, line)
End Function

Public Function FindNext(ByVal lineStart As Long, ByVal markerMask As Long) As Long
    FindNext = m_Sci.SciMsg(SCI_MARKERNEXT, lineStart, markerMask)
End Function

Public Function FindPrevious(ByVal lineStart As Long, ByVal markerMask As Long) As Long
    FindPrevious = m_Sci.SciMsg(SCI_MARKERPREVIOUS, lineStart, markerMask)
End Function

Public Sub SetAlpha(ByVal marker As Long, ByVal alpha As Long)
    m_Sci.SciMsg SCI_MARKERSETALPHA, marker, alpha
End Sub
 

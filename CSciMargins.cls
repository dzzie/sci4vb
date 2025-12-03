VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSciMargins"
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
' CSciMargins - Margin configuration
'=========================================================================
Option Explicit

Private m_Sci As SciWrapper

Private Const SCI_SETMARGINTYPEN = 2240
Private Const SCI_GETMARGINTYPEN = 2241
Private Const SCI_SETMARGINWIDTHN = 2242
Private Const SCI_GETMARGINWIDTHN = 2243
Private Const SCI_SETMARGINMASKN = 2244
Private Const SCI_GETMARGINMASKN = 2245
Private Const SCI_SETMARGINSENSITIVEN = 2246
Private Const SCI_GETMARGINSENSITIVEN = 2247
Private Const SCI_SETMARGINCURSORN = 2248
Private Const SCI_GETMARGINCURSORN = 2249
Private Const SCI_SETMARGINBACKN = 2250
Private Const SCI_GETMARGINBACKN = 2251
Private Const SCI_SETMARGINS = 2252
Private Const SCI_GETMARGINS = 2253
Private Const SC_MASK_FOLDERS = &HFE000000  'the mask for all fold markers.
                                            'Bit 26 = Marker 26 (SC_MARKNUM_FOLDEROPENMID)
                                            'Bit 27 = Marker 27 (SC_MARKNUM_FOLDERMIDTAIL)
                                            'Bit 28 = Marker 28 (SC_MARKNUM_FOLDERTAIL)
                                            'Bit 29 = Marker 29 (SC_MARKNUM_FOLDERSUB)
                                            'Bit 30 = Marker 30 (SC_MARKNUM_FOLDER)
                                            'Bit 31 = Marker 31 (SC_MARKNUM_FOLDEROPEN)

Public Enum eMarginType
    marginSymbol = 0         ' Symbol margin (markers)
    marginNumber = 1         ' Line number margin (text)
    marginBack = 2           ' Background color margin
    marginFore = 3           ' Foreground color margin
    marginText = 4           ' Text margin
    marginRText = 5          ' Right-justified text margin
End Enum

Public Enum eMarginIndex
    marginLineNumbers = 0    ' Text margin showing line numbers
    marginFoldMarkers = 2    ' Symbol margin for fold +/- icons
    marginSymbols = 1        ' Symbol margin for breakpoints/bookmarks
    marginCustom1 = 3        ' Available for custom use
    marginCustom2 = 4        ' Available for custom use
End Enum

Public Enum eMarginCursor
    cursorReversedArrow = -1
    cursorArrow = 2
End Enum

Friend Sub Init(sciCtrl As SciWrapper)
    Set m_Sci = sciCtrl
End Sub

'=========================================================================
' Margin Type
'=========================================================================
Public Function GetType(ByVal margin As eMarginIndex) As eMarginType
    GetType = m_Sci.SciMsg(SCI_GETMARGINTYPEN, margin)
End Function

Public Sub SetType(ByVal margin As eMarginIndex, ByVal marginType As eMarginType)
    m_Sci.SciMsg SCI_SETMARGINTYPEN, margin, marginType
End Sub

'=========================================================================
' Margin Width
'=========================================================================
Public Function GetWidth(ByVal margin As eMarginIndex) As Long
    GetWidth = m_Sci.SciMsg(SCI_GETMARGINWIDTHN, margin)
End Function

Public Sub SetWidth(ByVal margin As eMarginIndex, ByVal pixelWidth As Long)
    m_Sci.SciMsg SCI_SETMARGINWIDTHN, margin, pixelWidth
End Sub

'=========================================================================
' Margin Mask
'=========================================================================
Public Function GetMask(ByVal margin As eMarginIndex) As Long
    GetMask = m_Sci.SciMsg(SCI_GETMARGINMASKN, margin)
End Function

Public Sub SetMask(ByVal margin As eMarginIndex, ByVal mask As Long)
    m_Sci.SciMsg SCI_SETMARGINMASKN, margin, mask
End Sub

'=========================================================================
' Margin Sensitivity
'=========================================================================
Public Function GetSensitive(ByVal margin As eMarginIndex) As Boolean
    GetSensitive = (m_Sci.SciMsg(SCI_GETMARGINSENSITIVEN, margin) <> 0)
End Function

Public Sub SetSensitive(ByVal margin As eMarginIndex, ByVal sensitive As Boolean)
    m_Sci.SciMsg SCI_SETMARGINSENSITIVEN, margin, IIf(sensitive, 1, 0)
End Sub

'=========================================================================
' Margin Cursor
'=========================================================================
Public Function GetCursor(ByVal margin As eMarginIndex) As eMarginCursor
    GetCursor = m_Sci.SciMsg(SCI_GETMARGINCURSORN, margin)
End Function

Public Sub SetCursor(ByVal margin As eMarginIndex, ByVal cursor As eMarginCursor)
    m_Sci.SciMsg SCI_SETMARGINCURSORN, margin, cursor
End Sub

'=========================================================================
' Margin Background
'=========================================================================
Public Function GetBack(ByVal margin As eMarginIndex) As Long
    GetBack = m_Sci.SciMsg(SCI_GETMARGINBACKN, margin)
End Function

Public Sub SetBack(ByVal margin As eMarginIndex, ByVal color As Long)
    m_Sci.SciMsg SCI_SETMARGINBACKN, margin, color
End Sub

'=========================================================================
' Margin Count
'=========================================================================
Public Property Get Count() As Long
    Count = m_Sci.SciMsg(SCI_GETMARGINS)
End Property

Public Property Let Count(ByVal marginCount As Long)
    m_Sci.SciMsg SCI_SETMARGINS, marginCount
End Property

'=========================================================================
' Helper Functions first three are up for elimination as bulky generic
'=========================================================================
Public Sub ConfigureLineNumbers(ByVal margin As eMarginIndex, Optional ByVal Width As Long = 40)
    SetType margin, marginNumber
    SetWidth margin, Width
    SetSensitive margin, False
End Sub

Public Sub ConfigureFolding(ByVal margin As eMarginIndex, Optional ByVal Width As Long = 16)
    SetType margin, marginSymbol
    SetMask margin, SC_MASK_FOLDERS
    SetWidth margin, Width
    SetSensitive margin, True
End Sub

Public Sub ConfigureMarkers(ByVal margin As eMarginIndex, ByVal markerMask As Long, Optional ByVal Width As Long = 16)
    SetType margin, marginSymbol
    SetMask margin, markerMask
    SetWidth margin, Width
    SetSensitive margin, True
End Sub

Public Sub ShowLineNumbers(Optional ByVal Width As Long = 40)
    SetWidth marginLineNumbers, Width
    SetType marginLineNumbers, marginNumber
End Sub

Public Sub HideLineNumbers()
    SetWidth marginLineNumbers, 0
End Sub

Public Sub ShowFoldMargin(Optional ByVal Width As Long = 16)
    SetWidth marginFoldMarkers, Width
    SetType marginFoldMarkers, marginSymbol
    SetMask marginFoldMarkers, &H2000000  ' SC_MASK_FOLDERS
    SetSensitive marginFoldMarkers, True
End Sub

Public Sub HideFoldMargin()
    SetWidth marginFoldMarkers, 0
End Sub

Public Sub ShowBreakpointMargin(Optional ByVal Width As Long = 16)
    SetWidth marginSymbols, Width
    SetType marginSymbols, marginSymbol
    SetSensitive marginSymbols, True
End Sub

Public Sub HideBreakpointMargin()
    SetWidth marginSymbols, 0
End Sub


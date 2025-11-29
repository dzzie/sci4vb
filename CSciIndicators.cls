VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSciIndicators"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Author:  David Zimmer <dzzie@yahoo.com> + Claude.ai
'Site:    http://sandsprite.com
'License: MIT
'---------------------------------------------------


'=========================================================================
' CSciIndicators - Indicator management for highlighting
'=========================================================================
Option Explicit

Private m_Sci As SciWrapper

Private Const SCI_INDICSETSTYLE = 2080
Private Const SCI_INDICGETSTYLE = 2081
Private Const SCI_INDICSETFORE = 2082
Private Const SCI_INDICGETFORE = 2083
Private Const SCI_INDICSETUNDER = 2510
Private Const SCI_INDICGETUNDER = 2511
Private Const SCI_INDICSETALPHA = 2523
Private Const SCI_INDICGETALPHA = 2524
Private Const SCI_SETINDICATORCURRENT = 2500
Private Const SCI_GETINDICATORCURRENT = 2501
Private Const SCI_SETINDICATORVALUE = 2502
Private Const SCI_GETINDICATORVALUE = 2503
Private Const SCI_INDICATORFILLRANGE = 2504
Private Const SCI_INDICATORCLEARRANGE = 2505
Private Const SCI_INDICATORALLONFOR = 2506
Private Const SCI_INDICATORVALUEAT = 2507
Private Const SCI_INDICATORSTART = 2508
Private Const SCI_INDICATOREND = 2509

Public Enum eIndicatorStyle
    indicPlain = 0
    indicSquiggle = 1
    indicTT = 2
    indicDiagonal = 3
    indicStrike = 4
    indicHidden = 5
    indicBox = 6
    indicRoundBox = 7
    indicStraightBox = 8
    indicDash = 9
    indicDots = 10
    indicSquiggleLow = 11
    indicDotBox = 12
    indicSquigglePixmap = 13
    indicCompositionThick = 14
    indicCompositionThin = 15
    indicFullBox = 16
    indicTextFore = 17
    indicPoint = 18
    indicPointCharacter = 19
End Enum

Friend Sub Init(sciCtrl As SciWrapper)
    Set m_Sci = sciCtrl
End Sub

Public Sub SetStyle(ByVal indicator As Long, ByVal style As eIndicatorStyle)
    m_Sci.SciMsg SCI_INDICSETSTYLE, indicator, style
End Sub

Public Function GetStyle(ByVal indicator As Long) As eIndicatorStyle
    GetStyle = m_Sci.SciMsg(SCI_INDICGETSTYLE, indicator)
End Function

Public Sub SetFore(ByVal indicator As Long, ByVal color As Long)
    m_Sci.SciMsg SCI_INDICSETFORE, indicator, color
End Sub

Public Function GetFore(ByVal indicator As Long) As Long
    GetFore = m_Sci.SciMsg(SCI_INDICGETFORE, indicator)
End Function

Public Sub SetUnder(ByVal indicator As Long, ByVal under As Boolean)
    m_Sci.SciMsg SCI_INDICSETUNDER, indicator, IIf(under, 1, 0)
End Sub

Public Function GetUnder(ByVal indicator As Long) As Boolean
    GetUnder = (m_Sci.SciMsg(SCI_INDICGETUNDER, indicator) <> 0)
End Function

Public Sub SetAlpha(ByVal indicator As Long, ByVal alpha As Long)
    m_Sci.SciMsg SCI_INDICSETALPHA, indicator, alpha
End Sub

Public Function GetAlpha(ByVal indicator As Long) As Long
    GetAlpha = m_Sci.SciMsg(SCI_INDICGETALPHA, indicator)
End Function

Public Property Get Current() As Long
    Current = m_Sci.SciMsg(SCI_GETINDICATORCURRENT)
End Property

Public Property Let Current(ByVal indicator As Long)
    m_Sci.SciMsg SCI_SETINDICATORCURRENT, indicator
End Property

Public Property Get value() As Long
    value = m_Sci.SciMsg(SCI_GETINDICATORVALUE)
End Property

Public Property Let value(ByVal val As Long)
    m_Sci.SciMsg SCI_SETINDICATORVALUE, val
End Property

Public Sub FillRange(ByVal position As Long, ByVal fillLength As Long)
    m_Sci.SciMsg SCI_INDICATORFILLRANGE, position, fillLength
End Sub

Public Sub ClearRange(ByVal position As Long, ByVal clearLength As Long)
    m_Sci.SciMsg SCI_INDICATORCLEARRANGE, position, clearLength
End Sub

Public Function AllOnFor(ByVal position As Long) As Long
    AllOnFor = m_Sci.SciMsg(SCI_INDICATORALLONFOR, position)
End Function

Public Function ValueAt(ByVal indicator As Long, ByVal position As Long) As Long
    ValueAt = m_Sci.SciMsg(SCI_INDICATORVALUEAT, indicator, position)
End Function

Public Function GetStart(ByVal indicator As Long, ByVal position As Long) As Long
    GetStart = m_Sci.SciMsg(SCI_INDICATORSTART, indicator, position)
End Function

Public Function GetEnd(ByVal indicator As Long, ByVal position As Long) As Long
    GetEnd = m_Sci.SciMsg(SCI_INDICATOREND, indicator, position)
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSciView"
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
' CSciView - Display and rendering configuration
'=========================================================================
Option Explicit

Private m_Sci As SciWrapper

' View messages
Private Const SCI_SETZOOM = 2373
Private Const SCI_GETZOOM = 2374
Private Const SCI_ZOOMIN = 2333
Private Const SCI_ZOOMOUT = 2334
Private Const SCI_SETVIEWWS = 2021
Private Const SCI_GETVIEWWS = 2020
Private Const SCI_SETWHITESPACEFORE = 2084
Private Const SCI_SETWHITESPACEBACK = 2085
Private Const SCI_SETWHITESPACESIZE = 2086
Private Const SCI_GETWHITESPACESIZE = 2087
Private Const SCI_SETHIDESELECTION = 2173
Private Const SCI_GETHIDESELECTION = 2174
Private Const SCI_SETCARETLINEVISIBLE = 2096
Private Const SCI_GETCARETLINEVISIBLE = 2095
Private Const SCI_SETCARETLINEBACK = 2098
Private Const SCI_GETCARETLINEBACK = 2097
Private Const SCI_SETCARETLINEBACKALPHA = 2470
Private Const SCI_GETCARETLINEBACKALPHA = 2471
Private Const SCI_SETCARETFORE = 2069
Private Const SCI_GETCARETFORE = 2138
Private Const SCI_SETCARETWIDTH = 2188
Private Const SCI_GETCARETWIDTH = 2189
Private Const SCI_SETCARETSTYLE = 2512
Private Const SCI_GETCARETSTYLE = 2513
Private Const SCI_SETCARETPERIOD = 2076
Private Const SCI_GETCARETPERIOD = 2075
Private Const SCI_SETEDGEMODE = 2363
Private Const SCI_GETEDGEMODE = 2362
Private Const SCI_SETEDGECOLUMN = 2361
Private Const SCI_GETEDGECOLUMN = 2360
Private Const SCI_SETEDGECOLOUR = 2365
Private Const SCI_GETEDGECOLOUR = 2364
Private Const SCI_SETWRAPMODE = 2268
Private Const SCI_GETWRAPMODE = 2269
Private Const SCI_SETWRAPVISUALFLAGS = 2460
Private Const SCI_GETWRAPVISUALFLAGS = 2461
Private Const SCI_SETWRAPINDENTMODE = 2472
Private Const SCI_GETWRAPINDENTMODE = 2473
Private Const SCI_SETSCROLLWIDTH = 2274
Private Const SCI_GETSCROLLWIDTH = 2275
Private Const SCI_SETENDATLASTLINE = 2277
Private Const SCI_GETENDATLASTLINE = 2278
Private Const SCI_SETHSCROLLBAR = 2130
Private Const SCI_GETHSCROLLBAR = 2131
Private Const SCI_SETVSCROLLBAR = 2280
Private Const SCI_GETVSCROLLBAR = 2281

' Whitespace visibility
Public Enum eWhitespace
    wsInvisible = 0
    wsVisibleAlways = 1
    wsVisibleAfterIndent = 2
End Enum

' Edge mode
Public Enum eEdgeMode
    edgeNone = 0
    edgeLine = 1
    edgeBackground = 2
End Enum

' Wrap mode
Public Enum eWrapMode
    wrapNone = 0
    wrapWord = 1
    wrapChar = 2
    wrapWhitespace = 3
End Enum

' Wrap visual flags
Public Enum eWrapVisualFlags
    wrapFlagNone = &H0
    wrapFlagEnd = &H1
    wrapFlagStart = &H2
    wrapFlagIndent = &H4
End Enum

' Caret style
Public Enum eCaretStyle
    caretInvisible = 0
    caretLine = 1
    caretBlock = 2
End Enum

Friend Sub Init(sciCtrl As SciWrapper)
    Set m_Sci = sciCtrl
End Sub

Property Let IndentationGuides(show As Boolean)

    Const SC_IV_NONE = 0 'No indentation guides are shown.
    Const SC_IV_REAL = 1 'Indentation guides are shown inside real indentation white space.

    m_Sci.SciMsg 2132, IIf(show, SC_IV_REAL, SC_IV_NONE), 0

End Property

'=========================================================================
' Zoom
'=========================================================================
Public Property Get Zoom() As Long
    Zoom = m_Sci.SciMsg(SCI_GETZOOM)
End Property

Public Property Let Zoom(ByVal points As Long)
    m_Sci.SciMsg SCI_SETZOOM, points
End Property

Public Sub ZoomIn()
    m_Sci.SciMsg SCI_ZOOMIN
End Sub

Public Sub ZoomOut()
    m_Sci.SciMsg SCI_ZOOMOUT
End Sub

'=========================================================================
' Whitespace
'=========================================================================
Public Property Get ViewWhitespace() As eWhitespace
    ViewWhitespace = m_Sci.SciMsg(SCI_GETVIEWWS)
End Property

Public Property Let ViewWhitespace(ByVal mode As eWhitespace)
    m_Sci.SciMsg SCI_SETVIEWWS, mode
End Property

Public Sub SetWhitespaceFore(ByVal useCustom As Boolean, ByVal color As Long)
    m_Sci.SciMsg SCI_SETWHITESPACEFORE, IIf(useCustom, 1, 0), color
End Sub

Public Sub SetWhitespaceBack(ByVal useCustom As Boolean, ByVal color As Long)
    m_Sci.SciMsg SCI_SETWHITESPACEBACK, IIf(useCustom, 1, 0), color
End Sub

Public Property Get WhitespaceSize() As Long
    WhitespaceSize = m_Sci.SciMsg(SCI_GETWHITESPACESIZE)
End Property

Public Property Let WhitespaceSize(ByVal size As Long)
    m_Sci.SciMsg SCI_SETWHITESPACESIZE, size
End Property

'=========================================================================
' Selection Display
'=========================================================================

' Control whether selection is visible when control loses focus
Public Property Get HideSelection() As Boolean
    HideSelection = (m_Sci.SciMsg(SCI_GETHIDESELECTION) <> 0)
End Property

Public Property Let HideSelection(ByVal hide As Boolean)
    m_Sci.SciMsg SCI_SETHIDESELECTION, IIf(hide, 1, 0)
End Property

'=========================================================================
' Caret Line
'=========================================================================
Public Property Get CaretLineVisible() As Boolean
    CaretLineVisible = (m_Sci.SciMsg(SCI_GETCARETLINEVISIBLE) <> 0)
End Property

Public Property Let CaretLineVisible(ByVal Visible As Boolean)
    m_Sci.SciMsg SCI_SETCARETLINEVISIBLE, IIf(Visible, 1, 0)
End Property

Public Property Get CaretLineBack() As Long
    CaretLineBack = m_Sci.SciMsg(SCI_GETCARETLINEBACK)
End Property

Public Property Let CaretLineBack(ByVal color As Long)
    m_Sci.SciMsg SCI_SETCARETLINEBACK, color
End Property

Public Property Get CaretLineBackAlpha() As Long
    CaretLineBackAlpha = m_Sci.SciMsg(SCI_GETCARETLINEBACKALPHA)
End Property

Public Property Let CaretLineBackAlpha(ByVal alpha As Long)
    m_Sci.SciMsg SCI_SETCARETLINEBACKALPHA, alpha
End Property

'=========================================================================
' Caret
'=========================================================================
Public Property Get CaretFore() As Long
    CaretFore = m_Sci.SciMsg(SCI_GETCARETFORE)
End Property

Public Property Let CaretFore(ByVal color As Long)
    m_Sci.SciMsg SCI_SETCARETFORE, color
End Property

Public Property Get CaretWidth() As Long
    CaretWidth = m_Sci.SciMsg(SCI_GETCARETWIDTH)
End Property

Public Property Let CaretWidth(ByVal pixelWidth As Long)
    m_Sci.SciMsg SCI_SETCARETWIDTH, pixelWidth
End Property

Public Property Get CaretStyle() As eCaretStyle
    CaretStyle = m_Sci.SciMsg(SCI_GETCARETSTYLE)
End Property

Public Property Let CaretStyle(ByVal style As eCaretStyle)
    m_Sci.SciMsg SCI_SETCARETSTYLE, style
End Property

Public Property Get CaretPeriod() As Long
    CaretPeriod = m_Sci.SciMsg(SCI_GETCARETPERIOD)
End Property

Public Property Let CaretPeriod(ByVal milliseconds As Long)
    m_Sci.SciMsg SCI_SETCARETPERIOD, milliseconds
End Property

'=========================================================================
' Edge
'=========================================================================
Public Property Get EdgeMode() As eEdgeMode
    EdgeMode = m_Sci.SciMsg(SCI_GETEDGEMODE)
End Property

Public Property Let EdgeMode(ByVal mode As eEdgeMode)
    m_Sci.SciMsg SCI_SETEDGEMODE, mode
End Property

Public Property Get EdgeColumn() As Long
    EdgeColumn = m_Sci.SciMsg(SCI_GETEDGECOLUMN)
End Property

Public Property Let EdgeColumn(ByVal Column As Long)
    m_Sci.SciMsg SCI_SETEDGECOLUMN, Column
End Property

Public Property Get EdgeColor() As Long
    EdgeColor = m_Sci.SciMsg(SCI_GETEDGECOLOUR)
End Property

Public Property Let EdgeColor(ByVal color As Long)
    m_Sci.SciMsg SCI_SETEDGECOLOUR, color
End Property

'=========================================================================
' Wrap
'=========================================================================
Public Property Get WrapMode() As eWrapMode
    WrapMode = m_Sci.SciMsg(SCI_GETWRAPMODE)
End Property

Public Property Let WrapMode(ByVal mode As eWrapMode)
    m_Sci.SciMsg SCI_SETWRAPMODE, mode
End Property

Public Property Get WrapVisualFlags() As eWrapVisualFlags
    WrapVisualFlags = m_Sci.SciMsg(SCI_GETWRAPVISUALFLAGS)
End Property

Public Property Let WrapVisualFlags(ByVal flags As eWrapVisualFlags)
    m_Sci.SciMsg SCI_SETWRAPVISUALFLAGS, flags
End Property

Public Property Get WrapIndentMode() As Long
    WrapIndentMode = m_Sci.SciMsg(SCI_GETWRAPINDENTMODE)
End Property

Public Property Let WrapIndentMode(ByVal indent As Long)
    m_Sci.SciMsg SCI_SETWRAPINDENTMODE, indent
End Property

'=========================================================================
' Scrollbars
'=========================================================================
Public Property Get ScrollWidth() As Long
    ScrollWidth = m_Sci.SciMsg(SCI_GETSCROLLWIDTH)
End Property

Public Property Let ScrollWidth(ByVal pixelWidth As Long)
    m_Sci.SciMsg SCI_SETSCROLLWIDTH, pixelWidth
End Property

Public Property Get EndAtLastLine() As Boolean
    EndAtLastLine = (m_Sci.SciMsg(SCI_GETENDATLASTLINE) <> 0)
End Property

Public Property Let EndAtLastLine(ByVal endAtLast As Boolean)
    m_Sci.SciMsg SCI_SETENDATLASTLINE, IIf(endAtLast, 1, 0)
End Property

Public Property Get HScrollBar() As Boolean
    HScrollBar = (m_Sci.SciMsg(SCI_GETHSCROLLBAR) <> 0)
End Property

Public Property Let HScrollBar(ByVal Visible As Boolean)
    m_Sci.SciMsg SCI_SETHSCROLLBAR, IIf(Visible, 1, 0)
End Property

Public Property Get VScrollBar() As Boolean
    VScrollBar = (m_Sci.SciMsg(SCI_GETVSCROLLBAR) <> 0)
End Property

Public Property Let VScrollBar(ByVal Visible As Boolean)
    m_Sci.SciMsg SCI_SETVSCROLLBAR, IIf(Visible, 1, 0)
End Property

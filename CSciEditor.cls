VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSciEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Author:  David Zimmer <dzzie@yahoo.com> + Claude.ai
'Site:    http://sandsprite.com
'License: MIT
'---------------------------------------------------


'=========================================================================
' CSciEditor - Core editing operations
'=========================================================================
Option Explicit

Private m_Sci As SciWrapper

' Scintilla messages
Private Const SCI_CUT = 2177
Private Const SCI_COPY = 2178
Private Const SCI_PASTE = 2179
Private Const SCI_CANPASTE = 2173
Private Const SCI_COPYTEXT = 2420
Private Const SCI_COPYRANGE = 2419
Private Const SCI_CLEAR = 2180
Private Const SCI_REPLACESEL = 2170
Private Const SCI_SETTABWIDTH = 2036
Private Const SCI_GETTABWIDTH = 2121
Private Const SCI_SETUSETABS = 2124
Private Const SCI_GETUSETABS = 2125
Private Const SCI_SETINDENT = 2122
Private Const SCI_GETINDENT = 2123
Private Const SCI_SETTABINDENTS = 2260
Private Const SCI_GETTABINDENTS = 2261
Private Const SCI_SETBACKSPACEUNINDENTS = 2262
Private Const SCI_GETBACKSPACEUNINDENTS = 2263
Private Const SCI_LINEINDENT = 2126
Private Const SCI_LINEDEDENT = 2127
Private Const SCI_SETEOLMODE = 2031
Private Const SCI_GETEOLMODE = 2030
Private Const SCI_CONVERTEOLS = 2029
Private Const SCI_SETVIEWEOL = 2356
Private Const SCI_GETVIEWEOL = 2355
Private Const SCI_SETOVERTYPE = 2186
Private Const SCI_GETOVERTYPE = 2187
Private Const SCI_SETVIRTUALSPACEOPTIONS = 2596
Private Const SCI_GETVIRTUALSPACEOPTIONS = 2597

' EOL modes
Public Enum eEOLMode
    eolCRLF = 0
    eolCR = 1
    eolLF = 2
End Enum

' Virtual space options
Public Enum eVirtualSpace
    vsNone = 0
    vsRectangularSelection = 1
    vsUserAccessible = 2
End Enum

Friend Sub Init(sciCtrl As SciWrapper)
    Set m_Sci = sciCtrl
End Sub

'=========================================================================
' Clipboard Operations
'=========================================================================
Public Sub Cut()
    m_Sci.SciMsg SCI_CUT
End Sub

Public Sub Copy()
    m_Sci.SciMsg SCI_COPY
End Sub

Public Sub Paste()
    m_Sci.SciMsg SCI_PASTE
End Sub

Public Function CanPaste() As Boolean
    CanPaste = (m_Sci.SciMsg(SCI_CANPASTE) <> 0)
End Function

Public Sub CopyText(ByVal length As Long, ByVal text As String)
    m_Sci.SciMsgStr SCI_COPYTEXT, length, text
End Sub

Public Sub CopyRange(ByVal start As Long, ByVal endPos As Long)
    m_Sci.SciMsg SCI_COPYRANGE, start, endPos
End Sub

Public Sub Clear()
    m_Sci.SciMsg SCI_CLEAR
End Sub

Public Sub ReplaceSelection(ByVal text As String)
    m_Sci.SciMsgStr SCI_REPLACESEL, 0, text
End Sub

'=========================================================================
' Tab and Indentation
'=========================================================================
Public Property Get TabWidth() As Long
    TabWidth = m_Sci.SciMsg(SCI_GETTABWIDTH)
End Property

Public Property Let TabWidth(ByVal value As Long)
    m_Sci.SciMsg SCI_SETTABWIDTH, value
End Property

Public Property Get UseTabs() As Boolean
    UseTabs = (m_Sci.SciMsg(SCI_GETUSETABS) <> 0)
End Property

Public Property Let UseTabs(ByVal value As Boolean)
    m_Sci.SciMsg SCI_SETUSETABS, IIf(value, 1, 0)
End Property

Public Property Get indent() As Long
    indent = m_Sci.SciMsg(SCI_GETINDENT)
End Property

Public Property Let indent(ByVal value As Long)
    m_Sci.SciMsg SCI_SETINDENT, value
End Property

Public Property Get TabIndents() As Boolean
    TabIndents = (m_Sci.SciMsg(SCI_GETTABINDENTS) <> 0)
End Property

Public Property Let TabIndents(ByVal value As Boolean)
    m_Sci.SciMsg SCI_SETTABINDENTS, IIf(value, 1, 0)
End Property

Public Property Get BackspaceUnindents() As Boolean
    BackspaceUnindents = (m_Sci.SciMsg(SCI_GETBACKSPACEUNINDENTS) <> 0)
End Property

Public Property Let BackspaceUnindents(ByVal value As Boolean)
    m_Sci.SciMsg SCI_SETBACKSPACEUNINDENTS, IIf(value, 1, 0)
End Property

Public Sub IndentLine(ByVal line As Long)
    m_Sci.SciMsg SCI_LINEINDENT, line
End Sub

Public Sub DedentLine(ByVal line As Long)
    m_Sci.SciMsg SCI_LINEDEDENT, line
End Sub

'=========================================================================
' EOL (End of Line)
'=========================================================================
Public Property Get eolMode() As eEOLMode
    eolMode = m_Sci.SciMsg(SCI_GETEOLMODE)
End Property

Public Property Let eolMode(ByVal value As eEOLMode)
    m_Sci.SciMsg SCI_SETEOLMODE, value
End Property

Public Sub ConvertEOLs(ByVal eolMode As eEOLMode)
    m_Sci.SciMsg SCI_CONVERTEOLS, eolMode
End Sub

Public Property Get ViewEOL() As Boolean
    ViewEOL = (m_Sci.SciMsg(SCI_GETVIEWEOL) <> 0)
End Property

Public Property Let ViewEOL(ByVal value As Boolean)
    m_Sci.SciMsg SCI_SETVIEWEOL, IIf(value, 1, 0)
End Property

'=========================================================================
' Overtype Mode
'=========================================================================
Public Property Get Overtype() As Boolean
    Overtype = (m_Sci.SciMsg(SCI_GETOVERTYPE) <> 0)
End Property

Public Property Let Overtype(ByVal value As Boolean)
    m_Sci.SciMsg SCI_SETOVERTYPE, IIf(value, 1, 0)
End Property

'=========================================================================
' Virtual Space
'=========================================================================
Public Property Get VirtualSpaceOptions() As eVirtualSpace
    VirtualSpaceOptions = m_Sci.SciMsg(SCI_GETVIRTUALSPACEOPTIONS)
End Property

Public Property Let VirtualSpaceOptions(ByVal value As eVirtualSpace)
    m_Sci.SciMsg SCI_SETVIRTUALSPACEOPTIONS, value
End Property

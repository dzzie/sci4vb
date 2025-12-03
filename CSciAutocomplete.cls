VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSciAutocomplete"
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
' CSciAutocomplete - Autocomplete and call tip management
'=========================================================================
Option Explicit

Private m_Sci As SciWrapper

' Autocomplete messages
Private Const SCI_AUTOCSHOW = 2100
Private Const SCI_AUTOCCANCEL = 2101
Private Const SCI_AUTOCACTIVE = 2102
Private Const SCI_AUTOCPOSSTART = 2103
Private Const SCI_AUTOCCOMPLETE = 2104
Private Const SCI_AUTOCSTOPS = 2105
Private Const SCI_AUTOCSETSEPARATOR = 2106
Private Const SCI_AUTOCGETSEPARATOR = 2107
Private Const SCI_AUTOCSELECT = 2108
Private Const SCI_AUTOCSETCANCELATSTART = 2110
Private Const SCI_AUTOCGETCANCELATSTART = 2111
Private Const SCI_AUTOCSETFILLUPS = 2112
Private Const SCI_AUTOCSETCHOOSESINGLE = 2113
Private Const SCI_AUTOCGETCHOOSESINGLE = 2114
Private Const SCI_AUTOCSETIGNORECASE = 2115
Private Const SCI_AUTOCGETIGNORECASE = 2116
Private Const SCI_AUTOCSETAUTOHIDE = 2118
Private Const SCI_AUTOCGETAUTOHIDE = 2119
Private Const SCI_AUTOCSETDROPRESTOFWORD = 2270
Private Const SCI_AUTOCGETDROPRESTOFWORD = 2271
Private Const SCI_AUTOCSETMAXHEIGHT = 2210
Private Const SCI_AUTOCGETMAXHEIGHT = 2211
Private Const SCI_AUTOCSETMAXWIDTH = 2208
Private Const SCI_AUTOCGETMAXWIDTH = 2209

' Call tip messages
Private Const SCI_CALLTIPSHOW = 2200
Private Const SCI_CALLTIPCANCEL = 2201
Private Const SCI_CALLTIPACTIVE = 2202
Private Const SCI_CALLTIPPOSSTART = 2203
Private Const SCI_CALLTIPSETHLT = 2204
Private Const SCI_CALLTIPSETBACK = 2205
Private Const SCI_CALLTIPSETFORE = 2206
Private Const SCI_CALLTIPSETFOREHLT = 2207
Private Const SCI_CALLTIPUSESTYLE = 2212

Friend Sub Init(sciCtrl As SciWrapper)
    Set m_Sci = sciCtrl
End Sub

'=========================================================================
' Autocomplete Operations
'=========================================================================
Public Sub show(ByVal lenEntered As Long, ByVal itemList As String)
    m_Sci.SciMsgStr SCI_AUTOCSHOW, lenEntered, itemList
End Sub

Public Sub Cancel()
    m_Sci.SciMsg SCI_AUTOCCANCEL
End Sub

Public Function isActive() As Boolean
    isActive = (m_Sci.SciMsg(SCI_AUTOCACTIVE) <> 0)
End Function

Public Function GetPosStart() As Long
    GetPosStart = m_Sci.SciMsg(SCI_AUTOCPOSSTART)
End Function

Public Sub Complete()
    m_Sci.SciMsg SCI_AUTOCCOMPLETE
End Sub

Public Sub SetStops(ByVal characterSet As String)
    m_Sci.SciMsgStr SCI_AUTOCSTOPS, 0, characterSet
End Sub

Public Sub SetFillups(ByVal characterSet As String)
    m_Sci.SciMsgStr SCI_AUTOCSETFILLUPS, 0, characterSet
End Sub

Public Sub SelectItem(ByVal text As String)
    m_Sci.SciMsgStr SCI_AUTOCSELECT, 0, text
End Sub

'=========================================================================
' Autocomplete Configuration
'=========================================================================
Public Property Get Separator() As Long
    Separator = m_Sci.SciMsg(SCI_AUTOCGETSEPARATOR)
End Property

Public Property Let Separator(ByVal charCode As Long)
    m_Sci.SciMsg SCI_AUTOCSETSEPARATOR, charCode
End Property

Public Property Get CancelAtStart() As Boolean
    CancelAtStart = (m_Sci.SciMsg(SCI_AUTOCGETCANCELATSTART) <> 0)
End Property

Public Property Let CancelAtStart(ByVal value As Boolean)
    m_Sci.SciMsg SCI_AUTOCSETCANCELATSTART, IIf(value, 1, 0)
End Property

Public Property Get ChooseSingle() As Boolean
    ChooseSingle = (m_Sci.SciMsg(SCI_AUTOCGETCHOOSESINGLE) <> 0)
End Property

Public Property Let ChooseSingle(ByVal value As Boolean)
    m_Sci.SciMsg SCI_AUTOCSETCHOOSESINGLE, IIf(value, 1, 0)
End Property

Public Property Get IgnoreCase() As Boolean
    IgnoreCase = (m_Sci.SciMsg(SCI_AUTOCGETIGNORECASE) <> 0)
End Property

Public Property Let IgnoreCase(ByVal value As Boolean)
    m_Sci.SciMsg SCI_AUTOCSETIGNORECASE, IIf(value, 1, 0)
End Property

Public Property Get AutoHide() As Boolean
    AutoHide = (m_Sci.SciMsg(SCI_AUTOCGETAUTOHIDE) <> 0)
End Property

Public Property Let AutoHide(ByVal value As Boolean)
    m_Sci.SciMsg SCI_AUTOCSETAUTOHIDE, IIf(value, 1, 0)
End Property

Public Property Get DropRestOfWord() As Boolean
    DropRestOfWord = (m_Sci.SciMsg(SCI_AUTOCGETDROPRESTOFWORD) <> 0)
End Property

Public Property Let DropRestOfWord(ByVal value As Boolean)
    m_Sci.SciMsg SCI_AUTOCSETDROPRESTOFWORD, IIf(value, 1, 0)
End Property

Public Property Get MaxHeight() As Long
    MaxHeight = m_Sci.SciMsg(SCI_AUTOCGETMAXHEIGHT)
End Property

Public Property Let MaxHeight(ByVal rowCount As Long)
    m_Sci.SciMsg SCI_AUTOCSETMAXHEIGHT, rowCount
End Property

Public Property Get MaxWidth() As Long
    MaxWidth = m_Sci.SciMsg(SCI_AUTOCGETMAXWIDTH)
End Property

Public Property Let MaxWidth(ByVal characterCount As Long)
    m_Sci.SciMsg SCI_AUTOCSETMAXWIDTH, characterCount
End Property

'=========================================================================
' Call Tips
'=========================================================================
Public Sub ShowCallTip(ByVal pos As Long, ByVal definition As String)
    m_Sci.SciMsgStr SCI_CALLTIPSHOW, pos, definition
End Sub

Public Sub CancelCallTip()
    m_Sci.SciMsg SCI_CALLTIPCANCEL
End Sub

Public Function IsCallTipActive() As Boolean
    IsCallTipActive = (m_Sci.SciMsg(SCI_CALLTIPACTIVE) <> 0)
End Function

Public Function GetCallTipPosStart() As Long
    GetCallTipPosStart = m_Sci.SciMsg(SCI_CALLTIPPOSSTART)
End Function

Public Sub SetCallTipHighlight(ByVal highlightStart As Long, ByVal highlightEnd As Long)
    m_Sci.SciMsg SCI_CALLTIPSETHLT, highlightStart, highlightEnd
End Sub

Public Sub SetCallTipBack(ByVal color As Long)
    m_Sci.SciMsg SCI_CALLTIPSETBACK, color
End Sub

Public Sub SetCallTipFore(ByVal color As Long)
    m_Sci.SciMsg SCI_CALLTIPSETFORE, color
End Sub

Public Sub SetCallTipForeHlt(ByVal color As Long)
    m_Sci.SciMsg SCI_CALLTIPSETFOREHLT, color
End Sub

Public Sub UseCallTipStyle(ByVal tabSize As Long)
    m_Sci.SciMsg SCI_CALLTIPUSESTYLE, tabSize
End Sub


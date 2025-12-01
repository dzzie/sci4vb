VERSION 5.00
Begin VB.UserControl SciWrapper 
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7755
   ForwardFocus    =   -1  'True
   ScaleHeight     =   4800
   ScaleWidth      =   7755
   ToolboxBitmap   =   "SciWrapper.ctx":0000
End
Attribute VB_Name = "SciWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Author:  David Zimmer <dzzie@yahoo.com> + Claude.ai
'Site:    http://sandsprite.com
'License: MIT
'---------------------------------------------------



'=========================================================================
' SciWrapper - Professional VB6 Scintilla Wrapper
' Clean room implementation - No code derived from other wrappers
'=========================================================================
Option Explicit

Implements iSubclass

'=========================================================================
' Public Classes - Feature Modules (short names for less typing!)
'=========================================================================
Public doc As New CSciDocument
Public edit As New CSciEditor
Public sel As New CSciSelection
Public lines As New CSciLines
Public style As New CSciStyling
Public Fold As New CSciFolding
Public Mark As New CSciMarkers
Public Indic As New CSciIndicators
Public Autoc As New CSciAutocomplete
Public Margins As New CSciMargins
Public View As New CSciView
Public Search As New CSciSearch
Public Helper As New CSciHelpers
Public Intellisense As New CIntellisenseManager
Public breakpoints As New CBreakPoint

'=========================================================================
' Events
'=========================================================================
Event UpdateUI(updated As Long)
Event Modified(position As Long, modificationType As Long, text As String, length As Long, linesAdded As Long)
Event Notify(notificationCode As Long)
Event MarginClick(margin As Long, position As Long, modifiers As Long)
Event KeyPressed(key As Long, modifiers As Long, ByRef handled As Boolean)
Event CharAdded(ch As Long)
Event SavePointReached()
Event SavePointLeft()
Event StyleNeeded(position As Long)
Event DoubleClick(position As Long, line As Long)
Event DwellStart(position As Long, X As Long, Y As Long)
Event DwellEnd(position As Long, X As Long, Y As Long)
Event AutoCSelection(text As String, position As Long)
Event AutoCCancelled()
Event CallTipClick(position As Long)
Event AutoCompleteEvent(className As String) 'duplicate now?
Event WordHighlighted(word As String, instances As Long)
Event UserBreakpointToggle(line As Long, isAdding As Boolean, ByRef Cancel As Boolean)

'=========================================================================
' Private Members
'=========================================================================
Private m_hSci As Long                  ' Scintilla window handle
Private m_hSciLexer As Long             ' SciLexer.dll module handle
Private m_DirectPtr As Long             ' Direct function pointer
Private m_DirectPtrLo As Long           ' Low word of direct pointer
Private m_DirectPtrHi As Long           ' High word of direct pointer
Private m_Subclass As cSubclass         ' Subclass manager
Private m_Initialized As Boolean        ' Init flag
Private m_ActiveCallTipFunction As String
Private m_ActiveCallTipParenPos As Long
Private m_CallTipsEnabled As Boolean
Private m_WordHighlightEnabled As Boolean
Private m_WordHighlightIndicator As Long
Private m_LastHighlightedWord As String
Private m_ActiveCallTipObject As String  ' Track which object we're calling method on
Private m_LastAutocList As String  ' Add to private members

'=========================================================================
' Types for Scintilla Notifications
'=========================================================================
Private Type nmhdr
    hwndFrom As Long
    idFrom As Long
    code As Long
End Type

Private Type SCNotification
    nmhdr As nmhdr
    position As Long
    ch As Long
    modifiers As Long
    modificationType As Long
    text As Long
    length As Long
    linesAdded As Long
    message As Long
    wParam As Long
    lParam As Long
    line As Long
    foldLevelNow As Long
    foldLevelPrev As Long
    margin As Long
    listType As Long
    X As Long
    Y As Long
End Type

'=========================================================================
' Win32 API Declarations
'=========================================================================
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SendMessagePtr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal m As Long, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, ByVal flags As Long) As Long

'=========================================================================
' Constants
'=========================================================================
Private Const WS_CHILD = &H40000000
Private Const WS_VISIBLE = &H10000000
Private Const WS_TABSTOP = &H10000
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_VSCROLL = &H200000
Private Const WS_HSCROLL = &H100000
Private Const WS_EX_CLIENTEDGE = &H200
   
Private Const WM_NOTIFY = &H4E
Private Const WM_DESTROY = &H2
Private Const WM_SIZE = &H5
Private Const WM_SETFOCUS = &H7
Private Const WM_KEYDOWN = &H100
Private Const WM_CHAR = &H102

Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOCOPYBITS = &H100

' Scintilla message base
Private Const SCI_START = 2000

' Scintilla notification codes
Private Const SCN_STYLENEEDED = 2000
Private Const SCN_CHARADDED = 2001
Private Const SCN_SAVEPOINTREACHED = 2002
Private Const SCN_SAVEPOINTLEFT = 2003
Private Const SCN_MODIFYATTEMPTRO = 2004
Private Const SCN_KEY = 2005
Private Const SCN_DOUBLECLICK = 2006
Private Const SCN_UPDATEUI = 2007
Private Const SCN_MODIFIED = 2008
Private Const SCN_MACRORECORD = 2009
Private Const SCN_MARGINCLICK = 2010
Private Const SCN_NEEDSHOWN = 2011
Private Const SCN_DWELLSTART = 2016
Private Const SCN_DWELLEND = 2017
Private Const SCN_ZOOM = 2018
Private Const SCN_HOTSPOTCLICK = 2019
Private Const SCN_HOTSPOTDOUBLECLICK = 2020
Private Const SCN_CALLTIPCLICK = 2021
Private Const SCN_AUTOCSELECTION = 2022
Private Const SCN_INDICATORCLICK = 2023
Private Const SCN_INDICATORRELEASE = 2024
Private Const SCN_AUTOCCANCELLED = 2025
Private Const SCN_AUTOCCHARDELETED = 2026
Private Const SC_MARK_CIRCLE = 0
Private Const SC_MARK_ARROW = 2
Private Const SC_MARK_BACKGROUND = 22
Private Const SC_CP_UTF8 = 65001
Private Const SC_CP_DBCS = 1

Private Const STYLE_DEFAULT As Long = 32
Private Const STYLE_LINENUMBER As Long = 33
Private Const STYLE_BRACELIGHT As Long = 34
Private Const STYLE_BRACEBAD As Long = 35
   
' EOL mode
Private Const SC_EOL_CRLF = 0
Private Const SC_EOL_CR = 1
Private Const SC_EOL_LF = 2

' Message constants
Private Const SCI_GETDIRECTFUNCTION = 2184
Private Const SCI_GETDIRECTPOINTER = 2185
Private Const SCI_SETCODEPAGE = 2037
Private Const SCI_SETEOLMODE = 2031
Private Const SCI_SETCARETLINEVISIBLE = 2096
Private Const SCI_SETCARETLINEBACK = 2098
Private Const SCI_SETTABWIDTH = 2036
Private Const SCI_SETUSETABS = 2124

Private Const WM_GETDLGCODE = &H87
Private Const DLGC_WANTTAB = &H2
Private Const DLGC_WANTALLKEYS = &H4
Private Const DLGC_WANTARROWS = &H1

'=========================================================================
' Properties
'=========================================================================
Public Property Get SciHwnd() As Long
    SciHwnd = m_hSci
End Property

Public Property Get DirectPtr() As Long
    DirectPtr = m_DirectPtr
End Property

Public Property Get IsInitialized() As Boolean
    IsInitialized = m_Initialized
End Property

Public Property Get CallTipsEnabled() As Boolean
    CallTipsEnabled = m_CallTipsEnabled
End Property

Public Property Let CallTipsEnabled(ByVal value As Boolean)
    m_CallTipsEnabled = value
End Property

'not sure i can bear typing that out everytime..
Property Get bp() As CBreakPoint
    Set bp = breakpoints
End Property

'this is only used internally on margin click, not a user api
Private Sub ToggleBreakpoint(ByVal line As Long)
    Dim isAdding As Boolean, Cancel As Boolean
    
    isAdding = (breakpoints.isSet(line) = False)
    Cancel = False
    
    RaiseEvent UserBreakpointToggle(line, isAdding, Cancel)
    
    If Cancel Then Exit Sub  ' Host said no
    
    If isAdding Then
        breakpoints.add line
    Else
        breakpoints.remove line
    End If
End Sub

'=========================================================================
' Core Scintilla Message Interface
'=========================================================================
Public Function SciMsg(ByVal msg As Long, Optional ByVal wParam As Long = 0, Optional ByVal lParam As Long = 0) As Long
    If m_hSci = 0 Then Exit Function
    SciMsg = SendMessage(m_hSci, msg, wParam, lParam)
End Function

Public Function SciMsgStr(ByVal msg As Long, Optional ByVal wParam As Long = 0, Optional ByVal str As String = "") As Long
    If m_hSci = 0 Then Exit Function
    SciMsgStr = SendMessageStr(m_hSci, msg, wParam, str)
End Function

'this one is a utf8 unicode booby trap dont use...
Public Function SciMsgPtr(ByVal msg As Long, Optional ByVal wParam As Long = 0, Optional ByVal ptrVal As Long = 0) As Long
    If m_hSci = 0 Then Exit Function
    SciMsgPtr = SendMessagePtr(m_hSci, msg, wParam, ByVal ptrVal)
End Function

Public Function SciMsgStrUTF8(ByVal msg As Long, Optional ByVal wParam As Long = 0, Optional ByVal str As String = "") As Long
    Dim utf8Bytes() As Byte
    If m_hSci = 0 Then Exit Function
    If Len(str) = 0 Then
        SciMsgStrUTF8 = SendMessage(m_hSci, msg, wParam, 0)
    Else
        utf8Bytes = StrConv(str, vbFromUnicode)
        SciMsgStrUTF8 = SendMessagePtr(m_hSci, msg, wParam, ByVal VarPtr(utf8Bytes(0)))
    End If
End Function

Public Function SciSetProperty(ByVal propertyName As String, ByVal value As String) As Long
    Dim propName() As Byte, propValue() As Byte
    If m_hSci = 0 Then Exit Function
    propName = StrConv(propertyName, vbFromUnicode)
    propValue = StrConv(value, vbFromUnicode)
    SciSetProperty = SendMessagePtr(m_hSci, 4004, VarPtr(propName(0)), ByVal VarPtr(propValue(0)))
End Function

Public Function SciGetProperty(ByVal propertyName As String) As String
    Dim propName() As Byte, buffer() As Byte, length As Long
    If m_hSci = 0 Then Exit Function
    propName = StrConv(propertyName, vbFromUnicode)
    ReDim buffer(0 To 255) As Byte
    length = SendMessagePtr(m_hSci, 4008, VarPtr(propName(0)), ByVal VarPtr(buffer(0)))
    If length > 0 Then SciGetProperty = UTF8BytesToString(buffer, length)
End Function

'=========================================================================
' UserControl Events
'=========================================================================
Private Sub UserControl_Initialize()
    m_Initialized = False
End Sub

Private Sub UserControl_InitProperties()
    InitializeScintilla
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    InitializeScintilla
    UserControl.Extender.TabStop = False
End Sub

Private Sub UserControl_Resize()
    If m_hSci <> 0 Then
        'MoveWindow m_hSci, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 1
        On Error Resume Next
        SetWindowPos m_hSci, 0, 0, 0, UserControl.ScaleWidth / Screen.TwipsPerPixelX, UserControl.ScaleHeight / Screen.TwipsPerPixelY, SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
    End If
End Sub

Private Sub UserControl_Terminate()
    Shutdown
End Sub

'=========================================================================
' Initialization and Cleanup
'=========================================================================
Private Sub InitializeScintilla()
    Dim hInst As Long
    Dim style As Long
    
    If m_Initialized Then Exit Sub
    
    ' Load the Scintilla library
    m_hSciLexer = LoadLibrary("SciLexer.dll")
    If m_hSciLexer = 0 Then m_hSciLexer = LoadLibrary(App.path & "\SciLexer.dll")
    
    If m_hSciLexer = 0 Then
        Err.Raise 53, "SciWrapper", "SciLexer.dll not found. Please ensure it is in the application directory or system path."
        Exit Sub
    End If
    
    hInst = App.hInstance
    m_hSci = CreateWindowEx(WS_EX_CLIENTEDGE, "Scintilla", "sci4vb", WS_CHILD Or WS_VISIBLE, _
                            0, 0, 200, 200, _
                            UserControl.hwnd, 0, App.hInstance, 0)

    If m_hSci = 0 Then
        FreeLibrary m_hSciLexer
        m_hSciLexer = 0
        Err.Raise 5, "SciWrapper", "Failed to create Scintilla window."
        Exit Sub
    End If
    
    ' Get direct function pointer for faster message dispatch
    m_DirectPtr = SciMsg(SCI_GETDIRECTFUNCTION)
    m_DirectPtrLo = SciMsg(SCI_GETDIRECTPOINTER)
    
    ' Initialize subclassing
    Set m_Subclass = New cSubclass
    With m_Subclass
        .Subclass UserControl.hwnd, Me
        .AddMsg UserControl.hwnd, WM_NOTIFY, MSG_AFTER
        .AddMsg UserControl.hwnd, WM_SIZE, MSG_AFTER
        .AddMsg UserControl.hwnd, WM_SETFOCUS, MSG_AFTER
        .AddMsg UserControl.hwnd, WM_KEYUP, MSG_BEFORE
        .AddMsg UserControl.hwnd, WM_CHAR, MSG_BEFORE
        
        .Subclass m_hSci, Me
        .AddMsg m_hSci, WM_RBUTTONDOWN, MSG_AFTER
        .AddMsg m_hSci, WM_LBUTTONDOWN, MSG_AFTER
        .AddMsg m_hSci, WM_KEYDOWN, MSG_BEFORE '_AND_AFTER
        .AddMsg m_hSci, WM_KEYUP, MSG_BEFORE
        .AddMsg m_hSci, WM_LBUTTONUP, MSG_AFTER
        .AddMsg m_hSci, WM_RBUTTONUP, MSG_AFTER
        .AddMsg m_hSci, WM_CHAR, MSG_BEFORE
        .AddMsg m_hSci, WM_COMMAND, MSG_BEFORE
        .AddMsg m_hSci, WM_GETDLGCODE, MSG_BEFORE
    End With
    
    InitializeModules
    ApplyDefaults
    UserControl_Resize
    m_Initialized = True
    
End Sub

Private Sub InitializeModules()
    ' Pass control reference to all feature modules
    breakpoints.Init Me
    doc.Init Me
    edit.Init Me
    sel.Init Me
    lines.Init Me
    style.Init Me
    Fold.Init Me
    Mark.Init Me
    Indic.Init Me
    Autoc.Init Me
    Margins.Init Me
    View.Init Me
    Search.Init Me
    Helper.Init Me
End Sub

Private Sub ApplyDefaults()
    ' Set sensible defaults
    SciMsg SCI_SETCODEPAGE, SC_CP_UTF8
    SciMsg SCI_SETEOLMODE, SC_EOL_CRLF
    SciMsg SCI_SETTABWIDTH, 4
    SciMsg SCI_SETUSETABS, 0  ' Use spaces by default
    
    m_CallTipsEnabled = True  ' Enable by default
    m_WordHighlightEnabled = True
    m_WordHighlightIndicator = 8
    
    ' Configure indicator
    Indic.Current = 8
    Indic.SetStyle 8, indicRoundBox
    Indic.SetFore 8, RGB(255, 255, 0)  ' Yellow
    Indic.SetAlpha 8, 100
    Indic.SetUnder 8, True
    
    ' Setup margins for line numbers, folding, and breakpoints
    Margins.ConfigureLineNumbers 0, 40
    Margins.ConfigureFolding 2, 16
    
    Margins.SetSensitive 0, True
    Margins.SetSensitive 1, True
    Margins.SetSensitive 2, True
    
    ' Visual settings
    View.CaretLineVisible = False 'this hides our debugger yellow background
    View.CaretLineBack = &HE8E8E8
    View.EdgeMode = edgeLine
    View.EdgeColumn = 80
    View.EdgeColor = &HE0E0E0
    View.HScrollBar = True
    View.VScrollBar = True
    
    ' Editor settings
    edit.TabWidth = 4
    edit.UseTabs = False
    edit.TabIndents = True
    edit.BackspaceUnindents = True
    
    ' Setup default styling
    style.SetFont 32, "Consolas"  ' STYLE_DEFAULT
    style.SetSize 32, 10
    style.clearAll

    ' Autocomplete settings
    Autoc.IgnoreCase = True
    Autoc.AutoHide = True
    
    style.SetLanguage langJavaScript
    
    ' Set brace match style (bold, colored)
    style.SetFore 34, &HFF0000  ' STYLE_BRACELIGHT - Red
    style.SetBold 34, True
    style.SetBack 34, &HFFFFFF  ' White background
    
    ' Set unmatched brace style
    style.SetFore 35, &HFF      ' STYLE_BRACEBAD - Blue
    style.SetBold 35, True
    style.SetBack 35, &HFFFF00  ' Yellow background
       
    doc.PasteConvertEndings True
    doc.WordWrap = False
    doc.Folding = False
    doc.LineNumbers = True
    doc.ReadOnly = False
 
    With Mark
        .Define 2, SC_MARK_ARROW
        .SetFore 2, vbBlack  'current eip
        .SetBack 2, vbYellow
    
        ' Setup margin 1 for breakpoints (shown by default)
        Margins.SetType 1, marginSymbol
        Margins.SetWidth 1, 16
        Margins.SetSensitive 1, True
        Margins.SetMask 1, (2 Or 4 Or 8)  ' marker 1 (bit 1) + marker 2 (bit 2) + marker 3 (bit 3)
        
        ' Configure breakpoint marker appearance
        .Define BREAKPOINT_MARKER, markerCircle  'BREAKPOINT_MARKER=1
        .SetFore BREAKPOINT_MARKER, &HFFFFFF  ' White
        .SetBack BREAKPOINT_MARKER, &HFF      ' Red
    
        .Define 3, SC_MARK_BACKGROUND
        .SetFore 3, vbBlack 'current eip
        .SetBack 3, vbYellow
    End With
        
End Sub

Private Sub Shutdown()
    On Error Resume Next
    
    If Not m_Subclass Is Nothing Then
        m_Subclass.UnSubAll
        Set m_Subclass = Nothing
    End If
    
    If m_hSci <> 0 Then
        DestroyWindow m_hSci
        m_hSci = 0
    End If
    
    If m_hSciLexer <> 0 Then
        FreeLibrary m_hSciLexer
        m_hSciLexer = 0
    End If
    
    m_Initialized = False
End Sub

'=========================================================================
' Subclass Message Handler
'=========================================================================
Private Sub iSubclass_WndProc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    On Error Resume Next
    
    Dim word As String
    Dim wasHandled As Boolean

     If uMsg = WM_GETDLGCODE Then
        lReturn = DLGC_WANTALLKEYS Or DLGC_WANTARROWS Or DLGC_WANTTAB
        bHandled = True
        Exit Sub
    End If
    
    If uMsg = WM_NOTIFY Then
        HandleNotification lParam
    ElseIf uMsg = WM_SIZE Then
        UserControl_Resize
    ElseIf uMsg = WM_SETFOCUS Then
        If m_hSci <> 0 Then SetFocus m_hSci
    ElseIf uMsg = WM_LBUTTONUP Then
        If m_WordHighlightEnabled Then HandleSelectionHighlight
        
    ElseIf uMsg = WM_KEYDOWN Then 'happens before WM_CHAR, wasHandled ineffectual cant eat here..
                                           
        RaiseEvent KeyPressed(wParam, GetModifierState(), wasHandled)

        If Not wasHandled Then
            If IsCtrlOnly() Then
                If wParam = Asc("F") Or wParam = Asc("H") Then frmReplace.LaunchReplaceForm Me
                If Asc("G") = wParam Then Call ShowGoto
            End If
        Else
            ' Parent handled it - eat the message
            bHandled = True
            lReturn = 0
        End If
                    
    ElseIf uMsg = WM_KEYUP Then
    
        If wParam = 190 Then ' period key
            word = Helper.GetWordBeforeDot()
            If Len(word) > 0 Then
                RaiseEvent AutoCompleteEvent(word)
            End If
        End If

    ElseIf uMsg = WM_CHAR Then

        'RaiseEvent CharAdded(wParam) handled in handlenotification from sci, but no special keys like F7 there
        
        If IsCtrlOnly() Then
        
            'eat these control messages - normal Ctrl A/C/X/Y/Z handled by sci internally
            'Ctrl+F = 0x06 (ACK - Acknowledge), Ctrl+G = 0x07 (BEL - Bell), Ctrl+H = 0x08 (BS - Backspace)
            If wParam = 6 Or wParam = 7 Or wParam = 8 Then GoTo wasHandled

            If wParam = 32 Then  'ctrl + space
                If Intellisense.isLoaded Then HandleCtrlSpace
                word = Helper.WordAtCaret()
                RaiseEvent AutoCompleteEvent(word) 'in case user implements their own intellisense design
                GoTo wasHandled
            End If
            
        End If
        
    End If
    
    
Exit Sub
wasHandled:
    bHandled = True
    lReturn = 0
    wParam = 0
    
End Sub

Public Sub ShowGoto()
    On Error Resume Next
    Dim sline As Long
    Dim line As Long
    sline = Trim(InputBox("Goto Line:"))
    If Len(sline) <> 0 Then
        line = CLng(sline)
        If Err.Number = 0 Then lines.GotoLineCentered line
    End If
End Sub

Private Sub HandleNotification(ByVal lParam As Long)
    Dim scn As SCNotification
    Dim hdr As nmhdr
    Dim txt As String
    Dim line As Long, col As Long, pos As Long
     
    On Error Resume Next
    
    ' Copy notification header
    CopyMemory hdr, ByVal lParam, Len(hdr)
    
    ' Only handle notifications from our Scintilla control
    If hdr.hwndFrom <> m_hSci Then Exit Sub
    
    ' Copy full notification structure
    CopyMemory scn, ByVal lParam, Len(scn)
    
    ' Dispatch to appropriate event
    Select Case hdr.code
        Case SCN_UPDATEUI
        
            If Autoc.IsCallTipActive() Then
                UpdateCallTipHighlightInternal
            End If

            pos = sel.currentPos
            line = lines.LineFromPosition(pos)
            col = lines.GetColumn(pos)
            HighlightMatchingBrace pos
    
            RaiseEvent UpdateUI(scn.nmhdr.code)
            
        Case SCN_MODIFIED
            If scn.text <> 0 Then
                txt = String$(scn.length, 0)
                CopyMemory ByVal txt, ByVal scn.text, scn.length
            End If
            RaiseEvent Modified(scn.position, scn.modificationType, txt, scn.length, scn.linesAdded)

        Case SCN_CHARADDED
            HandleCallTips scn.ch        ' Call tips first
            If Intellisense.isLoaded Then HandleIntellisense scn.ch
            RaiseEvent CharAdded(scn.ch)
            
        Case SCN_SAVEPOINTREACHED
            RaiseEvent SavePointReached
            
        Case SCN_SAVEPOINTLEFT
            RaiseEvent SavePointLeft
            
        Case SCN_MARGINCLICK
            ' Handle breakpoint toggle on margin 1
            If scn.margin = 1 Then
                line = lines.LineFromPosition(scn.position)
                ToggleBreakpoint line
            End If
            RaiseEvent MarginClick(scn.margin, scn.position, scn.modifiers)
            
        Case SCN_STYLENEEDED
            RaiseEvent StyleNeeded(scn.position)
            
        Case SCN_DOUBLECLICK
            If m_WordHighlightEnabled Then HandleWordHighlight
            RaiseEvent DoubleClick(scn.position, scn.line)
            
        Case SCN_DWELLSTART
            RaiseEvent DwellStart(scn.position, scn.X, scn.Y)
            
        Case SCN_DWELLEND
            RaiseEvent DwellEnd(scn.position, scn.X, scn.Y)
            
        Case SCN_AUTOCSELECTION
            If scn.text <> 0 Then
                txt = String$(scn.length, 0)
                CopyMemory ByVal txt, ByVal scn.text, scn.length
            End If
            RaiseEvent AutoCSelection(txt, scn.position)
            
        Case SCN_AUTOCCANCELLED
            RaiseEvent AutoCCancelled
            
        Case SCN_CALLTIPCLICK
            RaiseEvent CallTipClick(scn.position)
            
        Case Else
            RaiseEvent Notify(hdr.code)
    End Select
    
End Sub
 
Private Sub HandleIntellisense(ch As Long)

    Dim currentWord As String
    Dim prevWord As String
    Dim matches As String
    Dim parenPos As Long
    Dim funcName As String
    Dim tip As String
    
    If Not Autoc.isActive() Then m_LastAutocList = Empty
    
    If Chr$(ch) = "." Then
        ' Show object members when dot is typed
        prevWord = Helper.PreviousWord()
        'currentWord = sci.Helper.WordAtCaret()
        
        ' Get methods for this object from Intellisense manager
        matches = Intellisense.GetMethods(prevWord)

        If Len(matches) > 0 Then
             m_LastAutocList = matches
             m_ActiveCallTipObject = prevWord
             Autoc.show 0, matches
'            wont work to soon?
'            If Len(currentWord) > 0 And InStr(1, matches, currentWord, vbTextCompare) > 0 Then
'                SciMsgStr 2111, 0, currentWord     ' SCI_AUTOCSELECT2208
'            End If
        Else
            ' No matches - cancel autocomplete
            m_LastAutocList = Empty
        End If
        
    ElseIf IsAlphaNumeric(Chr$(ch)) Then
    
        'If Autoc.isActive() Then Exit Sub
        
        ' Show filtered list while typing
        currentWord = Helper.WordAtCaret()
        
        ' Check if we're after a dot (filtering object members)
        Dim pos As Long
        pos = sel.currentPos - Len(currentWord) - 1
        
        If pos > 0 And Mid$(doc.text, pos + 1, 1) = "." Then
            ' Filtering object members
            prevWord = Intellisense.GetObjectBeforeDot(doc.text, pos + 1)
            matches = Intellisense.FindMethods(prevWord, currentWord)
            
            If Len(matches) > 0 Then
                If matches <> m_LastAutocList Then
                    m_LastAutocList = matches
                    Autoc.show Len(currentWord), matches
                Else
                    'Debug.Print "autoc show skipped no change: " & matches
                End If
            Else
                ' No matches - cancel autocomplete
                m_LastAutocList = Empty
                Autoc.Cancel
            End If
        End If
    End If
    
End Sub

Private Sub HandleCallTips(ByVal ch As Long)

    If Not m_CallTipsEnabled Then Exit Sub
    
    Select Case ch
        Case Asc("(")
            ShowCallTipInternal
        Case Asc(")")
            HideCallTipInternal
        Case Else
            If Autoc.IsCallTipActive() Then
                UpdateCallTipHighlightInternal
            End If
    End Select
    
End Sub

Private Sub ShowCallTipInternal()
    Dim text As String, currentPos As Long, parenPos As Long
    Dim funcName As String, tipText As String
    
    On Error Resume Next
    currentPos = sel.currentPos
    text = doc.text
    parenPos = currentPos '- 1
    
    funcName = Intellisense.GetFunctionNameBeforeParen(text, parenPos)
    If Len(funcName) = 0 Then Exit Sub
    
    'tipText = Intellisense.GetCallTip(funcName)
    tipText = Intellisense.GetCallTipEx(funcName, m_ActiveCallTipObject)
    If Len(tipText) = 0 Then Exit Sub
    
    m_ActiveCallTipFunction = funcName
    m_ActiveCallTipParenPos = parenPos
    
    Autoc.ShowCallTip parenPos, tipText
    SetParameterHighlightInternal tipText, 0
End Sub

Private Sub UpdateCallTipHighlightInternal()
    Dim text As String, currentPos As Long, openParen As Long
    Dim argIndex As Long, tipText As String
    
    On Error Resume Next
    text = doc.text
    currentPos = sel.currentPos
    openParen = Intellisense.FindOpenParen(text, currentPos)
    
    If openParen <> m_ActiveCallTipParenPos Then
HideCallTipInternal:         Exit Sub
    End If
    
    argIndex = Intellisense.GetCurrentArgIndex(text, m_ActiveCallTipParenPos, currentPos)
    'tipText = Intellisense.GetCallTip(m_ActiveCallTipFunction)
    tipText = Intellisense.GetCallTipEx(m_ActiveCallTipFunction, m_ActiveCallTipFunction)
    If Len(tipText) > 0 Then SetParameterHighlightInternal tipText, argIndex
End Sub

Private Sub HideCallTipInternal()
    On Error Resume Next
    Autoc.CancelCallTip
    m_ActiveCallTipFunction = ""
    m_ActiveCallTipParenPos = 0
    m_LastAutocList = Empty
End Sub

Private Sub SetParameterHighlightInternal(ByVal tipText As String, ByVal argIndex As Long)
    Dim parenStart As Long, parenEnd As Long, Params As String
    Dim paramAry() As String, i As Long, highlightStart As Long
    Dim highlightEnd As Long, leadingSpaces As Long
    
    On Error Resume Next
    parenStart = InStr(tipText, "(")
    parenEnd = InStr(parenStart, tipText, ")")
    If parenStart = 0 Or parenEnd = 0 Then Exit Sub
    
    Params = Trim$(Mid$(tipText, parenStart + 1, parenEnd - parenStart - 1))
    If Len(Params) = 0 Then Exit Sub
    
    paramAry = Split(Params, ",")
    If argIndex < 0 Or argIndex > UBound(paramAry) Then Exit Sub
    
    highlightStart = parenStart
    For i = 0 To argIndex - 1
        highlightStart = highlightStart + Len(paramAry(i)) + 1
    Next
    
    leadingSpaces = Len(paramAry(argIndex)) - Len(LTrim$(paramAry(argIndex)))
    highlightStart = highlightStart + leadingSpaces
    highlightEnd = highlightStart + Len(Trim$(paramAry(argIndex)))
    
    Autoc.SetCallTipHighlight highlightStart, highlightEnd
End Sub

Private Sub HandleCtrlSpace()
    Dim currentWord As String, prevWord As String
    Dim iStart As Long, iEnd As Long, methods As String
    Dim matches As String
    
    currentWord = Helper.currentWord(iStart, iEnd)
    prevWord = Helper.PreviousWord()
    
     If Len(currentWord) = 0 Then 'show top level objects
        matches = Intellisense.GetObjectNames()
        If Len(matches) > 0 Then
            Autoc.show 0, matches
            Exit Sub
        End If
     Else
        matches = Intellisense.GetObjectNames(currentWord)
        If Len(matches) > 0 Then
            Autoc.show Len(currentWord), matches
            Exit Sub
        End If
     End If
     
    ' Get methods for the object
    methods = Intellisense.GetMethods(prevWord)
    If Len(methods) = 0 Then Exit Sub
    
    If Len(currentWord) > 0 Then
        SmartComplete currentWord, methods, iStart, iEnd
    Else
        Autoc.show 0, methods
    End If
End Sub

Private Sub SmartComplete(ByVal currentWord As String, ByVal methodList As String, ByVal wordStart As Long, ByVal wordEnd As Long)
    Dim methods() As String, matches() As String
    Dim matchCount As Long, i As Long, lowerCurrent As String
    
    methods = Split(Trim$(methodList), " ")
    lowerCurrent = LCase$(currentWord)
    matchCount = 0
    ReDim matches(0 To UBound(methods))
    
    ' Find matches
    For i = 0 To UBound(methods)
        If Len(methods(i)) > 0 Then
            If LCase$(methods(i)) = lowerCurrent Or _
               Left$(LCase$(methods(i)), Len(currentWord)) = lowerCurrent Then
                matches(matchCount) = methods(i)
                matchCount = matchCount + 1
            End If
        End If
    Next
    
    If matchCount = 1 Then
        ' Single match - auto-complete!
        AutoCompleteWord matches(0), wordStart, wordEnd
    ElseIf matchCount > 1 Then
        ' Multiple matches - show filtered list
        Dim filtered As String
        For i = 0 To matchCount - 1
            If i > 0 Then filtered = filtered & " "
            filtered = filtered & matches(i)
        Next
        Autoc.show Len(currentWord), filtered
    Else
        ' No matches - show all
        Autoc.show Len(currentWord), methodList
    End If
End Sub

Private Sub AutoCompleteWord(ByVal correctWord As String, ByVal wordStart As Long, ByVal wordEnd As Long)
    Dim currentLine As Long, lineStart As Long
    
    Autoc.Cancel  ' Close any open autocomplete
    m_LastAutocList = Empty
    
    currentLine = lines.LineFromPosition(sel.currentPos)
    lineStart = lines.PositionFromLine(currentLine)
    
    ' Replace word with case-corrected version
    sel.SetSelection wordStart, wordEnd
    sel.ReplaceSelection correctWord
    sel.currentPos = wordStart + Len(correctWord)
    sel.anchor = sel.currentPos
End Sub

Private Sub HandleWordHighlight()
    Dim word As String
    Dim currentCaretPos As Long
    
    word = Helper.WordAtCaret()
    
    If Len(word) = 0 Or Len(word) > 20 Or word = m_LastHighlightedWord Then Exit Sub
    
    ' Save caret position
    currentCaretPos = sel.currentPos
    
    ' Highlight all instances of the word
    Helper.HighlightWord word, m_WordHighlightIndicator
    
    ' Clear any selection - just leave caret at current position
    sel.SetEmptySelection currentCaretPos
    
    m_LastHighlightedWord = word
    RaiseEvent WordHighlighted(word, CountInstances(word))
End Sub

Private Sub HandleSelectionHighlight()
    Dim selectedText As String
    Dim currentCaretPos As Long
    
    selectedText = sel.GetSelectedText()
    
    If Len(selectedText) = 0 Or Len(selectedText) > 20 Or selectedText = m_LastHighlightedWord Then Exit Sub
    
    ' Save caret position
    currentCaretPos = sel.currentPos
    
    ' Highlight all instances of the selected text
    Helper.HighlightWord selectedText, m_WordHighlightIndicator
    
    ' Clear any selection - just leave caret at current position
    'sel.SetEmptySelection currentCaretPos
    
    m_LastHighlightedWord = selectedText
    RaiseEvent WordHighlighted(selectedText, CountInstances(selectedText))
End Sub

Private Function CountInstances(ByVal word As String) As Long
    Dim pos As Long, Count As Long
    pos = 0
    Do
        pos = Search.Find(word, pos, False, True, False, False)
        If pos >= 0 Then
            Count = Count + 1
            pos = pos + Len(word)
        End If
    Loop While pos >= 0 And pos < doc.TextLength
    CountInstances = Count
End Function

Public Sub ShowObjectBrowser()
    frmObjBrowser.Init Intellisense.Objects
    frmObjBrowser.show vbModeless
End Sub

Function ShowFind(Optional ByVal findText As String) As Object
    frmReplace.LaunchReplaceForm Me
    If Len(findText) > 0 Then frmReplace.SetFindText findText
    Set ShowFind = frmReplace
End Function

Sub showAbout()
    frmAbout.LaunchForm Me
End Sub

Private Sub HighlightMatchingBrace(ByVal pos As Long)
    Dim bracePos As Long
    Dim charAtPos As String
    Dim charBefore As String
    
    ' Clear any previous brace highlighting
    SciMsg 2351, -1, -1   ' SCI_BRACEHIGHLIGHT with -1 clears
    
    ' Check character at current position
    If pos < doc.TextLength Then
        charAtPos = Chr$(SciMsg(2007, pos, 0))  ' SCI_GETCHARAT
        If IsBrace(charAtPos) Then
            bracePos = SciMsg(2353, pos, 0)  ' SCI_BRACEMATCH
            If bracePos >= 0 Then
                SciMsg 2351, pos, bracePos   ' SCI_BRACEHIGHLIGHT
                Exit Sub
            Else
                SciMsg 2352, pos, 0  ' SCI_BRACEBADLIGHT (no match)
                Exit Sub
            End If
        End If
    End If
    
    ' Check character before cursor
    If pos > 0 Then
        charBefore = Chr$(SciMsg(2007, pos - 1, 0))
        If IsBrace(charBefore) Then
            bracePos = SciMsg(2353, pos - 1, 0)
            If bracePos >= 0 Then
                SciMsg 2351, pos - 1, bracePos
            Else
                SciMsg 2352, pos - 1, 0
            End If
        End If
    End If
End Sub

Private Function IsBrace(ch As String) As Boolean
    IsBrace = (InStr("(){}[]<>", ch) > 0)
End Function

Private Function IsAlphaNumeric(ch As String) As Boolean
    Dim code As Integer
    code = Asc(ch)
    IsAlphaNumeric = (code >= 48 And code <= 57) Or _
                     (code >= 65 And code <= 90) Or _
                     (code >= 97 And code <= 122) Or _
                     code = 95  ' underscore
End Function



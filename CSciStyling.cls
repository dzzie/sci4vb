VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSciStyling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Author:  David Zimmer <dzzie@yahoo.com> + Claude.ai
'Site:    http://sandsprite.com
'License: MIT
'---------------------------------------------------


'=========================================================================
' CSciStyling - Styling and lexer configuration
'=========================================================================
Option Explicit

Private m_Sci As SciWrapper

' Scintilla messages
Private Const SCI_SETLEXER = 4001
Private Const SCI_GETLEXER = 4002
Private Const SCI_SETLEXERLANGUAGE = 4006
Private Const SCI_COLOURISE = 4003
Private Const SCI_SETKEYWORDS = 4005
Private Const SCI_SETPROPERTY = 4004
Private Const SCI_GETPROPERTY = 4008
Private Const SCI_STYLECLEARALL = 2050
Private Const SCI_STYLERESETDEFAULT = 2058
Private Const SCI_STYLESETFORE = 2051
Private Const SCI_STYLEGETFORE = 2481
Private Const SCI_STYLESETBACK = 2052
Private Const SCI_STYLEGETBACK = 2482
Private Const SCI_STYLESETBOLD = 2053
Private Const SCI_STYLEGETBOLD = 2483
Private Const SCI_STYLESETITALIC = 2054
Private Const SCI_STYLEGETITALIC = 2484
Private Const SCI_STYLESETSIZE = 2055
Private Const SCI_STYLEGETSIZE = 2485
Private Const SCI_STYLESETFONT = 2056
Private Const SCI_STYLEGETFONT = 2486
Private Const SCI_STYLESETEOLFILLED = 2057
Private Const SCI_STYLEGETEOLFILLED = 2487
Private Const SCI_STYLESETUNDERLINE = 2059
Private Const SCI_STYLEGETUNDERLINE = 2488
Private Const SCI_STYLESETCASE = 2060
Private Const SCI_STYLEGETCASE = 2489
Private Const SCI_STYLESETCHARACTERSET = 2066
Private Const SCI_STYLEGETCHARACTERSET = 2490
Private Const SCI_STYLESETVISIBLE = 2074
Private Const SCI_STYLEGETVISIBLE = 2491
Private Const SCI_STYLESETCHANGEABLE = 2099
Private Const SCI_STYLEGETCHANGEABLE = 2492
Private Const SCI_STYLESETHOTSPOT = 2409
Private Const SCI_STYLEGETHOTSPOT = 2493

' Lexer constants
Public Enum eLexer
    lexNull = 0
    lexPython = 2
    lexCPP = 3
    lexHTML = 4
    lexXML = 5
    lexPerl = 6
    lexSQL = 7
    lexVB = 8
    lexProperties = 9
    lexErrorList = 10
    lexMakefile = 11
    lexBatch = 12
    lexPascal = 18
    lexLua = 15
    lexDiff = 16
    lexBash = 62
    lexJavaScript = 58
    lexJSON = 120
    lexMarkdown = 22
End Enum

' Language preset enum
Public Enum eLanguage
    langNone = 0
    langVB = 1
    langJavaScript = 2
    langSQL = 3
    langPython = 4
    langHTML = 5
    langVBPCode = 6
End Enum

' Case styles
Public Enum eStyleCase
    caseMixed = 0
    caseUpper = 1
    caseLower = 2
    caseCamel = 3
End Enum

Friend Sub Init(sciCtrl As SciWrapper)
    Set m_Sci = sciCtrl
End Sub

'=========================================================================
' Lexer Configuration
'=========================================================================
Public Property Get Lexer() As eLexer
    Lexer = m_Sci.SciMsg(SCI_GETLEXER)
End Property

Public Property Let Lexer(ByVal value As eLexer)
    m_Sci.SciMsg SCI_SETLEXER, value
End Property

Public Sub SetLexerLanguage(ByVal language As String)
    m_Sci.SciMsgStr SCI_SETLEXERLANGUAGE, 0, language
End Sub

Public Sub Colorise(ByVal startPos As Long, ByVal endPos As Long)
    m_Sci.SciMsg SCI_COLOURISE, startPos, endPos
End Sub

Public Sub SetKeywords(ByVal keywordSet As Long, ByVal KEYWORDS As String)
    m_Sci.SciMsgStr SCI_SETKEYWORDS, keywordSet, KEYWORDS
End Sub

Public Sub SetProperty(ByVal key As String, ByVal value As String)
    ' Property setting requires special handling
    Dim combinedStr As String
    combinedStr = key & Chr$(0) & value
    m_Sci.SciMsgStr SCI_SETPROPERTY, 0, combinedStr
End Sub

Public Function GetProperty(ByVal key As String) As String
    Dim buffer As String
    buffer = String$(256, 0)
    m_Sci.SciMsgStr SCI_GETPROPERTY, 0, key
    GetProperty = buffer
End Function

'=========================================================================
' Style Management
'=========================================================================
Public Sub clearAll()
    m_Sci.SciMsg SCI_STYLECLEARALL
End Sub

Public Sub ResetDefault()
    m_Sci.SciMsg SCI_STYLERESETDEFAULT
End Sub

'=========================================================================
' Style Foreground and Background
'=========================================================================
Public Sub SetFore(ByVal style As Long, ByVal color As Long)
    m_Sci.SciMsg SCI_STYLESETFORE, style, color
End Sub

Public Function GetFore(ByVal style As Long) As Long
    GetFore = m_Sci.SciMsg(SCI_STYLEGETFORE, style)
End Function

Public Sub SetBack(ByVal style As Long, ByVal color As Long)
    m_Sci.SciMsg SCI_STYLESETBACK, style, color
End Sub

Public Function GetBack(ByVal style As Long) As Long
    GetBack = m_Sci.SciMsg(SCI_STYLEGETBACK, style)
End Function

'=========================================================================
' Font Attributes
'=========================================================================
Public Sub SetBold(ByVal style As Long, ByVal bold As Boolean)
    m_Sci.SciMsg SCI_STYLESETBOLD, style, IIf(bold, 1, 0)
End Sub

Public Function GetBold(ByVal style As Long) As Boolean
    GetBold = (m_Sci.SciMsg(SCI_STYLEGETBOLD, style) <> 0)
End Function

Public Sub SetItalic(ByVal style As Long, ByVal italic As Boolean)
    m_Sci.SciMsg SCI_STYLESETITALIC, style, IIf(italic, 1, 0)
End Sub

Public Function GetItalic(ByVal style As Long) As Boolean
    GetItalic = (m_Sci.SciMsg(SCI_STYLEGETITALIC, style) <> 0)
End Function

Public Sub SetUnderline(ByVal style As Long, ByVal underline As Boolean)
    m_Sci.SciMsg SCI_STYLESETUNDERLINE, style, IIf(underline, 1, 0)
End Sub

Public Function GetUnderline(ByVal style As Long) As Boolean
    GetUnderline = (m_Sci.SciMsg(SCI_STYLEGETUNDERLINE, style) <> 0)
End Function

Public Sub SetSize(ByVal style As Long, ByVal sizePoints As Long)
    m_Sci.SciMsg SCI_STYLESETSIZE, style, sizePoints
End Sub

Public Function GetSize(ByVal style As Long) As Long
    GetSize = m_Sci.SciMsg(SCI_STYLEGETSIZE, style)
End Function

Public Sub SetFont(ByVal style As Long, ByVal fontName As String)
    m_Sci.SciMsgStr SCI_STYLESETFONT, style, fontName
End Sub

Public Function GetFont(ByVal style As Long) As String
    Dim buffer As String
    buffer = String$(256, 0)
    m_Sci.SciMsgStr SCI_STYLEGETFONT, style, buffer
    GetFont = Left$(buffer, InStr(buffer, Chr$(0)) - 1)
End Function

'=========================================================================
' Style Attributes
'=========================================================================
Public Sub SetEOLFilled(ByVal style As Long, ByVal filled As Boolean)
    m_Sci.SciMsg SCI_STYLESETEOLFILLED, style, IIf(filled, 1, 0)
End Sub

Public Function GetEOLFilled(ByVal style As Long) As Boolean
    GetEOLFilled = (m_Sci.SciMsg(SCI_STYLEGETEOLFILLED, style) <> 0)
End Function

Public Sub SetCase(ByVal style As Long, ByVal caseStyle As eStyleCase)
    m_Sci.SciMsg SCI_STYLESETCASE, style, caseStyle
End Sub

Public Function GetCase(ByVal style As Long) As eStyleCase
    GetCase = m_Sci.SciMsg(SCI_STYLEGETCASE, style)
End Function

Public Sub SetVisible(ByVal style As Long, ByVal Visible As Boolean)
    m_Sci.SciMsg SCI_STYLESETVISIBLE, style, IIf(Visible, 1, 0)
End Sub

Public Function GetVisible(ByVal style As Long) As Boolean
    GetVisible = (m_Sci.SciMsg(SCI_STYLEGETVISIBLE, style) <> 0)
End Function

Public Sub SetChangeable(ByVal style As Long, ByVal changeable As Boolean)
    m_Sci.SciMsg SCI_STYLESETCHANGEABLE, style, IIf(changeable, 1, 0)
End Sub

Public Function GetChangeable(ByVal style As Long) As Boolean
    GetChangeable = (m_Sci.SciMsg(SCI_STYLEGETCHANGEABLE, style) <> 0)
End Function

Public Sub SetHotspot(ByVal style As Long, ByVal hotspot As Boolean)
    m_Sci.SciMsg SCI_STYLESETHOTSPOT, style, IIf(hotspot, 1, 0)
End Sub

Public Function GetHotspot(ByVal style As Long) As Boolean
    GetHotspot = (m_Sci.SciMsg(SCI_STYLEGETHOTSPOT, style) <> 0)
End Function

'=========================================================================
' Quick Configuration Helpers
'=========================================================================
Public Sub ConfigureStyle(ByVal style As Long, ByVal foreColor As Long, ByVal backColor As Long, Optional ByVal fontName As String = "", Optional ByVal fontSize As Long = 0, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False)
    SetFore style, foreColor
    SetBack style, backColor
    If Len(fontName) > 0 Then SetFont style, fontName
    If fontSize > 0 Then SetSize style, fontSize
    SetBold style, bold
    SetItalic style, italic
End Sub

'=========================================================================
' Language Preset Wrapper
'=========================================================================
Public Function SetLanguage(ByVal language As eLanguage, Optional ByVal additionalKeywords As String = "") As Boolean
    ' Simple wrapper to apply predefined language configurations
    ' Uses the ModHighlighter module functions
    ' note the fold properties are specific to language lexer
    
    On Error GoTo ErrorHandler
    
    Select Case language
        Case langVB
            SetLanguage = SetVBHighlighter(m_Sci)
            ' VB folding
            m_Sci.SciSetProperty "fold", "1"
            m_Sci.SciSetProperty "fold.compact", "0"
            m_Sci.SciSetProperty "fold.comment", "1"
            m_Sci.SciSetProperty "fold.vb.preprocessor", "1"
            
        Case langJavaScript
            SetLanguage = SetJavaHighlighter(m_Sci)
            ' C/C++/JavaScript folding
            m_Sci.SciSetProperty "fold", "1"
            m_Sci.SciSetProperty "fold.compact", "0"
            m_Sci.SciSetProperty "fold.comment", "1"
            m_Sci.SciSetProperty "fold.preprocessor", "1"
            m_Sci.SciSetProperty "fold.at.else", "1"
            
        Case langSQL
            SetLanguage = SetSqlHighlighter(m_Sci)
            ' SQL folding
            m_Sci.SciSetProperty "fold", "1"
            m_Sci.SciSetProperty "fold.compact", "0"
            m_Sci.SciSetProperty "fold.comment", "1"
            m_Sci.SciSetProperty "fold.sql.only.begin", "0"
            
        Case langPython
            SetLanguage = SetPythonHighlighter(m_Sci)
            ' Python folding
            m_Sci.SciSetProperty "fold", "1"
            m_Sci.SciSetProperty "fold.compact", "0"
            m_Sci.SciSetProperty "fold.comment", "1"
            m_Sci.SciSetProperty "fold.quotes.python", "1"
            
        Case langHTML
            SetLanguage = SetHTMLHighlighter(m_Sci)
            ' HTML folding
            m_Sci.SciSetProperty "fold", "1"
            m_Sci.SciSetProperty "fold.compact", "0"
            m_Sci.SciSetProperty "fold.html", "1"
            m_Sci.SciSetProperty "fold.html.preprocessor", "1"
            m_Sci.SciSetProperty "fold.hypertext.comment", "1"
            m_Sci.SciSetProperty "fold.hypertext.heredoc", "1"
            
        Case langVBPCode
            SetVBPCodeHighlighter m_Sci, additionalKeywords
            ' VBP uses VB lexer, so same folding as VB
            m_Sci.SciSetProperty "fold", "1"
            m_Sci.SciSetProperty "fold.compact", "0"
            m_Sci.SciSetProperty "fold.comment", "1"
            m_Sci.SciSetProperty "fold.vb.preprocessor", "1"
            SetLanguage = True
            
        Case langNone
            Lexer = lexNull
            clearAll
            SetLanguage = True
            
        Case Else
            SetLanguage = False
    End Select
    
    Exit Function
    
ErrorHandler:
    SetLanguage = False
End Function



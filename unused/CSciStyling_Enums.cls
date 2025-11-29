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

'=========================================================================
' Scintilla Style Constants - Common across all lexers
'=========================================================================
Public Enum eStyleCommon
    STYLE_DEFAULT = 32
    STYLE_LINENUMBER = 33
    STYLE_BRACELIGHT = 34
    STYLE_BRACEBAD = 35
    STYLE_CONTROLCHAR = 36
    STYLE_INDENTGUIDE = 37
    STYLE_CALLTIP = 38
    STYLE_FOLDDISPLAYTEXT = 39
    STYLE_LASTPREDEFINED = 39
End Enum

'=========================================================================
' CPP/C/C++/JavaScript Lexer Styles (SCLEX_CPP)
'=========================================================================
Public Enum eStyleCPP
    SCE_C_DEFAULT = 0
    SCE_C_COMMENT = 1
    SCE_C_COMMENTLINE = 2
    SCE_C_COMMENTDOC = 3
    SCE_C_NUMBER = 4
    SCE_C_WORD = 5
    SCE_C_STRING = 6
    SCE_C_CHARACTER = 7
    SCE_C_UUID = 8
    SCE_C_PREPROCESSOR = 9
    SCE_C_OPERATOR = 10
    SCE_C_IDENTIFIER = 11
    SCE_C_STRINGEOL = 12
    SCE_C_VERBATIM = 13
    SCE_C_REGEX = 14
    SCE_C_COMMENTLINEDOC = 15
    SCE_C_WORD2 = 16
    SCE_C_COMMENTDOCKEYWORD = 17
    SCE_C_COMMENTDOCKEYWORDERROR = 18
    SCE_C_GLOBALCLASS = 19
    SCE_C_STRINGRAW = 20
    SCE_C_TRIPLEVERBATIM = 21
    SCE_C_HASHQUOTEDSTRING = 22
    SCE_C_PREPROCESSORCOMMENT = 23
    SCE_C_PREPROCESSORCOMMENTDOC = 24
    SCE_C_USERLITERAL = 25
    SCE_C_TASKMARKER = 26
    SCE_C_ESCAPESEQUENCE = 27
End Enum

'=========================================================================
' Python Lexer Styles (SCLEX_PYTHON)
'=========================================================================
Public Enum eStylePython
    SCE_P_DEFAULT = 0
    SCE_P_COMMENTLINE = 1
    SCE_P_NUMBER = 2
    SCE_P_STRING = 3
    SCE_P_CHARACTER = 4
    SCE_P_WORD = 5
    SCE_P_TRIPLE = 6
    SCE_P_TRIPLEDOUBLE = 7
    SCE_P_CLASSNAME = 8
    SCE_P_DEFNAME = 9
    SCE_P_OPERATOR = 10
    SCE_P_IDENTIFIER = 11
    SCE_P_COMMENTBLOCK = 12
    SCE_P_STRINGEOL = 13
    SCE_P_WORD2 = 14
    SCE_P_DECORATOR = 15
    SCE_P_FSTRING = 16
    SCE_P_FCHARACTER = 17
    SCE_P_FTRIPLE = 18
    SCE_P_FTRIPLEDOUBLE = 19
End Enum

'=========================================================================
' VB/VBA/VBScript Lexer Styles (SCLEX_VB)
'=========================================================================
Public Enum eStyleVB
    SCE_B_DEFAULT = 0
    SCE_B_COMMENT = 1
    SCE_B_NUMBER = 2
    SCE_B_KEYWORD = 3
    SCE_B_STRING = 4
    SCE_B_PREPROCESSOR = 5
    SCE_B_OPERATOR = 6
    SCE_B_IDENTIFIER = 7
    SCE_B_DATE = 8
    SCE_B_STRINGEOL = 9
    SCE_B_KEYWORD2 = 10
    SCE_B_KEYWORD3 = 11
    SCE_B_KEYWORD4 = 12
    SCE_B_CONSTANT = 13
    SCE_B_ASM = 14
    SCE_B_LABEL = 15
    SCE_B_ERROR = 16
    SCE_B_HEXNUMBER = 17
    SCE_B_BINNUMBER = 18
    SCE_B_COMMENTBLOCK = 19
    SCE_B_DOCLINE = 20
    SCE_B_DOCBLOCK = 21
    SCE_B_DOCKEYWORD = 22
End Enum

'=========================================================================
' HTML/XML Lexer Styles (SCLEX_HTML)
'=========================================================================
Public Enum eStyleHTML
    SCE_H_DEFAULT = 0
    SCE_H_TAG = 1
    SCE_H_UNKNOWNTAG = 2
    SCE_H_ATTRIBUTE = 3
    SCE_H_ATTRIBUTEUNKNOWN = 4
    SCE_H_NUMBER = 5
    SCE_H_DOUBLESTRING = 6
    SCE_H_SINGLESTRING = 7
    SCE_H_OTHER = 8
    SCE_H_COMMENT = 9
    SCE_H_ENTITY = 10
    SCE_H_TAGEND = 11
    SCE_H_XMLSTART = 12
    SCE_H_XMLEND = 13
    SCE_H_SCRIPT = 14
    SCE_H_ASP = 15
    SCE_H_ASPAT = 16
    SCE_H_CDATA = 17
    SCE_H_QUESTION = 18
    SCE_H_VALUE = 19
    SCE_H_XCCOMMENT = 20
End Enum

'=========================================================================
' SQL Lexer Styles (SCLEX_SQL)
'=========================================================================
Public Enum eStyleSQL
    SCE_SQL_DEFAULT = 0
    SCE_SQL_COMMENT = 1
    SCE_SQL_COMMENTLINE = 2
    SCE_SQL_COMMENTDOC = 3
    SCE_SQL_NUMBER = 4
    SCE_SQL_WORD = 5
    SCE_SQL_STRING = 6
    SCE_SQL_CHARACTER = 7
    SCE_SQL_SQLPLUS = 8
    SCE_SQL_SQLPLUS_PROMPT = 9
    SCE_SQL_OPERATOR = 10
    SCE_SQL_IDENTIFIER = 11
    SCE_SQL_SQLPLUS_COMMENT = 13
    SCE_SQL_COMMENTLINEDOC = 15
    SCE_SQL_WORD2 = 16
    SCE_SQL_COMMENTDOCKEYWORD = 17
    SCE_SQL_COMMENTDOCKEYWORDERROR = 18
    SCE_SQL_USER1 = 19
    SCE_SQL_USER2 = 20
    SCE_SQL_USER3 = 21
    SCE_SQL_USER4 = 22
    SCE_SQL_QUOTEDIDENTIFIER = 23
End Enum

'=========================================================================
' JSON Lexer Styles (SCLEX_JSON)
'=========================================================================
Public Enum eStyleJSON
    SCE_JSON_DEFAULT = 0
    SCE_JSON_NUMBER = 1
    SCE_JSON_STRING = 2
    SCE_JSON_STRINGEOL = 3
    SCE_JSON_PROPERTYNAME = 4
    SCE_JSON_ESCAPESEQUENCE = 5
    SCE_JSON_LINECOMMENT = 6
    SCE_JSON_BLOCKCOMMENT = 7
    SCE_JSON_OPERATOR = 8
    SCE_JSON_URI = 9
    SCE_JSON_COMPACTIRI = 10
    SCE_JSON_KEYWORD = 11
    SCE_JSON_LDKEYWORD = 12
    SCE_JSON_ERROR = 13
End Enum

'=========================================================================
' Batch File Lexer Styles (SCLEX_BATCH)
'=========================================================================
Public Enum eStyleBatch
    SCE_BAT_DEFAULT = 0
    SCE_BAT_COMMENT = 1
    SCE_BAT_WORD = 2
    SCE_BAT_LABEL = 3
    SCE_BAT_HIDE = 4
    SCE_BAT_COMMAND = 5
    SCE_BAT_IDENTIFIER = 6
    SCE_BAT_OPERATOR = 7
End Enum

'=========================================================================
' Makefile Lexer Styles (SCLEX_MAKEFILE)
'=========================================================================
Public Enum eStyleMake
    SCE_MAKE_DEFAULT = 0
    SCE_MAKE_COMMENT = 1
    SCE_MAKE_PREPROCESSOR = 2
    SCE_MAKE_IDENTIFIER = 3
    SCE_MAKE_OPERATOR = 4
    SCE_MAKE_TARGET = 5
    SCE_MAKE_IDEOL = 9
End Enum

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

' Case styles
Public Enum eStyleCase
    caseMixed = 0
    caseUpper = 1
    caseLower = 2
    caseCamel = 3
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

Public Sub SetKeywords(ByVal keywordSet As Long, ByVal keywords As String)
    m_Sci.SciMsgStr SCI_SETKEYWORDS, keywordSet, keywords
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
Public Sub ClearAll()
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

Public Sub SetVisible(ByVal style As Long, ByVal visible As Boolean)
    m_Sci.SciMsg SCI_STYLESETVISIBLE, style, IIf(visible, 1, 0)
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
    
    On Error GoTo ErrorHandler
    
    Select Case language
        Case langVB
            SetLanguage = SetVBHighlighter(m_Sci)
            
        Case langJavaScript
            SetLanguage = SetJavaHighlighter(m_Sci)
            
        Case langSQL
            SetLanguage = SetSqlHighlighter(m_Sci)
            
        Case langPython
            SetLanguage = SetPythonHighlighter(m_Sci)
            
        Case langHTML
            SetLanguage = SetHTMLHighlighter(m_Sci)
            
        Case langVBPCode
            SetVBPCodeHighlighter m_Sci, additionalKeywords
            SetLanguage = True
            
        Case langNone
            Lexer = lexNull
            ClearAll
            SetLanguage = True
            
        Case Else
            SetLanguage = False
    End Select
    
    Exit Function
    
ErrorHandler:
    SetLanguage = False
End Function

'=========================================================================
' Style Constant Helper Methods
'=========================================================================

' Apply style using enum - CPP/C/JavaScript
Public Sub ApplyCPPStyle(ByVal styleEnum As eStyleCPP, ByVal foreColor As Long, Optional ByVal backColor As Long = &HFFFFFF, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False)
    ConfigureStyle styleEnum, foreColor, backColor, , , bold, italic
End Sub

' Apply style using enum - Python
Public Sub ApplyPythonStyle(ByVal styleEnum As eStylePython, ByVal foreColor As Long, Optional ByVal backColor As Long = &HFFFFFF, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False)
    ConfigureStyle styleEnum, foreColor, backColor, , , bold, italic
End Sub

' Apply style using enum - VB
Public Sub ApplyVBStyle(ByVal styleEnum As eStyleVB, ByVal foreColor As Long, Optional ByVal backColor As Long = &HFFFFFF, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False)
    ConfigureStyle styleEnum, foreColor, backColor, , , bold, italic
End Sub

' Apply style using enum - HTML
Public Sub ApplyHTMLStyle(ByVal styleEnum As eStyleHTML, ByVal foreColor As Long, Optional ByVal backColor As Long = &HFFFFFF, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False)
    ConfigureStyle styleEnum, foreColor, backColor, , , bold, italic
End Sub

' Apply style using enum - SQL
Public Sub ApplySQLStyle(ByVal styleEnum As eStyleSQL, ByVal foreColor As Long, Optional ByVal backColor As Long = &HFFFFFF, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False)
    ConfigureStyle styleEnum, foreColor, backColor, , , bold, italic
End Sub

' Apply style using enum - JSON
Public Sub ApplyJSONStyle(ByVal styleEnum As eStyleJSON, ByVal foreColor As Long, Optional ByVal backColor As Long = &HFFFFFF, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False)
    ConfigureStyle styleEnum, foreColor, backColor, , , bold, italic
End Sub

'=========================================================================
' Quick Common Style Configurations
'=========================================================================

' Configure line number margin style
Public Sub ConfigureLineNumberStyle(Optional ByVal foreColor As Long = &H808080, Optional ByVal backColor As Long = &HF0F0F0, Optional ByVal fontName As String = "Courier New", Optional ByVal fontSize As Long = 9)
    ConfigureStyle STYLE_LINENUMBER, foreColor, backColor, fontName, fontSize
End Sub

' Configure brace match/bad styles
Public Sub ConfigureBraceStyles(Optional ByVal matchFore As Long = &HFF, Optional ByVal matchBack As Long = &HFFFFFF, Optional ByVal badFore As Long = &HFF0000, Optional ByVal badBack As Long = &HFFFFFF, Optional ByVal bold As Boolean = True)
    ConfigureStyle STYLE_BRACELIGHT, matchFore, matchBack, , , bold
    ConfigureStyle STYLE_BRACEBAD, badFore, badBack, , , bold
End Sub

' Configure indent guide style
Public Sub ConfigureIndentGuideStyle(Optional ByVal foreColor As Long = &HC0C0C0, Optional ByVal backColor As Long = &HFFFFFF)
    SetFore STYLE_INDENTGUIDE, foreColor
    SetBack STYLE_INDENTGUIDE, backColor
End Sub

'=========================================================================
' Style Bits Configuration
'=========================================================================

' Set style bits (required for some lexers, especially HTML)
Public Sub SetStyleBits(ByVal bits As Long)
    m_Sci.SciMsg 2090, bits  ' SCI_SETSTYLEBITS
End Sub

Public Function GetStyleBits() As Long
    GetStyleBits = m_Sci.SciMsg(2091)  ' SCI_GETSTYLEBITS
End Function

'=========================================================================
' Style Inspection Helper
'=========================================================================

' Get the style at a specific position
Public Function GetStyleAt(ByVal position As Long) As Long
    GetStyleAt = m_Sci.SciMsg(2010, position)  ' SCI_GETSTYLEAT
End Function

' Get the style name for a given style number (for debugging)
Public Function GetStyleName(ByVal style As Long, ByVal currentLexer As eLexer) As String
    Select Case currentLexer
        Case lexCPP, lexJavaScript
            Select Case style
                Case SCE_C_DEFAULT: GetStyleName = "Default"
                Case SCE_C_COMMENT: GetStyleName = "Comment"
                Case SCE_C_COMMENTLINE: GetStyleName = "Comment Line"
                Case SCE_C_COMMENTDOC: GetStyleName = "Comment Doc"
                Case SCE_C_NUMBER: GetStyleName = "Number"
                Case SCE_C_WORD: GetStyleName = "Keyword"
                Case SCE_C_STRING: GetStyleName = "String"
                Case SCE_C_CHARACTER: GetStyleName = "Character"
                Case SCE_C_UUID: GetStyleName = "UUID"
                Case SCE_C_PREPROCESSOR: GetStyleName = "Preprocessor"
                Case SCE_C_OPERATOR: GetStyleName = "Operator"
                Case SCE_C_IDENTIFIER: GetStyleName = "Identifier"
                Case SCE_C_STRINGEOL: GetStyleName = "String EOL"
                Case SCE_C_WORD2: GetStyleName = "Keyword2"
                Case SCE_C_GLOBALCLASS: GetStyleName = "Global Class"
                Case Else: GetStyleName = "Style " & style
            End Select
            
        Case lexVB
            Select Case style
                Case SCE_B_DEFAULT: GetStyleName = "Default"
                Case SCE_B_COMMENT: GetStyleName = "Comment"
                Case SCE_B_NUMBER: GetStyleName = "Number"
                Case SCE_B_KEYWORD: GetStyleName = "Keyword"
                Case SCE_B_STRING: GetStyleName = "String"
                Case SCE_B_PREPROCESSOR: GetStyleName = "Preprocessor"
                Case SCE_B_OPERATOR: GetStyleName = "Operator"
                Case SCE_B_IDENTIFIER: GetStyleName = "Identifier"
                Case SCE_B_DATE: GetStyleName = "Date"
                Case SCE_B_STRINGEOL: GetStyleName = "String EOL"
                Case Else: GetStyleName = "Style " & style
            End Select
            
        Case lexPython
            Select Case style
                Case SCE_P_DEFAULT: GetStyleName = "Default"
                Case SCE_P_COMMENTLINE: GetStyleName = "Comment"
                Case SCE_P_NUMBER: GetStyleName = "Number"
                Case SCE_P_STRING: GetStyleName = "String"
                Case SCE_P_CHARACTER: GetStyleName = "Character"
                Case SCE_P_WORD: GetStyleName = "Keyword"
                Case SCE_P_TRIPLE: GetStyleName = "Triple Quote"
                Case SCE_P_TRIPLEDOUBLE: GetStyleName = "Triple Double"
                Case SCE_P_CLASSNAME: GetStyleName = "Class Name"
                Case SCE_P_DEFNAME: GetStyleName = "Function Name"
                Case SCE_P_OPERATOR: GetStyleName = "Operator"
                Case SCE_P_IDENTIFIER: GetStyleName = "Identifier"
                Case Else: GetStyleName = "Style " & style
            End Select
            
        Case Else
            GetStyleName = "Style " & style
    End Select
End Function

'=========================================================================
' Batch Style Operations
'=========================================================================

' Copy all style attributes from one style to another
Public Sub CopyStyle(ByVal sourceStyle As Long, ByVal destStyle As Long)
    SetFore destStyle, GetFore(sourceStyle)
    SetBack destStyle, GetBack(sourceStyle)
    SetBold destStyle, GetBold(sourceStyle)
    SetItalic destStyle, GetItalic(sourceStyle)
    SetUnderline destStyle, GetUnderline(sourceStyle)
    SetSize destStyle, GetSize(sourceStyle)
    SetFont destStyle, GetFont(sourceStyle)
    SetEOLFilled destStyle, GetEOLFilled(sourceStyle)
    SetCase destStyle, GetCase(sourceStyle)
    SetVisible destStyle, GetVisible(sourceStyle)
    SetChangeable destStyle, GetChangeable(sourceStyle)
    SetHotspot destStyle, GetHotspot(sourceStyle)
End Sub

' Export current style configuration to string (for debugging/saving)
Public Function ExportStyleConfig(ByVal style As Long) As String
    Dim parts(11) As String
    
    parts(0) = IIf(GetBold(style), "1", "0")
    parts(1) = IIf(GetItalic(style), "1", "0")
    parts(2) = IIf(GetUnderline(style), "1", "0")
    parts(3) = IIf(GetVisible(style), "1", "0")
    parts(4) = GetFont(style)
    parts(5) = CStr(GetFore(style))
    parts(6) = CStr(GetBack(style))
    parts(7) = CStr(GetSize(style))
    parts(8) = IIf(GetEOLFilled(style), "1", "0")
    parts(9) = CStr(GetCase(style))
    parts(10) = IIf(GetChangeable(style), "1", "0")
    parts(11) = IIf(GetHotspot(style), "1", "0")
    
    ExportStyleConfig = Join(parts, ",")
End Function

' Import style configuration from string
Public Sub ImportStyleConfig(ByVal style As Long, ByVal config As String)
    Dim parts() As String
    parts = Split(config, ",")
    
    If UBound(parts) >= 8 Then
        SetBold style, (parts(0) = "1")
        SetItalic style, (parts(1) = "1")
        SetUnderline style, (parts(2) = "1")
        SetVisible style, (parts(3) = "1")
        If Len(parts(4)) > 0 Then SetFont style, parts(4)
        SetFore style, CLng(parts(5))
        SetBack style, CLng(parts(6))
        SetSize style, CLng(parts(7))
        SetEOLFilled style, (parts(8) = "1")
        
        ' Optional newer fields
        If UBound(parts) >= 9 Then SetCase style, CLng(parts(9))
        If UBound(parts) >= 10 Then SetChangeable style, (parts(10) = "1")
        If UBound(parts) >= 11 Then SetHotspot style, (parts(11) = "1")
    End If
End Sub

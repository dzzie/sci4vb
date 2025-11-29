Attribute VB_Name = "ModHighlighter"
'Author:  David Zimmer <dzzie@yahoo.com> + Claude.ai
'Site:    http://sandsprite.com
'License: MIT
'---------------------------------------------------

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Option Explicit

'=========================================================================
' ModHighlighter - Professional syntax highlighting presets for SciWrapper
' Adapted to work with modular API architecture
'=========================================================================

' Scintilla CPP lexer style constants
Private Const SCE_C_DEFAULT = 0
Private Const SCE_C_COMMENT = 1
Private Const SCE_C_COMMENTLINE = 2
Private Const SCE_C_COMMENTDOC = 3
Private Const SCE_C_NUMBER = 4
Private Const SCE_C_WORD = 5
Private Const SCE_C_STRING = 6
Private Const SCE_C_CHARACTER = 7
Private Const SCE_C_UUID = 8
Private Const SCE_C_PREPROCESSOR = 9
Private Const SCE_C_OPERATOR = 10
Private Const SCE_C_IDENTIFIER = 11
Private Const SCE_C_STRINGEOL = 12
Private Const SCE_C_VERBATIM = 13
Private Const SCE_C_REGEX = 14
Private Const SCE_C_COMMENTLINEDOC = 15
Private Const SCE_C_WORD2 = 16
Private Const SCE_C_COMMENTDOCKEYWORD = 17
Private Const SCE_C_COMMENTDOCKEYWORDERROR = 18
Private Const SCE_C_GLOBALCLASS = 19

Global Const LANG_US = &H409

Enum shellOpenState
    so_Hidden = 0
    so_Min = 2
    so_Max = 3
    so_Norm = 4
End Enum

Public Function ShellExec(path As String, _
                         Optional ByVal action As String = "Open", _
                         Optional ByVal Params As String = vbNullString, _
                         Optional ByVal Directory As String = vbNullString, _
                         Optional ByVal State As shellOpenState = so_Norm) As Boolean
                         
    ShellExec = (ShellExecute(0, action, path, Params, Directory, State) >= 33)
    
End Function

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo Init
    Dim X
       
    X = UBound(ary)
    ReDim Preserve ary(X + 1)
    
    If IsObject(value) Then
        Set ary(X + 1) = value
    Else
        ary(X + 1) = value
    End If
    
    Exit Sub
Init:
    ReDim ary(0)
    If IsObject(value) Then
        Set ary(0) = value
    Else
        ary(0) = value
    End If
End Sub

Function pad(v, Optional L As Long = 8, Optional char As String = " ", Optional padRight As Boolean = True)
    On Error GoTo hell
    Dim X As Long
    X = Len(v)
    If X < L Then
        If padRight Then
             pad = v & String(L - X, char)
        Else
             pad = String(L - X, char) & v
        End If
    Else
hell:
        pad = v
    End If
End Function

Function FileExists(path) As Boolean
  On Error GoTo hell
    
  '.(0), ..(0) etc cause dir to read it as cwd!
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If InStr(path, Chr(0)) > 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function


Public Function FileSize(fPath As String) As String
    Dim fsize As Long
    Dim szName As String
    On Error GoTo hell
    
    fsize = FileLen(fPath)
    
    szName = " bytes"
    If fsize > 1024 Then
        fsize = fsize / 1024
        szName = " Kb"
    End If
    
    If fsize > 1024 Then
        fsize = fsize / 1024
        szName = " Mb"
    End If
    
    FileSize = fsize & szName
    
    Exit Function
hell:
    
End Function

'=========================================================================
' VB Syntax Highlighting
'=========================================================================
Public Function SetVBHighlighter(sci As SciWrapper) As Boolean
    On Error GoTo ErrorHandler
    
    ' ========================================================================
    ' Configure Scintilla for Visual Basic syntax highlighting
    ' Uses the VB lexer for VB6/VBA syntax
    ' ========================================================================
    
    ' VB Keywords - Primary keyword set (flow control, declarations, etc.)
    Const KEYWORDS_PRIMARY = "and begin case call class continue do each else " & _
                            "elseif end erase error event exit false for function " & _
                            "get gosub goto if implement in load loop lset me mid " & _
                            "new next not nothing on or property raiseevent rem " & _
                            "resume return rset select set stop sub then to true " & _
                            "unload until wend while with withevents attribute alias " & _
                            "as boolean byref byte byval const compare currency date " & _
                            "declare dim double enum explicit friend global integer " & _
                            "let lib long module object option optional preserve " & _
                            "private public redim single static string type variant"
    
    ' VB Keywords - Secondary keyword set (statement keywords)
    Const KEYWORDS_STATEMENT = "and begin case call class continue do each else " & _
                              "elseif end erase error event exit false for function " & _
                              "get gosub goto if implement in load loop lset me mid " & _
                              "new next not nothing on or property raiseevent rem " & _
                              "resume return rset select set stop sub then to true"
    
    ' Color constants for readability (BGR format)
    Const CLR_WHITE = &HFFFFFF
    Const CLR_BLACK = &H0
    Const CLR_DARK_GREEN = &H5500
    Const CLR_BLUE = &HFF
    Const CLR_DARK_BLUE = &HA00000
    Const CLR_PURPLE = &HC01090
    Const CLR_BROWN = &H8CA0
    Const CLR_TEAL = &H808000
    Const CLR_ORANGE = &H8080
    Const CLR_RED = &HFF0000
    Const CLR_TAN = &HBD7373
    Const CLR_LINE_NUM_BACK = &HEFEFF3
    
    With sci
        ' Initialize VB lexer
        .style.Lexer = lexVB
        .style.ClearAll
        .SciMsg 2090, 5  ' SCI_SETSTYLEBITS = 5
        
        ' Set keyword lists
        .style.SetKeywords 0, KEYWORDS_PRIMARY
        .style.SetKeywords 2, KEYWORDS_STATEMENT
        
        ' ------------------------------------------------------------------------
        ' Configure Default Style (STYLE_DEFAULT = 32)
        ' All other styles inherit from this unless explicitly overridden
        ' ------------------------------------------------------------------------
        .style.SetFont 32, "Courier New"
        .style.SetSize 32, 10
        .style.SetFore 32, CLR_BLACK
        .style.SetBack 32, CLR_WHITE
        .style.SetBold 32, False
        .style.SetItalic 32, False
        .style.SetUnderline 32, False
        .style.SetVisible 32, True
        .style.SetEOLFilled 32, False
        .style.ClearAll  ' Propagate default style to all
        
        ' ------------------------------------------------------------------------
        ' SCE_VB_DEFAULT (0) - Default text
        ' ------------------------------------------------------------------------
        .style.SetFore 0, CLR_BLACK
        
        ' ------------------------------------------------------------------------
        ' SCE_VB_COMMENT (1) - Comments (REM or ')
        ' ------------------------------------------------------------------------
        .style.SetFore 1, CLR_DARK_GREEN
        
        ' ------------------------------------------------------------------------
        ' SCE_VB_NUMBER (2) - Numeric literals
        ' ------------------------------------------------------------------------
        .style.SetFore 2, CLR_BLUE
        
        ' ------------------------------------------------------------------------
        ' SCE_VB_KEYWORD (3) - Language keywords (bold dark blue)
        ' ------------------------------------------------------------------------
        .style.SetFore 3, CLR_DARK_BLUE
        .style.SetBold 3, True
        
        ' ------------------------------------------------------------------------
        ' SCE_VB_STRING (4) - String literals "..." (italic purple)
        ' ------------------------------------------------------------------------
        .style.SetFore 4, CLR_PURPLE
        .style.SetItalic 4, True
        
        ' ------------------------------------------------------------------------
        ' SCE_VB_PREPROCESSOR (5) - Preprocessor directives (#If, #Const, etc.)
        ' ------------------------------------------------------------------------
        .style.SetFore 5, CLR_BROWN
        
        ' ------------------------------------------------------------------------
        ' SCE_VB_OPERATOR (6) - Operators (+, -, *, /, =, etc.) (bold teal)
        ' ------------------------------------------------------------------------
        .style.SetFore 6, CLR_TEAL
        .style.SetBold 6, True
        
        ' ------------------------------------------------------------------------
        ' SCE_VB_IDENTIFIER (7) - Identifiers (variable/function names)
        ' ------------------------------------------------------------------------
        .style.SetFore 7, CLR_BLACK
        
        ' ------------------------------------------------------------------------
        ' SCE_VB_STRINGEOL (8) - Unclosed string at EOL (error state)
        ' ------------------------------------------------------------------------
        .style.SetFore 8, CLR_ORANGE
        
        ' ------------------------------------------------------------------------
        ' STYLE_LINENUMBER (33) - Line number margin
        ' ------------------------------------------------------------------------
        .style.SetFore 33, CLR_TAN
        .style.SetBack 33, CLR_LINE_NUM_BACK
        .style.SetSize 33, 8  ' Smaller font for line numbers
        
        ' ------------------------------------------------------------------------
        ' STYLE_BRACELIGHT (34) - Matching brace highlight
        ' ------------------------------------------------------------------------
        .style.SetFore 34, CLR_RED
        .style.SetBack 34, CLR_WHITE
        .style.SetBold 34, True
        
        ' ------------------------------------------------------------------------
        ' STYLE_BRACEBAD (35) - Unmatched brace highlight
        ' ------------------------------------------------------------------------
        .style.SetFore 35, CLR_BLUE
        .style.SetBack 35, CLR_WHITE
        .style.SetBold 35, True
        
        ' Apply syntax highlighting to entire document
        .style.Colorise 0, -1
    End With
    
    SetVBHighlighter = True
    Exit Function
    
ErrorHandler:
    SetVBHighlighter = False
End Function

'=========================================================================
' JavaScript/Java Syntax Highlighting
'=========================================================================
Public Function SetJavaHighlighter(sci As SciWrapper) As Boolean
    On Error GoTo ErrorHandler
    
    ' ========================================================================
    ' Configure Scintilla for JavaScript/Java syntax highlighting
    ' Uses the CPP lexer which handles C-style syntax
    ' ========================================================================
    
    Dim i As Long
    
    ' JavaScript/Java Keywords (style 5)
    Const KEYWORDS = "abstract boolean break byte case catch char class const " & _
                     "continue debugger default delete do double else enum export " & _
                     "extends final finally float for function goto if implements " & _
                     "import in instanceof int interface long native new package " & _
                     "private protected public return short static super switch " & _
                     "synchronized this throw throws transient try typeof var void " & _
                     "volatile while with true false null"
    
    ' Color constants for readability
    Const CLR_WHITE = &HFFFFFF
    Const CLR_BLACK = &H0
    Const CLR_GRAY = &H808080
    Const CLR_DARK_GREEN = &H8000
    Const CLR_DARK_BLUE = &H800000
    Const CLR_PURPLE = &H800080
    Const CLR_TEAL = &H808000
    Const CLR_CYAN = &HC08000
    Const CLR_LIGHT_CYAN = &HC0FFC0
    Const CLR_LIGHT_YELLOW = &HE0FFE0
    Const CLR_LIGHT_PINK = &HE0E0FF
    Const CLR_LIGHT_BLUE = &HFFC0C0
    Const CLR_RED = &HFF0000
    Const CLR_BLUE = &HFF
    Const CLR_ORANGE = &H5080
    Const CLR_BROWN = &H7080
    Const CLR_TAN = &HBD7373
    Const CLR_LINE_NUM_BACK = &HEFEFF3
    
    With sci
        ' Initialize lexer
        .style.Lexer = lexCPP
        .style.ClearAll
        .SciMsg 2090, 5  ' SCI_SETSTYLEBITS = 5
        
        ' Set keyword list
        .style.SetKeywords 0, KEYWORDS
        
        ' ------------------------------------------------------------------------
        ' Configure Default Style (STYLE_DEFAULT = 32)
        ' All other styles inherit from this unless explicitly overridden
        ' ------------------------------------------------------------------------
        .style.SetFont 32, "Courier New"
        .style.SetSize 32, 11
        .style.SetFore 32, CLR_BLACK
        .style.SetBack 32, CLR_WHITE
        .style.SetBold 32, False
        .style.SetItalic 32, False
        .style.SetUnderline 32, False
        .style.SetVisible 32, True
        .style.SetEOLFilled 32, False
        .style.ClearAll  ' Propagate default style to all
        
        ' ------------------------------------------------------------------------
        ' SCE_C_DEFAULT (0) - Default text
        ' ------------------------------------------------------------------------
        .style.SetFore 0, CLR_GRAY
        
        ' ------------------------------------------------------------------------
        ' SCE_C_COMMENT (1) - Block comments /* ... */
        ' ------------------------------------------------------------------------
        .style.SetFore 1, CLR_DARK_GREEN
        
        ' ------------------------------------------------------------------------
        ' SCE_C_COMMENTLINE (2) - Line comments //
        ' ------------------------------------------------------------------------
        .style.SetFore 2, &H7F00  ' Dark green-blue
        
        ' ------------------------------------------------------------------------
        ' SCE_C_COMMENTDOC (3) - Documentation comments /** ... */
        ' ------------------------------------------------------------------------
        .style.SetFore 3, &H555555  ' Medium gray
        
        ' ------------------------------------------------------------------------
        ' SCE_C_NUMBER (4) - Numeric literals
        ' ------------------------------------------------------------------------
        .style.SetFore 4, &H808000  ' Teal
        
        ' ------------------------------------------------------------------------
        ' SCE_C_WORD (5) - Keywords (bold dark blue)
        ' ------------------------------------------------------------------------
        .style.SetFore 5, CLR_DARK_BLUE
        .style.SetBold 5, True
        
        ' ------------------------------------------------------------------------
        ' SCE_C_STRING (6) - String literals "..." (italic purple)
        ' ------------------------------------------------------------------------
        .style.SetFore 6, &H850080  ' Purple
        .style.SetItalic 6, True
        
        ' ------------------------------------------------------------------------
        ' SCE_C_CHARACTER (7) - Character literals '...' (italic)
        ' ------------------------------------------------------------------------
        .style.SetFore 7, &H7EFFFF  ' Light yellow
        .style.SetItalic 7, True
        
        ' ------------------------------------------------------------------------
        ' SCE_C_UUID (8) - UUIDs
        ' ------------------------------------------------------------------------
        .style.SetFore 8, &HC08000  ' Light cyan
        
        ' ------------------------------------------------------------------------
        ' SCE_C_PREPROCESSOR (9) - Preprocessor directives
        ' ------------------------------------------------------------------------
        .style.SetFore 9, &H8080  ' Orange
        
        ' ------------------------------------------------------------------------
        ' SCE_C_OPERATOR (10) - Operators and punctuation
        ' ------------------------------------------------------------------------
        .style.SetFore 10, &H808000  ' Teal
        
        ' ------------------------------------------------------------------------
        ' SCE_C_STRINGEOL (12) - Unclosed string at EOL (error state)
        ' ------------------------------------------------------------------------
        .style.SetFore 12, CLR_BLACK
        .style.SetBack 12, &HFFC1FF  ' Light pink background
        .style.SetEOLFilled 12, True
        
        ' ------------------------------------------------------------------------
        ' SCE_C_VERBATIM (13) - Verbatim strings @"..."
        ' ------------------------------------------------------------------------
        .style.SetFore 13, CLR_DARK_GREEN
        .style.SetBack 13, &HE0E0E0  ' Light gray background
        
        ' ------------------------------------------------------------------------
        ' SCE_C_REGEX (14) - Regular expressions /pattern/
        ' ------------------------------------------------------------------------
        .style.SetFore 14, &HC000  ' Dark cyan
        .style.SetBack 14, &HE0FFE0  ' Very light green background
        
        ' ------------------------------------------------------------------------
        ' SCE_C_COMMENTLINEDOC (15) - Documentation line comments ///
        ' ------------------------------------------------------------------------
        .style.SetFore 15, CLR_DARK_GREEN
        
        ' ------------------------------------------------------------------------
        ' SCE_C_WORD2 (16) - Secondary keyword set
        ' ------------------------------------------------------------------------
        .style.SetFore 16, &HA00030  ' Dark magenta
        
        ' ------------------------------------------------------------------------
        ' SCE_C_COMMENTDOCKEYWORD (17) - Keywords in doc comments
        ' ------------------------------------------------------------------------
        .style.SetFore 17, &H7080  ' Brown
        
        ' ------------------------------------------------------------------------
        ' SCE_C_COMMENTDOCKEYWORDERROR (18) - Unrecognized doc keywords
        ' ------------------------------------------------------------------------
        .style.SetFore 18, &HA00000  ' Dark red
        
        ' ------------------------------------------------------------------------
        ' STYLE_LINENUMBER (33) - Line number margin
        ' ------------------------------------------------------------------------
        .style.SetFore 33, CLR_TAN
        .style.SetBack 33, CLR_LINE_NUM_BACK
        
        ' ------------------------------------------------------------------------
        ' STYLE_BRACELIGHT (34) - Matching brace highlight
        ' ------------------------------------------------------------------------
        .style.SetFore 34, CLR_RED
        .style.SetBack 34, CLR_WHITE
        .style.SetBold 34, True
        
        ' ------------------------------------------------------------------------
        ' STYLE_BRACEBAD (35) - Unmatched brace highlight
        ' ------------------------------------------------------------------------
        .style.SetFore 35, CLR_BLUE
        .style.SetBack 35, CLR_WHITE
        .style.SetBold 35, True
        
        ' Apply syntax highlighting to entire document
        .style.Colorise 0, -1
    End With
    
    SetJavaHighlighter = True
    Exit Function
    
ErrorHandler:
    SetJavaHighlighter = False
End Function


Public Function SetSqlHighlighter(sci As SciWrapper) As Boolean
    On Error GoTo ErrorHandler
    
    ' ========================================================================
    ' Configure Scintilla for SQL syntax highlighting
    ' Supports T-SQL, PL/SQL, and standard SQL syntax
    ' ========================================================================
    
    ' SQL Keywords - Commands, clauses, operators, and data types
    Const SQL_KEYWORDS = "select from where and or not in like between is null " & _
                         "order by group having insert update delete create alter drop " & _
                         "table index view primary key foreign references constraint " & _
                         "inner outer left right join union distinct as " & _
                         "count sum avg max min desc asc into values set " & _
                         "declare begin end if else while loop cursor open fetch close " & _
                         "deallocate procedure function returns return exec execute trigger " & _
                         "varchar int bigint datetime char text nvarchar nchar " & _
                         "float decimal numeric bit date time timestamp"
    
    ' Color constants (BGR format)
    Const CLR_WHITE = &HFFFFFF
    Const CLR_BLACK = &H0
    Const CLR_GREEN = &H8000          ' Dark green for comments
    Const CLR_BLUE = &HFF             ' Bright blue for keywords
    Const CLR_PURPLE = &H850080       ' Purple for strings
    Const CLR_TEAL = &H808000         ' Teal for numbers/operators
    Const CLR_RED = &HFF0000          ' Red for brace matching
    Const CLR_TAN = &HBD7373          ' Tan for line numbers
    Const CLR_LINE_NUM_BACK = &HEFEFF3
    
    With sci
        ' Initialize SQL lexer
        .style.Lexer = lexSQL
        .style.ClearAll
        .SciMsg 2090, 5  ' SCI_SETSTYLEBITS = 5
        
        ' Set SQL keyword list
        .style.SetKeywords 0, SQL_KEYWORDS
        
        ' ------------------------------------------------------------------------
        ' Configure Default Style (STYLE_DEFAULT = 32)
        ' All other styles inherit from this unless explicitly overridden
        ' ------------------------------------------------------------------------
        .style.SetFont 32, "Courier New"
        .style.SetSize 32, 10
        .style.SetFore 32, CLR_BLACK
        .style.SetBack 32, CLR_WHITE
        .style.SetBold 32, False
        .style.SetItalic 32, False
        .style.SetUnderline 32, False
        .style.SetVisible 32, True
        .style.SetEOLFilled 32, False
        .style.ClearAll  ' Propagate default style to all
        
        ' ------------------------------------------------------------------------
        ' SCE_SQL_DEFAULT (0) - Default text
        ' ------------------------------------------------------------------------
        .style.SetFore 0, CLR_BLACK
        
        ' ------------------------------------------------------------------------
        ' SCE_SQL_COMMENT (1) - Block comments /* ... */
        ' ------------------------------------------------------------------------
        .style.SetFore 1, CLR_GREEN
        
        ' ------------------------------------------------------------------------
        ' SCE_SQL_COMMENTLINE (2) - Line comments --
        ' ------------------------------------------------------------------------
        .style.SetFore 2, CLR_GREEN
        
        ' ------------------------------------------------------------------------
        ' SCE_SQL_NUMBER (3) - Numeric literals
        ' ------------------------------------------------------------------------
        .style.SetFore 3, CLR_TEAL
        
        ' ------------------------------------------------------------------------
        ' SCE_SQL_STRING (4) - String literals '...' (italic purple)
        ' ------------------------------------------------------------------------
        .style.SetFore 4, CLR_PURPLE
        .style.SetItalic 4, True
        
        ' ------------------------------------------------------------------------
        ' SCE_SQL_WORD (5) - SQL keywords (bold blue)
        ' SELECT, FROM, WHERE, JOIN, etc.
        ' ------------------------------------------------------------------------
        .style.SetFore 5, CLR_BLUE
        .style.SetBold 5, True
        
        ' ------------------------------------------------------------------------
        ' SCE_SQL_OPERATOR (6) - Operators and punctuation
        ' =, <>, +, -, *, /, %, etc.
        ' ------------------------------------------------------------------------
        .style.SetFore 6, CLR_TEAL
        
        ' ------------------------------------------------------------------------
        ' SCE_SQL_IDENTIFIER (7) - Identifiers (table/column names)
        ' ------------------------------------------------------------------------
        .style.SetFore 7, CLR_BLACK
        
        ' ------------------------------------------------------------------------
        ' STYLE_LINENUMBER (33) - Line number margin
        ' ------------------------------------------------------------------------
        .style.SetFore 33, CLR_TAN
        .style.SetBack 33, CLR_LINE_NUM_BACK
        .style.SetSize 33, 8  ' Smaller font for line numbers
        
        ' ------------------------------------------------------------------------
        ' STYLE_BRACELIGHT (34) - Matching parenthesis/bracket highlight
        ' ------------------------------------------------------------------------
        .style.SetFore 34, CLR_RED
        .style.SetBack 34, CLR_WHITE
        .style.SetBold 34, True
        
        ' ------------------------------------------------------------------------
        ' STYLE_BRACEBAD (35) - Unmatched parenthesis/bracket highlight
        ' ------------------------------------------------------------------------
        .style.SetFore 35, CLR_BLUE
        .style.SetBack 35, CLR_WHITE
        .style.SetBold 35, True
        
        ' Apply syntax highlighting to entire document
        .style.Colorise 0, -1
    End With
    
    SetSqlHighlighter = True
    Exit Function
    
ErrorHandler:
    SetSqlHighlighter = False
End Function


'=========================================================================
' VB P-Code Disassembly Highlighting (for your malware analysis work!)
'=========================================================================
Public Sub SetVBPCodeHighlighter(sci As SciWrapper, Optional additionalKeywords As String = "")
    Dim callOpcodes As String
    Dim branchOpcodes As String
    Dim mathFileOpcodes As String
    
    ' Set CPP lexer for C-style syntax
    sci.style.Lexer = lexCPP
    sci.style.ClearAll
    
    ' Configure base font
    sci.style.SetFont 32, "Courier New"
    sci.style.SetSize 32, 10
    sci.style.SetBack 32, &HFFFFFF
    sci.style.SetFore 32, &H0
    sci.style.ClearAll
    
    ' Function call opcodes (keywords set 0)
    callOpcodes = "CallI2 CallI4 CallR4 CallR8 CallCy CallVar CallStr CallBool CallDate " & _
                  "CallFunc CallHresult CallSt CallStByref CallFFP4 CallFFP8 ReturnHresult " & _
                  "FStCall FStCallFFP4 FStCallFFP8 QuoteRem"
    
    ' Branch/comparison opcodes (keywords set 1)
    branchOpcodes = "BranchF BranchT Branch BranchFVar OnErrorGoto OnErrorResumeNext BoS BranchFVarFree BranchTVar " & _
                   "BranchTVarFree Gosub Resume OnGosub OnGoto EqI2 EqI4 EqR8 EqCy EqVar EqStr EqTextVar EqTextStr " & _
                   "EqVarBool EqTextVarBool EqCyR8 NeUI1 NeI4 NeR4 NeCy NeVar NeStr NeTextVar NeTextStr NeVarBool " & _
                   "NeTextVarBool NeCyR8 LeUI1 LeI2 LeI4 LeR4 LeCy LeVar LeStr LeTextVar LeTextStr LeVarBool " & _
                   "LeTextVarBool LeCyR8 GeUI1 GeI2 GeI4 GeR4 GeCy GeVar GeStr GeTextVar GeTextStr GeVarBool " & _
                   "GeTextVarBool GeCyR8 LtUI1 LtI2 LtI4 LtR4 LtCy LtVar LtStr LtTextVar LtTextStr LtVarBool " & _
                   "LtTextVarBool LtCyR8 GtUI1 GtI2 GtI4 GtR4 GtCy GtVar GtStr GtTextVar GtTextStr GtVarBool " & _
                   "GtTextVarBool GtCyR8 LikeVar LikeStr LikeTextVar LikeTextStr LikeVarBool LikeTextVarBool " & _
                   "BetweenUI1 BetweenI2 BetweenI4 BetweenR4 BetweenCy BetweenVar BetweenStr BetweenTextVar " & _
                   "BetweenTextStr EqR4 NeI2 NeR8 LtR8 LeR8 GeR8 ForUI1 ForI2 ForI4 ForR4 ForR8 ForCy ForVar " & _
                   "ForStepUI1 ForStepI2 ForStepI4 ForStepR4 ForStepR8 ForStepCy ForStepVar ForEachCollVar " & _
                   "NextEachCollVar ForEachCollAd NextEachCollAd ForEachAryVar NextEachAryVar NextUI1 NextI2 " & _
                   "NextI4 NextStepR4 NextR8 NextStepCy NextStepVar InvalidExcode NextStepUI1 NextStepI2 " & _
                   "NextStepI4 ForEachCollObj ForEachVar ForEachVarFree NextEachCollObj NextEachVar NextStepR8 " & _
                   "ExitForCollObj ExitForAryVar ExitForVar"
    
    ' Math/file operations (keywords set 3)
    If Len(additionalKeywords) = 0 Then
        mathFileOpcodes = "ModI2 ModI4 NotI4 AndI4 OrI4 XorI4 XorVar OrI2 OrVar AndUI1 AndVar ModUI1 ModVar NotUI1 " & _
                         "SeekFile NameFile OpenFile LockFile PrintFile WriteFile InputFile Input InputDone " & _
                         "InputItemUI1 InputItemI2 InputItemI4 InputItemR4 InputItemR8 InputItemCy InputItemVar " & _
                         "InputItemStr InputItemBool InputItemDate LineInputVar LineInputStr WriteChan Close " & _
                         "CloseAll GetRec3 GetRec4 PutRec3 PutRec4 GetRecOwner3 GetRecOwner4 PutRecOwner3 " & _
                         "PutRecOwner4 GetRecOwn3 GetRecOwn4 PutRecOwn3 PutRecOwn4"
    Else
        mathFileOpcodes = additionalKeywords
    End If
    
    ' Set keywords
    sci.style.SetKeywords 0, callOpcodes
    sci.style.SetKeywords 1, branchOpcodes
    sci.style.SetKeywords 3, mathFileOpcodes
    
    ' Set colors - optimized for malware analysis readability
    sci.style.SetFore SCE_C_COMMENT, &H8000        ' Green comments
    sci.style.SetFore SCE_C_COMMENTLINE, &H8000
    sci.style.SetFore SCE_C_WORD, &H800000         ' Dark red for calls
    sci.style.SetFore SCE_C_WORD2, vbRed           ' Red for branches
    sci.style.SetFore SCE_C_GLOBALCLASS, &H25208D  ' Purple for custom keywords
    sci.style.SetFore SCE_C_STRING, &H800080       ' Purple strings
    sci.style.SetFore SCE_C_CHARACTER, &H800080
    sci.style.SetFore SCE_C_NUMBER, &H8000         ' Blue numbers
    
    ' Set bold for emphasis
    sci.style.SetBold SCE_C_WORD, True
    sci.style.SetBold SCE_C_WORD2, True
    sci.style.SetBold SCE_C_GLOBALCLASS, True
    
    sci.style.Colorise 0, -1
End Sub

Public Function SetPythonHighlighter(sci As SciWrapper) As Boolean
    On Error GoTo ErrorHandler
    
    ' ========================================================================
    ' Configure Scintilla for Python syntax highlighting
    ' Supports Python 2.x and 3.x syntax
    ' ========================================================================
    
    ' Python Keywords - Language keywords and built-in constants
    Const PYTHON_KEYWORDS = "and as assert break class continue def del elif else except " & _
                           "exec finally for from global if import in is lambda not or " & _
                           "pass print raise return try while with yield True False None"
    
    ' Color constants (BGR format)
    Const CLR_WHITE = &HFFFFFF
    Const CLR_BLACK = &H0
    Const CLR_DARK_GREEN = &H8000       ' Comments
    Const CLR_DARK_CYAN = &H808000      ' Numbers
    Const CLR_PURPLE = &H800080         ' Strings
    Const CLR_GRAY = &H808080           ' Triple-quoted strings/docstrings
    Const CLR_RED = &HFF0000            ' Keywords
    Const CLR_DARK_BLUE = &H800000      ' Function/class names
    Const CLR_BLUE = &HFF               ' Brace mismatch
    Const CLR_TAN = &HBD7373            ' Line numbers
    Const CLR_LINE_NUM_BACK = &HEFEFF3
    
    With sci
        ' Initialize Python lexer
        .style.Lexer = lexPython
        .style.ClearAll
        
        ' Set Python keyword list
        .style.SetKeywords 0, PYTHON_KEYWORDS
        
        ' ------------------------------------------------------------------------
        ' Configure Default Style (STYLE_DEFAULT = 32)
        ' Using Consolas for better readability with Python's significant whitespace
        ' ------------------------------------------------------------------------
        .style.SetFont 32, "Consolas"
        .style.SetSize 32, 10
        .style.SetFore 32, CLR_BLACK
        .style.SetBack 32, CLR_WHITE
        .style.SetBold 32, False
        .style.SetItalic 32, False
        .style.SetUnderline 32, False
        .style.SetVisible 32, True
        .style.SetEOLFilled 32, False
        .style.ClearAll  ' Propagate default style to all
        
        ' ------------------------------------------------------------------------
        ' SCE_P_DEFAULT (0) - Default text
        ' ------------------------------------------------------------------------
        .style.SetFore 0, CLR_BLACK
        
        ' ------------------------------------------------------------------------
        ' SCE_P_COMMENTLINE (1) - Comments (# ...)
        ' ------------------------------------------------------------------------
        .style.SetFore 1, CLR_DARK_GREEN
        .style.SetItalic 1, True
        
        ' ------------------------------------------------------------------------
        ' SCE_P_NUMBER (2) - Numeric literals (int, float, hex, octal, binary)
        ' ------------------------------------------------------------------------
        .style.SetFore 2, CLR_DARK_CYAN
        
        ' ------------------------------------------------------------------------
        ' SCE_P_STRING (3) - String literals (double quotes "...")
        ' ------------------------------------------------------------------------
        .style.SetFore 3, CLR_PURPLE
        
        ' ------------------------------------------------------------------------
        ' SCE_P_CHARACTER (4) - String literals (single quotes '...')
        ' ------------------------------------------------------------------------
        .style.SetFore 4, CLR_PURPLE
        
        ' ------------------------------------------------------------------------
        ' SCE_P_WORD (5) - Keywords (bold red)
        ' if, for, while, def, class, import, etc.
        ' ------------------------------------------------------------------------
        .style.SetFore 5, CLR_RED
        .style.SetBold 5, True
        
        ' ------------------------------------------------------------------------
        ' SCE_P_TRIPLE (6) - Triple single-quoted strings ('''...''')
        ' Used for multi-line strings and docstrings
        ' ------------------------------------------------------------------------
        .style.SetFore 6, CLR_GRAY
        
        ' ------------------------------------------------------------------------
        ' SCE_P_TRIPLEDOUBLE (7) - Triple double-quoted strings ("""...""")
        ' Used for multi-line strings and docstrings
        ' ------------------------------------------------------------------------
        .style.SetFore 7, CLR_GRAY
        
        ' ------------------------------------------------------------------------
        ' SCE_P_CLASSNAME (8) - Class name definition (bold dark green)
        ' Name following 'class' keyword
        ' ------------------------------------------------------------------------
        .style.SetFore 8, CLR_DARK_GREEN
        .style.SetBold 8, True
        
        ' ------------------------------------------------------------------------
        ' SCE_P_DEFNAME (9) - Function/method name definition (bold dark blue)
        ' Name following 'def' keyword
        ' ------------------------------------------------------------------------
        .style.SetFore 9, CLR_DARK_BLUE
        .style.SetBold 9, True
        
        ' ------------------------------------------------------------------------
        ' SCE_P_OPERATOR (10) - Operators and punctuation
        ' +, -, *, /, =, ==, :, etc.
        ' ------------------------------------------------------------------------
        .style.SetFore 10, CLR_BLACK
        
        ' ------------------------------------------------------------------------
        ' SCE_P_IDENTIFIER (11) - Identifiers (variable names, function calls)
        ' ------------------------------------------------------------------------
        .style.SetFore 11, CLR_BLACK
        
        ' ------------------------------------------------------------------------
        ' STYLE_LINENUMBER (33) - Line number margin
        ' ------------------------------------------------------------------------
        .style.SetFore 33, CLR_TAN
        .style.SetBack 33, CLR_LINE_NUM_BACK
        .style.SetSize 33, 8  ' Smaller font for line numbers
        
        ' ------------------------------------------------------------------------
        ' STYLE_BRACELIGHT (34) - Matching parenthesis/bracket highlight
        ' Important for Python's tuple/list/dict syntax
        ' ------------------------------------------------------------------------
        .style.SetFore 34, CLR_RED
        .style.SetBack 34, CLR_WHITE
        .style.SetBold 34, True
        
        ' ------------------------------------------------------------------------
        ' STYLE_BRACEBAD (35) - Unmatched parenthesis/bracket highlight
        ' ------------------------------------------------------------------------
        .style.SetFore 35, CLR_BLUE
        .style.SetBack 35, CLR_WHITE
        .style.SetBold 35, True
        
        ' Apply syntax highlighting to entire document
        .style.Colorise 0, -1
    End With
    
    SetPythonHighlighter = True
    Exit Function
    
ErrorHandler:
    SetPythonHighlighter = False
End Function


'=========================================================================
' HTML/XML Syntax Highlighting
'=========================================================================
Public Function SetHTMLHighlighter(sci As SciWrapper) As Boolean
    On Error GoTo ErrorHandler
    
    sci.style.Lexer = lexHTML
    sci.style.ClearAll
    
    sci.SciMsg 2090, 7  ' SCI_SETSTYLEBITS = 7 for HTML
    
    sci.style.SetFont 32, "Consolas"
    sci.style.SetSize 32, 10
    sci.style.SetBack 32, &HFFFFFF
    sci.style.ClearAll
    
    ' HTML styles
    sci.style.SetFore 1, &H8000        ' Tag - green
    sci.style.SetFore 2, &HFF0000      ' Unknown tag - blue
    sci.style.SetFore 3, &H800080      ' Attribute - purple
    sci.style.SetFore 4, &H8000        ' Unknown attribute
    sci.style.SetFore 5, &HFF          ' Number - red
    sci.style.SetFore 6, &H800080      ' Double string
    sci.style.SetFore 7, &H800080      ' Single string
    sci.style.SetFore 8, &H808080      ' Other
    sci.style.SetFore 9, &H8000        ' Comment - green
    sci.style.SetItalic 9, True
    
    sci.style.Colorise 0, -1
    
    SetHTMLHighlighter = True
    Exit Function
    
ErrorHandler:
    SetHTMLHighlighter = False
End Function

'=========================================================================
' Utility: Configure Default Editor Settings
'=========================================================================
Public Sub ConfigureDefaultEditorSettings(sci As SciWrapper)
    With sci
        ' Margins
        .Margins.ConfigureLineNumbers 0, 40
        .Margins.ConfigureFolding 1, 16
        
        ' Editor preferences
        .edit.TabWidth = 4
        .edit.UseTabs = False
        .edit.TabIndents = True
        .edit.BackspaceUnindents = True
        .edit.eolMode = eolCRLF  ' VB6 standard
        
        ' View settings
        .View.CaretLineVisible = True
        .View.CaretLineBack = &HE8E8E8
        .View.EdgeMode = edgeLine
        .View.EdgeColumn = 80
        .View.EdgeColor = &HE0E0E0
        
        ' Whitespace
        .View.ViewWhitespace = wsInvisible
        
        ' Caret
        .View.CaretWidth = 1
        .View.CaretPeriod = 500
    End With
End Sub

'=========================================================================
' Utility: Apply Dark Theme (for late-night reverse engineering!)
'=========================================================================
Public Sub ApplyDarkTheme(sci As SciWrapper)
    With sci
        ' Dark background, light text
        .style.SetBack 32, &H1E1E1E   ' Dark gray background
        .style.SetFore 32, &HE0E0E0   ' Light gray text
        .style.SetFont 32, "Consolas"
        .style.SetSize 32, 10
        .style.ClearAll
        
        ' Line number margin
        .Margins.SetBack 0, &H2D2D2D
        
        ' Caret line
        .View.CaretLineVisible = True
        .View.CaretLineBack = &H2D2D30
        .View.CaretLineBackAlpha = 128
        
        ' Caret color
        .View.CaretFore = &HFFFFFF
        
        ' Edge column
        .View.EdgeColor = &H3E3E42
        
        ' Selection colors
        .SciMsg 2067, &H3E3E42  ' SCI_SETSELBACK with useSetting=True
        .SciMsg 2068, &HFFFFFF  ' SCI_SETSELFORE
    End With
End Sub

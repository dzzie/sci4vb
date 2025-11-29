Attribute VB_Name = "modStylingExamples"
Option Explicit

'=========================================================================
' CSciStyling Enum Usage Examples
' Demonstrates the new type-safe styling with IntelliSense support
'=========================================================================

'=========================================================================
' EXAMPLE 1: Using Style Enums for Type Safety
'=========================================================================
Public Sub ConfigureVBStylingWithEnums(sci As SciWrapper)
    With sci.Styling
        ' Set lexer
        .Lexer = lexVB
        .ClearAll
        
        ' Use VB-specific enum constants with IntelliSense support
        .ApplyVBStyle SCE_B_DEFAULT, &H0, &HFFFFFF
        .ApplyVBStyle SCE_B_COMMENT, &H8000, , False, True  ' Green italic
        .ApplyVBStyle SCE_B_KEYWORD, &HFF0000, , True       ' Red bold
        .ApplyVBStyle SCE_B_STRING, &H800080, , False, True ' Purple italic
        .ApplyVBStyle SCE_B_NUMBER, &H8000                  ' Blue
        .ApplyVBStyle SCE_B_OPERATOR, &H0
        
        ' Configure common styles using enum
        .ConfigureBraceStyles &HFF, &HFFFFFF, &HFF0000, &HFFFFFF, True
        .ConfigureLineNumberStyle &H808080, &HF0F0F0, "Consolas", 9
        
        .Colorise 0, -1
    End With
End Sub

'=========================================================================
' EXAMPLE 2: CPP/JavaScript Styling with Type-Safe Constants
'=========================================================================
Public Sub ConfigureJavaScriptWithEnums(sci As SciWrapper)
    With sci.Styling
        .Lexer = lexJavaScript
        .ClearAll
        
        ' Type-safe style configuration using CPP enum
        .ApplyCPPStyle SCE_C_DEFAULT, &H0, &HFFFFFF
        .ApplyCPPStyle SCE_C_COMMENT, &H8000, , False, True
        .ApplyCPPStyle SCE_C_COMMENTLINE, &H8000, , False, True
        .ApplyCPPStyle SCE_C_WORD, &HFF0000, , True          ' Keywords bold red
        .ApplyCPPStyle SCE_C_WORD2, &H800000, , True         ' Keywords2 bold dark red
        .ApplyCPPStyle SCE_C_STRING, &H800080                ' Purple strings
        .ApplyCPPStyle SCE_C_NUMBER, &H8000                  ' Blue numbers
        .ApplyCPPStyle SCE_C_OPERATOR, &H0
        .ApplyCPPStyle SCE_C_PREPROCESSOR, &H808080          ' Gray
        .ApplyCPPStyle SCE_C_GLOBALCLASS, &H8000, , True     ' Blue bold
        
        ' JavaScript keywords
        Dim keywords As String
        keywords = "function var let const if else for while do break continue " & _
                   "return new this try catch finally throw typeof instanceof"
        .SetKeywords 0, keywords
        
        .Colorise 0, -1
    End With
End Sub

'=========================================================================
' EXAMPLE 3: Python with Enum Constants
'=========================================================================
Public Sub ConfigurePythonWithEnums(sci As SciWrapper)
    With sci.Styling
        .Lexer = lexPython
        .ClearAll
        
        ' Python-specific enums provide IntelliSense
        .ApplyPythonStyle SCE_P_DEFAULT, &H0
        .ApplyPythonStyle SCE_P_COMMENTLINE, &H8000, , False, True
        .ApplyPythonStyle SCE_P_WORD, &HFF0000, , True       ' Keywords
        .ApplyPythonStyle SCE_P_STRING, &H800080
        .ApplyPythonStyle SCE_P_CHARACTER, &H800080
        .ApplyPythonStyle SCE_P_NUMBER, &H8000
        .ApplyPythonStyle SCE_P_TRIPLE, &H808080
        .ApplyPythonStyle SCE_P_TRIPLEDOUBLE, &H808080
        .ApplyPythonStyle SCE_P_CLASSNAME, &H8000, , True    ' Class names bold
        .ApplyPythonStyle SCE_P_DEFNAME, &H800000, , True    ' Function names bold
        .ApplyPythonStyle SCE_P_OPERATOR, &H0
        .ApplyPythonStyle SCE_P_IDENTIFIER, &H0
        
        .Colorise 0, -1
    End With
End Sub

'=========================================================================
' EXAMPLE 4: HTML/XML with Enum Constants
'=========================================================================
Public Sub ConfigureHTMLWithEnums(sci As SciWrapper)
    With sci.Styling
        .Lexer = lexHTML
        .ClearAll
        
        .SetStyleBits 7  ' HTML requires 7 style bits
        
        ' HTML-specific enums
        .ApplyHTMLStyle SCE_H_DEFAULT, &H0
        .ApplyHTMLStyle SCE_H_TAG, &H8000, , True            ' Tags green bold
        .ApplyHTMLStyle SCE_H_UNKNOWNTAG, &HFF0000           ' Unknown tags red
        .ApplyHTMLStyle SCE_H_ATTRIBUTE, &H800080            ' Attributes purple
        .ApplyHTMLStyle SCE_H_NUMBER, &H8000
        .ApplyHTMLStyle SCE_H_DOUBLESTRING, &H800080
        .ApplyHTMLStyle SCE_H_SINGLESTRING, &H800080
        .ApplyHTMLStyle SCE_H_COMMENT, &H8000, , False, True ' Comments green italic
        .ApplyHTMLStyle SCE_H_ENTITY, &H800000               ' Entities dark red
        
        .Colorise 0, -1
    End With
End Sub

'=========================================================================
' EXAMPLE 5: SQL with Enum Constants
'=========================================================================
Public Sub ConfigureSQLWithEnums(sci As SciWrapper)
    With sci.Styling
        .Lexer = lexSQL
        .ClearAll
        
        .ApplySQLStyle SCE_SQL_DEFAULT, &H0
        .ApplySQLStyle SCE_SQL_COMMENT, &H8000, , False, True
        .ApplySQLStyle SCE_SQL_COMMENTLINE, &H8000, , False, True
        .ApplySQLStyle SCE_SQL_WORD, &HFF0000, , True        ' Keywords bold red
        .ApplySQLStyle SCE_SQL_STRING, &H800080
        .ApplySQLStyle SCE_SQL_NUMBER, &H8000
        .ApplySQLStyle SCE_SQL_OPERATOR, &H0
        .ApplySQLStyle SCE_SQL_IDENTIFIER, &H0
        
        Dim keywords As String
        keywords = "select from where and or not in like between is null " & _
                   "order by group having insert update delete create alter drop"
        .SetKeywords 0, keywords
        
        .Colorise 0, -1
    End With
End Sub

'=========================================================================
' EXAMPLE 6: JSON with Enum Constants
'=========================================================================
Public Sub ConfigureJSONWithEnums(sci As SciWrapper)
    With sci.Styling
        .Lexer = lexJSON
        .ClearAll
        
        .ApplyJSONStyle SCE_JSON_DEFAULT, &H0
        .ApplyJSONStyle SCE_JSON_NUMBER, &H8000
        .ApplyJSONStyle SCE_JSON_STRING, &H800080
        .ApplyJSONStyle SCE_JSON_PROPERTYNAME, &H800000, , True  ' Property names bold
        .ApplyJSONStyle SCE_JSON_KEYWORD, &HFF0000, , True       ' true/false/null bold
        .ApplyJSONStyle SCE_JSON_LINECOMMENT, &H8000, , False, True
        .ApplyJSONStyle SCE_JSON_BLOCKCOMMENT, &H8000, , False, True
        .ApplyJSONStyle SCE_JSON_OPERATOR, &H0
        .ApplyJSONStyle SCE_JSON_ERROR, &HFF, &HFFCCCC         ' Errors red on pink
        
        .Colorise 0, -1
    End With
End Sub

'=========================================================================
' EXAMPLE 7: Using Common Style Constants
'=========================================================================
Public Sub ConfigureCommonStyles(sci As SciWrapper)
    With sci.Styling
        ' Configure default style using enum constant
        .SetFont STYLE_DEFAULT, "Consolas"
        .SetSize STYLE_DEFAULT, 10
        .SetFore STYLE_DEFAULT, &H0
        .SetBack STYLE_DEFAULT, &HFFFFFF
        .ClearAll
        
        ' Configure line numbers using helper
        .ConfigureLineNumberStyle &H808080, &HF0F0F0, "Consolas", 9
        
        ' Configure brace matching using helper
        .ConfigureBraceStyles matchFore:=&HFF, matchBack:=&HFFFFFF, _
                             badFore:=&HFF0000, badBack:=&HFFFFFF, bold:=True
        
        ' Configure indent guides using helper
        .ConfigureIndentGuideStyle &HC0C0C0, &HFFFFFF
        
        ' Configure call tip style
        .SetBack STYLE_CALLTIP, &HFFFFCC  ' Light yellow background
        .SetFore STYLE_CALLTIP, &H0       ' Black text
    End With
End Sub

'=========================================================================
' EXAMPLE 8: Exporting/Importing Style Configurations
'=========================================================================
Public Sub ExportAndImportStyleConfig(sci As SciWrapper)
    Dim config As String
    
    With sci.Styling
        ' Export a style configuration to string
        config = .ExportStyleConfig(SCE_C_COMMENT)
        Debug.Print "Comment style config: " & config
        
        ' Import it to another style
        .ImportStyleConfig SCE_C_COMMENTLINE, config
        
        ' Copy all attributes from one style to another
        .CopyStyle SCE_C_WORD, SCE_C_WORD2
    End With
End Sub

'=========================================================================
' EXAMPLE 9: Style Inspection and Debugging
'=========================================================================
Public Sub InspectStylesAtPosition(sci As SciWrapper, position As Long)
    Dim styleNum As Long
    Dim styleName As String
    Dim foreColor As Long
    Dim backColor As Long
    Dim isBold As Boolean
    Dim isItalic As Boolean
    
    With sci.Styling
        ' Get style at position
        styleNum = .GetStyleAt(position)
        
        ' Get style name (for debugging)
        styleName = .GetStyleName(styleNum, .Lexer)
        
        ' Get style attributes
        foreColor = .GetFore(styleNum)
        backColor = .GetBack(styleNum)
        isBold = .GetBold(styleNum)
        isItalic = .GetItalic(styleNum)
        
        Debug.Print "Position " & position & ":"
        Debug.Print "  Style: " & styleNum & " (" & styleName & ")"
        Debug.Print "  Foreground: &H" & Hex$(foreColor)
        Debug.Print "  Background: &H" & Hex$(backColor)
        Debug.Print "  Bold: " & isBold
        Debug.Print "  Italic: " & isItalic
    End With
End Sub

'=========================================================================
' EXAMPLE 10: VB P-Code Disassembly with Enums (for malware analysis)
'=========================================================================
Public Sub ConfigureVBPCodeWithEnums(sci As SciWrapper)
    With sci.Styling
        .Lexer = lexCPP
        .ClearAll
        
        ' Use CPP enum constants for P-code highlighting
        .ApplyCPPStyle SCE_C_DEFAULT, &H0, &HFFFFFF
        .ApplyCPPStyle SCE_C_COMMENT, &H8000, , False, True
        .ApplyCPPStyle SCE_C_COMMENTLINE, &H8000, , False, True
        .ApplyCPPStyle SCE_C_WORD, &H800000, , True          ' Call opcodes (dark red bold)
        .ApplyCPPStyle SCE_C_WORD2, &HFF, , True             ' Branch opcodes (red bold)
        .ApplyCPPStyle SCE_C_GLOBALCLASS, &H25208D, , True   ' Custom opcodes (purple bold)
        .ApplyCPPStyle SCE_C_STRING, &H800080                ' Strings
        .ApplyCPPStyle SCE_C_NUMBER, &H8000                  ' Numbers (blue)
        .ApplyCPPStyle SCE_C_OPERATOR, &H0
        
        ' Set P-code specific keywords
        Dim callOpcodes As String
        callOpcodes = "CallI2 CallI4 CallR4 CallR8 CallCy CallVar CallStr CallBool " & _
                      "CallFunc CallHresult FStCall ReturnHresult"
        .SetKeywords 0, callOpcodes
        
        Dim branchOpcodes As String
        branchOpcodes = "BranchF BranchT Branch EqI2 EqI4 NeI2 NeI4 LtI2 LtI4 " & _
                        "LeI2 LeI4 GtI2 GtI4 GeI2 GeI4 ForI2 ForI4 NextI2 NextI4"
        .SetKeywords 1, branchOpcodes
        
        .Colorise 0, -1
    End With
End Sub

'=========================================================================
' EXAMPLE 11: Quick Theme Switching
'=========================================================================
Public Sub SwitchToLightTheme(sci As SciWrapper)
    With sci.Styling
        .SetBack STYLE_DEFAULT, &HFFFFFF  ' White
        .SetFore STYLE_DEFAULT, &H0       ' Black
        .ClearAll
        
        sci.View.CaretLineBack = &HE8E8E8
        sci.Margins.SetBack 0, &HF0F0F0
    End With
End Sub

Public Sub SwitchToDarkTheme(sci As SciWrapper)
    With sci.Styling
        .SetBack STYLE_DEFAULT, &H1E1E1E  ' Dark gray
        .SetFore STYLE_DEFAULT, &HE0E0E0  ' Light gray
        .ClearAll
        
        sci.View.CaretLineBack = &H2D2D30
        sci.View.CaretFore = &HFFFFFF
        sci.Margins.SetBack 0, &H2D2D2D
    End With
End Sub

'=========================================================================
' Usage Summary
'=========================================================================
' The enums provide several benefits:
'
' 1. IntelliSense Support: Type "SCE_C_" and get autocomplete
' 2. Type Safety: Can't accidentally use wrong style number
' 3. Readability: SCE_C_COMMENT is clearer than "1"
' 4. Maintainability: Changes to constants propagate automatically
' 5. Self-Documenting: Code is more obvious about intent
'
' Example of clarity improvement:
'
' OLD WAY (magic numbers):
'   .SetFore 5, &HFF0000
'   .SetBold 5, True
'
' NEW WAY (enum constants):
'   .ApplyCPPStyle SCE_C_WORD, &HFF0000, , True
'
' Much clearer that we're styling keywords!
'=========================================================================

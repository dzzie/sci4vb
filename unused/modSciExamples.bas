Attribute VB_Name = "Module2"
'=========================================================================
' modSciExamples - Advanced Usage Examples for SciWrapper
'=========================================================================
Option Explicit


'=========================================================================
' Scintilla Constants (subset for property exposure)
'=========================================================================
' Code page
Public Const SC_CP_UTF8 = 65001
Public Const SC_CP_DBCS = 1

' EOL mode
Public Const SC_EOL_CRLF = 0
Public Const SC_EOL_CR = 1
Public Const SC_EOL_LF = 2

' Message constants
Public Const SCI_GETDIRECTFUNCTION = 2184
Public Const SCI_GETDIRECTPOINTER = 2185
Public Const SCI_SETCODEPAGE = 2037
Public Const SCI_SETEOLMODE = 2031
Public Const SCI_SETCARETLINEVISIBLE = 2096
Public Const SCI_SETCARETLINEBACK = 2098
Public Const SCI_SETTABWIDTH = 2036
Public Const SCI_SETUSETABS = 2124

'=========================================================================
' EXAMPLE 1: Advanced Syntax Highlighting Configuration
'=========================================================================
Public Sub ConfigureAdvancedCppHighlighting(sci As SciWrapper)
    With sci.style
        ' Set lexer
        .Lexer = lexCPP
        
        ' Configure base styles
        .ConfigureStyle 0, &H0, &HFFFFFF, "Consolas", 10  ' Default
        .ConfigureStyle 1, &H8000, &HFFFFFF, "Consolas", 10, False, True    ' Comment (green, italic)
        .ConfigureStyle 2, &H8080, &HFFFFFF, "Consolas", 10    ' Line comment
        .ConfigureStyle 3, &H808080, &HFFFFFF, "Consolas", 10  ' Doc comment
        .ConfigureStyle 4, &H8000, &HFFFFFF, "Consolas", 10  ' Number
        .ConfigureStyle 5, &H8000, &HFFFFFF, "Consolas", 10  ' Keyword (blue)
        .ConfigureStyle 6, &H8000, &HFFFFFF, "Consolas", 10  ' String
        
        ' Set keywords
        Dim keywords As String
        keywords = "class public private protected virtual static const " & _
                   "void int long char bool float double return if else " & _
                   "while for switch case break continue delete new this"
        .SetKeywords 0, keywords
        
        ' Set keyword set 2 (types)
        keywords = "string vector map set list iostream fstream"
        .SetKeywords 1, keywords
        
        ' Apply coloring
        .Colorise 0, sci.doc.TextLength
    End With
End Sub

'=========================================================================
' EXAMPLE 2: Smart Indentation
'=========================================================================
Public Sub SmartIndent(sci As SciWrapper)
    Dim currentLine As Long
    Dim previousLine As Long
    Dim previousIndent As Long
    Dim lineText As String
    
    currentLine = sci.lines.LineFromPosition(sci.sel.currentPos)
    
    If currentLine > 0 Then
        previousLine = currentLine - 1
        previousIndent = sci.lines.GetLineIndentation(previousLine)
        lineText = Trim$(sci.lines.GetLine(previousLine))
        
        ' Check if previous line should increase indent
        If Right$(lineText, 1) = "{" Or _
           Right$(lineText, 5) = "Then" Or _
           Left$(lineText, 3) = "If " Then
            previousIndent = previousIndent + sci.edit.TabWidth
        End If
        
        ' Apply indent to current line
        sci.lines.SetLineIndentation currentLine, previousIndent
        
        ' Move caret to end of indentation
        Dim indentPos As Long
        indentPos = sci.lines.GetLineIndentPosition(currentLine)
        sci.sel.GotoPos indentPos
    End If
End Sub

'=========================================================================
' EXAMPLE 3: Bracket Matching
'=========================================================================
Public Sub HighlightMatchingBracket(sci As SciWrapper)
    Dim pos As Long
    Dim matchPos As Long
    Dim ch As String
    
    pos = sci.sel.currentPos
    
    ' Check character at position
    If pos > 0 And pos <= sci.doc.TextLength Then
        ch = Mid$(sci.doc.text, pos, 1)
        
        If InStr("(){}[]", ch) > 0 Then
            matchPos = FindMatchingBracket(sci, pos, ch)
            
            If matchPos >= 0 Then
                ' Highlight both brackets using indicators
                sci.Indic.Current = 0
                sci.Indic.SetStyle 0, indicBox
                sci.Indic.SetFore 0, &HFF00  ' Green
                
                sci.Indic.FillRange pos - 1, 1
                sci.Indic.FillRange matchPos, 1
            End If
        End If
    End If
End Sub

Private Function FindMatchingBracket(sci As SciWrapper, startPos As Long, bracket As String) As Long
    Dim text As String
    Dim pos As Long
    Dim depth As Long
    Dim searchChar As String
    Dim matchChar As String
    Dim direction As Long
    
    text = sci.doc.text
    
    ' Determine search direction and matching bracket
    Select Case bracket
        Case "(": matchChar = ")": direction = 1
        Case ")": matchChar = "(": direction = -1
        Case "{": matchChar = "}": direction = 1
        Case "}": matchChar = "{": direction = -1
        Case "[": matchChar = "]": direction = 1
        Case "]": matchChar = "[": direction = -1
        Case Else: FindMatchingBracket = -1: Exit Function
    End Select
    
    depth = 1
    pos = startPos + direction
    
    Do While pos > 0 And pos <= Len(text)
        searchChar = Mid$(text, pos, 1)
        
        If searchChar = bracket Then
            depth = depth + 1
        ElseIf searchChar = matchChar Then
            depth = depth - 1
            If depth = 0 Then
                FindMatchingBracket = pos - 1
                Exit Function
            End If
        End If
        
        pos = pos + direction
    Loop
    
    FindMatchingBracket = -1
End Function

'=========================================================================
' EXAMPLE 4: Error Marker with Indicator
'=========================================================================
Public Sub MarkError(sci As SciWrapper, startPos As Long, length As Long, errorMsg As String)
    ' Configure error indicator
    sci.Indic.SetStyle 1, indicSquiggle
    sci.Indic.SetFore 1, &HFF  ' Red squiggle
    sci.Indic.SetUnder 1, True
    
    ' Apply indicator
    sci.Indic.Current = 1
    sci.Indic.FillRange startPos, length
    
    ' Optionally add a marker in the margin
    Dim line As Long
    line = sci.lines.LineFromPosition(startPos)
    
    sci.Mark.Define 2, markerCircle
    sci.Mark.SetBack 2, &HFF  ' Red
    sci.Mark.Add line, 2
End Sub

Public Sub ClearAllErrors(sci As SciWrapper)
    ' Clear error indicators
    sci.Indic.Current = 1
    sci.Indic.ClearRange 0, sci.doc.TextLength
    
    ' Clear error markers
    sci.Mark.DeleteAll 2
End Sub

'=========================================================================
' EXAMPLE 5: Advanced Find/Replace with Progress
'=========================================================================
Public Function AdvancedReplaceAll(sci As SciWrapper, findText As String, replaceText As String, Optional showProgress As Boolean = False) As Long
    Dim count As Long
    Dim pos As Long
    Dim totalMatches As Long
    Dim progressForm As Form
    
    ' First pass: count matches
    If showProgress Then
        pos = 0
        Do
            pos = sci.Search.Find(findText, pos, True, False, False, False)
            If pos >= 0 Then
                totalMatches = totalMatches + 1
                pos = pos + 1
            End If
        Loop While pos >= 0
        
        ' Show progress form (pseudo-code)
        ' Set progressForm = New frmProgress
        ' progressForm.Maximum = totalMatches
        ' progressForm.Show vbModeless
    End If
    
    ' Begin undo action for atomic operation
    sci.doc.BeginUndoAction
    
    pos = 0
    count = 0
    
    Do
        pos = sci.Search.Find(findText, pos, True, False, False, False)
        If pos >= 0 Then
            sci.Search.SetTargetRange pos, pos + Len(findText)
            sci.Search.ReplaceTarget replaceText
            
            count = count + 1
            pos = pos + Len(replaceText)
            
            ' Update progress
            If showProgress Then
                ' progressForm.Value = count
                DoEvents
            End If
        End If
    Loop While pos >= 0
    
    sci.doc.EndUndoAction
    
    ' Close progress
    If showProgress Then
        ' Unload progressForm
    End If
    
    AdvancedReplaceAll = count
End Function

'=========================================================================
' EXAMPLE 6: Code Outlining (Folding) Setup
'=========================================================================
Public Sub SetupCodeFolding(sci As SciWrapper)
    With sci
        ' Configure folding margin
        .Margins.ConfigureFolding 2, 16
        .Margins.SetSensitive 2, True
        
        ' Configure fold markers
        .Mark.Define 25, markerBoxPlus
        .Mark.Define 26, markerBoxMinus
        .Mark.Define 27, markerVLine
        .Mark.Define 28, markerLCorner
        .Mark.Define 29, markerTCorner
        .Mark.Define 30, markerBoxPlusConnected
        .Mark.Define 31, markerBoxMinusConnected
        
        ' Set marker colors
        Dim i As Long
        For i = 25 To 31
            .Mark.SetFore i, &HFFFFFF
            .Mark.SetBack i, &H808080
        Next i
        
        ' Set fold flags
        .Fold.SetFoldFlags foldLineBeforeExpanded Or foldLineAfterExpanded
    End With
End Sub

'=========================================================================
' EXAMPLE 7: Intelligent Autocomplete
'=========================================================================
Public Sub ShowIntelligentAutocomplete(sci As SciWrapper)
    Dim currentPos As Long
    Dim lineStart As Long
    Dim line As Long
    Dim lineText As String
    Dim wordStart As Long
    Dim partial As String
    Dim completions As String
    
    currentPos = sci.sel.currentPos
    line = sci.lines.LineFromPosition(currentPos)
    lineStart = sci.lines.PositionFromLine(line)
    lineText = sci.lines.GetLine(line)
    
    ' Find start of current word
    wordStart = currentPos
    Do While wordStart > lineStart
        Dim ch As String
        ch = Mid$(sci.doc.text, wordStart, 1)
        If Not IsWordChar(ch) Then Exit Do
        wordStart = wordStart - 1
    Loop
    
    ' Get partial word
    partial = Mid$(sci.doc.text, wordStart + 1, currentPos - wordStart)
    
    ' Build completion list (in real app, parse document for identifiers)
    completions = BuildCompletionList(partial)
    
    If Len(completions) > 0 Then
        sci.Autoc.IgnoreCase = True
        sci.Autoc.AutoHide = True
        sci.Autoc.Show Len(partial), completions
    End If
End Sub

Private Function IsWordChar(ch As String) As Boolean
    IsWordChar = (ch >= "A" And ch <= "Z") Or _
                 (ch >= "a" And ch <= "z") Or _
                 (ch >= "0" And ch <= "9") Or _
                 ch = "_"
End Function

Private Function BuildCompletionList(partial As String) As String
    ' Dummy implementation - in real app, scan document for identifiers
    BuildCompletionList = "Function Sub Property Dim Private Public " & _
                          "Integer Long String Boolean Variant " & _
                          "If Then Else ElseIf End While For Next"
End Function

'=========================================================================
' EXAMPLE 8: Multi-Select and Rectangular Selection
'=========================================================================
Public Sub DemoMultipleSelections(sci As SciWrapper)
    ' Enable multiple selections
    sci.sel.SelectionMode = selRectangle
    
    ' Programmatically add multiple selections
    sci.sel.ClearSelections
    sci.sel.AddSelection 100, 95
    sci.sel.AddSelection 200, 195
    sci.sel.AddSelection 300, 295
    
    ' Main selection
    sci.sel.MainSelection = 0
End Sub

'=========================================================================
' EXAMPLE 9: Custom Context Menu Handler
'=========================================================================
Public Sub ShowCustomContextMenu(sci As SciWrapper, x As Long, y As Long)
    ' Get word at click position
    Dim pos As Long
    pos = sci.SciMsg(2022, x, y)  ' SCI_POSITIONFROMPOINT
    
    Dim wordStart As Long
    Dim wordEnd As Long
    wordStart = sci.SciMsg(2266, pos, 1)  ' SCI_WORDSTARTPOSITION
    wordEnd = sci.SciMsg(2267, pos, 1)    ' SCI_WORDENDPOSITION
    
    Dim word As String
    If wordEnd > wordStart Then
        word = Mid$(sci.doc.text, wordStart + 1, wordEnd - wordStart)
    End If
    
    ' Show context menu with word-specific actions
    ' (Would use PopupMenu in real implementation)
    Debug.Print "Context menu for word: " & word
End Sub

'=========================================================================
' EXAMPLE 10: Document Comparison Highlighting
'=========================================================================
Public Sub HighlightDifferences(sci As SciWrapper, originalText As String)
    Dim currentText As String
    Dim i As Long
    Dim startDiff As Long
    Dim inDiff As Boolean
    
    currentText = sci.doc.text
    startDiff = -1
    inDiff = False
    
    ' Simple character-by-character comparison
    For i = 1 To IIf(Len(currentText) > Len(originalText), Len(currentText), Len(originalText))
        Dim currentChar As String
        Dim originalChar As String
        
        If i <= Len(currentText) Then currentChar = Mid$(currentText, i, 1)
        If i <= Len(originalText) Then originalChar = Mid$(originalText, i, 1)
        
        If currentChar <> originalChar Then
            If Not inDiff Then
                startDiff = i - 1
                inDiff = True
            End If
        Else
            If inDiff Then
                ' Mark difference region
                sci.Indic.Current = 2
                sci.Indic.SetStyle 2, indicRoundBox
                sci.Indic.SetFore 2, &HFFFF  ' Yellow
                sci.Indic.FillRange startDiff, i - startDiff - 1
                inDiff = False
            End If
        End If
    Next i
    
    ' Handle trailing difference
    If inDiff Then
        sci.Indic.Current = 2
        sci.Indic.FillRange startDiff, Len(currentText) - startDiff
    End If
End Sub

'=========================================================================
' EXAMPLE 11: Performance: Batch Updates
'=========================================================================
Public Sub BatchUpdate(sci As SciWrapper, lines() As String)
    Dim i As Long
    Dim text As String
    
    ' Disable visual updates
    sci.SciMsg 2011, 0, 0  ' SCI_SETUNDOCOLLECTION = false
    
    ' Build text
    For i = LBound(lines) To UBound(lines)
        text = text & lines(i) & vbCrLf
    Next i
    
    ' Set all at once
    sci.doc.text = text
    
    ' Re-enable undo collection
    sci.SciMsg 2011, 1, 0  ' SCI_SETUNDOCOLLECTION = true
    
    ' Trigger syntax coloring
    sci.style.Colorise 0, sci.doc.TextLength
End Sub

'=========================================================================
' EXAMPLE 12: Export to HTML with Syntax Highlighting
'=========================================================================
Public Function ExportToHTML(sci As SciWrapper) As String
    Dim html As String
    Dim i As Long
    Dim length As Long
    Dim style As Long
    Dim ch As String
    Dim currentStyle As Long
    
    html = "<pre style='font-family: Consolas, monospace; font-size: 10pt;'>"
    
    length = sci.doc.TextLength
    currentStyle = -1
    
    For i = 0 To length - 1
        style = sci.SciMsg(2010, i, 0)  ' SCI_GETSTYLEAT
        
        If style <> currentStyle Then
            If currentStyle <> -1 Then html = html & "</span>"
            html = html & "<span style='color: " & ColorToHTML(sci.style.GetFore(style)) & ";'>"
            currentStyle = style
        End If
        
        ch = Mid$(sci.doc.text, i + 1, 1)
        
        Select Case ch
            Case "<": html = html & "&lt;"
            Case ">": html = html & "&gt;"
            Case "&": html = html & "&amp;"
            Case vbCr: ' Skip
            Case vbLf: html = html & "<br>" & vbCrLf
            Case Else: html = html & ch
        End Select
    Next i
    
    If currentStyle <> -1 Then html = html & "</span>"
    html = html & "</pre>"
    
    ExportToHTML = html
End Function

Private Function ColorToHTML(color As Long) As String
    Dim r As Long, g As Long, b As Long
    r = color And &HFF
    g = (color \ &H100) And &HFF
    b = (color \ &H10000) And &HFF
    ColorToHTML = "#" & Right$("0" & Hex$(r), 2) & Right$("0" & Hex$(g), 2) & Right$("0" & Hex$(b), 2)
End Function

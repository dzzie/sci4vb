Attribute VB_Name = "modBreakpointExample"
Option Explicit

'=========================================================================
' Breakpoint Management Examples
' Click margin to toggle breakpoints automatically!
'=========================================================================

'=========================================================================
' Example 1: Basic Setup - Enable Breakpoint Margin
'=========================================================================
Public Sub SetupBreakpoints()
    ' One line to configure breakpoint margin
    sci.ConfigureBreakpointMargin
    
    ' Now clicking margin 2 will automatically toggle breakpoints!
End Sub

'=========================================================================
' Example 2: Programmatically Add/Remove Breakpoints
'=========================================================================
Public Sub ManageBreakpoints()
    ' Add breakpoint at line 10
    sci.AddBreakpoint 10
    
    ' Remove breakpoint at line 10
    sci.RemoveBreakpoint 10
    
    ' Check if line has breakpoint
    If sci.HasBreakpoint(10) Then
        Debug.Print "Line 10 has a breakpoint"
    End If
    
    ' Clear all breakpoints
    sci.ClearAllBreakpoints
End Sub

'=========================================================================
' Example 3: Navigate Between Breakpoints
'=========================================================================
Public Sub GotoNextBreakpoint()
    Dim currentLine As Long
    Dim nextBP As Long
    
    currentLine = sci.Lines.LineFromPosition(sci.sel.CurrentPos)
    nextBP = sci.GetNextBreakpoint(currentLine + 1)
    
    If nextBP >= 0 Then
        sci.sel.GotoLine nextBP
        sci.Lines.ScrollCaret
    Else
        MsgBox "No more breakpoints below", vbInformation
    End If
End Sub

Public Sub GotoPreviousBreakpoint()
    Dim currentLine As Long
    Dim prevBP As Long
    
    currentLine = sci.Lines.LineFromPosition(sci.sel.CurrentPos)
    prevBP = sci.GetPreviousBreakpoint(currentLine - 1)
    
    If prevBP >= 0 Then
        sci.sel.GotoLine prevBP
        sci.Lines.ScrollCaret
    Else
        MsgBox "No more breakpoints above", vbInformation
    End If
End Sub

'=========================================================================
' Example 4: Get All Breakpoint Lines
'=========================================================================
Public Sub ListAllBreakpoints()
    Dim lines() As Long
    Dim i As Long
    Dim msg As String
    
    lines = sci.GetBreakpointLines()
    
    If UBound(lines) >= LBound(lines) Then
        msg = "Breakpoints at lines:" & vbCrLf
        For i = LBound(lines) To UBound(lines)
            msg = msg & "  Line " & (lines(i) + 1) & vbCrLf
        Next i
        MsgBox msg, vbInformation, "Breakpoints"
    Else
        MsgBox "No breakpoints set", vbInformation
    End If
End Sub

'=========================================================================
' Example 5: Save/Load Breakpoints to File
'=========================================================================
Public Sub SaveBreakpoints(filename As String)
    Dim lines() As Long
    Dim i As Long
    Dim fileNum As Integer
    
    lines = sci.GetBreakpointLines()
    
    If UBound(lines) >= LBound(lines) Then
        fileNum = FreeFile
        Open filename For Output As #fileNum
        
        For i = LBound(lines) To UBound(lines)
            Print #fileNum, lines(i)
        Next i
        
        Close #fileNum
    End If
End Sub

Public Sub LoadBreakpoints(filename As String)
    Dim fileNum As Integer
    Dim line As Long
    
    On Error Resume Next
    
    ' Clear existing breakpoints
    sci.ClearAllBreakpoints
    
    ' Load from file
    fileNum = FreeFile
    Open filename For Input As #fileNum
    
    Do While Not EOF(fileNum)
        Input #fileNum, line
        If Err.Number = 0 Then
            sci.AddBreakpoint line
        End If
    Loop
    
    Close #fileNum
End Sub

'=========================================================================
' Example 6: Toggle Breakpoint with Keyboard Shortcut (F9)
'=========================================================================
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF9 Then
        ' F9 to toggle breakpoint on current line
        Dim line As Long
        line = sci.Lines.LineFromPosition(sci.sel.CurrentPos)
        
        If sci.HasBreakpoint(line) Then
            sci.RemoveBreakpoint line
        Else
            sci.AddBreakpoint line
        End If
    ElseIf KeyCode = vbKeyF8 And Shift = 0 Then
        ' F8 to go to next breakpoint
        GotoNextBreakpoint
    ElseIf KeyCode = vbKeyF8 And Shift = vbShiftMask Then
        ' Shift+F8 to go to previous breakpoint
        GotoPreviousBreakpoint
    End If
End Sub

'=========================================================================
' Example 7: Conditional Breakpoints (Track in Collection)
'=========================================================================
Private Type ConditionalBreakpoint
    line As Long
    condition As String
End Type

Private m_ConditionalBPs As Collection

Public Sub AddConditionalBreakpoint(line As Long, condition As String)
    Dim bp As ConditionalBreakpoint
    
    If m_ConditionalBPs Is Nothing Then
        Set m_ConditionalBPs = New Collection
    End If
    
    bp.line = line
    bp.condition = condition
    
    m_ConditionalBPs.Add bp, "L" & line
    sci.AddBreakpoint line
End Sub

Public Function GetBreakpointCondition(line As Long) As String
    Dim bp As ConditionalBreakpoint
    
    On Error Resume Next
    bp = m_ConditionalBPs("L" & line)
    
    If Err.Number = 0 Then
        GetBreakpointCondition = bp.condition
    Else
        GetBreakpointCondition = ""
    End If
End Function

'=========================================================================
' Example 8: Visual Breakpoint Styles (Different Colors)
'=========================================================================
Public Sub ConfigureBreakpointStyles()
    ' Regular breakpoint (marker 1) - Red
    sci.Mark.Define 1, markerCircle
    sci.Mark.SetFore 1, &HFFFFFF
    sci.Mark.SetBack 1, &HFF
    
    ' Disabled breakpoint (marker 2) - Gray
    sci.Mark.Define 2, markerCircle
    sci.Mark.SetFore 2, &HFFFFFF
    sci.Mark.SetBack 2, &H808080
    
    ' Conditional breakpoint (marker 3) - Yellow
    sci.Mark.Define 3, markerCircle
    sci.Mark.SetFore 3, &H0
    sci.Mark.SetBack 3, &HFFFF
End Sub

'=========================================================================
' Example 9: Breakpoint Hit Counter
'=========================================================================
Private m_BreakpointHits As Collection

Public Sub InitializeBreakpointTracking()
    Set m_BreakpointHits = New Collection
End Sub

Public Sub OnBreakpointHit(line As Long)
    Dim hitCount As Long
    
    On Error Resume Next
    hitCount = m_BreakpointHits("L" & line)
    
    If Err.Number <> 0 Then
        hitCount = 0
        m_BreakpointHits.Add 0, "L" & line
    End If
    
    hitCount = hitCount + 1
    m_BreakpointHits.Remove "L" & line
    m_BreakpointHits.Add hitCount, "L" & line
    
    Debug.Print "Breakpoint at line " & line & " hit " & hitCount & " times"
End Sub

'=========================================================================
' Example 10: Breakpoint Panel/List
'=========================================================================
Public Sub PopulateBreakpointList(lstBreakpoints As ListBox)
    Dim lines() As Long
    Dim i As Long
    Dim lineText As String
    
    lstBreakpoints.Clear
    lines = sci.GetBreakpointLines()
    
    If UBound(lines) >= LBound(lines) Then
        For i = LBound(lines) To UBound(lines)
            ' Get line text
            lineText = Trim$(sci.Lines.GetLine(lines(i)))
            
            ' Add to list with line number
            lstBreakpoints.AddItem "Line " & (lines(i) + 1) & ": " & Left$(lineText, 50)
            lstBreakpoints.ItemData(lstBreakpoints.NewIndex) = lines(i)
        Next i
    End If
End Sub

Private Sub lstBreakpoints_DblClick()
    ' Jump to breakpoint on double-click
    If lstBreakpoints.ListIndex >= 0 Then
        Dim line As Long
        line = lstBreakpoints.ItemData(lstBreakpoints.ListIndex)
        sci.sel.GotoLine line
        sci.Lines.ScrollCaret
    End If
End Sub

'=========================================================================
' Example 11: Context Menu for Breakpoint Margin
'=========================================================================
Private Sub ShowBreakpointContextMenu(line As Long)
    ' Show popup menu with breakpoint options
    Dim menu As String
    
    If sci.HasBreakpoint(line) Then
        menu = "Remove Breakpoint" & vbCrLf & _
               "Disable Breakpoint" & vbCrLf & _
               "Edit Condition..."
    Else
        menu = "Add Breakpoint" & vbCrLf & _
               "Add Conditional Breakpoint..."
    End If
    
    ' Would show PopupMenu here
    Debug.Print menu
End Sub

'=========================================================================
' Example 12: Complete Integration - Debugger-Style UI
'=========================================================================
Public Sub SetupDebuggerUI()
    ' Configure margins
    sci.Margins.ConfigureLineNumbers 0, 40
    sci.ConfigureBreakpointMargin
    
    ' Configure different breakpoint styles
    ConfigureBreakpointStyles
    
    ' Initialize tracking
    InitializeBreakpointTracking
    
    ' Form keyboard shortcuts
    ' F9 = Toggle breakpoint
    ' F8 = Next breakpoint
    ' Shift+F8 = Previous breakpoint
    ' Ctrl+Shift+F9 = Clear all breakpoints
End Sub

'=========================================================================
' Summary of Breakpoint API
'=========================================================================
'
' Setup:
'   sci.ConfigureBreakpointMargin           ' Enable breakpoint margin
'
' Add/Remove:
'   sci.AddBreakpoint line                  ' Add breakpoint
'   sci.RemoveBreakpoint line               ' Remove breakpoint
'   sci.ClearAllBreakpoints                 ' Clear all
'
' Query:
'   has = sci.HasBreakpoint(line)           ' Check if exists
'   lines() = sci.GetBreakpointLines()      ' Get all breakpoint lines
'
' Navigate:
'   line = sci.GetNextBreakpoint(fromLine)      ' Find next
'   line = sci.GetPreviousBreakpoint(fromLine)  ' Find previous
'
' Automatic:
'   Click margin 2 to toggle breakpoint (handled automatically!)
'
'=========================================================================

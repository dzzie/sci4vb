VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBreakPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Author:  David Zimmer <dzzie@yahoo.com> + Claude.ai
'Site:    http://sandsprite.com
'License: MIT
'---------------------------------------------------

'=========================================================================
' Public Breakpoint API
'=========================================================================

'scintinilla keeps track of the breakpoint list for us internally
 
'SciWrapper handles the margin clicks to add/remove
'   If the user clicks on the margin to set/remove in the UI,
'   the wrapper will raiseEvent UserBreakpointToggle to let you sync your execution engine (if running)

'we dont fireevents on user api set bp and dont include a toggle because thats a margin click thing

Option Explicit

Private m_Sci As SciWrapper


Friend Sub Init(sciCtrl As SciWrapper)
    Set m_Sci = sciCtrl
End Sub

Property Let Visible(show As Boolean)
    If show Then
        m_Sci.Margins.ShowBreakpointMargin
    Else
        m_Sci.Margins.HideBreakpointMargin
    End If
End Property

Public Sub add(line As Long)
     m_Sci.Mark.add line, BREAKPOINT_MARKER
End Sub

Public Sub remove(line As Long)
     m_Sci.Mark.Delete line, BREAKPOINT_MARKER
End Sub

Public Sub clearAll()
     m_Sci.Mark.DeleteAll BREAKPOINT_MARKER
End Sub

Public Function isSet(ByVal line As Long) As Boolean
    Dim markers As Long
    markers = m_Sci.Mark.GetMarkers(line)
    isSet = ((markers And BREAKPOINT_MASK) <> 0)
End Function

Public Function nextBp(ByVal fromLine As Long) As Long
    nextBp = m_Sci.Mark.FindNext(fromLine, BREAKPOINT_MASK)
End Function

Public Function prevBp(ByVal fromLine As Long) As Long
    prevBp = m_Sci.Mark.FindPrevious(fromLine, BREAKPOINT_MASK)
End Function

Public Function getAll() As Collection
    ' Return collection of all lines with breakpoints
    Dim c As New Collection
    Dim Count As Long
    Dim line As Long
    Dim maxLine As Long
 
    line = 0
    maxLine = m_Sci.lines.Count - 1
    
    Do
        line = m_Sci.Mark.FindNext(line, BREAKPOINT_MASK)
        If line >= 0 And line <= maxLine Then
            c.add line
            line = line + 1
        End If
    Loop While line >= 0 And line <= maxLine
    
    Set getAll = c

End Function

'this is the user api version
'Sub Toggle(ByVal line As Long)
'    Dim markers As Long, isAdding As Boolean, Cancel As Boolean
'
'    markers = m_Sci.Mark.GetMarkers(line)
'    isAdding = ((markers And BREAKPOINT_MASK) = 0)
'
'    If isAdding Then
'        m_Sci.Mark.Add line, BREAKPOINT_MARKER
'    Else
'        m_Sci.Mark.Delete line, BREAKPOINT_MARKER
'    End If
'End Sub



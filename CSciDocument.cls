VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSciDocument"
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
' CSciDocument - Document-level text operations
'=========================================================================
Option Explicit

Private m_Sci As SciWrapper
Private m_WordWrap As Boolean

' Scintilla messages
Private Const SCI_GETTEXT = 2182
Private Const SCI_SETTEXT = 2181
Private Const SCI_GETTEXTLENGTH = 2183
Private Const SCI_CLEARALL = 2004
Private Const SCI_ADDTEXT = 2001
Private Const SCI_INSERTTEXT = 2003
Private Const SCI_APPENDTEXT = 2282
Private Const SCI_GETCHARACTERPOINTER = 2520
Private Const SCI_ALLOCATE = 2446
Private Const SCI_SETREADONLY = 2171
Private Const SCI_GETREADONLY = 2140
Private Const SCI_CANUNDO = 2174
Private Const SCI_CANREDO = 2016
Private Const SCI_UNDO = 2176
Private Const SCI_REDO = 2011
Private Const SCI_EMPTYUNDOBUFFER = 2175
Private Const SCI_BEGINUNDOACTION = 2078
Private Const SCI_ENDUNDOACTION = 2079
Private Const SCI_SETSAVEPOINT = 2014
Private Const SCI_GETMODIFY = 2159
Private Const SCI_GETCODEPAGE = 2137
Private Const SC_CP_UTF8 = 65001


Friend Sub Init(sciCtrl As SciWrapper)
    Set m_Sci = sciCtrl
End Sub

Public Sub PasteConvertEndings(convert As Boolean)
  m_Sci.SciMsg 2467, convert, 0
End Sub

Public Property Get WordWrap() As Boolean 'If set to true the document will wrap lines which are longer than itself.  If false then it will dsiplay normally.
    WordWrap = IIf(m_WordWrap = 0, False, True)
End Property

Public Property Let WordWrap(ByVal wrap As Boolean)
    m_WordWrap = IIf(wrap, 1, 0)
    m_Sci.SciMsg 268, m_WordWrap, 0
End Property

Property Let LineNumbers(show As Boolean)
    If show Then
        m_Sci.Margins.ShowLineNumbers
    Else
        m_Sci.Margins.HideLineNumbers
    End If
End Property

Property Let MouseDwellTime(ms As Long)
    m_Sci.SciMsg 2264, ms, 0
End Property

Public Property Let Folding(ByVal enabled As Boolean)
    If enabled Then
        m_Sci.Margins.SetWidth 2, 13
    Else
        m_Sci.Margins.SetWidth 2, 0
    End If
    'InitFolding enabled
End Property

Property Get Folding() As Boolean
    Folding = (m_Sci.Margins.GetWidth(2) = 13) 'this matches the hardcoded let width 0/13
End Property


Sub ClearLastExec()
    SetExecLine -1, False, False
End Sub

Sub SetExecLine(ByVal curLine As Long, Optional isActive As Boolean = True, Optional jumpTo As Boolean = True)
    
    Dim startPos As Long, endPos As Long
    Static lastEIP As Long
    
    If curLine = -1 Then curLine = lastEIP
    
    With m_Sci
        
        If isActive Then
        
            If lastEIP >= 0 Then
                .Mark.Delete lastEIP, 2
                .Mark.Delete lastEIP, 3
                ' Refresh old line
                startPos = .lines.PositionFromLine(lastEIP)
                endPos = .lines.PositionFromLine(lastEIP + 1)
                .style.Colorise startPos, endPos
            End If
        
            .Mark.add curLine, 2
            .Mark.add curLine, 3
        Else
            .Mark.Delete curLine, 2
            .Mark.Delete curLine, 3
        End If
        
        'force a refresh of the specified line or it might not catch it..
        startPos = .lines.PositionFromLine(curLine)
        endPos = .lines.PositionFromLine(curLine + 1)
        .style.Colorise startPos, endPos
    
        If jumpTo Then .sel.GotoLine curLine
        lastEIP = curLine
    End With
    
End Sub



'=========================================================================
' Text Content - FIXED VERSION
'=========================================================================
Public Property Get text() As String
    Dim length As Long
    Dim buffer() As Byte
    Dim codePage As Long
    Dim result As String
    
    ' Get text length (in bytes for UTF-8)
    length = m_Sci.SciMsg(SCI_GETTEXTLENGTH)
    If length = 0 Then
        text = ""
        Exit Property
    End If
    
    ' Check if Scintilla is using UTF-8
    codePage = m_Sci.SciMsg(SCI_GETCODEPAGE)
    
    ' Allocate byte buffer for UTF-8 text
    ReDim buffer(0 To length) As Byte
    
    ' Get text as bytes (Scintilla returns UTF-8)
    m_Sci.SciMsgPtr SCI_GETTEXT, length + 1, VarPtr(buffer(0))
    
    ' Convert from UTF-8 to VB6 Unicode string
    If codePage = SC_CP_UTF8 Then
        ' UTF-8 mode - convert properly
        result = UTF8BytesToString(buffer, length)
    Else
        ' ANSI mode - use StrConv
        result = StrConv(buffer, vbUnicode)
        result = Left$(result, length)
    End If
    
    text = result
End Property

Public Property Let text(ByVal newText As String)
    Dim utf8Bytes() As Byte
    Dim codePage As Long
    
    If Len(newText) = 0 Then
        m_Sci.SciMsg SCI_CLEARALL
        Exit Property
    End If
    
    codePage = m_Sci.SciMsg(SCI_GETCODEPAGE)
    
    If codePage = SC_CP_UTF8 Then
        ' Convert VB6 Unicode string to UTF-8 bytes
        utf8Bytes = StringToUTF8Bytes(newText)
        m_Sci.SciMsgPtr SCI_SETTEXT, 0, VarPtr(utf8Bytes(0))
    Else
        ' ANSI mode - use normal string
        m_Sci.SciMsgStr SCI_SETTEXT, 0, newText
    End If
End Property

Public Property Get TextLength() As Long
    TextLength = m_Sci.SciMsg(SCI_GETTEXTLENGTH)
End Property


'=========================================================================
' Document Modification
'=========================================================================
Public Sub Clear()
    m_Sci.SciMsg SCI_CLEARALL
End Sub

Public Sub AddText(ByVal text As String)
    Dim utf8Bytes() As Byte
    Dim codePage As Long
    
    If Len(text) = 0 Then Exit Sub
    
    codePage = m_Sci.SciMsg(SCI_GETCODEPAGE)
    
    If codePage = SC_CP_UTF8 Then
        ' Convert to UTF-8
        utf8Bytes = StringToUTF8Bytes(text)
        m_Sci.SciMsgPtr SCI_ADDTEXT, UBound(utf8Bytes), VarPtr(utf8Bytes(0))
    Else
        ' ANSI mode
        m_Sci.SciMsgStr SCI_ADDTEXT, Len(text), text
    End If
End Sub

Public Sub InsertText(ByVal position As Long, ByVal text As String)
    Dim utf8Bytes() As Byte
    Dim codePage As Long
    
    If Len(text) = 0 Then Exit Sub
    
    codePage = m_Sci.SciMsg(SCI_GETCODEPAGE)
    
    If codePage = SC_CP_UTF8 Then
        ' Convert to UTF-8
        utf8Bytes = StringToUTF8Bytes(text)
        m_Sci.SciMsgPtr SCI_INSERTTEXT, position, VarPtr(utf8Bytes(0))
    Else
        ' ANSI mode
        m_Sci.SciMsgStr SCI_INSERTTEXT, position, text
    End If
End Sub

Public Sub AppendText(ByVal text As String)
    Dim utf8Bytes() As Byte
    Dim codePage As Long
    
    If Len(text) = 0 Then Exit Sub
    
    codePage = m_Sci.SciMsg(SCI_GETCODEPAGE)
    
    If codePage = SC_CP_UTF8 Then
        ' Convert to UTF-8
        utf8Bytes = StringToUTF8Bytes(text)
        m_Sci.SciMsgPtr SCI_APPENDTEXT, UBound(utf8Bytes), VarPtr(utf8Bytes(0))
    Else
        ' ANSI mode
        m_Sci.SciMsgStr SCI_APPENDTEXT, Len(text), text
    End If
End Sub

'=========================================================================
' Read-Only
'=========================================================================
Public Property Get ReadOnly() As Boolean
    ReadOnly = (m_Sci.SciMsg(SCI_GETREADONLY) <> 0)
End Property

Public Property Let ReadOnly(ByVal value As Boolean)
    m_Sci.SciMsg SCI_SETREADONLY, IIf(value, 1, 0)
End Property

'=========================================================================
' Undo/Redo
'=========================================================================
Public Property Get CanUndo() As Boolean
    CanUndo = (m_Sci.SciMsg(SCI_CANUNDO) <> 0)
End Property

Public Property Get CanRedo() As Boolean
    CanRedo = (m_Sci.SciMsg(SCI_CANREDO) <> 0)
End Property

Public Sub Undo()
    m_Sci.SciMsg SCI_UNDO
End Sub

Public Sub Redo()
    m_Sci.SciMsg SCI_REDO
End Sub

Public Sub EmptyUndoBuffer()
    m_Sci.SciMsg SCI_EMPTYUNDOBUFFER
End Sub

Public Sub BeginUndoAction()
    m_Sci.SciMsg SCI_BEGINUNDOACTION
End Sub

Public Sub EndUndoAction()
    m_Sci.SciMsg SCI_ENDUNDOACTION
End Sub

'=========================================================================
' Save Point
'=========================================================================
Public Sub SetSavePoint()
    m_Sci.SciMsg SCI_SETSAVEPOINT
End Sub

Public Property Get IsModified() As Boolean
    IsModified = (m_Sci.SciMsg(SCI_GETMODIFY) <> 0)
End Property

'=========================================================================
' Buffer Management
'=========================================================================
Public Sub Allocate(ByVal bytes As Long)
    m_Sci.SciMsg SCI_ALLOCATE, bytes
End Sub

Public Function GetCharacterPointer() As Long
    GetCharacterPointer = m_Sci.SciMsg(SCI_GETCHARACTERPOINTER)
End Function

Public Function LoadFile(fileName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    Dim fileContent As String
    
    fileNum = FreeFile
    Open fileName For Binary As #fileNum
    fileContent = Space$(LOF(fileNum))
    Get #fileNum, , fileContent
    Close #fileNum
    
    text = fileContent
    SetSavePoint
    
    LoadFile = True
    Exit Function
    
ErrorHandler:
    LoadFile = False
End Function

Public Function SaveFile(fileName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    
    fileNum = FreeFile
    Open fileName For Binary As #fileNum
    Put #fileNum, , Me.text
    Close #fileNum
    
    SetSavePoint
    
    SaveFile = True
    Exit Function
    
ErrorHandler:
    SaveFile = False
End Function


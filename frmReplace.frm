VERSION 5.00
Begin VB.Form frmReplace 
   Caption         =   "Find/Replace"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11160
   LinkTopic       =   "Form3"
   ScaleHeight     =   2640
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   5355
      TabIndex        =   17
      Top             =   0
      Width           =   5775
   End
   Begin VB.CommandButton cmdFindAll 
      Caption         =   "Find All"
      Height          =   375
      Left            =   3960
      TabIndex        =   16
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox chkUnescape 
      Caption         =   "Use %xx for hex character values"
      Height          =   240
      Left            =   1035
      TabIndex        =   15
      Top             =   945
      Width           =   2685
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "Find Next"
      Height          =   375
      Left            =   3960
      TabIndex        =   14
      Top             =   1350
      Width           =   1335
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find First"
      Height          =   375
      Left            =   3960
      TabIndex        =   13
      Top             =   900
      Width           =   1335
   End
   Begin VB.CheckBox chkCaseSensitive 
      Caption         =   "Case Sensitive"
      Height          =   255
      Left            =   2040
      TabIndex        =   11
      Top             =   2160
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Selection"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   1440
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Whole Text"
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   1800
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   600
      TabIndex        =   7
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   600
      TabIndex        =   6
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   2250
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label lblSelSize 
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "Hex"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Char"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Replace"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Find"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyAll 
         Caption         =   "Copy All"
      End
   End
End
Attribute VB_Name = "frmReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author:  David Zimmer <dzzie@yahoo.com>
'AI:      Claude.ai
'Site:    http://sandsprite.com
'License: MIT
'---------------------------------------------------

Option Explicit

Public sci As SciWrapper
Dim lastkey As Integer
Dim lastIndex As Long
Dim lastsearch As String
Dim Init As Long

Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_SHOWWINDOW = &H40

'=========================================================================
' Public Interface
'=========================================================================
Public Sub LaunchReplaceForm(txtObj As SciWrapper)
    On Error Resume Next
    Set sci = txtObj
    
    If Len(txtObj.sel.GetSelectedText()) > 1 Then
        lblSelSize = "Selection Size: " & Len(txtObj.sel.GetSelectedText())
        Text1 = txtObj.sel.GetSelectedText()
    End If
    
    cmdFindAll.Visible = True
    Me.show
    If Not Init Then Form_Load
    Form_Resize
End Sub

Friend Sub SetFindText(X As String)
    Text1 = X
    
    ' Scan text2 (replace text) for numeric extension increment if found
    On Error Resume Next
    Dim ext, a, c, i
    
    ext = ""
    X = Trim(Text2)
    a = Len(X)
    
    If a = 0 Then Exit Sub
    
    ' Must be at least one numeric char at end
    If Not IsNumeric(Mid(X, a)) Then Exit Sub
    
    For i = 0 To 10
        c = Mid(X, a - i)
        If Not IsNumeric(c) Then
            Exit For
        Else
            ext = c
        End If
    Next
    
    If Len(ext) > 0 Then
        ext = CLng(ext)
        Text2 = Mid(X, 1, Len(X) - Len(ext)) & ext + 1
    End If
End Sub

'=========================================================================
' Find First
'=========================================================================
Private Sub cmdFind_Click()
    On Error Resume Next
    
    Dim f As String
    Dim matchCase As Boolean
    Dim pos As Long
    Dim line As Long
    
    ' Get search text
    If chkUnescape.value = 1 Then
        f = unescape(Text1)
    Else
        f = Text1
    End If
    
    lastsearch = f
    matchCase = (chkCaseSensitive.value = 1)
    
    ' Use new Search API
    pos = sci.Search.Find(f, 0, matchCase, False, False, False)
    
    If pos >= 0 Then
        lastIndex = pos + Len(f)
        sci.sel.SetSelection pos, pos + Len(f)
        
        line = sci.lines.LineFromPosition(pos)
        sci.sel.GotoLine line
        sci.lines.ScrollCaret
        
        Me.Caption = "Line: " & (line + 1) & " CharPos: " & pos
    Else
        lastIndex = 0
        MsgBox "Not found", vbInformation
    End If
End Sub

'=========================================================================
' Find Next
'=========================================================================
Private Sub cmdFindNext_Click()
    On Error Resume Next
    
    Dim f As String
    Dim matchCase As Boolean
    Dim pos As Long
    Dim line As Long
    
    ' Get search text
    If chkUnescape.value = 1 Then
        f = unescape(Text1)
    Else
        f = Text1
    End If
    
    ' If search text changed, do Find First
    If lastsearch <> f Then
        cmdFind_Click
        Exit Sub
    End If
    
    ' Check if we're at the end
    If lastIndex >= sci.doc.TextLength Then
        MsgBox "Reached end of text, no more matches", vbInformation
        Exit Sub
    End If
    
    matchCase = (chkCaseSensitive.value = 1)
    
    ' Find next occurrence
    pos = sci.Search.Find(f, lastIndex, matchCase, False, False, False)
    
    If pos >= 0 Then
        lastIndex = pos + Len(f)
        sci.sel.SetSelection pos, pos + Len(f)
        
        line = sci.lines.LineFromPosition(pos)
        sci.sel.GotoLine line
        sci.lines.ScrollCaret
        
        Me.Caption = "Line: " & (line + 1) & " CharPos: " & pos
    Else
        MsgBox "No more matches found", vbInformation
    End If
End Sub

'=========================================================================
' Find All
'=========================================================================
Public Sub cmdFindAll_Click()
    On Error Resume Next
    
    Dim f As String
    Dim matchCase As Boolean
    Dim pos As Long
    Dim line As Long
    Dim lineText As String
    Dim Count As Long
    Dim FirstVisibleLine As Long
    
    If Me.Width < 10440 Then Me.Width = 10440
    List1.Clear
    
    ' Get search text
    If chkUnescape.value = 1 Then
        f = unescape(Text1)
    Else
        f = Text1
    End If
    
    matchCase = (chkCaseSensitive.value = 1)
    FirstVisibleLine = sci.lines.FirstVisibleLine
    
    ' Find all occurrences
    pos = 0
    Do
        pos = sci.Search.Find(f, pos, matchCase, False, False, False)
        
        If pos >= 0 Then
            line = sci.lines.LineFromPosition(pos)
            lineText = sci.lines.GetLine(line)
            
            ' Add to list: "LineNum: LineText"
            List1.AddItem (line + 1) & ": " & Trim$(lineText)
            Count = Count + 1
            
            pos = pos + Len(f)
        End If
    Loop While pos >= 0 And pos < sci.doc.TextLength
    
    ' Restore scroll position
    sci.lines.FirstVisibleLine = FirstVisibleLine
    
    Me.Caption = "Find All: " & Count & " matches found"
End Sub

'=========================================================================
' Replace
'=========================================================================
Private Sub cmdReplace_Click()
    On Error Resume Next
    
    Dim f As String, r As String
    Dim matchCase As Boolean
    Dim Count As Long
    Dim curLine As Long
    
    ' Get search and replace text
    If chkUnescape.value = 1 Then
        f = unescape(Text1)
        r = unescape(Text2)
    Else
        f = Text1
        r = Text2
    End If
    
    matchCase = (chkCaseSensitive.value = 1)
    
    If Option1.value Then  ' Whole text
        curLine = sci.lines.GetFirstVisibleLine()
        Count = sci.Search.ReplaceAll(f, r, matchCase, False, False, False)
        sci.lines.SetFirstVisibleLine curLine
        MsgBox Count & " replacements made", vbInformation
    Else  ' Selection only
        Dim selText As String
        Dim newText As String
        Dim anchor As Long, caret As Long
        
        selText = sci.sel.GetSelectedText()
        anchor = sci.sel.anchor
        
        ' Do simple string replace on selection
        If matchCase Then
            newText = Replace(selText, f, r, , , vbBinaryCompare)
        Else
            newText = Replace(selText, f, r, , , vbTextCompare)
        End If
        
        sci.edit.ReplaceSelection newText
        sci.sel.SetSelection anchor, anchor + Len(newText)
        
        lblSelSize = "Selection Size: " & Len(newText)
    End If
End Sub

'=========================================================================
' List Navigation
'=========================================================================
Private Sub List1_Click()
    On Error Resume Next
    
    Dim tmp As String
    Dim line As Long
    Dim index As Long
    
    index = ListSelIndex(List1)
    
    If index >= 0 Then
        tmp = List1.List(index)
        If InStr(1, tmp, ":") > 0 Then
            line = CLng(Split(tmp, ":")(0)) - 1  ' Convert to 0-based
            sci.sel.GotoLine line
            sci.lines.ScrollCaret
        End If
    End If
End Sub

Private Function ListSelIndex(lst As ListBox) As Long
    On Error GoTo hell
    Dim i As Long
    
    For i = 0 To List1.ListCount - 1
        If List1.selected(i) Then
            ListSelIndex = i
            Exit Function
        End If
    Next

hell:
    ListSelIndex = -1
End Function

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub mnuCopyAll_Click()
    On Error Resume Next
    Dim X As String, i As Long
    For i = 0 To List1.ListCount - 1
        X = X & List1.List(i) & vbCrLf
    Next
    Clipboard.Clear
    Clipboard.SetText X
    MsgBox Len(X) & " bytes copied", vbInformation
End Sub

'=========================================================================
' Form Events
'=========================================================================
Private Sub Form_Load()
    Init = True
    ' FormPos Me, True  ' Uncomment if you have FormPos function
    SetWindowPos Me.hwnd, HWND_TOPMOST, Me.Left / 15, Me.Top / 15, Me.Width / 15, Me.Height / 15, SWP_SHOWWINDOW
    
    ' Load last settings (if you have registry functions)
    ' If Len(Text1) = 0 Then Text1 = GetMySetting("lastFind")
    ' Text2 = GetMySetting("lastReplace")
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    List1.Width = Me.Width - List1.Left - 200
    List1.Height = Me.Height - List1.Top - 300
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' FormPos Me, True, True  ' Uncomment if you have FormPos
    ' SaveMySetting "lastFind", Text1
    ' SaveMySetting "lastReplace", Text2
End Sub

'=========================================================================
' Hex/Char Conversion Helpers
'=========================================================================
Private Sub Text3_KeyPress(KeyAscii As Integer)
    lastkey = KeyAscii
End Sub

Private Sub Text3_KeyUp(KeyAscii As Integer, Shift As Integer)
    Dim X As String
    X = Hex(lastkey)
    If Len(X) = 1 Then X = "0" & X
    Text4 = X
    Text3 = Chr(lastkey)
End Sub

'=========================================================================
' URL Unescape Functions (%xx and %uxxxx)
'=========================================================================
Public Function isHexChar(hexValue As String, Optional b As Byte) As Boolean
    On Error Resume Next
    Dim v As Long
    
    If Len(hexValue) = 0 Then GoTo nope
    If Len(hexValue) > 2 Then GoTo nope
    
    v = CLng("&h" & hexValue)
    If Err.Number <> 0 Then GoTo nope
    
    b = CByte(v)
    If Err.Number <> 0 Then GoTo nope

    isHexChar = True
    Exit Function
    
nope:
    Err.Clear
    isHexChar = False
End Function

Private Function hex_bpush(bAry() As Byte, hexValue As String) As Boolean
    On Error Resume Next
    Dim b As Byte
    If Not isHexChar(hexValue, b) Then Exit Function
    bpush bAry, b
    hex_bpush = True
End Function

Function unescape(X) As String  ' %uxxxx and %xx
    On Error GoTo hell
    
    Dim tmp() As String
    Dim b1 As String, b2 As String
    Dim i As Long
    Dim r() As Byte
    Dim t
    
    tmp = Split(X, "%")
    
    s_bpush r(), tmp(0)  ' Any prefix before encoded part
    
    For i = 1 To UBound(tmp)
        t = tmp(i)
        
        If LCase(VBA.Left(t, 1)) = "u" Then
            If Len(t) < 5 Then  ' %u21 -> %u0021
                t = "u" & String(5 - Len(t), "0") & Mid(t, 2)
            End If

            b1 = Mid(t, 2, 2)
            b2 = Mid(t, 4, 2)
            
            If isHexChar(b1) And isHexChar(b2) Then
                hex_bpush r(), b2
                hex_bpush r(), b1
            Else
                s_bpush r(), "%u" & b1 & b2
            End If
            
            If Len(t) > 5 Then s_bpush r(), Mid(t, 6)
        Else
            b1 = Mid(t, 1, 2)
            If Not hex_bpush(r(), b1) Then s_bpush r(), "%" & b1
            If Len(t) > 2 Then s_bpush r(), Mid(t, 3)
        End If
    Next
            
hell:
    unescape = StrConv(r(), vbUnicode, LANG_US)
     
    If Err.Number <> 0 Then
        MsgBox "Error in unescape: " & Err.Description
    End If
End Function

Private Sub s_bpush(bAry() As Byte, sValue As String)
    Dim tmp() As Byte
    Dim i As Long
    tmp() = StrConv(sValue, vbFromUnicode, LANG_US)
    For i = 0 To UBound(tmp)
        bpush bAry, tmp(i)
    Next
End Sub

Private Sub bpush(bAry() As Byte, b As Byte)
    On Error GoTo Init
    Dim X As Long
    
    X = UBound(bAry)
    ReDim Preserve bAry(UBound(bAry) + 1)
    bAry(UBound(bAry)) = b
    Exit Sub

Init:
    ReDim bAry(0)
    bAry(0) = b
End Sub

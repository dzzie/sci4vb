VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmObjBrowser 
   Caption         =   "Object Browser"
   ClientHeight    =   5370
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   9840
   Icon            =   "frmObjBrowser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin sci4vb.ucFilterList lvMembers 
      Height          =   3735
      Left            =   3180
      TabIndex        =   2
      Top             =   60
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6588
   End
   Begin VB.TextBox txtDescription 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3195
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3915
      Width           =   6435
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   180
      Top             =   420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObjBrowser.frx":1782
            Key             =   "prop"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObjBrowser.frx":1BD4
            Key             =   "enum"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObjBrowser.frx":3366
            Key             =   "func"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmObjBrowser.frx":4AF8
            Key             =   "class"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvClasses 
      Height          =   5235
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   9234
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Objects"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Begin VB.Menu mnuOpenFile 
         Caption         =   "Open File"
      End
   End
End
Attribute VB_Name = "frmObjBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public isense As Collection 'of CIntellisenseItem
Dim selClass As CIntellisenseItem

Sub Init(isen As Collection)

    Dim ii As CIntellisenseItem
    Dim li As ListItem
    
    lvClasses.ListItems.Clear
    lvMembers.ListItems.Clear
    
    Set isense = isen
    
    For Each ii In isense
        Set li = lvClasses.ListItems.Add(, , " " & ii.objName, , "class")
        Set li.Tag = ii
    Next
    
    txtDescription = "Right click on class list for menu"
    lvClasses.ColumnHeaders(1).Width = lvClasses.Width
    FormPos Me, True
    Me.visible = True
    
End Sub

Private Sub Form_Load()
    mnuPopup.visible = False
    lvMembers.SetColumnHeaders "Members"
    Set lvMembers.mainLV.SmallIcons = ImageList1
    Set lvMembers.filtLV.SmallIcons = ImageList1
    lvMembers.SetFont "Courier", 12
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lvClasses.Height = Me.Height - lvClasses.Top - 600
    txtDescription.Top = Me.Height - txtDescription.Height - 600
    lvMembers.Height = txtDescription.Top - lvMembers.Top - 200
    lvMembers.Width = Me.Width - lvClasses.Left - lvClasses.Width - 350
    txtDescription.Width = lvMembers.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FormPos Me, True, True
End Sub

Private Sub lvClasses_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvColumnSort lvClasses, ColumnHeader
End Sub

Private Sub lvClasses_ItemClick(ByVal Item As MSComctlLib.ListItem)
       
       On Error Resume Next
        
       Dim ii As CIntellisenseItem, X, li As ListItem, tmp() As String, i As Long
       Dim ico As String, methodArr() As String, proto As String
       
       lvMembers.ListItems.Clear
       Set ii = Item.Tag
       Set selClass = ii
    
       methodArr = ii.GetMethodArray()
       txtDescription.text = selClass.Description
       
       For i = LBound(methodArr) To UBound(methodArr)
            ico = IIf(ii.isProp(methodArr(i)), "prop", "func")
            proto = ii.GetCallTip(methodArr(i))
            If Len(proto) = 0 Then proto = methodArr(i)
            Set li = lvMembers.ListItems.Add(, , " " & proto, , ico)
       Next
       
       If Err.Number <> 0 Then
            txtDescription.text = txtDescription.text & vbCrLf & " (Err: " & Err.Description & ")"
       End If
       
End Sub

Private Sub lvClasses_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub lvMembers_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvColumnSort lvMembers, ColumnHeader
End Sub

Private Sub lvMembers_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    txtDescription = Empty
    txtDescription = selClass.GetRawLine(Item.text) & vbCrLf & vbCrLf
    'txtDescription = txtDescription & selClass.comments(Item.Index) 'index wont work on filter list!
    txtDescription = txtDescription & selClass.GetComment(Item.text)
End Sub
 
'Private Sub mnuOpenFile_Click()
'    On Error Resume Next
'    Dim li As ListItem, ii As CIntellisenseItem
'    Set li = lvClasses.SelectedItem
'    Set ii = li.Tag
'    If Len(ii.path) = 0 Then
'        MsgBox "No file was set for this item", vbInformation
'        Exit Sub
'    End If
'    If fso.FileExists(ii.path) Then
'        Shell "notepad.exe " & fso.GetShortName(ii.path), vbNormalFocus
'    End If
'    If Err.Number <> 0 Then MsgBox Err.Description
'End Sub

Public Sub lvColumnSort(ListViewControl As Object, Column As Object)
    On Error Resume Next
    Const lvwAscending As Long = 0
    Const lvwDescending As Long = 1
     
    With ListViewControl
       If .SortKey <> Column.index - 1 Then
             .SortKey = Column.index - 1
             .SortOrder = lvwAscending
       Else
             If .SortOrder = lvwAscending Then
              .SortOrder = lvwDescending
             Else
              .SortOrder = lvwAscending
             End If
       End If
       .Sorted = -1
    End With
End Sub

Sub FormPos(fform As Object, Optional andSize As Boolean = False, Optional save_mode As Boolean = False)
    
    On Error Resume Next
    
    Dim f, sz, i, ff, def
    f = Split(",Left,Top,Height,Width", ",")
    
    If fform.WindowState = vbMinimized Then Exit Sub
    If andSize = False Then sz = 2 Else sz = 4
    
    For i = 1 To sz
        If save_mode Then
            ff = CallByName(fform, f(i), VbGet)
            SaveSetting App.EXEName, fform.name & ".FormPos", f(i), ff
        Else
            def = CallByName(fform, f(i), VbGet)
            ff = GetSetting(App.EXEName, fform.name & ".FormPos", f(i), def)
            CallByName fform, f(i), VbLet, ff
        End If
    Next
    
End Sub


VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About sci4vb"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDesc 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3060
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   7965
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   6960
      TabIndex        =   0
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lblURL 
      Caption         =   "https://scintilla.org/"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   60
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   3
      Top             =   3420
      Width           =   4335
   End
   Begin VB.Label lblURL 
      Caption         =   "https://sandsprite.com"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   0
      Left            =   60
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   2
      Top             =   3120
      Width           =   4335
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, ByVal Source As Long, ByVal length As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long


Private Sub cmdOK_Click()
      Unload Me
End Sub

Public Function LaunchForm(owner As SciWrapper)

     txtDesc = CompileVersionInfo(owner) & vbCrLf & vbCrLf & _
            "sci4vb is an easy to use ActiveX control that wraps Scintilla." & vbCrLf & _
            vbCrLf & _
            "Scintilla is an excellent open source component which supports " & vbCrLf & _
            "syntax highlighting, folding, auto complete, code tips, and more." & vbCrLf & _
            vbCrLf & _
            "sci4vb was created by David Zimmer and claude.ai"
        
     Me.visible = True
     
End Function
 

Private Sub lblURL_Click(index As Integer)
        ShellExec lblURL(index).Caption
End Sub


Public Function CompileVersionInfo(owner As SciWrapper) As String
    On Error Resume Next
    Dim dllVer As String
    Dim dllPath As String
    Dim ret() As String
    Dim hIndex As Long
    Dim hlNames As String
    Dim i As Long
    
    push ret, pad("sci4vb:", 13) & App.Major & "." & App.Minor & "." & App.Revision & "  (" & FileSize(App.path & "\sci4vb.ocx") & ")"
    
    dllPath = GetLoadedSciLexerPath()
    If FileExists(dllPath) Then
        dllVer = GetFileVersion(dllPath)
        If Len(dllVer) > 0 Then push ret, pad("Lexer:", 13) & dllVer & "    (" & FileSize(dllPath) & ")"
        push ret, "sci4vb Path: " & App.path
        push ret, "Lexer Path:  " & dllPath
        
    Else
        push ret, "sci4vb Path: " & App.path
        push ret, "SciLexer.dll: NOT FOUND!"
    End If
    
    CompileVersionInfo = Join(ret, vbCrLf)
    
End Function

Public Function GetLoadedSciLexerPath() As String
     Dim h As Long, ret As String
     ret = Space(500)
     h = GetModuleHandle("SciLexer.dll")
     h = GetModuleFileName(h, ret, 500)
     If h > 0 Then ret = Mid(ret, 1, h)
     GetLoadedSciLexerPath = ret
End Function

Public Function GetFileVersion(Optional ByVal PathWithFilename As String) As String
    ' return file-properties of given file  (EXE , DLL , OCX)
    'http://support.microsoft.com/default.aspx?scid=kb;en-us;160042
    
    If Len(PathWithFilename) = 0 Then Exit Function
    
    Dim lngBufferlen As Long
    Dim lngDummy As Long
    Dim lngRc As Long
    Dim lngVerPointer As Long
    Dim lngHexNumber As Long
    Dim b() As Byte
    Dim b2() As Byte
    Dim strBuffer As String
    Dim strLangCharset As String
    Dim strTemp As String
    Dim n As Long
    
    ReDim b2(500)
    
    lngBufferlen = GetFileVersionInfoSize(PathWithFilename, lngDummy)
    If lngBufferlen <= 0 Then Exit Function
    
    ReDim b(lngBufferlen)
    lngRc = GetFileVersionInfo(PathWithFilename, 0&, lngBufferlen, b(0))
    If lngRc = 0 Then Exit Function
    
    lngRc = VerQueryValue(b(0), "\VarFileInfo\Translation", lngVerPointer, lngBufferlen)
    If lngRc = 0 Then Exit Function
    
    MoveMemory b2(0), lngVerPointer, lngBufferlen
    lngHexNumber = b2(2) + b2(3) * &H100 + b2(0) * &H10000 + b2(1) * &H1000000
    strLangCharset = Right("0000000" & Hex(lngHexNumber), 8)
    
    strBuffer = String$(800, 0)
    strTemp = "\StringFileInfo\" & strLangCharset & "\FileVersion"
    lngRc = VerQueryValue(b(0), strTemp, lngVerPointer, lngBufferlen)
    If lngRc = 0 Then Exit Function
    
    lstrcpy strBuffer, lngVerPointer
    n = InStr(strBuffer, Chr(0)) - 1
    If n > 0 Then
        strBuffer = Mid$(strBuffer, 1, n)
        GetFileVersion = strBuffer
    End If
   
End Function

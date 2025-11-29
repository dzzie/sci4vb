VERSION 5.00
Object = "{0925C46F-375B-45F9-A43F-379F6C87DDE5}#9.0#0"; "sci4vb.ocx"
Begin VB.Form frmSciDemo 
   Caption         =   "Scintilla Wrapper Demo"
   ClientHeight    =   7530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12600
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   12600
   StartUpPosition =   2  'CenterScreen
   Begin sci4vb.SciWrapper sci 
      Height          =   5835
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   10292
   End
   Begin VB.CommandButton cmdObjBrowser 
      Caption         =   "ObjBrowser"
      Height          =   315
      Left            =   9060
      TabIndex        =   0
      Top             =   7140
      Width           =   1155
   End
   Begin VB.CommandButton cmdRndExecLine 
      Caption         =   "Set ExecLine"
      Height          =   315
      Left            =   10380
      TabIndex        =   9
      Top             =   7140
      Width           =   1275
   End
   Begin VB.CheckBox chkViewWhite 
      Caption         =   "View Whitespace"
      Height          =   255
      Left            =   6660
      TabIndex        =   8
      Top             =   7200
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox chkIndentGuides 
      Caption         =   "Indent Guides"
      Height          =   255
      Left            =   5040
      TabIndex        =   7
      Top             =   7200
      Value           =   1  'Checked
      Width           =   1395
   End
   Begin VB.CheckBox chkBreakPoints 
      Caption         =   "BreakPoints"
      Height          =   255
      Left            =   3540
      TabIndex        =   6
      Top             =   7200
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CommandButton cmdZoomOut 
      Caption         =   "-"
      Height          =   315
      Left            =   12180
      TabIndex        =   5
      Top             =   7140
      Width           =   315
   End
   Begin VB.CommandButton cmdZoomIn 
      Caption         =   "+"
      Height          =   315
      Left            =   11820
      TabIndex        =   4
      Top             =   7140
      Width           =   315
   End
   Begin VB.CheckBox chkLineNumbers 
      Caption         =   "Line Numbers"
      Height          =   255
      Left            =   1980
      TabIndex        =   3
      Top             =   7200
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox chkFolding 
      Caption         =   "Code Folding"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   7200
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.TextBox txtStatus 
      Height          =   1035
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   6060
      Width           =   12135
   End
End
Attribute VB_Name = "frmSciDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 'Author:  David Zimmer <dzzie@yahoo.com> + Claude.ai
'Site:    http://sandsprite.com
'License: MIT
'---------------------------------------------------

'=========================================================================
' frmSciDemo - Demonstration of SciWrapper capabilities
'=========================================================================
Option Explicit

Private Sub chkBreakPoints_Click()
    sci.doc.BreakPoints = (chkBreakPoints.Value = 1)
End Sub

Private Sub chkIndentGuides_Click()
    sci.View.IndentationGuides = (chkIndentGuides.Value = 1)
End Sub

Private Sub chkViewWhite_Click()
    sci.View.ViewWhitespace = IIf(chkViewWhite.Value = 1, wsVisibleAlways, wsInvisible)
End Sub

Private Sub cmdObjBrowser_Click()
    sci.ShowObjectBrowser
End Sub

Private Sub cmdRndExecLine_Click()
    sci.doc.SetExecLine RandBetween1AndX(10)
End Sub

Function RandBetween1AndX(ByVal X As Integer) As Integer
    Randomize
    RandBetween1AndX = Int(Rnd * X) + 1
End Function

Private Sub Form_Load()

    InitializeIntellisense
    
    With sci
        .doc.MouseDwellTime = 600 'disabled by default
        .doc.Folding = True
        .View.ViewWhitespace = wsVisibleAlways
        .View.IndentationGuides = True
        .doc.text = GetSampleCode()
        .Style.SetLanguage langJavaScript
        .doc.SetExecLine 2
        .AddBreakpoint 5
    End With
    
End Sub

Private Function GetSampleCode() As String
    Dim code As String
    
    code = "// IDA JScript API Demo" & vbCrLf
    code = code & "// Try typing 'ida.' to see all API methods!" & vbCrLf
    code = code & "// Hover over function names for tooltips" & vbCrLf
    code = code & "// Type '(' after function name to see arguments!" & vbCrLf & vbCrLf
    code = code & "function analyzeFunction() {" & vbCrLf
    code = code & "    // Type 'ida.' below to see autocomplete!" & vbCrLf
    code = code & "    var count = ida.funcCount();" & vbCrLf
    code = code & "    ida.message('Found ' + count + ' functions');" & vbCrLf & vbCrLf
    code = code & "    for (var i = 0; i < count; i++) {" & vbCrLf
    code = code & "        var start = ida.functionStart(i);" & vbCrLf
    code = code & "        var end = ida.functionEnd(i);" & vbCrLf
    code = code & "        var name = ida.functionName(i);" & vbCrLf & vbCrLf
    code = code & "        ida.addAddr(start, name, end);" & vbCrLf
    code = code & "    }" & vbCrLf & vbCrLf
    code = code & "    ida.showList();" & vbCrLf
    code = code & "}" & vbCrLf & vbCrLf
    code = code & "// Type 'ida.' to see 100+ API methods!" & vbCrLf
    code = code & "// Try: ida.read, ida.func, ida.get, ida.add" & String(12, vbCrLf)

    GetSampleCode = code

End Function

Private Sub InitializeIntellisense()

    'a basic flat implementation of intellisense has been built into the main ctl you dont have to use it
    'it does not support nesting, and all objects draw from the same LoadCallTips list. (dont duplicate names across objects)
    'if you need more complex dont use this, and implement it raw in the host form layer. I keep most of my api flat where i can.
    
    'loading from file is easier and contains more metadata like comments for objbrowser
    'and tooltips extracted from the prototypes themselves
    sci.Intellisense.ParseFromFile App.Path & "\dbg.js" 'this loads method names and per object call tips from the prototypes


    'or we can do it manually in a simplier style - you will load calltips manually too see below
    sci.Intellisense.Add "ida", _
        "isUp is32Bit message makeStr makeUnk loadedFile patchString patchByte getAsm instSize " & _
        "xRefsTo xRefsFrom getName functionName hideBlock showBlock setName addComment getComment addCodeXRef addDataXRef " & _
        "delCodeXRef delDataXRef funcVAByName renameFunc find decompile jump jumpRVA refresh undefine showEA hideEA " & _
        "removeName makeCode funcIndexFromVA nextEA prevEA funcCount numFuncs functionStart functionEnd readByte " & _
        "originalByte imageBase screenEA quickCall clearDecompilerCache isCode isData readLong readShort readQWord " & _
        "dumpFunc dumpFuncBytes getopv add_enum get_enum add_enum_member importFile addSect sectExists delSect getSects " & _
        "getFunc readStr OffsetToVa VaToOffset getRes version bits CloseIDA strings xRefsToCnt " & _
        "readFile writeFile appendFile fileExists deleteFile writeBin getBin md5File md5 rc4 expandPath " & _
        "toHex hexstr toBytes hexDump add subtract intToHex alert confirm askValue openFileDialog saveFileDialog " & _
        "copyAll clear addAddr setColHeaders showList hideList ScanProcess ResolveExport loadToJS benchMark" _
        , _
        "this is the description for the ida class in objbrowser"
    
    sci.Intellisense.Add "ida2", "test"
    
    '
    sci.Intellisense.Add "Math", _
        "abs acos asin atan atan2 ceil cos exp floor log max min pow random round sin sqrt tan E LN10 LN2 LOG10E LOG2E PI SQRT1_2 SQRT2"
    
    sci.Intellisense.Add "String", _
        "charAt charCodeAt concat indexOf lastIndexOf match replace search slice split substr substring toLowerCase toUpperCase trim length"
    
    sci.Intellisense.Add "Array", _
        "concat join pop push reverse shift slice sort splice unshift forEach map filter reduce every some indexOf lastIndexOf length"
    
    On Error Resume Next
    'for simplicity we can also just include all the calltips in a flat file
    sci.Intellisense.LoadCallTips App.Path & "\api.txt"
    
End Sub

Private Sub chkLineNumbers_Click()
    sci.doc.LineNumbers = (chkLineNumbers.Value = 1)
End Sub

Private Sub chkFolding_Click()
    sci.doc.Folding = (chkFolding.Value = 1)
End Sub

Private Sub cmdZoomIn_Click()
    sci.View.ZoomIn
    AddStatus "Zoom: " & sci.View.Zoom
End Sub

Private Sub cmdZoomOut_Click()
    sci.View.ZoomOut
    AddStatus "Zoom: " & sci.View.Zoom
End Sub

Private Sub sci_AutoCompleteEvent(className As String)
     
    'if you loaded the internal intellisense list there is probably  nothing to do here.
    'you might need to check the followign to make sure you dont wack internal displays..
    'If Not sci.Autoc.isActive Then
    
End Sub

Private Sub sci_UpdateUI(updated As Long)
    ' Update status when caret moves
    Dim line As Long
    Dim col As Long
    Dim pos As Long
    
    pos = sci.sel.currentPos
    line = sci.lines.LineFromPosition(pos)
    col = sci.lines.GetColumn(pos)
    
    Me.Caption = "Scintilla Wrapper Demo - Line: " & (line + 1) & " Col: " & (col + 1) & " Pos: " & pos
End Sub

Private Sub sci_Modified(position As Long, modificationType As Long, text As String, length As Long, linesAdded As Long)
    ' Track modifications
    Static modCount As Long
    modCount = modCount + 1
End Sub

Private Sub sci_MarginClick(margin As Long, position As Long, modifiers As Long)
    Dim line As Long
    
    If margin = 2 Then  ' Folding margin
        line = sci.lines.LineFromPosition(position)
        sci.Fold.ToggleFold line
        AddStatus "Toggled fold at line " & (line + 1)
    End If
    
End Sub

Private Sub sci_DwellStart(position As Long, X As Long, Y As Long)
    ' Show tooltip when hovering over function names
    Dim word As String
    Dim tip As String
    
    word = sci.Helper.WordUnderMouse(position)
    
    If Len(word) > 0 Then
        tip = sci.Intellisense.GetCallTip(word)
        
        If Len(tip) > 0 Then
            sci.Autoc.ShowCallTip position, tip
        End If
    End If
    
End Sub

Private Sub sci_DwellEnd(position As Long, X As Long, Y As Long)
    ' Hide tooltip when mouse moves away
    If Not sci.Helper.IsMouseOverCallTip() Then
        sci.Autoc.CancelCallTip
    End If
End Sub

Private Sub AddStatus(msg As String)
    txtStatus.text = txtStatus.text & Format$(Now, "hh:nn:ss") & " - " & msg & vbCrLf
    txtStatus.SelStart = Len(txtStatus.text)
End Sub

 

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIntellisenseManager"
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
' CIntellisenseManager - Manages collection of intellisense objects
'=========================================================================
Option Explicit

Private m_Objects As Collection        ' CIntellisenseItem objects
Private m_CallTips As Collection       ' Function name -> call tip
Private m_CurrentArgIndex As Long      ' Current argument being typed

Property Get isLoaded() As Boolean
    isLoaded = (m_Objects.Count > 0)
End Property

'=========================================================================
' Initialize
'=========================================================================
Private Sub Class_Initialize()
    Set m_Objects = New Collection
    Set m_CallTips = New Collection
    m_CurrentArgIndex = 0
End Sub

Function GetObjectNames(Optional partial As String) As String
    Dim Item As New CIntellisenseItem
    Dim tmp As String, sz As Long
    For Each Item In m_Objects
        sz = Len(partial)
        If sz > 0 Then
            If LCase(Left(Item.objName, sz)) = LCase(partial) Then
                tmp = tmp & Item.objName & " "
            End If
        Else
            tmp = tmp & Item.objName & " "
        End If
    Next
    GetObjectNames = Trim(tmp)
End Function

'=========================================================================
' Add an intellisense object
'=========================================================================
Public Sub add(objName As String, methodList As String, Optional desc As String = "")
    Dim Item As New CIntellisenseItem
    
    Item.objName = objName
    Item.methods = methodList
    Item.Description = desc
    Item.ParseMethods
    
    ' Add or replace
    On Error Resume Next
    m_Objects.remove objName
    On Error GoTo 0
    
    m_Objects.add Item, objName
End Sub

Friend Property Get Objects() As Collection
    Set Objects = m_Objects
End Property

'=========================================================================
' Get object by name
'=========================================================================
Public Function GetObject(objName As String) As CIntellisenseItem
    On Error Resume Next
    Set GetObject = m_Objects(objName)
End Function

'=========================================================================
' Get methods for object
'=========================================================================
Public Function GetMethods(objName As String) As String
    Dim Item As CIntellisenseItem
    
    Set Item = GetObject(objName)
    
    If Not Item Is Nothing Then
        GetMethods = Item.methods
    End If
End Function

'=========================================================================
' Find methods for object with partial match
'=========================================================================
Public Function FindMethods(objName As String, partial As String) As String
    Dim Item As CIntellisenseItem
    
    Set Item = GetObject(objName)
    
    If Not Item Is Nothing Then
        FindMethods = Item.FindMethods(partial)
    End If
End Function

'=========================================================================
' Load call tips from file
'=========================================================================
Public Function LoadCallTips(fileName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    Dim line As String
    Dim funcName As String
    Dim parenPos As Long
    
    fileNum = FreeFile
    Open fileName For Input As #fileNum
    
    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        line = Trim$(line)
        
        ' Skip empty lines and comments
        If Len(line) > 0 And Left$(line, 1) <> "'" And Left$(line, 2) <> "//" Then
            ' Parse: functionName(params) description
            parenPos = InStr(line, "(")
            
            If parenPos > 0 Then
                funcName = Trim$(Left$(line, parenPos - 1))
                
                ' Add to call tips
                On Error Resume Next
                m_CallTips.remove funcName
                On Error GoTo ErrorHandler
                
                m_CallTips.add line, funcName
                'If funcName = "addAddr" Then Stop
            End If
        End If
    Loop
    
    Close #fileNum
    LoadCallTips = True
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    Close #fileNum
    LoadCallTips = False
End Function

'=========================================================================
' Get call tip for function (GLOBALS ONLY)
'=========================================================================
Public Function GetCallTip(functionName As String) As String
    On Error Resume Next
    GetCallTip = m_CallTips(functionName)
    
    If Err.Number <> 0 Then
        GetCallTip = ""
    End If
End Function

'=========================================================================
' Get call tip with current argument highlighted
'=========================================================================
'Public Function GetCallTipWithHighlight(functionName As String, argIndex As Long) As String
'    Dim tip As String
'    Dim parenPos As Long
'    Dim paramsStart As Long
'    Dim paramsEnd As Long
'    Dim params As String
'    Dim paramAry() As String
'    Dim i As Long
'    Dim result As String
'    Dim currentParam As String
'
'    tip = GetCallTip(functionName)
'    If Len(tip) = 0 Then
'        GetCallTipWithHighlight = ""
'        Exit Function
'    End If
'
'    ' Find parameter section
'    parenPos = InStr(tip, "(")
'    If parenPos = 0 Then
'        GetCallTipWithHighlight = tip
'        Exit Function
'    End If
'
'    paramsStart = parenPos + 1
'    paramsEnd = InStr(paramsStart, tip, ")")
'
'    If paramsEnd = 0 Then paramsEnd = Len(tip)
'
'    params = Mid$(tip, paramsStart, paramsEnd - paramsStart)
'
'    If Len(Trim$(params)) = 0 Then
'        GetCallTipWithHighlight = tip
'        Exit Function
'    End If
'
'    ' Split parameters
'    paramAry = Split(params, ",")
'
'    ' Build result with highlighted parameter
'    result = Left$(tip, parenPos)
'
'    For i = 0 To UBound(paramAry)
'        If i > 0 Then result = result & ","
'
'        currentParam = Trim$(paramAry(i))
'
'        If i = argIndex Then
'            ' Highlight current parameter (using \001 and \002 as markers)
'            result = result & Chr(1) & currentParam & Chr(2)
'        Else
'            result = result & currentParam
'        End If
'    Next i
'
'    result = result & Mid$(tip, paramsEnd)
'
'    GetCallTipWithHighlight = result
'End Function

'=========================================================================
' Calculate current argument index from position
'=========================================================================
Public Function GetCurrentArgIndex(text As String, parenPos As Long, currentPos As Long) As Long
    Dim i As Long
    Dim commaCount As Long
    Dim depth As Long
    Dim ch As String
    
    ' Count commas at depth 0 between paren and current position
    For i = parenPos + 1 To currentPos
        If i > Len(text) Then Exit For
        
        ch = Mid$(text, i, 1)
        
        Select Case ch
            Case "("
                depth = depth + 1
            Case ")"
                depth = depth - 1
                If depth < 0 Then Exit For
            Case ","
                If depth = 0 Then commaCount = commaCount + 1
        End Select
    Next i
    
    GetCurrentArgIndex = commaCount
End Function

'=========================================================================
' Find opening paren for function call
'=========================================================================
Public Function FindOpenParen(text As String, currentPos As Long) As Long
    Dim i As Long
    Dim depth As Long
    Dim ch As String
    
    ' Search backwards from current position
    For i = currentPos - 1 To 1 Step -1
        ch = Mid$(text, i, 1)
        
        Select Case ch
            Case ")"
                depth = depth + 1
            Case "("
                If depth = 0 Then
                    FindOpenParen = i
                    Exit Function
                Else
                    depth = depth - 1
                End If
            Case ";", vbCr, vbLf
                ' Stop at statement boundary
                Exit For
        End Select
    Next i
    
    FindOpenParen = -1
End Function

'=========================================================================
' Check if we're inside a function call
'=========================================================================
Public Function IsInFunctionCall(text As String, currentPos As Long) As Boolean
    IsInFunctionCall = (FindOpenParen(text, currentPos) >= 0)
End Function

'=========================================================================
' Get function name before opening paren
'=========================================================================
Public Function GetFunctionNameBeforeParen(text As String, parenPos As Long) As String
    Dim i As Long
    Dim ch As String
    Dim funcName As String
    
    ' Read backwards to get function name
    For i = parenPos - 1 To 1 Step -1
        ch = Mid$(text, i, 1)
        
        If (ch >= "a" And ch <= "z") Or _
           (ch >= "A" And ch <= "Z") Or _
           (ch >= "0" And ch <= "9") Or _
           ch = "_" Or ch = "$" Then
            funcName = ch & funcName
        Else
            Exit For
        End If
    Next i
    
    GetFunctionNameBeforeParen = funcName
End Function

'=========================================================================
' Get object name before dot
'=========================================================================
Public Function GetObjectBeforeDot(text As String, dotPos As Long) As String
    Dim i As Long
    Dim ch As String
    Dim objName As String
    
    ' Read backwards to get object name
    For i = dotPos - 1 To 1 Step -1
        ch = Mid$(text, i, 1)
        
        If (ch >= "a" And ch <= "z") Or _
           (ch >= "A" And ch <= "Z") Or _
           (ch >= "0" And ch <= "9") Or _
           ch = "_" Or ch = "$" Then
            objName = ch & objName
        Else
            Exit For
        End If
    Next i
    
    GetObjectBeforeDot = objName
End Function

'=========================================================================
' Clear all
'=========================================================================
Public Sub Clear()
    Set m_Objects = New Collection
    Set m_CallTips = New Collection
End Sub

'=========================================================================
' Get object count
'=========================================================================
Public Function GetObjectCount() As Long
    GetObjectCount = m_Objects.Count
End Function

'=========================================================================
' Get all object names
'=========================================================================
'Public Function GetObjectNames() As String
'    Dim item As CIntellisenseItem
'    Dim result As String
'
'    For Each item In m_Objects
'        result = result & item.objName & " "
'    Next
'
'    GetObjectNames = Trim$(result)
'End Function

'=========================================================================
' Parse intellisense from file
' Object name = file basename (without extension)
' Supports both simple "func(args)" and VB6 "Function func(args)" formats
'=========================================================================
Public Function ParseFromFile(filePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    Dim line As String
    Dim methodName As String
    Dim methodList As String
    Dim objName As String
    Dim parenPos As Long
    Dim calltip As String
    Dim Item As New CIntellisenseItem
    Dim intoDefs As Boolean
    Dim desc As String
    Dim comment As String
    Dim commentPos As Long


    ' Extract object name from filename
    objName = GetBaseName(filePath)
    If Len(objName) = 0 Then Exit Function
 
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    
    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        line = Trim$(Replace(line, vbTab, Empty))
        
        ' Skip empty lines and comments
        If Len(line) = 0 Then GoTo NextLine
        If Left$(line, 2) = "/*" Then GoTo NextLine
        If Left$(line, 2) = "\*" Then GoTo NextLine
        
        If Left$(line, 1) = "'" Then
            If Not intoDefs Then desc = desc & Mid(line, 2) & vbCrLf
            GoTo NextLine
        End If
        
        ' Strip VB6 keywords and get clean call tip
        calltip = StripVBKeywords(line)
        
        comment = Empty
        commentPos = InStr(calltip, "'")
        If commentPos > 0 Then
            comment = Trim$(Mid$(calltip, commentPos + 1))
            calltip = Trim$(Left$(calltip, commentPos - 1))
        End If

        ' Extract method name (before opening paren)
        parenPos = InStr(calltip, "(")
        If parenPos > 0 Then
            methodName = Trim$(Left$(calltip, parenPos - 1))
        Else
            methodName = calltip 'property no ()
        End If
        
        ' Skip if no valid method name
        If Len(methodName) > 0 And IsValidIdentifier(methodName) Then
        
            ' Add to method list (avoid duplicates from property get/let/set)
            If InStr(1, " " & methodList & " ", " " & methodName & " ", vbTextCompare) = 0 Then
                methodList = methodList & methodName & " "
            End If

            intoDefs = True
            Item.AddItem methodName, calltip, comment, line

            ' Also add to global tips for backward compat
            On Error Resume Next
            m_CallTips.remove methodName
            On Error GoTo ErrorHandler
            m_CallTips.add calltip, methodName
            
        End If
         
NextLine:
    Loop
    
    Close #fileNum
    
    ' Only add if we found methods
    If Len(methodList) > 0 Then
        Item.objName = objName
        Item.methods = Trim$(methodList)
        Item.path = filePath
        Item.Description = desc
        Item.ParseMethods
        
        ' Add or replace in collection
        On Error Resume Next
        m_Objects.remove objName
        On Error GoTo 0
        
        m_Objects.add Item, objName
        ParseFromFile = True
    End If
    
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    MsgBox "Error loading: " & filePath & " Desc: " & Err.Description
    Close #fileNum
    ParseFromFile = False
End Function

'=========================================================================
' Load all intellisense files from folder
'=========================================================================
Public Function ParseFromFolder(folderPath As String, Optional ext As String = "*.txt") As Long
    Dim fileName As String
    Dim Count As Long
    
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    fileName = Dir$(folderPath & ext)
    Do While Len(fileName) > 0
        If ParseFromFile(folderPath & fileName) Then
            Count = Count + 1
        End If
        fileName = Dir$
    Loop
    
    ParseFromFolder = Count
End Function

'=========================================================================
' Get call tip - checks object-specific first, then global
'=========================================================================
Public Function GetCallTipEx(functionName As String, Optional objName As String) As String
    Dim Item As CIntellisenseItem
    Dim tip As String
    
    ' Try object-specific first
    If Len(objName) > 0 Then
        Set Item = GetObject(objName)
        If Not Item Is Nothing Then
            tip = Item.GetCallTip(functionName)
            If Len(tip) > 0 Then
                GetCallTipEx = tip
                Exit Function
            End If
        End If
    End If
    
    ' Fall back to global
    On Error Resume Next
    GetCallTipEx = m_CallTips(functionName)
End Function

'=========================================================================
' Private helpers
'=========================================================================
Private Function GetBaseName(filePath As String) As String
    Dim fileName As String
    Dim dotPos As Long, slashPos As Long
    
    slashPos = InStrRev(filePath, "\")
    If slashPos = 0 Then slashPos = InStrRev(filePath, "/")
    
    If slashPos > 0 Then
        fileName = Mid$(filePath, slashPos + 1)
    Else
        fileName = filePath
    End If
    
    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then
        GetBaseName = Left$(fileName, dotPos - 1)
    Else
        GetBaseName = fileName
    End If
End Function

Private Function StripVBKeywords(line As String) As String
    Dim lower As String, a As Long
    Dim result As String
    
    result = line
    lower = LCase$(line)
    lower = Trim(Replace(lower, "public", "", , , vbTextCompare))
    
    If Left$(lower, 13) = "property get " Then
        result = Mid$(line, 14)
    ElseIf Left$(lower, 13) = "property let " Then
        result = Mid$(line, 14)
    ElseIf Left$(lower, 13) = "property set " Then
        result = Mid$(line, 14)
    ElseIf Left$(lower, 8) = "property" Then
        'im not going to get to crazy parsing this..
        result = Trim(Mid$(line, 9))
        If InStr(1, result, "get/let", vbTextCompare) > 0 Then result = Trim(Replace(result, "get/let", Empty, , 1, vbTextCompare))
        If InStr(1, result, "let/get", vbTextCompare) > 0 Then result = Trim(Replace(result, "let/get", Empty, , 1, vbTextCompare))
    ElseIf Left$(lower, 9) = "function " Then
        result = Mid$(line, 10)
    ElseIf Left$(lower, 4) = "sub " Then
        result = Mid$(line, 5)
    End If

    If InStr(1, line, "property", vbTextCompare) > 0 Then 'no () for properties
        a = InStr(result, "(")
        If a > 1 Then result = Mid(result, 1, a - 1)
    End If

    StripVBKeywords = Trim$(result)
End Function

Private Function IsValidIdentifier(s As String) As Boolean
    Dim i As Long, ch As String
    
    If Len(s) = 0 Then Exit Function
    
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If Not ((ch >= "a" And ch <= "z") Or _
                (ch >= "A" And ch <= "Z") Or _
                (ch >= "0" And ch <= "9") Or _
                ch = "_" Or ch = "$") Then
            Exit Function
        End If
    Next i
    
    IsValidIdentifier = True
End Function


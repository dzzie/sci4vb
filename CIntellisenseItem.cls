VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIntellisenseItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Author:  David Zimmer <dzzie@yahoo.com> + Claude.ai
'Site:    http://sandsprite.com
'License: MIT
'---------------------------------------------------



'=========================================================================
' CIntellisenseItem - Represents an object with autocomplete methods
'=========================================================================
Option Explicit

Public objName As String           ' Object name (e.g., "ida", "Math")
Public methods As String           ' Space-separated method list
Public Description As String       ' Object description
Public path As String              ' Optional file path

Private m_MethodArray() As String  ' Parsed method array
Private m_CallTips As Collection   ' Method name -> call tip signature
Private m_Comments As Collection   ' Method name -> comment
Private m_rawLine As Collection

'=========================================================================
' Initialize
'=========================================================================
Private Sub Class_Initialize()
    Set m_CallTips = New Collection
    Set m_Comments = New Collection
    Set m_rawLine = New Collection
End Sub

Public Function isProp(methodName As String) As Boolean
    On Error Resume Next
    Dim r As String
    r = m_rawLine(Trim(methodName))
    If InStr(1, r, "prop", vbTextCompare) > 0 Then isProp = True
End Function

Public Sub AddItem(methodName As String, ByVal calltip As String, ByVal comment As String, ByVal rawline As String)
    AddCallTip methodName, calltip
    AddComment methodName, comment
    AddRawLine methodName, rawline
End Sub

Public Sub AddRawLine(methodName As String, ByVal rawline As String)
    On Error Resume Next
    rawline = Trim(rawline)
    If Len(rawline) = 0 Then Exit Sub
    m_rawLine.Remove methodName
    m_rawLine.Add rawline, methodName
End Sub

Public Function GetRawLine(ByVal methodName As String, Optional withComment As Boolean = False) As String
    On Error Resume Next
    Dim a As Long
    methodName = Trim(methodName)
    a = InStr(methodName, "(")
    If a > 1 Then methodName = Mid(methodName, 1, a - 1)
    GetRawLine = m_rawLine(Trim(methodName))
    If Not withComment Then
        a = InStr(GetRawLine, "'")
        If a > 1 Then GetRawLine = Mid(GetRawLine, 1, a - 1)
    End If
End Function

Public Sub AddComment(methodName As String, ByVal comment As String)
    On Error Resume Next
    comment = Trim(comment)
    If Len(comment) = 0 Then Exit Sub
    m_Comments.Remove methodName
    m_Comments.Add comment, methodName
End Sub

Public Function GetComment(ByVal methodName As String) As String
    On Error Resume Next
    Dim a As Long
    methodName = Trim(methodName)
    a = InStr(methodName, "(")
    If a > 1 Then methodName = Mid(methodName, 1, a - 1)
    GetComment = m_Comments(Trim(methodName))
End Function

'=========================================================================
' Parse methods string into array
'=========================================================================
Public Sub ParseMethods()
    If Len(methods) > 0 Then
        m_MethodArray = Split(Trim$(methods), " ")
    Else
        ReDim m_MethodArray(-1 To -1)
    End If
End Sub

'=========================================================================
' Get method array
'=========================================================================
Public Function GetMethodArray() As String()
    If UBound(m_MethodArray) < 0 Then ParseMethods
    GetMethodArray = m_MethodArray
End Function

'=========================================================================
' Get method count
'=========================================================================
Public Function GetMethodCount() As Long
    If UBound(m_MethodArray) < 0 Then ParseMethods
    
    If UBound(m_MethodArray) >= LBound(m_MethodArray) Then
        GetMethodCount = UBound(m_MethodArray) - LBound(m_MethodArray) + 1
    Else
        GetMethodCount = 0
    End If
End Function

'=========================================================================
' Add call tip for a method
'=========================================================================
Public Sub AddCallTip(methodName As String, signature As String)
    On Error Resume Next
    m_CallTips.Add signature, methodName
End Sub

'=========================================================================
' Get call tip for a method
'=========================================================================
Public Function GetCallTip(methodName As String) As String
    On Error Resume Next
    GetCallTip = m_CallTips(methodName)
End Function

'=========================================================================
' Check if method exists
'=========================================================================
Public Function HasMethod(methodName As String) As Boolean
    Dim i As Long
    Dim arr() As String
    
    arr = GetMethodArray()
    
    For i = LBound(arr) To UBound(arr)
        If StrComp(arr(i), methodName, vbTextCompare) = 0 Then
            HasMethod = True
            Exit Function
        End If
    Next i
End Function

'=========================================================================
' Find methods by partial match
'=========================================================================
Public Function FindMethods(partial As String) As String
    Dim i As Long
    Dim arr() As String
    Dim result As String
    Dim lowerPartial As String
    
    arr = GetMethodArray()
    lowerPartial = LCase$(partial)
    
    For i = LBound(arr) To UBound(arr)
        If InStr(1, LCase$(arr(i)), lowerPartial) = 1 Then
            result = result & arr(i) & " "
        End If
    Next i
    
    FindMethods = Trim$(result)
End Function

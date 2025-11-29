Attribute VB_Name = "modMockAPI"
Option Explicit

'=========================================================================
' Mock API Objects for Autocomplete/Tooltip Testing
' Based on IDA JScript API
'=========================================================================

Public Type APIFunction
    name As String
    signature As String
    description As String
End Type

Private m_Functions() As APIFunction
Private m_FunctionCount As Long

'=========================================================================
' Initialize Mock API from api.txt
'=========================================================================
Public Sub InitializeMockAPI()
    ReDim m_Functions(100)
    m_FunctionCount = 0
    
    ' Core functions
    AddFunc "getFunc", "getFunc(IndexVaOrName)", "Get function by index or name"
    AddFunc "getSects", "getSects([optSegNameBaseorIndex])", "Get sections/segments"
    AddFunc "delSect", "delSect(nameOrBase)", "Delete section"
    AddFunc "sectExists", "sectExists(nameOrBase)", "Check if section exists"
    AddFunc "addSect", "addSect(base, size, name)", "Add new section"
    AddFunc "importFile", "importFile(va, path, [newSegName])", "Import binary file"
    
    ' Enum functions
    AddFunc "add_enum", "add_enum(name)", "Create new enumeration"
    AddFunc "get_enum", "get_enum(name)", "Get enumeration by name"
    AddFunc "add_enum_member", "add_enum_member(id, name, value)", "Add enum member"
    
    ' Analysis functions
    AddFunc "getopv", "getopv(ea, index)", "Get operand value"
    AddFunc "dumpFunc", "dumpFunc(i, [flags])", "Dump function disassembly"
    AddFunc "dumpFuncBytes", "dumpFuncBytes(i)", "Dump function bytes"
    
    ' List operations
    AddFunc "copyAll", "copyAll()", "Copy all list items"
    AddFunc "clear", "clear()", "Clear list"
    AddFunc "addAddr", "addAddr(addr, [txt], [txt2], [txt3], [txt4])", "Add address to list"
    AddFunc "setColHeaders", "setColHeaders([txt1], [txt2], [txt3], [txt4])", "Set column headers"
    AddFunc "showList", "showList()", "Show results list"
    AddFunc "hideList", "hideList()", "Hide results list"
    
    ' Memory operations
    AddFunc "readQWord", "readQWord(va)", "Read 8 bytes at address"
    AddFunc "readLong", "readLong(va)", "Read 4 bytes at address"
    AddFunc "readShort", "readShort(va)", "Read 2 bytes at address"
    AddFunc "readByte", "readByte(va)", "Read 1 byte at address"
    AddFunc "readStr", "readStr(va, [isAscii], [maxLeng])", "Read string at address"
    AddFunc "originalByte", "originalByte(va)", "Get original byte before patching"
    
    ' Math operations
    AddFunc "toHex", "toHex(str1)", "Convert to hex string"
    AddFunc "add", "add(x64Str1, strorInt2)", "64-bit addition"
    AddFunc "subtract", "subtract(x64Str1, strorInt2)", "64-bit subtraction"
    AddFunc "intToHex", "intToHex(intVal)", "Convert integer to hex"
    
    ' UI functions
    AddFunc "alert", "alert(msg)", "Show alert dialog"
    AddFunc "message", "message(msg)", "Print message to log"
    AddFunc "t", "t(message)", "Trace message"
    AddFunc "confirm", "confirm(prompt)", "Show yes/no dialog"
    AddFunc "askValue", "askValue([prompt], [defVal])", "Prompt for value"
    
    ' File operations
    AddFunc "readFile", "readFile(filename)", "Read file contents"
    AddFunc "writeFile", "writeFile(path, data)", "Write data to file"
    AddFunc "appendFile", "appendFile(path, data)", "Append to file"
    AddFunc "fileExists", "fileExists(path)", "Check if file exists"
    AddFunc "deleteFile", "deleteFile(fpath)", "Delete file"
    AddFunc "writeBin", "writeBin(fpath, data)", "Write binary data"
    AddFunc "getBin", "getBin(fpath, [offset], [leng])", "Read binary data"
    AddFunc "md5File", "md5File(fpath)", "Calculate MD5 of file"
    
    ' Properties
    AddFunc "imagebase", "imagebase()", "Get image base address"
    AddFunc "loadedfile", "loadedfile", "Get loaded file path (property)"
    
    ' Navigation
    AddFunc "jump", "jump(va)", "Jump to virtual address"
    AddFunc "jumpRVA", "jumpRVA(rva)", "Jump to relative virtual address"
    AddFunc "screenEA", "screenEA()", "Get current screen address"
    
    ' Patching
    AddFunc "patchByte", "patchByte(va, newval)", "Patch byte at address"
    AddFunc "patchString", "patchString(va, str, [isUnicode])", "Patch string at address"
    
    ' Function analysis
    AddFunc "numFuncs", "numFuncs()", "Get number of functions"
    AddFunc "funcCount", "funcCount()", "Get function count"
    AddFunc "functionStart", "functionStart(functionIndex)", "Get function start address"
    AddFunc "functionEnd", "functionEnd(functionIndex)", "Get function end address"
    AddFunc "functionName", "functionName(functionIndex)", "Get function name"
    AddFunc "funcIndexFromVA", "funcIndexFromVA(va)", "Get function index from address"
    AddFunc "funcVAByName", "funcVAByName(name)", "Get function address by name"
    AddFunc "renameFunc", "renameFunc(oldname, newName)", "Rename function"
    
    ' Instruction operations
    AddFunc "instSize", "instSize(va)", "Get instruction size"
    AddFunc "getAsm", "getAsm(va)", "Get disassembly at address"
    AddFunc "makeCode", "makeCode(offset)", "Convert to code"
    AddFunc "isCode", "isCode(va)", "Check if address is code"
    AddFunc "isData", "isData(va)", "Check if address is data"
    
    ' Cross-references
    AddFunc "xRefsTo", "xRefsTo(va)", "Get references to address"
    AddFunc "xRefsFrom", "xRefsFrom(va)", "Get references from address"
    AddFunc "xRefsToCnt", "xRefsToCnt(offset)", "Count references to address"
    AddFunc "addCodeXRef", "addCodeXRef(offset, tova)", "Add code cross-reference"
    AddFunc "addDataXRef", "addDataXRef(offset, tova)", "Add data cross-reference"
    AddFunc "delCodeXRef", "delCodeXRef(offset, tova)", "Delete code cross-reference"
    AddFunc "delDataXRef", "delDataXRef(offset, tova)", "Delete data cross-reference"
    
    ' Names and comments
    AddFunc "setName", "setName(offset, name)", "Set name at address"
    AddFunc "removeName", "removeName(offset)", "Remove name"
    AddFunc "getName", "getName(offset)", "Get name at address"
    AddFunc "addComment", "addComment(offset, comment)", "Add comment"
    AddFunc "getComment", "getComment(offset)", "Get comment"
    
    ' Visibility
    AddFunc "hideEA", "hideEA(offset)", "Hide address"
    AddFunc "showEA", "showEA(offset)", "Show address"
    AddFunc "hideBlock", "hideBlock(offset, endAt)", "Hide block"
    AddFunc "showBlock", "showBlock(offset, endAt)", "Show block"
    AddFunc "undefine", "undefine(va)", "Undefine at address"
    
    ' Navigation
    AddFunc "nextEA", "nextEA(va)", "Get next address"
    AddFunc "prevEA", "prevEA(va)", "Get previous address"
    AddFunc "refresh", "refresh()", "Refresh display"
    
    ' Clipboard
    AddFunc "getClipboard", "getClipboard()", "Get clipboard text"
    AddFunc "setClipboard", "setClipboard(string)", "Set clipboard text"
    
    ' Dialogs
    AddFunc "openFileDialog", "openFileDialog()", "Show open file dialog"
    AddFunc "saveFileDialog", "saveFileDialog()", "Show save file dialog"
    
    ' Advanced
    AddFunc "decompile", "decompile(va)", "Decompile function (requires Hex-Rays)"
    AddFunc "quickCall", "quickCall(msg, arg1)", "Quick IDA API call"
    AddFunc "enableIDADebugMessages", "enableIDADebugMessages([enabled])", "Enable debug output"
    AddFunc "benchMark", "benchMark()", "Start/stop benchmark timer"
    
    ' Process scanning
    AddFunc "ScanProcess", "ScanProcess(pidOrName)", "Scan running process"
    AddFunc "ResolveExport", "ResolveExport(apiOrAddress)", "Resolve export"
    
    ' Utilities
    AddFunc "hexDump", "hexDump(x)", "Create hex dump"
    AddFunc "hexstr", "hexstr(x)", "Convert to hex string"
    AddFunc "toBytes", "toBytes(hexstr)", "Convert hex string to bytes"
    AddFunc "md5", "md5(Str)", "Calculate MD5 hash"
    AddFunc "rc4", "rc4(Str, Pass)", "RC4 encrypt/decrypt"
    AddFunc "expandPath", "expandPath(path)", "Expand environment variables"
    
    ' Address conversion
    AddFunc "OffsetToVa", "OffsetToVa(offset)", "Convert file offset to VA"
    AddFunc "VaToOffset", "VaToOffset(va)", "Convert VA to file offset"
    
    ' Script loading
    AddFunc "loadToJS", "loadToJS(fpath)", "Load file to JavaScript"
    AddFunc "getRes", "getRes(optPath)", "Get resource"
    
    ' Misc
    AddFunc "CloseIDA", "CloseIDA()", "Close IDA Pro"
    AddFunc "strings", "strings()", "Get all strings"
    AddFunc "find", "find(startea, endea, hexstr)", "Find hex pattern"
End Sub

Private Sub AddFunc(name As String, signature As String, desc As String)
    m_Functions(m_FunctionCount).name = name
    m_Functions(m_FunctionCount).signature = signature
    m_Functions(m_FunctionCount).description = desc
    m_FunctionCount = m_FunctionCount + 1
    
    If m_FunctionCount >= UBound(m_Functions) Then
        ReDim Preserve m_Functions(UBound(m_Functions) + 50)
    End If
End Sub

'=========================================================================
' Get autocomplete list (space-separated)
'=========================================================================
Public Function GetAutocompleteList() As String
    Dim i As Long
    Dim list As String
    
    For i = 0 To m_FunctionCount - 1
        list = list & m_Functions(i).name & " "
    Next i
    
    GetAutocompleteList = Trim$(list)
End Function

'=========================================================================
' Get call tip for function
'=========================================================================
Public Function GetCallTip(functionName As String) As String
    Dim i As Long
    
    For i = 0 To m_FunctionCount - 1
        If StrComp(m_Functions(i).name, functionName, vbTextCompare) = 0 Then
            GetCallTip = m_Functions(i).signature & vbCrLf & _
                        "  " & m_Functions(i).description
            Exit Function
        End If
    Next i
    
    GetCallTip = ""
End Function

'=========================================================================
' Find function by partial name (for autocomplete filtering)
'=========================================================================
Public Function FindFunctions(partial As String) As String
    Dim i As Long
    Dim list As String
    Dim lowerPartial As String
    
    lowerPartial = LCase$(partial)
    
    For i = 0 To m_FunctionCount - 1
        If InStr(1, LCase$(m_Functions(i).name), lowerPartial, vbTextCompare) = 1 Then
            list = list & m_Functions(i).name & " "
        End If
    Next i
    
    FindFunctions = Trim$(list)
End Function

'=========================================================================
' Get all function names as array
'=========================================================================
Public Function GetFunctionNames() As String()
    Dim names() As String
    Dim i As Long
    
    ReDim names(m_FunctionCount - 1)
    
    For i = 0 To m_FunctionCount - 1
        names(i) = m_Functions(i).name
    Next i
    
    GetFunctionNames = names
End Function

'=========================================================================
' Get function count
'=========================================================================
Public Function GetFunctionCount() As Long
    GetFunctionCount = m_FunctionCount
End Function

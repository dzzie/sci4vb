# Professional Scintilla Wrapper for VB6

## Overview

This is a **clean room implementation** of a professional-grade Scintilla text editor wrapper for Visual Basic 6. The wrapper provides comprehensive access to Scintilla's powerful features through an intuitive, well-organized object model.

## Architecture

### Core Design Philosophy

1. **Modular Feature Classes**: Functionality is organized into logical feature modules rather than exposing a monolithic API
2. **Clean Separation**: Each feature class (Document, Editor, Selection, etc.) handles a specific aspect of editing
3. **Type Safety**: Extensive use of enums for constants provides IntelliSense support and type safety
4. **Pro-Grade Subclassing**: Uses the proven cSubclass implementation for robust message handling
5. **Memory Efficient**: Direct function pointers where applicable for fast message dispatch

### Component Structure

```
SciWrapper (UserControl)
├── Document        - Text content and file operations
├── Editor          - Core editing operations (cut/copy/paste/indent)
├── Selection       - Caret and selection management
├── Lines           - Line-based operations
├── Styling         - Syntax highlighting and lexer configuration
├── Folding         - Code folding operations
├── Markers         - Visual markers in margins
├── Indicators      - Text highlighting overlays
├── Autocomplete    - Autocomplete and call tips
├── Margins         - Margin configuration
├── View            - Display settings (zoom, whitespace, caret)
└── Search          - Find and replace operations
```

## Key Features

### Document Management (CSciDocument)
- Get/Set text content
- Undo/Redo operations with undo action grouping
- Save point tracking for modification detection
- Read-only mode
- Text insertion at specific positions

### Editor Operations (CSciEditor)
- Clipboard operations (cut, copy, paste, clear)
- Tab and indentation management
- EOL mode configuration (CRLF, CR, LF)
- EOL conversion
- Overtype mode
- Virtual space options

### Selection Management (CSciSelection)
- Single and multiple selections
- Caret positioning
- Selection modes (stream, rectangle, lines)
- Caret policy configuration
- Goto operations

### Line Operations (CSciLines)
- Line counting and positioning
- Get line content
- Line indentation management
- Column operations
- Line manipulation (transpose, reverse, duplicate)
- Line visibility and scrolling

### Styling & Syntax (CSciStyling)
- Multiple lexer support (C++, Python, VB, JavaScript, HTML, XML, SQL, JSON, etc.)
- Comprehensive style configuration
- Font attributes (bold, italic, underline, size)
- Foreground and background colors
- Custom keyword sets
- Lexer properties

### Code Folding (CSciFolding)
- Fold level management
- Fold expansion/collapse
- Helper functions (CollapseAll, ExpandAll)
- Line visibility control

### Markers (CSciMarkers)
- 32 marker symbols available
- Custom colors per marker
- Find next/previous marker
- Marker management per line

### Indicators (CSciIndicators)
- 20 indicator styles (squiggle, box, underline, etc.)
- Custom colors and transparency
- Range filling and clearing
- Multiple indicators per position

### Autocomplete (CSciAutocomplete)
- Autocomplete list management
- Call tip display
- Customizable behavior (auto-hide, case sensitivity, etc.)
- Fill-up and stop characters

### Margins (CSciMargins)
- Up to 5 configurable margins
- Line numbers
- Folding controls
- Custom markers
- Margin colors and cursors
- Helper functions for common configurations

### View Settings (CSciView)
- Zoom in/out
- Whitespace visibility
- Caret line highlighting
- Edge column guide
- Word wrap configuration
- Scrollbar control

### Search & Replace (CSciSearch)
- Forward and backward search
- Regular expression support
- Match case, whole word, word start options
- Replace operations with regex support
- Replace all functionality
- Target range management

## Usage Examples

### Basic Setup

```vb
' Initialize in Form_Load
Private Sub Form_Load()
    With sci
        ' Configure margins
        .Margins.ConfigureLineNumbers 0, 40
        .Margins.ConfigureFolding 1, 16
        
        ' Set editor preferences
        .Editor.TabWidth = 4
        .Editor.UseTabs = False
        .Editor.TabIndents = True
        
        ' Configure visual settings
        .View.CaretLineVisible = True
        .View.CaretLineBack = &HE8E8E8
        .View.EdgeMode = edgeLine
        .View.EdgeColumn = 80
        
        ' Setup default font
        .Styling.SetFont 32, "Consolas"  ' STYLE_DEFAULT
        .Styling.SetSize 32, 10
        .Styling.ClearAll
    End With
End Sub
```

### Loading and Saving Files

```vb
' Load file
Private Sub LoadFile(filename As String)
    Dim fileNum As Integer
    Dim content As String
    
    fileNum = FreeFile
    Open filename For Binary As #fileNum
    content = Space$(LOF(fileNum))
    Get #fileNum, , content
    Close #fileNum
    
    sci.Document.Text = content
    sci.Document.SetSavePoint
End Sub

' Save file
Private Sub SaveFile(filename As String)
    Dim fileNum As Integer
    
    fileNum = FreeFile
    Open filename For Output As #fileNum
    Print #fileNum, sci.Document.Text;
    Close #fileNum
    
    sci.Document.SetSavePoint
End Sub
```

### Syntax Highlighting

```vb
' Configure VB syntax highlighting
Private Sub SetupVBHighlighting()
    With sci.Styling
        .Lexer = lexVB
        
        ' Configure colors
        .SetFore 0, &H0           ' Default
        .SetBack 0, &HFFFFFF
        .SetFore 1, &H808080      ' Comment - gray
        .SetFore 2, &H8000        ' Number - blue
        .SetFore 3, &H8000        ' String - blue
        .SetFore 4, &HFF0000      ' Keyword - red
        .SetBold 4, True
        
        ' Set VB keywords
        .SetKeywords 0, "if then else elseif end sub function private public dim as long string"
        
        ' Colorize document
        .Colorise 0, sci.Document.TextLength
    End With
End Sub
```

### Search and Replace

```vb
' Find next occurrence
Private Sub FindNext()
    Dim pos As Long
    
    pos = sci.Search.FindNext("searchterm", _
                              matchCase:=True, _
                              wholeWord:=False, _
                              useRegex:=False)
    
    If pos >= 0 Then
        sci.Selection.SetSelection pos, pos + Len("searchterm")
        sci.Lines.ScrollCaret
    End If
End Sub

' Replace all occurrences
Private Sub ReplaceAll()
    Dim count As Long
    
    count = sci.Search.ReplaceAll("find", "replace", _
                                  matchCase:=True, _
                                  wholeWord:=True)
    
    MsgBox "Replaced " & count & " occurrences"
End Sub
```

### Markers and Bookmarks

```vb
' Add bookmark marker
Private Sub AddBookmark(line As Long)
    Const BOOKMARK_MARKER = 1
    
    ' Define bookmark appearance
    sci.Markers.Define BOOKMARK_MARKER, markerCircle
    sci.Markers.SetBack BOOKMARK_MARKER, &HFF0000  ' Red
    
    ' Add marker to line
    sci.Markers.Add line, BOOKMARK_MARKER
End Sub

' Find next bookmark
Private Sub FindNextBookmark()
    Dim currentLine As Long
    Dim nextBookmark As Long
    
    currentLine = sci.Lines.LineFromPosition(sci.Selection.CurrentPos)
    nextBookmark = sci.Markers.FindNext(currentLine + 1, 2)  ' Mask = 2 (marker 1)
    
    If nextBookmark >= 0 Then
        sci.Selection.GotoLine nextBookmark
    End If
End Sub
```

### Code Folding

```vb
' Toggle fold at current line
Private Sub ToggleFold()
    Dim line As Long
    line = sci.Lines.LineFromPosition(sci.Selection.CurrentPos)
    
    If sci.Folding.IsFoldHeader(line) Then
        sci.Folding.ToggleFold line
    End If
End Sub

' Collapse all folds
Private Sub CollapseAll()
    sci.Folding.CollapseAll
End Sub

' Expand all folds
Private Sub ExpandAll()
    sci.Folding.ExpandAll
End Sub
```

### Autocomplete

```vb
' Show autocomplete list
Private Sub ShowAutocomplete()
    Dim wordStart As Long
    Dim currentPos As Long
    Dim lineText As String
    
    currentPos = sci.Selection.CurrentPos
    wordStart = sci.Selection.CurrentPos - GetCurrentWordLength()
    
    ' Build completion list (items separated by space)
    Dim items As String
    items = "Function Sub Property Dim Private Public"
    
    sci.Autocomplete.Show currentPos - wordStart, items
End Sub

' Show call tip
Private Sub ShowCallTip()
    Dim pos As Long
    pos = sci.Selection.CurrentPos
    
    sci.Autocomplete.ShowCallTip pos, "FunctionName(param1 As String, param2 As Long) As Boolean"
End Sub
```

## Event Handling

### Available Events

```vb
' Text modification
Private Sub sci_Modified(position As Long, modificationType As Long, text As String, length As Long, linesAdded As Long)
    ' Handle document modifications
End Sub

' UI update (caret move, selection change)
Private Sub sci_UpdateUI(updated As Long)
    ' Update status bar, etc.
End Sub

' Margin click (for folding, bookmarks)
Private Sub sci_MarginClick(margin As Long, position As Long, modifiers As Long)
    If margin = 1 Then  ' Folding margin
        Dim line As Long
        line = sci.Lines.LineFromPosition(position)
        sci.Folding.ToggleFold line
    End If
End Sub

' Character added
Private Sub sci_CharAdded(ch As Long)
    ' Trigger autocomplete, etc.
End Sub

' Save point reached/left
Private Sub sci_SavePointReached()
    ' Document saved
End Sub

Private Sub sci_SavePointLeft()
    ' Document modified after save
End Sub
```

## Requirements

- Visual Basic 6.0
- SciLexer.dll (must be in application directory or system path)
- Windows XP or later

## Installation

1. Add all class files to your VB6 project
2. Add the SciWrapper.ctl UserControl to your project
3. Ensure SciLexer.dll is available at runtime
4. Place the control on your form
5. Configure as needed in Form_Load

## Performance Considerations

- The wrapper uses direct function pointers where possible for optimal performance
- Message batching is supported via BeginUndoAction/EndUndoAction
- Large documents are handled efficiently through Scintilla's gap buffer implementation
- Styling operations can be deferred for better performance during bulk updates

## Thread Safety

Like all VB6 COM controls, this wrapper is designed for single-threaded apartments (STA). Do not call methods from multiple threads.

## Error Handling

The wrapper includes basic error handling. Production applications should wrap calls in proper error handlers:

```vb
On Error GoTo ErrorHandler
    sci.Document.Text = LoadLargeFile()
Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description
End Sub
```

## Extending the Wrapper

To add new Scintilla messages:

1. Add the message constant to the appropriate class
2. Create a property or method that calls SciMsg/SciMsgStr/SciMsgPtr
3. Add appropriate error handling
4. Update documentation

## License

This is a clean room implementation created from Scintilla documentation. The Scintilla library itself is licensed under a permissive license. See the Scintilla documentation for details.

## Credits

- Scintilla Editor: Neil Hodgson (https://www.scintilla.org/)
- Subclassing Implementation: Paul Caton
- Wrapper Architecture: David Zimmer 
- AI: Claude.ai

## Support

This is a professional implementation suitable for production use. The modular architecture makes it easy to extend and maintain.

For Scintilla-specific questions, consult the official Scintilla documentation at https://www.scintilla.org/ScintillaDoc.html

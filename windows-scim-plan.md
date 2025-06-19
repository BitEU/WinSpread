# Windows Terminal Spreadsheet Calculator (WTSC) - Development Plan

## Overview
A terminal-based spreadsheet calculator for Windows Command Prompt, inspired by sc-im but built using native Windows Console APIs. No .NET or external dependencies - only standard Windows C libraries and APIs.

## Core Architecture

### 1. Technology Stack
- **Language**: C (C99 standard)
- **Compiler**: MSVC (comes with Windows SDK) or MinGW
- **APIs**: Windows Console API, Win32 API
- **Build System**: Simple batch files or makefiles
- **No external dependencies**: Only Windows system libraries

### 2. Project Structure
```
wtsc/
├── src/
│   ├── main.c              # Entry point and main loop
│   ├── console.c/.h        # Windows console handling
│   ├── display.c/.h        # Screen rendering and UI
│   ├── input.c/.h          # Keyboard input handling
│   ├── sheet.c/.h          # Spreadsheet data structures
│   ├── cell.c/.h           # Cell data and operations
│   ├── formula.c/.h        # Formula parser and evaluator
│   ├── commands.c/.h       # Command mode implementation
│   ├── file_io.c/.h        # File operations (CSV, native format)
│   ├── clipboard.c/.h      # Windows clipboard integration
│   ├── history.c/.h        # Undo/redo functionality
│   ├── utils.c/.h          # Utility functions
│   └── config.h            # Configuration constants
├── tests/
│   └── test_*.c            # Unit tests
├── docs/
│   ├── user_manual.md
│   └── developer_guide.md
├── build.bat               # Build script
└── README.md
```

## Key Components

### 1. Console Management (console.c)
Windows Console API functions to use:
- `GetStdHandle()` - Get console handles
- `SetConsoleMode()` - Enable virtual terminal sequences
- `SetConsoleScreenBufferSize()` - Set buffer dimensions
- `SetConsoleCursorPosition()` - Move cursor
- `WriteConsoleOutput()` - Fast screen updates
- `SetConsoleTextAttribute()` - Colors
- `GetConsoleScreenBufferInfo()` - Get console info

Features:
- Double buffering for flicker-free updates
- Color support (16 base colors + intensity)
- Unicode support via wide character APIs
- Screen resize detection
- Cursor control

### 2. Display Engine (display.c)
Rendering system:
- Grid layout with row/column headers
- Status line at bottom
- Command line area
- Cell highlighting
- Scrolling viewport for large sheets
- Efficient partial screen updates

Layout example:
```
     A       B       C       D       E
  1  100     200     =A1+B1  400     500
  2  Text    50      =B2*2   100     150
  3  Date    =SUM(B1:B2)     250     300
  
[Sheet1] A1: 100                    NUM
```

### 3. Data Structures (sheet.c, cell.c)
```c
typedef enum {
    CELL_EMPTY,
    CELL_NUMBER,
    CELL_STRING,
    CELL_FORMULA,
    CELL_ERROR
} CellType;

typedef struct Cell {
    CellType type;
    union {
        double number;
        char* string;
        char* formula;
    } value;
    char* display_value;  // Cached display string
    int width;            // Column width
    int format;           // Number format flags
} Cell;

typedef struct Sheet {
    Cell** cells;         // 2D array of cells
    int max_row, max_col;
    int* col_widths;      // Column widths
    char* name;
    struct Sheet* next;   // Linked list for multiple sheets
} Sheet;
```

### 4. Input Handling (input.c)
Windows Console Input:
- `ReadConsoleInput()` - Read keyboard events
- Handle special keys (arrows, function keys)
- Modal operation (NORMAL, INSERT, COMMAND modes)
- Key mapping system

Key bindings (vim-like):
- `h,j,k,l` - Navigation
- `=` - Enter numeric formula
- `"` - Enter text
- `:` - Command mode
- `y/p` - Copy/paste
- `u/Ctrl+R` - Undo/redo

### 5. Formula Engine (formula.c)
Parser and evaluator:
- Tokenizer for formula parsing
- Recursive descent parser
- Support for basic operators: +, -, *, /, ^
- Functions: SUM, AVG, COUNT, MIN, MAX, IF
- Cell references (A1, B2:C5)
- Error handling (#DIV/0!, #REF!, #VALUE!)

Example grammar:
```
expression := term (('+' | '-') term)*
term := factor (('*' | '/') factor)*
factor := number | cell_ref | function | '(' expression ')'
function := FUNC_NAME '(' arg_list ')'
```

### 6. Command System (commands.c)
Command mode operations:
- `:w [filename]` - Save
- `:q` - Quit
- `:wq` - Save and quit
- `:set` - Configuration
- `:goto` - Jump to cell
- `:format` - Cell formatting

### 7. File I/O (file_io.c)
Supported formats:
1. **Native format** (.wts) - Binary format for full feature preservation
2. **CSV** - Import/export with proper escaping
3. **Tab-delimited** - Simple data exchange

Native format structure:
```
[Header]
- Magic number
- Version
- Sheet count
[Sheets]
- Sheet name
- Dimensions
- Cell data (compressed)
```

### 8. Windows-Specific Features

#### Clipboard Integration
```c
// Using Windows Clipboard API
OpenClipboard(NULL);
EmptyClipboard();
HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, size);
// Copy data...
SetClipboardData(CF_TEXT, hMem);
CloseClipboard();
```

#### Console Virtual Terminal Sequences
Enable ANSI escape codes on Windows 10:
```c
DWORD mode;
GetConsoleMode(hOut, &mode);
mode |= ENABLE_VIRTUAL_TERMINAL_PROCESSING;
SetConsoleMode(hOut, mode);
```

## Implementation Phases

### Phase 1: Core Infrastructure (2-3 weeks)
- [ ] Console initialization and management
- [ ] Basic display grid rendering
- [ ] Keyboard input handling
- [ ] Navigation (arrow keys, hjkl)
- [ ] Basic cell data structure

### Phase 2: Cell Operations (2-3 weeks)
- [ ] Cell input (numbers and text)
- [ ] Cell selection and highlighting
- [ ] Copy/paste operations
- [ ] Column width adjustment
- [ ] Basic formatting

### Phase 3: Formula Engine (3-4 weeks)
- [ ] Formula tokenizer
- [ ] Expression parser
- [ ] Basic arithmetic operations
- [ ] Cell references
- [ ] Common functions (SUM, AVG, etc.)

### Phase 4: File Operations (1-2 weeks)
- [ ] Native file format
- [ ] CSV import/export
- [ ] Autosave functionality

### Phase 5: Advanced Features (2-3 weeks)
- [ ] Multiple sheets
- [ ] Undo/redo system
- [ ] Search and replace
- [ ] Cell formatting options
- [ ] Status line and help

### Phase 6: Polish and Testing (1-2 weeks)
- [ ] Performance optimization
- [ ] Memory leak fixes
- [ ] Comprehensive testing
- [ ] Documentation

## Key Algorithms

### 1. Dependency Graph for Formulas
Track cell dependencies to handle updates:
```c
typedef struct DepNode {
    int row, col;
    struct DepNode** dependents;
    int dep_count;
} DepNode;
```

### 2. Efficient Screen Updates
Only redraw changed cells:
```c
typedef struct DirtyRegion {
    int min_row, max_row;
    int min_col, max_col;
} DirtyRegion;
```

### 3. Virtual Viewport
For large spreadsheets:
```c
typedef struct Viewport {
    int top_row, left_col;
    int height, width;
} Viewport;
```

## Windows Console Tips

### 1. Performance
- Use `WriteConsoleOutput` for bulk updates
- Implement dirty region tracking
- Cache formatted strings
- Use double buffering

### 2. Unicode Support
```c
// Use wide character versions
WriteConsoleW();
ReadConsoleW();
SetConsoleTitleW();
```

### 3. Colors
```c
// 16 colors + intensity bit
enum Colors {
    BLACK = 0,
    BLUE = 1,
    GREEN = 2,
    CYAN = 3,
    RED = 4,
    MAGENTA = 5,
    YELLOW = 6,
    WHITE = 7,
    INTENSITY = 8
};
```

## Build Configuration

### Using MSVC (build.bat):
```batch
@echo off
cl /O2 /W3 src/*.c /Fe:wtsc.exe /link user32.lib
```

### Using MinGW:
```batch
gcc -O2 -Wall src/*.c -o wtsc.exe -municode
```

## Testing Strategy

1. **Unit Tests**: Test individual components
2. **Integration Tests**: Test component interactions
3. **Stress Tests**: Large spreadsheets, complex formulas
4. **User Acceptance**: Real-world use cases

## Future Enhancements

1. **Scripting**: Embedded scripting language
2. **Charts**: ASCII-based charts
3. **Import**: Excel file reading (using Windows COM if available)
4. **Network**: Shared spreadsheets
5. **Macros**: Record and playback actions

## Example Code Snippet

```c
// Main loop structure
int main() {
    Console* con = console_init();
    Sheet* sheet = sheet_new(1000, 100);
    State state = {.mode = MODE_NORMAL};
    
    while (state.running) {
        display_render(con, sheet, &state);
        
        INPUT_RECORD input;
        if (console_read_input(con, &input)) {
            handle_input(&input, sheet, &state);
        }
        
        update_calculations(sheet);
    }
    
    sheet_free(sheet);
    console_cleanup(con);
    return 0;
}
```

## Conclusion

This plan provides a solid foundation for building a sc-im-like spreadsheet calculator for Windows. The key is to leverage Windows Console API effectively while maintaining clean architecture and efficient algorithms. The modular design allows for incremental development and easy maintenance.
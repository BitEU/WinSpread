# Windows Terminal Spreadsheet Calculator (WTSC)

A lightweight, terminal-based spreadsheet calculator for Windows Command Prompt, inspired by sc-im but built using native Windows Console APIs.

## Features

- **Vim-like navigation**: Use h,j,k,l or arrow keys to move around
- **Modal editing**: Normal, Insert, and Command modes
- **Cell types**: Numbers, text, and formulas
- **Basic formulas**: Cell references and simple calculations
- **Efficient rendering**: Double-buffered console output
- **No dependencies**: Uses only Windows system libraries

## Requirements

- Windows 10 or later
- C compiler (MSVC or MinGW)
- No external libraries required!

## Building

### Method 1: Using the batch file (Recommended)
```cmd
build.bat
```

### Method 2: Using Make (if you have MinGW or MSYS2)
```cmd
make
```

### Method 3: Manual compilation with MSVC
```cmd
cl /O2 /W3 /TC main.c console.c sheet.c /Fe:wtsc.exe /link user32.lib
```

### Method 4: Manual compilation with GCC
```cmd
gcc -O2 -Wall -std=c99 main.c console.c sheet.c -o wtsc.exe
```

## File Structure

```
WinSpread/
├── console.h      # Windows Console API wrapper declarations
├── console.c      # Console handling implementation
├── sheet.h        # Spreadsheet data structure declarations
├── sheet.c        # Spreadsheet logic implementation
├── main.c         # Main program and UI
├── build.bat      # Windows build script
├── Makefile       # Alternative build system
└── README.md      # This file
```

## Usage

### Navigation
- `h,j,k,l` or Arrow Keys - Move cursor
- `PageUp/PageDown` - Move 10 rows up/down
- `Ctrl+Q` - Quick quit

### Cell Input
- `=` - Enter a number or formula
- `"` - Enter text
- `x` - Clear current cell

### Command Mode
- `:` - Enter command mode
- `:q` - Quit
- `:w` - Save (not implemented yet)
- `:wq` - Save and quit

### Formulas
Currently supports:
- Simple numbers: `123.45`
- Cell references: `=A1`
- More complex formulas coming soon!

## Example Session

1. Run `wtsc.exe`
2. Press `"` to enter text mode
3. Type "Hello World" and press Enter
4. Press `j` to move down
5. Press `=` to enter a number
6. Type `42` and press Enter
7. Press `:q` to quit

## Future Enhancements

- [ ] Complex formula parsing (arithmetic expressions)
- [ ] More functions (SUM, AVG, etc.)
- [ ] CSV import/export
- [ ] Save/load native format
- [ ] Copy/paste operations
- [ ] Multiple sheets
- [ ] Undo/redo

## Known Limitations

- Basic formula parser (only handles single cell references)
- No file saving yet
- Fixed column width
- Limited to console buffer size

## Contributing

This is a learning project designed to demonstrate Windows Console API usage. Feel free to fork and experiment!

## License

This project is provided as-is for educational purposes.

## Troubleshooting

### "Cannot open source file: 'user32.lib'"
This error means the MSVC linker can't find Windows libraries. Make sure you're running from a Visual Studio Developer Command Prompt.

### Compilation errors about undefined functions
Make sure all .c files are included in the compilation command. The build script should handle this automatically.

### Display issues
The program requires a console window at least 80x25 characters. Maximize your command prompt window for best results.
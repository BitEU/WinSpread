# WinSpread - A Windows Terminal Spreadsheet Calculator

A powerful, vi-style spreadsheet application that runs entirely in the Windows terminal/console. WinSpread combines the functionality of a traditional spreadsheet with the efficiency of keyboard-driven navigation, making it perfect for quick calculations, data analysis, and spreadsheet work without leaving the command line.

![License](https://img.shields.io/badge/license-GPL--3.0-blue.svg)
![Platform](https://img.shields.io/badge/platform-Windows-blue.svg)
![Language](https://img.shields.io/badge/language-C-orange.svg)

## Features

### Core Functionality
- **Full spreadsheet capabilities** with 1000 rows × 100 columns
- **Multiple data types**: Numbers, strings, formulas, and error handling
- **Real-time formula evaluation** with automatic recalculation
- **Vi-style keyboard navigation** for efficient operation
- **Double-buffered rendering** for smooth 60 FPS display
- **Blinking cursor** with visual feedback

### Supported Functions
- **Mathematical**: `SUM`, `AVG`, `MAX`, `MIN`, `MEDIAN`, `MODE`, `POWER`
- **Conditional**: `IF(condition, true_value, false_value)`
- **Operators**: `+`, `-`, `*`, `/`, `>`, `<`, `>=`, `<=`, `=`, `<>`
- **Cell ranges**: `A1:A10`, `B1:C5` for aggregate functions

### Advanced Features
- **Copy/Paste**: Full cell copying with `Ctrl+C` and `Ctrl+V`
- **Formula dependencies**: Automatic dependency tracking and recalculation
- **Error handling**: Division by zero, reference errors, and parse errors
- **Cell formatting**: Width, precision, and alignment support
- **Command mode**: Vi-style commands for advanced operations

## Installation

### Prerequisites
- Windows 10/11 or Windows Server 2016+
- One of the following compilers:
- Visual Studio Build Tools

### Building from Source

1. **Clone the repository:**
   ```cmd
   git clone https://github.com/yourusername/WinSpread.git
   cd WinSpread
   ```

2. **Build using the provided script:**
   ```cmd
   build.bat
   ```

3. **Run the application:**
   ```cmd
   wtsc.exe
   ```

## Usage

### Basic Navigation
- **`h`** - Move left
- **`j`** - Move down  
- **`k`** - Move up
- **`l`** - Move right
- **Arrow keys** - Alternative navigation
- **`Page Up/Down`** - Jump 10 rows

### Data Entry
- **`=`** - Enter number or formula mode
- **`"`** - Enter text/string mode
- **`:`** - Enter command mode
- **`Enter`** - Confirm input
- **`Escape`** - Cancel input
- **`Backspace`** - Delete character during input

### Cell Operations
- **`x`** - Clear current cell
- **`Ctrl+C`** - Copy current cell
- **`Ctrl+V`** - Paste copied cell
- **`Ctrl+Q`** - Quick quit

### Command Mode
- **`:q`** or **`:quit`** - Quit application
- **`:w`** - Save (not yet implemented)
- **`:wq`** - Save and quit

## Functions

(COPILOT, SUM IS THE EXAMPLE. MODEL THE OTHER SEVEN AFTER IT. SHOW ALL POSSIBLE AND ADVANCED WAYS A FORUMLA COULD BE IMPLEMENTED)

1. SUM: 
    -`=SUM(A1:A6)`
    -`SUM(B33:G33)`
2. AVG:
3. MAX:
4. MIN: 
5. MEDIAN:
6. MODE:
7. IF:
8. POWER:

## File Structure

```
WinSpread/
├── main.c          # Main application logic and UI
├── sheet.h         # Spreadsheet engine and formula parser
├── console.h       # Windows console wrapper and input handling
├── compat.h        # Compatibility definitions
├── build.bat       # Build script for Windows
├── LICENSE         # GPL v3 license
└── README.md       # This file
```

## Architecture

### Core Components

1. **Application State (`main.c`)**
   - Manages UI rendering and user interaction
   - Handles cursor movement and blinking
   - Coordinates between console and spreadsheet

2. **Spreadsheet Engine (`sheet.h`)**
   - Cell storage and data management
   - Formula parsing and evaluation
   - Dependency tracking and recalculation
   - Built-in function implementations

3. **Console Wrapper (`console.h`)**
   - Windows Console API abstraction
   - Double-buffered rendering
   - Keyboard input handling
   - Color and cursor management

### Key Features

- **Memory efficient**: Uses sparse matrix for cell storage
- **Fast rendering**: Double-buffered display with only changed cells updated
- **Robust parsing**: Recursive descent parser for complex formulas
- **Dependency tracking**: Automatic recalculation when referenced cells change

## Error Handling

WinSpread provides comprehensive error handling:

- **`#DIV/0!`** - Division by zero
- **`#REF!`** - Invalid cell reference
- **`#VALUE!`** - Invalid value or type mismatch
- **`#PARSE!`** - Formula parsing error

## Future Enhancements

- [ ] Save/Load functionality (CSV, Excel formats)
- [ ] System clipboard integration
- [ ] Undo/Redo support
- [ ] Cell formatting (colors, borders)
- [ ] Additional mathematical functions
- [ ] Search and replace
- [ ] Sorting and filtering

## Contributing

Contributions are welcome! Please feel free to submit pull requests or open issues for bugs and feature requests.

### Development Setup
1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly on Windows
5. Submit a pull request

## License

This project is licensed under the GNU General Public License v3.0 - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Inspired by classical spreadsheet applications like VisiCalc and sc-im
- Vi-style navigation inspired by Vim text editor
- Built with the Windows Console API for optimal terminal performance

---

*WinSpread - Bringing spreadsheet power to your Windows terminal*

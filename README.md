# WinSpread - A Windows Terminal Spreadsheet Calculator

A powerful, vi-style spreadsheet application that runs entirely in the Windows terminal/console. WinSpread combines the functionality of a traditional spreadsheet with the efficiency of keyboard-driven navigation, making it perfect for quick calculations, data analysis, and spreadsheet work without leaving the command line.

![License](https://img.shields.io/badge/license-GPL--3.0-blue.svg)
![Platform](https://img.shields.io/badge/platform-Windows-blue.svg)
![Language](https://img.shields.io/badge/language-C-orange.svg)

## Table of Contents

- [Features](#features)
  - [Core Functionality](#core-functionality)
  - [Supported Functions](#supported-functions)
  - [Advanced Features](#advanced-features)
- [Installation](#installation)
  - [Prerequisites](#prerequisites)
  - [Building from Source](#building-from-source)
- [Usage](#usage)
  - [Basic Navigation](#basic-navigation)
  - [Data Entry](#data-entry)
  - [Cell Operations](#cell-operations)
  - [Command Mode](#command-mode)
- [CSV File Operations](#csv-file-operations)
  - [Saving to CSV](#saving-to-csv)
  - [Loading from CSV](#loading-from-csv)
  - [CSV Format Options](#csv-format-options)
- [Functions](#functions)
  - [1. SUM Function](#1-sum-function)
  - [2. AVG Function](#2-avg-function)
  - [3. MAX Function](#3-max-function)
  - [4. MIN Function](#4-min-function)
  - [5. MEDIAN Function](#5-median-function)
  - [6. MODE Function](#6-mode-function)
  - [7. IF Function](#7-if-function)
  - [8. POWER Function](#8-power-function)
  - [Mathematical Operators](#mathematical-operators)
  - [Range Notation](#range-notation)
- [File Structure](#file-structure)
- [Architecture](#architecture)
  - [Core Components](#core-components)
  - [Key Features](#key-features)
- [Error Handling](#error-handling)
- [Future Enhancements](#future-enhancements)
- [Contributing](#contributing)
  - [Development Setup](#development-setup)
- [License](#license)
- [Acknowledgments](#acknowledgments)

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
   git clone https://github.com/BitEU/WinSpread.git
   cd WinSpread
   ```

2. **Build using the provided script:**
   ```cmd
   .\build.bat
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
- **`Ctrl+C`** - Copy current cell (internal)
- **`Ctrl+V`** - Paste copied cell (internal)
- **`Ctrl+Shift+C`** - Copy current cell (external)
- **`Ctrl+Shift+V`** - Paste copied cell (external)
- **`Ctrl+Q`** - Quick quit

### Command Mode
- **`:q`** or **`:quit`** - Quit application

## CSV File Operations

WinSpread provides comprehensive CSV (Comma-Separated Values) file support, allowing you to import data from external sources and export your spreadsheet calculations for use in other applications.

### Saving to CSV

To save your current spreadsheet to a CSV file, use the `savecsv` command in command mode:

**Syntax:** `:savecsv <filename>`

**Examples:**
- `:savecsv data.csv` - Save to data.csv in the current directory
- `:savecsv C:\Documents\report.csv` - Save with full path
- `:savecsv budget_2024.csv` - Save with descriptive filename

When saving to CSV, you'll be prompted to choose between two formats:

#### CSV Format Options

**Flatten Mode (`f`)**: Save calculated values
- Formulas are evaluated and their results are saved
- Best for sharing data with other applications
- Preserves the visual output of your spreadsheet
- Example: A cell containing `=SUM(A1:A5)` saves as `150` (the calculated result)

**Preserve Mode (`p`)**: Save formulas as text
- Formulas are saved as text strings with their original expressions
- Best for backing up your work or sharing with other WinSpread users
- Maintains the logic and structure of your spreadsheet
- Example: A cell containing `=SUM(A1:A5)` saves as `"=SUM(A1:A5)"`

### Loading from CSV

To load data from a CSV file into your current spreadsheet, use the `loadcsv` command:

**Syntax:** `:loadcsv <filename>`

**Examples:**
- `:loadcsv data.csv` - Load from data.csv in the current directory
- `:loadcsv C:\Documents\import.csv` - Load with full path
- `:loadcsv sales_data.csv` - Load sales data

**Important Notes:**
- Loading CSV data will overwrite existing cells in the affected range
- The data will be loaded starting from cell A1
- Text fields containing formulas (from preserve mode) will be interpreted as formulas
- Numeric data will be automatically recognized and converted
- Empty cells in the CSV will clear corresponding spreadsheet cells

### CSV Format Compatibility

WinSpread's CSV implementation follows standard CSV formatting rules:

**Supported Features:**
- Comma-separated values
- Quoted text fields containing commas
- Escaped quotes within quoted fields (`""` becomes `"`)
- Mixed data types (numbers, text, formulas)
- Empty fields and rows
- Windows and Unix line endings

**Example CSV Format:**
```csv
Name,Age,Salary,Department
"Smith, John",35,50000,Engineering
"Doe, Jane",28,45000,"Human Resources"
Bob,42,55000,Sales
Alice,31,,Marketing
```

This CSV would load as:
- A1: `Name`, B1: `Age`, C1: `Salary`, D1: `Department`
- A2: `Smith, John`, B2: `35`, C2: `50000`, D2: `Engineering`
- A3: `Doe, Jane`, B3: `28`, C3: `45000`, D3: `Human Resources`
- A4: `Bob`, B4: `42`, C4: `55000`, D4: `Sales`
- A5: `Alice`, B5: `31`, C5: (empty), D5: `Marketing`

## Functions

WinSpread supports a comprehensive set of built-in functions for mathematical calculations, statistical analysis, and conditional logic. All functions are case-sensitive and must be entered in UPPERCASE.

### 1. SUM Function
**Syntax:** `=SUM(range)` or `=SUM(value1, value2, ...)`
**Description:** Calculates the sum of all numeric values in the specified range or arguments.

**Examples:**
- `=SUM(A1:A6)` - Sum cells A1 through A6
- `=SUM(B33:G33)` - Sum cells B33 through G33 (horizontal range)
- `=SUM(A1:C3)` - Sum all cells in the rectangular range A1 to C3
- `=SUM(A1:A10, C1:C5)` - Sum multiple ranges (if implemented)

### 2. AVG Function
**Syntax:** `=AVG(range)` or `=AVG(value1, value2, ...)`
**Description:** Calculates the arithmetic mean (average) of all numeric values in the specified range.

**Examples:**
- `=AVG(A1:A10)` - Average of cells A1 through A10
- `=AVG(B1:B20)` - Average of cells B1 through B20
- `=AVG(A1:E1)` - Average of cells A1 through E1 (horizontal range)
- `=AVG(A1:C3)` - Average of all cells in the rectangular range A1 to C3

### 3. MAX Function
**Syntax:** `=MAX(range)` or `=MAX(value1, value2, ...)`
**Description:** Returns the largest (maximum) value from the specified range or arguments.

**Examples:**
- `=MAX(A1:A10)` - Maximum value in cells A1 through A10
- `=MAX(B1:D1)` - Maximum value in cells B1 through D1
- `=MAX(A1:C5)` - Maximum value in the rectangular range A1 to C5
- `=MAX(A1:A5, C1:C5)` - Maximum value across multiple ranges

### 4. MIN Function
**Syntax:** `=MIN(range)` or `=MIN(value1, value2, ...)`
**Description:** Returns the smallest (minimum) value from the specified range or arguments.

**Examples:**
- `=MIN(A1:A10)` - Minimum value in cells A1 through A10
- `=MIN(B1:D1)` - Minimum value in cells B1 through D1
- `=MIN(A1:C5)` - Minimum value in the rectangular range A1 to C5
- `=MIN(A1:A5, C1:C5)` - Minimum value across multiple ranges

### 5. MEDIAN Function
**Syntax:** `=MEDIAN(range)` or `=MEDIAN(value1, value2, ...)`
**Description:** Returns the median (middle value) of the specified range. For even number of values, returns the average of the two middle values.

**Examples:**
- `=MEDIAN(A1:A9)` - Median of cells A1 through A9 (odd count)
- `=MEDIAN(A1:A10)` - Median of cells A1 through A10 (even count)
- `=MEDIAN(B1:E1)` - Median of horizontal range B1 through E1
- `=MEDIAN(A1:C3)` - Median of all values in rectangular range A1 to C3

### 6. MODE Function
**Syntax:** `=MODE(range)` or `=MODE(value1, value2, ...)`
**Description:** Returns the most frequently occurring value in the specified range. If multiple values have the same highest frequency, returns the first one encountered.

**Examples:**
- `=MODE(A1:A10)` - Most frequent value in cells A1 through A10
- `=MODE(B1:B20)` - Most frequent value in cells B1 through B20
- `=MODE(A1:E1)` - Most frequent value in horizontal range A1 through E1
- `=MODE(A1:C5)` - Most frequent value in rectangular range A1 to C5

### 7. IF Function
**Syntax:** `=IF(condition, true_value, false_value)`
**Description:** Evaluates a condition and returns one value if true, another if false. Supports both numeric and string values.

**Comparison Operators:**
- `=` (equals), `<>` (not equals)
- `<` (less than), `>` (greater than)
- `<=` (less than or equal), `>=` (greater than or equal)

**Examples:**
- `=IF(A1>10, "High", "Low")` - Return "High" if A1 > 10, otherwise "Low"
- `=IF(B1=0, "Zero", B1*2)` - Return "Zero" if B1 is 0, otherwise double B1
- `=IF(A1>=B1, A1, B1)` - Return the larger of A1 or B1
- `=IF(C1="Yes", 100, 0)` - Return 100 if C1 contains "Yes", otherwise 0
- `=IF(A1<>B1, "Different", "Same")` - Compare two cells for equality

**String Comparisons:**
- `=IF(A1="Apple", "Fruit", "Other")` - Check if A1 contains "Apple"
- `=IF(B1<>"", "Has Value", "Empty")` - Check if B1 is not empty

### 8. POWER Function
**Syntax:** `=POWER(base, exponent)`
**Description:** Raises a number to a specified power (base^exponent).

**Examples:**
- `=POWER(2, 3)` - Calculate 2³ = 8
- `=POWER(A1, 2)` - Square the value in A1
- `=POWER(10, B1)` - Calculate 10 raised to the power of B1
- `=POWER(A1, 0.5)` - Calculate square root of A1 (A1^0.5)
- `=POWER(A1, 1/3)` - Calculate cube root of A1

### Mathematical Operators

**Arithmetic Operators:**
- `+` Addition: `=A1+B1`, `=SUM(A1:A5)+10`
- `-` Subtraction: `=A1-B1`, `=MAX(A1:A5)-MIN(A1:A5)`
- `*` Multiplication: `=A1*B1`, `=A1*2.5`
- `/` Division: `=A1/B1`, `=SUM(A1:A5)/5`

**Complex Formula Examples:**
- `=SUM(A1:A10)/MAX(B1:B10)` - Ratio of sum to maximum
- `=IF(AVG(A1:A10)>50, "Pass", "Fail")` - Conditional based on average
- `=POWER(A1, 2) + POWER(B1, 2)` - Sum of squares
- `=IF(A1>0, POWER(A1, 0.5), 0)` - Square root if positive, else 0

### Range Notation

**Supported Range Formats:**
- `A1:A10` - Vertical range (column A, rows 1-10)
- `A1:E1` - Horizontal range (row 1, columns A-E)
- `A1:C3` - Rectangular range (3×3 block)
- `B5:D10` - Rectangular range (3×6 block)

**Range Examples:**
- Single column: `=SUM(A1:A100)`
- Single row: `=AVG(A1:Z1)`
- Rectangle: `=MAX(A1:E10)`
- Large range: `=MIN(A1:Z100)`

## File Structure

```
WinSpread/
├── main.c          # Main application logic and UI
├── sheet.h         # Spreadsheet engine and formula parser
├── console.h       # Windows console wrapper and input handling
├── compat.h        # Compatibility definitions
├── debug.h        # Manages debugging log
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

*WinSpread - Party like it's 1979*
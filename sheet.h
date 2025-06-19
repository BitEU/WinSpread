// sheet.h - Spreadsheet data structures and operations
#ifndef SHEET_H
#define SHEET_H

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <math.h>
#include <ctype.h>

// Cell types
typedef enum {
    CELL_EMPTY,
    CELL_NUMBER,
    CELL_STRING,
    CELL_FORMULA,
    CELL_ERROR
} CellType;

// Error types
typedef enum {
    ERROR_NONE,
    ERROR_DIV_ZERO,
    ERROR_REF,
    ERROR_VALUE,
    ERROR_PARSE
} ErrorType;

// Cell structure
typedef struct Cell {
    CellType type;
    union {
        double number;
        char* string;        struct {
            char* expression;
            double cached_value;
            char* cached_string;    // For IF function string results
            int is_string_result;   // Flag indicating if result is a string
            ErrorType error;
        } formula;
    } data;
    
    // Display properties
    int width;
    int precision;
    int align;  // 0=left, 1=center, 2=right
    
    // Dependencies
    struct Cell** depends_on;    // Cells this cell depends on
    int depends_count;
    struct Cell** dependents;    // Cells that depend on this cell
    int dependents_count;
    
    // Position (for dependency tracking)
    int row;
    int col;
} Cell;

// Sheet structure
typedef struct Sheet {
    Cell*** cells;      // 2D array of cell pointers
    int rows;
    int cols;
    int* col_widths;
    char* name;
    
    // Calculation state
    int needs_recalc;
    Cell** calc_order;  // Topological sort of cells for calculation
    int calc_count;
} Sheet;

// Function prototypes
Sheet* sheet_new(int rows, int cols);
void sheet_free(Sheet* sheet);
Cell* sheet_get_cell(Sheet* sheet, int row, int col);
Cell* sheet_get_or_create_cell(Sheet* sheet, int row, int col);
void sheet_set_number(Sheet* sheet, int row, int col, double value);
void sheet_set_string(Sheet* sheet, int row, int col, const char* str);
void sheet_set_formula(Sheet* sheet, int row, int col, const char* formula);
void sheet_clear_cell(Sheet* sheet, int row, int col);
char* sheet_get_display_value(Sheet* sheet, int row, int col);
void sheet_recalculate(Sheet* sheet);

// Copy/paste operations
void sheet_copy_cell(Sheet* sheet, int src_row, int src_col, int dest_row, int dest_col);
Cell* sheet_get_clipboard_cell(void);
void sheet_set_clipboard_cell(Cell* cell);

// Cell operations
Cell* cell_new(int row, int col);
void cell_free(Cell* cell);
void cell_set_number(Cell* cell, double value);
void cell_set_string(Cell* cell, const char* str);
void cell_set_formula(Cell* cell, const char* formula);
void cell_clear(Cell* cell);
char* cell_get_display_value(Cell* cell);

// Formula evaluation
double evaluate_formula(Sheet* sheet, const char* formula, ErrorType* error);
double evaluate_expression(Sheet* sheet, const char* expr, ErrorType* error);
double evaluate_comparison(Sheet* sheet, const char* expr, ErrorType* error);
int parse_cell_reference(const char* ref, int* row, int* col);
char* cell_reference_to_string(int row, int col);

// Skip whitespace in expression
void skip_whitespace(const char** expr);

// Implementation

Sheet* sheet_new(int rows, int cols) {
    Sheet* sheet = (Sheet*)calloc(1, sizeof(Sheet));
    if (!sheet) return NULL;
    
    sheet->rows = rows;
    sheet->cols = cols;
    sheet->name = _strdup("Sheet1");
    
    // Allocate cell array
    sheet->cells = (Cell***)calloc(rows, sizeof(Cell**));
    if (!sheet->cells) {
        free(sheet);
        return NULL;
    }
    
    for (int i = 0; i < rows; i++) {
        sheet->cells[i] = (Cell**)calloc(cols, sizeof(Cell*));
        if (!sheet->cells[i]) {
            // Cleanup on failure
            for (int j = 0; j < i; j++) {
                free(sheet->cells[j]);
            }
            free(sheet->cells);
            free(sheet);
            return NULL;
        }
    }
    
    // Initialize column widths
    sheet->col_widths = (int*)calloc(cols, sizeof(int));
    for (int i = 0; i < cols; i++) {
        sheet->col_widths[i] = 10;  // Default width
    }
    
    return sheet;
}

void sheet_free(Sheet* sheet) {
    if (!sheet) return;
    
    // Free all cells
    for (int i = 0; i < sheet->rows; i++) {
        for (int j = 0; j < sheet->cols; j++) {
            if (sheet->cells[i][j]) {
                cell_free(sheet->cells[i][j]);
            }
        }
        free(sheet->cells[i]);
    }
    free(sheet->cells);
    
    free(sheet->col_widths);
    free(sheet->name);
    free(sheet->calc_order);
    free(sheet);
}

Cell* sheet_get_cell(Sheet* sheet, int row, int col) {
    if (row < 0 || row >= sheet->rows || col < 0 || col >= sheet->cols) {
        return NULL;
    }
    return sheet->cells[row][col];
}

Cell* sheet_get_or_create_cell(Sheet* sheet, int row, int col) {
    if (row < 0 || row >= sheet->rows || col < 0 || col >= sheet->cols) {
        return NULL;
    }
    
    if (!sheet->cells[row][col]) {
        sheet->cells[row][col] = cell_new(row, col);
    }
    
    return sheet->cells[row][col];
}

Cell* cell_new(int row, int col) {
    Cell* cell = (Cell*)calloc(1, sizeof(Cell));
    if (!cell) return NULL;
    
    cell->type = CELL_EMPTY;
    cell->row = row;
    cell->col = col;
    cell->width = 10;
    cell->precision = 2;
    cell->align = 2;  // Right align for numbers by default
    
    return cell;
}

void cell_free(Cell* cell) {
    if (!cell) return;
      // Free string data
    if (cell->type == CELL_STRING && cell->data.string) {
        free(cell->data.string);
    } else if (cell->type == CELL_FORMULA) {
        if (cell->data.formula.expression) {
            free(cell->data.formula.expression);
        }
        if (cell->data.formula.cached_string) {
            free(cell->data.formula.cached_string);
        }
    }
    
    // Free dependency arrays
    free(cell->depends_on);
    free(cell->dependents);
    
    free(cell);
}

void cell_clear(Cell* cell) {
    if (!cell) return;
      // Free existing data
    if (cell->type == CELL_STRING && cell->data.string) {
        free(cell->data.string);
        cell->data.string = NULL;
    } else if (cell->type == CELL_FORMULA) {
        if (cell->data.formula.expression) {
            free(cell->data.formula.expression);
            cell->data.formula.expression = NULL;
        }
        if (cell->data.formula.cached_string) {
            free(cell->data.formula.cached_string);
            cell->data.formula.cached_string = NULL;
        }
    }
    
    cell->type = CELL_EMPTY;
}

void cell_set_number(Cell* cell, double value) {
    if (!cell) return;
    
    cell_clear(cell);
    cell->type = CELL_NUMBER;
    cell->data.number = value;
}

void cell_set_string(Cell* cell, const char* str) {
    if (!cell) return;
    
    cell_clear(cell);
    cell->type = CELL_STRING;
    cell->data.string = _strdup(str);
    cell->align = 0;  // Left align for strings
}

void cell_set_formula(Cell* cell, const char* formula) {
    if (!cell) return;
      cell_clear(cell);
    cell->type = CELL_FORMULA;
    cell->data.formula.expression = _strdup(formula);
    cell->data.formula.cached_value = 0.0;
    cell->data.formula.cached_string = NULL;
    cell->data.formula.is_string_result = 0;
    cell->data.formula.error = ERROR_NONE;
}

void sheet_set_number(Sheet* sheet, int row, int col, double value) {
    Cell* cell = sheet_get_or_create_cell(sheet, row, col);
    if (cell) {
        cell_set_number(cell, value);
        sheet->needs_recalc = 1;
    }
}

void sheet_set_string(Sheet* sheet, int row, int col, const char* str) {
    Cell* cell = sheet_get_or_create_cell(sheet, row, col);
    if (cell) {
        cell_set_string(cell, str);
    }
}

void sheet_set_formula(Sheet* sheet, int row, int col, const char* formula) {
    Cell* cell = sheet_get_or_create_cell(sheet, row, col);
    if (cell) {
        cell_set_formula(cell, formula);
        sheet->needs_recalc = 1;
    }
}

void sheet_clear_cell(Sheet* sheet, int row, int col) {
    Cell* cell = sheet_get_cell(sheet, row, col);
    if (cell) {
        cell_clear(cell);
        sheet->needs_recalc = 1;
    }
}

char* cell_get_display_value(Cell* cell) {
    static char buffer[256];
    
    if (!cell || cell->type == CELL_EMPTY) {
        return "";
    }
    
    switch (cell->type) {
        case CELL_NUMBER:
            snprintf(buffer, sizeof(buffer), "%.*f", cell->precision, cell->data.number);
            // Remove trailing zeros
            char* dot = strchr(buffer, '.');
            if (dot) {
                char* end = buffer + strlen(buffer) - 1;
                while (end > dot && *end == '0') *end-- = '\0';
                if (*end == '.') *end = '\0';
            }
            return buffer;
            
        case CELL_STRING:
            return cell->data.string;
              case CELL_FORMULA:
            if (cell->data.formula.error != ERROR_NONE) {
                switch (cell->data.formula.error) {
                    case ERROR_DIV_ZERO: return "#DIV/0!";
                    case ERROR_REF: return "#REF!";
                    case ERROR_VALUE: return "#VALUE!";
                    case ERROR_PARSE: return "#PARSE!";
                    default: return "#ERROR!";
                }
            }
            
            // Check if this formula has a string result (from IF function)
            if (cell->data.formula.is_string_result && cell->data.formula.cached_string) {
                return cell->data.formula.cached_string;
            }
            
            // Otherwise format as number
            snprintf(buffer, sizeof(buffer), "%.*f", cell->precision, cell->data.formula.cached_value);
            // Remove trailing zeros
            dot = strchr(buffer, '.');
            if (dot) {
                char* end = buffer + strlen(buffer) - 1;
                while (end > dot && *end == '0') *end-- = '\0';
                if (*end == '.') *end = '\0';
            }
            return buffer;
            
        case CELL_ERROR:
            return "#ERROR!";
            
        default:
            return "";
    }
}

char* sheet_get_display_value(Sheet* sheet, int row, int col) {
    Cell* cell = sheet_get_cell(sheet, row, col);
    return cell_get_display_value(cell);
}

// Convert column number to letter(s) (0 -> A, 25 -> Z, 26 -> AA, etc.)
char* cell_reference_to_string(int row, int col) {
    static char buffer[16];
    char colStr[8];
    int idx = 0;
    
    // Convert column number to letters
    col++;  // Make it 1-based for conversion
    while (col > 0) {
        col--;
        colStr[idx++] = 'A' + (col % 26);
        col /= 26;
    }
    
    // Reverse the string
    for (int i = 0; i < idx / 2; i++) {
        char temp = colStr[i];
        colStr[i] = colStr[idx - 1 - i];
        colStr[idx - 1 - i] = temp;
    }
    colStr[idx] = '\0';
    
    snprintf(buffer, sizeof(buffer), "%s%d", colStr, row + 1);
    return buffer;
}

// Parse cell reference like "A1" or "AB23"
int parse_cell_reference(const char* ref, int* row, int* col) {
    if (!ref || !row || !col) return 0;
    
    const char* p = ref;
    *col = 0;
    
    // Skip whitespace
    while (*p && isspace(*p)) p++;
    
    // Parse column letters
    if (!isalpha(*p)) return 0;
    
    while (*p && isalpha(*p)) {
        *col = *col * 26 + (toupper(*p) - 'A' + 1);
        p++;
    }
    (*col)--;  // Convert to 0-based
    
    // Parse row number
    if (!isdigit(*p)) return 0;
    
    *row = 0;
    while (*p && isdigit(*p)) {
        *row = *row * 10 + (*p - '0');
        p++;
    }
    (*row)--;  // Convert to 0-based
    
    // Check for trailing garbage
    while (*p && isspace(*p)) p++;
    if (*p) return 0;
    
    return 1;
}

// Forward declarations for expression parsing
double parse_arithmetic_expression(Sheet* sheet, const char** expr, ErrorType* error);
double parse_arithmetic_expression_old(Sheet* sheet, const char* expr, ErrorType* error);
double parse_term(Sheet* sheet, const char** expr, ErrorType* error);
double parse_factor(Sheet* sheet, const char** expr, ErrorType* error);
double parse_function(Sheet* sheet, const char** expr, ErrorType* error);
void skip_whitespace(const char** expr);

// Global variables for IF function string results
extern char g_if_result_string[256];
extern int g_if_result_is_string;
extern Cell* g_current_evaluating_cell;

// Range parsing structures
typedef struct {
    int start_row, start_col;
    int end_row, end_col;
} CellRange;

int parse_range(const char* range_str, CellRange* range);
int get_range_values(Sheet* sheet, const CellRange* range, double* values, int max_values);

// Function implementations
double func_sum(const double* values, int count);
double func_avg(const double* values, int count);
double func_max(const double* values, int count);
double func_min(const double* values, int count);
double func_median(double* values, int count);
double func_mode(const double* values, int count);
double func_if(double condition, double true_val, double false_val);
double func_if_enhanced(double condition, double true_val, double false_val, 
                       const char* true_str, const char* false_str);
double func_power(double base, double exponent);

// Parse range notation like "A1:A3" or "B2:D5"
int parse_range(const char* range_str, CellRange* range) {
    if (!range_str || !range) return 0;
    
    const char* colon = strchr(range_str, ':');
    if (!colon) return 0;
    
    // Parse start cell
    char start_ref[16];
    int start_len = (int)(colon - range_str);
    if (start_len >= sizeof(start_ref)) return 0;
    
    strncpy_s(start_ref, sizeof(start_ref), range_str, start_len);
    start_ref[start_len] = '\0';
    
    if (!parse_cell_reference(start_ref, &range->start_row, &range->start_col)) {
        return 0;
    }
    
    // Parse end cell
    const char* end_ref = colon + 1;
    if (!parse_cell_reference(end_ref, &range->end_row, &range->end_col)) {
        return 0;
    }
    
    // Ensure start <= end
    if (range->start_row > range->end_row) {
        int temp = range->start_row;
        range->start_row = range->end_row;
        range->end_row = temp;
    }
    if (range->start_col > range->end_col) {
        int temp = range->start_col;
        range->start_col = range->end_col;
        range->end_col = temp;
    }
    
    return 1;
}

// Get values from a range of cells
int get_range_values(Sheet* sheet, const CellRange* range, double* values, int max_values) {
    int count = 0;
    
    for (int row = range->start_row; row <= range->end_row; row++) {
        for (int col = range->start_col; col <= range->end_col; col++) {
            if (count >= max_values) break;
            
            Cell* cell = sheet_get_cell(sheet, row, col);
            if (cell) {
                switch (cell->type) {
                    case CELL_NUMBER:
                        values[count++] = cell->data.number;
                        break;
                    case CELL_FORMULA:
                        if (cell->data.formula.error == ERROR_NONE) {
                            values[count++] = cell->data.formula.cached_value;
                        }
                        break;
                    case CELL_EMPTY:
                        values[count++] = 0.0;
                        break;
                    default:
                        // Skip strings and errors
                        break;
                }
            } else {
                values[count++] = 0.0;
            }
        }
    }
    
    return count;
}

// Function implementations
double func_sum(const double* values, int count) {
    double sum = 0.0;
    for (int i = 0; i < count; i++) {
        sum += values[i];
    }
    return sum;
}

double func_avg(const double* values, int count) {
    if (count == 0) return 0.0;
    return func_sum(values, count) / count;
}

double func_max(const double* values, int count) {
    if (count == 0) return 0.0;
    double max = values[0];
    for (int i = 1; i < count; i++) {
        if (values[i] > max) max = values[i];
    }
    return max;
}

double func_min(const double* values, int count) {
    if (count == 0) return 0.0;
    double min = values[0];
    for (int i = 1; i < count; i++) {
        if (values[i] < min) min = values[i];
    }
    return min;
}

// Helper function for median calculation
int compare_double(const void* a, const void* b) {
    double da = *(const double*)a;
    double db = *(const double*)b;
    if (da < db) return -1;
    if (da > db) return 1;
    return 0;
}

double func_median(double* values, int count) {
    if (count == 0) return 0.0;
    
    // Sort the values
    qsort(values, count, sizeof(double), compare_double);
    
    if (count % 2 == 0) {
        // Even number of values - return average of middle two
        return (values[count/2 - 1] + values[count/2]) / 2.0;
    } else {
        // Odd number of values - return middle value
        return values[count/2];
    }
}

double func_mode(const double* values, int count) {
    if (count == 0) return 0.0;
    
    // Simple mode calculation - find most frequent value
    double mode = values[0];
    int max_count = 1;
    
    for (int i = 0; i < count; i++) {
        int current_count = 1;
        for (int j = i + 1; j < count; j++) {
            if (fabs(values[i] - values[j]) < 1e-10) {  // Compare with small tolerance
                current_count++;
            }
        }
        if (current_count > max_count) {
            max_count = current_count;
            mode = values[i];
        }
    }
    
    return mode;
}

double func_if(double condition, double true_val, double false_val) {
    return (condition != 0.0) ? true_val : false_val;
}

// Global variables for IF function string results
char g_if_result_string[256] = {0};
int g_if_result_is_string = 0;
Cell* g_current_evaluating_cell = NULL;  // Track current cell during evaluation

double func_if_enhanced(double condition, double true_val, double false_val, 
                       const char* true_str, const char* false_str) {
    // Reset global string result flag
    g_if_result_is_string = 0;
    g_if_result_string[0] = '\0';
    
    // Also reset the current cell's string result
    if (g_current_evaluating_cell) {
        if (g_current_evaluating_cell->data.formula.cached_string) {
            free(g_current_evaluating_cell->data.formula.cached_string);
            g_current_evaluating_cell->data.formula.cached_string = NULL;
        }
        g_current_evaluating_cell->data.formula.is_string_result = 0;
    }
    
    if (condition != 0.0) {
        // True condition
        if (true_str) {
            strcpy_s(g_if_result_string, sizeof(g_if_result_string), true_str);
            g_if_result_is_string = 1;
            
            // Store in current cell if available
            if (g_current_evaluating_cell) {
                g_current_evaluating_cell->data.formula.cached_string = _strdup(true_str);
                g_current_evaluating_cell->data.formula.is_string_result = 1;
            }
            
            return 1.0; // Return non-zero to indicate string result available
        } else {
            return true_val;
        }
    } else {
        // False condition
        if (false_str) {
            strcpy_s(g_if_result_string, sizeof(g_if_result_string), false_str);
            g_if_result_is_string = 1;
            
            // Store in current cell if available
            if (g_current_evaluating_cell) {
                g_current_evaluating_cell->data.formula.cached_string = _strdup(false_str);
                g_current_evaluating_cell->data.formula.is_string_result = 1;
            }
            
            return 0.0; // Return zero to indicate string result available
        } else {
            return false_val;
        }
    }
}

double func_power(double base, double exponent) {
    return pow(base, exponent);
}

// Skip whitespace in expression
void skip_whitespace(const char** expr) {
    while (**expr && isspace(**expr)) {
        (*expr)++;
    }
}

// Clipboard functionality
static Cell* clipboard_cell = NULL;

Cell* sheet_get_clipboard_cell(void) {
    return clipboard_cell;
}

void sheet_set_clipboard_cell(Cell* cell) {
    // Free existing clipboard cell
    if (clipboard_cell) {
        cell_free(clipboard_cell);
    }
    
    // Create a copy of the cell
    if (cell) {
        clipboard_cell = cell_new(cell->row, cell->col);
        switch (cell->type) {
            case CELL_NUMBER:
                cell_set_number(clipboard_cell, cell->data.number);
                break;
            case CELL_STRING:
                cell_set_string(clipboard_cell, cell->data.string);
                break;
            case CELL_FORMULA:
                cell_set_formula(clipboard_cell, cell->data.formula.expression);
                clipboard_cell->data.formula.cached_value = cell->data.formula.cached_value;
                clipboard_cell->data.formula.error = cell->data.formula.error;
                break;
            default:
                clipboard_cell->type = CELL_EMPTY;
                break;
        }
        
        // Copy display properties
        clipboard_cell->width = cell->width;
        clipboard_cell->precision = cell->precision;
        clipboard_cell->align = cell->align;
    } else {
        clipboard_cell = NULL;
    }
}

void sheet_copy_cell(Sheet* sheet, int src_row, int src_col, int dest_row, int dest_col) {
    Cell* src_cell = sheet_get_cell(sheet, src_row, src_col);
    
    if (!src_cell) {
        // Clear destination cell if source is empty
        sheet_clear_cell(sheet, dest_row, dest_col);
        return;
    }
    
    switch (src_cell->type) {
        case CELL_NUMBER:
            sheet_set_number(sheet, dest_row, dest_col, src_cell->data.number);
            break;
        case CELL_STRING:
            sheet_set_string(sheet, dest_row, dest_col, src_cell->data.string);
            break;
        case CELL_FORMULA:
            sheet_set_formula(sheet, dest_row, dest_col, src_cell->data.formula.expression);
            break;
        default:
            sheet_clear_cell(sheet, dest_row, dest_col);
            break;
    }
    
    // Copy display properties
    Cell* dest_cell = sheet_get_cell(sheet, dest_row, dest_col);
    if (dest_cell) {
        dest_cell->width = src_cell->width;
        dest_cell->precision = src_cell->precision;
        dest_cell->align = src_cell->align;
    }
    
    sheet_recalculate(sheet);
}

// Parse a factor (number, cell reference, or function call)
double parse_factor(Sheet* sheet, const char** expr, ErrorType* error) {
    skip_whitespace(expr);
    
    if (**expr == '(') {
        (*expr)++; // Skip '('
        double result = parse_arithmetic_expression(sheet, expr, error);
        if (*error != ERROR_NONE) return 0.0;
        
        skip_whitespace(expr);
        if (**expr != ')') {
            *error = ERROR_PARSE;
            return 0.0;
        }
        (*expr)++; // Skip ')'
        
        return result;
    }
    
    // Check if this could be a function call (letters followed by '(')
    const char* start = *expr;
    const char* lookahead = *expr;
    
    // Look ahead to see if this is a function call
    while (*lookahead && isalpha(*lookahead)) lookahead++;
    skip_whitespace(&lookahead);
    
    if (*lookahead == '(') {
        // This is a function call
        return parse_function(sheet, expr, error);
    }
    
    // Try to parse as cell reference or range
    char ref_buf[32];
    int i = 0;
    
    // Extract potential cell reference (letters followed by numbers, possibly with colon)
    while (**expr && (isalpha(**expr) || isdigit(**expr) || **expr == ':') && i < 31) {
        ref_buf[i++] = **expr;
        (*expr)++;
    }
    ref_buf[i] = '\0';
    
    if (i > 0) {
        // Check if it's a range (contains ':')
        if (strchr(ref_buf, ':')) {
            CellRange range;
            if (parse_range(ref_buf, &range)) {
                // For a range without a function, just return the sum
                double values[1000];
                int count = get_range_values(sheet, &range, values, 1000);
                return func_sum(values, count);
            } else {
                *error = ERROR_PARSE;
                return 0.0;
            }
        } else {
            // Single cell reference
            int row, col;
            if (parse_cell_reference(ref_buf, &row, &col)) {
                Cell* cell = sheet_get_cell(sheet, row, col);
                if (!cell) {
                    *error = ERROR_REF;
                    return 0.0;
                }
                
                switch (cell->type) {
                    case CELL_EMPTY:
                        return 0.0;
                    case CELL_NUMBER:
                        return cell->data.number;
                    case CELL_FORMULA:
                        if (cell->data.formula.error != ERROR_NONE) {
                            *error = cell->data.formula.error;
                            return 0.0;
                        }
                        return cell->data.formula.cached_value;
                    default:
                        *error = ERROR_VALUE;
                        return 0.0;
                }
            }
        }
    }
    
    // Reset and try to parse as number
    *expr = start;
    char* endptr;
    double value = strtod(*expr, &endptr);
    if (endptr != *expr) {
        *expr = endptr;
        return value;
    }
    
    *error = ERROR_PARSE;
    return 0.0;
}

// Parse term (handles * and /)
double parse_term(Sheet* sheet, const char** expr, ErrorType* error) {
    double result = parse_factor(sheet, expr, error);
    if (*error != ERROR_NONE) return 0.0;
    
    while (1) {
        skip_whitespace(expr);
        if (**expr == '*') {
            (*expr)++;
            double right = parse_factor(sheet, expr, error);
            if (*error != ERROR_NONE) return 0.0;
            result *= right;
        } else if (**expr == '/') {
            (*expr)++;
            double right = parse_factor(sheet, expr, error);
            if (*error != ERROR_NONE) return 0.0;
            if (right == 0.0) {
                *error = ERROR_DIV_ZERO;
                return 0.0;
            }
            result /= right;
        } else {
            break;
        }
    }
      return result;
}

// Wrapper for backward compatibility with old API
double parse_arithmetic_expression_old(Sheet* sheet, const char* expr, ErrorType* error) {
    const char* p = expr;
    double result = parse_arithmetic_expression(sheet, &p, error);
    
    // Check for remaining characters
    skip_whitespace(&p);
    if (*p != '\0') {
        *error = ERROR_PARSE;
        return 0.0;
    }
    
    return result;
}

// Parse function calls like SUM(A1:A3), AVG(B1:B5), IF(A1>0,B1,C1)
double parse_function(Sheet* sheet, const char** expr, ErrorType* error) {
    skip_whitespace(expr);
    
    // Extract function name
    const char* start = *expr;
    char func_name[32];
    int i = 0;
    
    while (**expr && isalpha(**expr) && i < 31) {
        func_name[i++] = toupper(**expr);
        (*expr)++;
    }
    func_name[i] = '\0';
    
    skip_whitespace(expr);
    if (**expr != '(') {
        *error = ERROR_PARSE;
        return 0.0;
    }
    (*expr)++; // Skip '('
    
    double values[1000];  // Max 1000 values in a range
    int value_count = 0;
    
    // Handle different functions
    if (strcmp(func_name, "SUM") == 0 || strcmp(func_name, "AVG") == 0 || 
        strcmp(func_name, "MAX") == 0 || strcmp(func_name, "MIN") == 0 ||
        strcmp(func_name, "MEDIAN") == 0 || strcmp(func_name, "MODE") == 0) {
        
        // Parse range or single values
        skip_whitespace(expr);
        const char* arg_start = *expr;
        
        // Find the end of the argument (look for closing parenthesis)
        int paren_count = 1;
        const char* arg_end = *expr;
        while (*arg_end && paren_count > 0) {
            if (*arg_end == '(') paren_count++;
            else if (*arg_end == ')') paren_count--;
            if (paren_count > 0) arg_end++;
        }
        
        // Extract the argument
        int arg_len = (int)(arg_end - arg_start);
        char arg[256];
        if (arg_len >= sizeof(arg)) {
            *error = ERROR_PARSE;
            return 0.0;
        }
        
        strncpy_s(arg, sizeof(arg), arg_start, arg_len);
        arg[arg_len] = '\0';
        
        // Check if it's a range (contains ':')
        if (strchr(arg, ':')) {
            CellRange range;
            if (parse_range(arg, &range)) {
                value_count = get_range_values(sheet, &range, values, 1000);
            } else {
                *error = ERROR_PARSE;
                return 0.0;
            }
        } else {
            // Single cell reference or value
            int row, col;
            if (parse_cell_reference(arg, &row, &col)) {
                Cell* cell = sheet_get_cell(sheet, row, col);
                if (cell) {
                    switch (cell->type) {
                        case CELL_NUMBER:
                            values[0] = cell->data.number;
                            value_count = 1;
                            break;
                        case CELL_FORMULA:
                            if (cell->data.formula.error == ERROR_NONE) {
                                values[0] = cell->data.formula.cached_value;
                                value_count = 1;
                            }
                            break;
                        case CELL_EMPTY:
                            values[0] = 0.0;
                            value_count = 1;
                            break;
                        default:
                            *error = ERROR_VALUE;
                            return 0.0;
                    }
                } else {
                    values[0] = 0.0;
                    value_count = 1;
                }
            } else {
                // Try to parse as number
                char* endptr;
                double val = strtod(arg, &endptr);
                if (*endptr == '\0') {
                    values[0] = val;
                    value_count = 1;
                } else {
                    *error = ERROR_PARSE;
                    return 0.0;
                }
            }
        }
        
        *expr = arg_end + 1; // Skip closing ')'
        
        // Call appropriate function
        if (strcmp(func_name, "SUM") == 0) {
            return func_sum(values, value_count);
        } else if (strcmp(func_name, "AVG") == 0) {
            return func_avg(values, value_count);
        } else if (strcmp(func_name, "MAX") == 0) {
            return func_max(values, value_count);
        } else if (strcmp(func_name, "MIN") == 0) {
            return func_min(values, value_count);
        } else if (strcmp(func_name, "MEDIAN") == 0) {
            return func_median(values, value_count);        } else if (strcmp(func_name, "MODE") == 0) {
            return func_mode(values, value_count);
        }
        
    } else if (strcmp(func_name, "POWER") == 0) {
        // POWER function: POWER(base, exponent)
        skip_whitespace(expr);
        
        // Parse first argument (base)
        double base = parse_arithmetic_expression(sheet, expr, error);
        if (*error != ERROR_NONE) return 0.0;
        
        skip_whitespace(expr);
        if (**expr != ',') {
            *error = ERROR_PARSE;
            return 0.0;
        }
        (*expr)++; // Skip comma
        
        // Parse second argument (exponent)  
        double exponent = parse_arithmetic_expression(sheet, expr, error);
        if (*error != ERROR_NONE) return 0.0;
        
        skip_whitespace(expr);
        if (**expr != ')') {
            *error = ERROR_PARSE;
            return 0.0;
        }
        (*expr)++; // Skip closing parenthesis
        
        return func_power(base, exponent);      } else if (strcmp(func_name, "IF") == 0) {
        // IF function: IF(condition, true_value, false_value)
        // Enhanced to support string literals and string comparisons
        skip_whitespace(expr);
        
        // Extract the entire condition expression more carefully
        const char* condition_start = *expr;
        int paren_count = 0;
        const char* condition_end = *expr;
        
        // Find the comma that separates condition from true_value
        while (*condition_end && (paren_count > 0 || *condition_end != ',')) {
            if (*condition_end == '(') paren_count++;
            else if (*condition_end == ')') paren_count--;
            condition_end++;
        }
        
        if (*condition_end != ',') {
            *error = ERROR_PARSE;
            return 0.0;
        }
        
        // Extract condition string
        int condition_len = (int)(condition_end - condition_start);
        char condition_expr[256];
        if (condition_len >= sizeof(condition_expr)) {
            *error = ERROR_PARSE;
            return 0.0;
        }
        
        strncpy_s(condition_expr, sizeof(condition_expr), condition_start, condition_len);
        condition_expr[condition_len] = '\0';
        
        // Evaluate the condition
        double condition = evaluate_comparison(sheet, condition_expr, error);
        if (*error != ERROR_NONE) return 0.0;
        
        // Move past the comma
        *expr = condition_end + 1;
        
        // Enhanced parsing for true value (can be string or number)
        skip_whitespace(expr);
        double true_val = 0.0;
        char true_str[256] = {0};
        int is_true_string = 0;
        
        if (**expr == '"') {
            // Parse string literal
            (*expr)++; // Skip opening quote
            int i = 0;
            while (**expr && **expr != '"' && i < 255) {
                true_str[i++] = **expr;
                (*expr)++;
            }
            if (**expr == '"') {
                (*expr)++; // Skip closing quote
                is_true_string = 1;
            } else {
                *error = ERROR_PARSE;
                return 0.0;
            }
        } else {
            // Parse as numeric expression
            true_val = parse_arithmetic_expression(sheet, expr, error);
            if (*error != ERROR_NONE) return 0.0;
        }
        
        skip_whitespace(expr);
        if (**expr != ',') {
            *error = ERROR_PARSE;
            return 0.0;
        }
        (*expr)++; // Skip comma
        
        // Enhanced parsing for false value (can be string or number)
        skip_whitespace(expr);
        double false_val = 0.0;
        char false_str[256] = {0};
        int is_false_string = 0;
        
        if (**expr == '"') {
            // Parse string literal
            (*expr)++; // Skip opening quote
            int i = 0;
            while (**expr && **expr != '"' && i < 255) {
                false_str[i++] = **expr;
                (*expr)++;
            }
            if (**expr == '"') {
                (*expr)++; // Skip closing quote
                is_false_string = 1;
            } else {
                *error = ERROR_PARSE;
                return 0.0;
            }
        } else {
            // Parse as numeric expression
            false_val = parse_arithmetic_expression(sheet, expr, error);
            if (*error != ERROR_NONE) return 0.0;
        }
        
        skip_whitespace(expr);
        if (**expr != ')') {
            *error = ERROR_PARSE;
            return 0.0;
        }
        (*expr)++; // Skip closing parenthesis
        
        // Handle the IF logic with string support
        return func_if_enhanced(condition, true_val, false_val, 
                               is_true_string ? true_str : NULL,
                               is_false_string ? false_str : NULL);
    }
    
    *error = ERROR_PARSE;
    return 0.0;
}

// Simple formula evaluator (basic arithmetic and cell references)
double evaluate_formula(Sheet* sheet, const char* formula, ErrorType* error) {
    *error = ERROR_NONE;
    
    // Skip leading '='
    if (*formula == '=') formula++;
    
    // Try comparison expressions first (for IF function support)
    return evaluate_comparison(sheet, formula, error);
}

// Helper function to get string value from a cell reference
char* get_cell_string_value(Sheet* sheet, const char* ref) {
    int row, col;
    if (!parse_cell_reference(ref, &row, &col)) {
        return NULL;
    }
    
    Cell* cell = sheet_get_cell(sheet, row, col);
    if (!cell) return NULL;
    
    switch (cell->type) {
        case CELL_STRING:
            return cell->data.string;
        case CELL_FORMULA:
            if (cell->data.formula.is_string_result && cell->data.formula.cached_string) {
                return cell->data.formula.cached_string;
            }
            break;
        default:
            break;
    }
    return NULL;
}

// Enhanced comparison evaluation with string support
double evaluate_comparison(Sheet* sheet, const char* expr, ErrorType* error) {
    const char* p = expr;
    const char* start_left = p;
    
    // First, try to identify if we have a string comparison
    // Look for pattern: cell_ref = "string" or "string" = cell_ref
    
    // Check if left side is a cell reference
    char left_ref[32] = {0};
    int left_ref_len = 0;
    const char* temp_p = p;
    
    // Extract potential cell reference
    while (*temp_p && (isalpha(*temp_p) || isdigit(*temp_p)) && left_ref_len < 31) {
        left_ref[left_ref_len++] = *temp_p;
        temp_p++;
    }
    left_ref[left_ref_len] = '\0';
    
    // Check if it's a valid cell reference
    int is_left_cell_ref = 0;
    int left_row, left_col;
    if (left_ref_len > 0 && parse_cell_reference(left_ref, &left_row, &left_col)) {
        is_left_cell_ref = 1;
    }
    
    // Skip whitespace and look for comparison operator
    skip_whitespace(&temp_p);
    char comparison_op[3] = {0};
    int op_len = 0;
    
    if (*temp_p == '=' || *temp_p == '<' || *temp_p == '>') {
        comparison_op[op_len++] = *temp_p++;
        if (*temp_p == '=' || (*temp_p == '>' && comparison_op[0] == '<')) {
            comparison_op[op_len++] = *temp_p++;
        }
        comparison_op[op_len] = '\0';
    }
    
    // Check if right side is a string literal
    skip_whitespace(&temp_p);
    int is_right_string = (*temp_p == '"');
    
    // If we have cell_ref = "string", handle as string comparison
    if (is_left_cell_ref && strlen(comparison_op) > 0 && is_right_string) {
        // Get the string value from the cell
        char* left_str = get_cell_string_value(sheet, left_ref);
        
        // Parse the string literal
        temp_p++; // Skip opening quote
        char right_str[256] = {0};
        int right_len = 0;
        while (*temp_p && *temp_p != '"' && right_len < 255) {
            right_str[right_len++] = *temp_p++;
        }
        if (*temp_p == '"') temp_p++; // Skip closing quote
        
        // Perform string comparison
        int cmp_result = 0;
        if (left_str) {
            cmp_result = strcmp(left_str, right_str);
        } else {
            // If cell is empty or not string, compare with empty string
            cmp_result = strcmp("", right_str);
        }
        
        // Apply comparison operator
        if (strcmp(comparison_op, "=") == 0) {
            return (cmp_result == 0) ? 1.0 : 0.0;
        } else if (strcmp(comparison_op, "<>") == 0) {
            return (cmp_result != 0) ? 1.0 : 0.0;
        } else if (strcmp(comparison_op, "<") == 0) {
            return (cmp_result < 0) ? 1.0 : 0.0;
        } else if (strcmp(comparison_op, "<=") == 0) {
            return (cmp_result <= 0) ? 1.0 : 0.0;
        } else if (strcmp(comparison_op, ">") == 0) {
            return (cmp_result > 0) ? 1.0 : 0.0;
        } else if (strcmp(comparison_op, ">=") == 0) {
            return (cmp_result >= 0) ? 1.0 : 0.0;
        }
        
        *error = ERROR_PARSE;
        return 0.0;
    }
    
    // Fall back to numeric comparison
    double left = parse_arithmetic_expression(sheet, &p, error);
    if (*error != ERROR_NONE) return left;
    
    skip_whitespace(&p);
    
    // Check for comparison operators
    if (*p == '>') {
        p++;
        if (*p == '=') {
            p++; // >=
            double right = parse_arithmetic_expression(sheet, &p, error);
            return (left >= right) ? 1.0 : 0.0;
        } else {
            // >
            double right = parse_arithmetic_expression(sheet, &p, error);
            return (left > right) ? 1.0 : 0.0;
        }
    } else if (*p == '<') {
        p++;
        if (*p == '=') {
            p++; // <=
            double right = parse_arithmetic_expression(sheet, &p, error);
            return (left <= right) ? 1.0 : 0.0;
        } else if (*p == '>') {
            p++; // <>
            double right = parse_arithmetic_expression(sheet, &p, error);
            return (left != right) ? 1.0 : 0.0;
        } else {
            // <
            double right = parse_arithmetic_expression(sheet, &p, error);
            return (left < right) ? 1.0 : 0.0;
        }
    } else if (*p == '=') {
        p++; // =
        double right = parse_arithmetic_expression(sheet, &p, error);
        return (fabs(left - right) < 1e-10) ? 1.0 : 0.0; // Use tolerance for floating point comparison
    }
    
    // No comparison operator found, return the arithmetic expression result
    return left;
}

// Enhanced arithmetic expression parser (with updated signature)
double parse_arithmetic_expression(Sheet* sheet, const char** expr, ErrorType* error) {
    double result = parse_term(sheet, expr, error);
    if (*error != ERROR_NONE) return 0.0;
    
    while (1) {
        skip_whitespace(expr);
        if (**expr == '+') {
            (*expr)++;
            double right = parse_term(sheet, expr, error);
            if (*error != ERROR_NONE) return 0.0;
            result += right;
        } else if (**expr == '-') {
            (*expr)++;
            double right = parse_term(sheet, expr, error);
            if (*error != ERROR_NONE) return 0.0;
            result -= right;
        } else {
            break;
        }
    }
    
    return result;
}

// Recalculate all formulas in the sheet
void sheet_recalculate(Sheet* sheet) {
    if (!sheet->needs_recalc) return;
    
    // Simple recalculation - just evaluate all formulas
    // TODO: Implement proper dependency tracking and topological sort
    
    for (int row = 0; row < sheet->rows; row++) {
        for (int col = 0; col < sheet->cols; col++) {            Cell* cell = sheet_get_cell(sheet, row, col);
            if (cell && cell->type == CELL_FORMULA) {
                ErrorType error;
                g_current_evaluating_cell = cell;  // Set global context
                double value = evaluate_formula(sheet, cell->data.formula.expression, &error);
                cell->data.formula.cached_value = value;
                cell->data.formula.error = error;
                g_current_evaluating_cell = NULL;   // Clear global context
            }
        }
    }
    
    sheet->needs_recalc = 0;
}

// Demo: Create a simple spreadsheet with some data
void demo_spreadsheet() {
    Sheet* sheet = sheet_new(100, 26);
    
    // Add some sample data
    sheet_set_string(sheet, 0, 0, "Item");
    sheet_set_string(sheet, 0, 1, "Quantity");
    sheet_set_string(sheet, 0, 2, "Price");
    sheet_set_string(sheet, 0, 3, "Total");
    
    sheet_set_string(sheet, 1, 0, "Apples");
    sheet_set_number(sheet, 1, 1, 10);
    sheet_set_number(sheet, 1, 2, 0.5);
    sheet_set_formula(sheet, 1, 3, "=B2*C2");  // Would calculate to 5.0
    
    sheet_set_string(sheet, 2, 0, "Oranges");
    sheet_set_number(sheet, 2, 1, 15);
    sheet_set_number(sheet, 2, 2, 0.75);
    sheet_set_formula(sheet, 2, 3, "=B3*C3");  // Would calculate to 11.25
    
    // Recalculate formulas
    sheet_recalculate(sheet);
    
    // Print the sheet
    printf("\nSimple Spreadsheet Demo:\n");
    printf("------------------------\n");
    
    for (int row = 0; row <= 3; row++) {
        for (int col = 0; col <= 3; col++) {
            char* value = sheet_get_display_value(sheet, row, col);
            printf("%-12s", value);
        }
        printf("\n");
    }
    
    sheet_free(sheet);
}

#endif // SHEET_H
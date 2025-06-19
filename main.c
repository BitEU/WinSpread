// main.c - Main program for WinSpread
#include <windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>

#include "console.h"
#include "sheet.h"
#include "debug.h"

// Helper macros - use min/max from Windows headers

// Application state
typedef enum {
    MODE_NORMAL,
    MODE_INSERT_NUMBER,
    MODE_INSERT_STRING,
    MODE_INSERT_FORMULA,
    MODE_COMMAND
} AppMode;

typedef struct {
    Sheet* sheet;
    Console* console;
    AppMode mode;
    int cursor_row;
    int cursor_col;
    int view_top;     // Top visible row
    int view_left;    // Leftmost visible column
    char input_buffer[256];
    int input_pos;
    char status_message[256];
    BOOL running;
    
    // Cursor blinking state
    DWORD cursor_blink_time;    // Last time cursor blink state changed
    BOOL cursor_visible;        // Current cursor visibility state
    DWORD cursor_blink_rate;    // Blink rate in milliseconds
} AppState;

// Function prototypes
void app_init(AppState* state);
void app_cleanup(AppState* state);
void app_render(AppState* state);
void app_handle_input(AppState* state, KeyEvent* key);
void app_execute_command(AppState* state, const char* command);
void app_start_input(AppState* state, AppMode mode);
void app_finish_input(AppState* state);
void app_cancel_input(AppState* state);
void app_copy_cell(AppState* state);
void app_paste_cell(AppState* state);
void app_copy_to_system_clipboard(AppState* state);
void app_paste_from_system_clipboard(AppState* state);
BOOL set_system_clipboard_text(const char* text);
char* get_system_clipboard_text(void);
void app_update_cursor_blink(AppState* state);
void app_copy_cell(AppState* state);
void app_paste_cell(AppState* state);
void app_copy_to_system_clipboard(AppState* state);
void app_paste_from_system_clipboard(AppState* state);    // Initialize application
void app_init(AppState* state) {
    debug_log("Starting app_init");
    
    debug_log("Creating sheet with 1000x100 dimensions");
    state->sheet = sheet_new(1000, 100);
    if (!state->sheet) {
        debug_log("ERROR: Failed to create sheet");
    } else {
        debug_log("Sheet created successfully");
    }
      debug_log("Initializing console");
    state->console = console_init();
    if (!state->console) {
        debug_log("ERROR: Failed to initialize console");
    } else {
        debug_log("Console initialized successfully, size: %dx%d", 
                 state->console->width, state->console->height);
        
        // Check if console is large enough
        if (state->console->height < 10 || state->console->width < 40) {
            debug_log("Console too small! Minimum size required: 40x10, actual: %dx%d", 
                      state->console->width, state->console->height);
            printf("ERROR: Console window too small!\n");
            printf("Current size: %dx%d\n", state->console->width, state->console->height);
            printf("Minimum required: 40x10\n");
            printf("Please resize your console window and try again.\n");
            console_cleanup(state->console);
            state->console = NULL;
        }
    }
    
    // Check if initialization was successful
    if (!state->sheet || !state->console) {
        debug_log("ERROR: Initialization failed, cleaning up");
        if (state->sheet) {
            sheet_free(state->sheet);
            debug_log("Sheet freed");
        }
        if (state->console) {
            console_cleanup(state->console);
            debug_log("Console cleaned up");
        }
        state->running = FALSE;
        return;
    }
    
    debug_log("Setting initial state");
    state->mode = MODE_NORMAL;
    state->cursor_row = 0;
    state->cursor_col = 0;
    state->view_top = 0;
    state->view_left = 0;
    state->input_buffer[0] = '\0';
    state->input_pos = 0;
    strcpy_s(state->status_message, sizeof(state->status_message), "Ready");
    state->running = TRUE;    
    // Initialize cursor blinking
    debug_log("Initializing cursor blinking");
    state->cursor_blink_time = GetTickCount();
    state->cursor_visible = TRUE;
    state->cursor_blink_rate = 500;  // 500ms blink rate
    
    debug_log("Hiding console cursor");
    console_hide_cursor(state->console);
    
    debug_log("Adding sample data to sheet");
    // Add some sample data
    sheet_set_string(state->sheet, 0, 0, "Windows Terminal Spreadsheet");
    sheet_set_string(state->sheet, 2, 0, "Commands:");
    sheet_set_string(state->sheet, 3, 0, "= - Enter number/formula");
    sheet_set_string(state->sheet, 4, 0, "\" - Enter text");
    sheet_set_string(state->sheet, 5, 0, ": - Command mode");
    sheet_set_string(state->sheet, 6, 0, "hjkl - Navigate");    sheet_set_string(state->sheet, 7, 0, "Ctrl+C - Copy cell (internal)");
    sheet_set_string(state->sheet, 8, 0, "Ctrl+V - Paste cell (internal)");
    sheet_set_string(state->sheet, 9, 0, "Ctrl+Shift+C - Copy to system clipboard");
    sheet_set_string(state->sheet, 10, 0, "Ctrl+Shift+V - Paste from system clipboard");
    sheet_set_string(state->sheet, 11, 0, "Functions: SUM, AVG, MAX, MIN, MEDIAN, MODE");
    sheet_set_string(state->sheet, 12, 0, "IF(condition, true_val, false_val)");
    sheet_set_string(state->sheet, 13, 0, "POWER(base, exponent)");
    sheet_set_string(state->sheet, 14, 0, "Operators: +, -, *, /, >, <, >=, <=, =, <>");
    
    debug_log("app_init completed successfully");
}

// Cleanup
void app_cleanup(AppState* state) {
    if (state->sheet) {
        sheet_free(state->sheet);
        state->sheet = NULL;
    }
    if (state->console) {
        console_cleanup(state->console);
        state->console = NULL;
    }
}

// Update cursor blink state
void app_update_cursor_blink(AppState* state) {
    DWORD current_time = GetTickCount();
    if (current_time - state->cursor_blink_time > state->cursor_blink_rate) {
        state->cursor_visible = !state->cursor_visible;
        state->cursor_blink_time = current_time;
    }
}

// Render the spreadsheet
void app_render(AppState* state) {
    debug_log("Starting app_render");
    Console* con = state->console;
    
    // Safety check
    if (!con || !con->backBuffer) {
        debug_log("ERROR: Console or backBuffer is NULL");
        return;
    }
    
    debug_log("Console size: %dx%d", con->width, con->height);
    
    // Colors
    WORD headerColor = MAKE_COLOR(COLOR_BLACK, COLOR_WHITE);
    WORD cellColor = MAKE_COLOR(COLOR_WHITE, COLOR_BLACK);
    WORD selectedColor = MAKE_COLOR(COLOR_BLACK, COLOR_CYAN);
    WORD gridColor = MAKE_COLOR(COLOR_WHITE | COLOR_BRIGHT, COLOR_BLACK);
    
    debug_log("Clearing back buffer");
    // Clear back buffer
    for (int i = 0; i < con->width * con->height; i++) {
        con->backBuffer[i].Char.AsciiChar = ' ';
        con->backBuffer[i].Attributes = cellColor;
    }
      debug_log("Calculating visible area");
    // Calculate visible area
    int col_header_width = 4;
    int status_height = 2;
    int visible_cols = (con->width - col_header_width) / 10;
    int visible_rows = con->height - status_height - 1;
    
    debug_log("Visible area: %dx%d", visible_cols, visible_rows);
    debug_log("View position: top=%d, left=%d", state->view_top, state->view_left);
    debug_log("Cursor position: row=%d, col=%d", state->cursor_row, state->cursor_col);
    
    // Safety check for minimum viable display area
    if (visible_rows < 1 || visible_cols < 1) {
        debug_log("ERROR: Visible area too small: %dx%d", visible_cols, visible_rows);
        return;
    }
    debug_log("Drawing column headers");
    // Draw column headers
    for (int i = 0; i < visible_cols && state->view_left + i < state->sheet->cols; i++) {
        char colName[10];
        int col = state->view_left + i;
        
        // Convert column number to letters
        if (col < 26) {
            sprintf_s(colName, sizeof(colName), "%c", 'A' + col);
        } else {
            sprintf_s(colName, sizeof(colName), "%c%c", 'A' + (col / 26) - 1, 'A' + (col % 26));
        }
        
        int x = col_header_width + i * 10 + 4;
        console_write_string(con, x, 0, colName, headerColor);
    }
    
    debug_log("Drawing row headers");
    // Draw row headers
    for (int i = 0; i < visible_rows && state->view_top + i < state->sheet->rows; i++) {
        char rowNum[10];
        sprintf_s(rowNum, sizeof(rowNum), "%3d", state->view_top + i + 1);
        console_write_string(con, 0, i + 1, rowNum, headerColor);
    }
    
    debug_log("Drawing grid and cell contents");
    // Draw grid and cell contents
    for (int row = 0; row < visible_rows && state->view_top + row < state->sheet->rows; row++) {
        for (int col = 0; col < visible_cols && state->view_left + col < state->sheet->cols; col++) {
            int x = col_header_width + col * 10;
            int y = row + 1;
            
            // Draw vertical grid line
            console_write_char(con, x, y, '|', gridColor);
            
            // Get cell value
            char* value = sheet_get_display_value(state->sheet, 
                                                  state->view_top + row, 
                                                  state->view_left + col);
            
            // Truncate if too long
            char display[10];
            strncpy_s(display, sizeof(display), value, 9);
            display[9] = '\0';
              // Determine color
            WORD color = cellColor;
            BOOL is_current_cell = (state->view_top + row == state->cursor_row && 
                                    state->view_left + col == state->cursor_col);
            
            if (is_current_cell) {
                // For the current cell, use different colors based on cursor visibility
                if (state->cursor_visible) {
                    color = selectedColor;
                } else {
                    // When cursor is "off", use a dimmed selection color
                    color = MAKE_COLOR(COLOR_WHITE, COLOR_BLUE);
                }
            }
            
            // Draw cell content
            console_write_string(con, x + 1, y, display, color);
            
            // Draw blinking cursor indicator in the current cell
            if (is_current_cell && state->cursor_visible) {
                // Add a cursor indicator at the end of the cell content
                int cursor_x = x + 1 + (int)strlen(display);
                if (cursor_x < x + 9) {  // Make sure we don't go outside cell bounds
                    console_write_char(con, cursor_x, y, '_', selectedColor);
                }
            }
        }    }
    
    debug_log("Drawing horizontal line");
    // Draw horizontal line above status
    int status_y = con->height - status_height;
    for (int x = 0; x < con->width; x++) {
        console_write_char(con, x, status_y, '-', headerColor);
    }    
    
    debug_log("Drawing status line");
    // Draw status line
    char status[256];
    char* cellRef = cell_reference_to_string(state->cursor_row, state->cursor_col);
    Cell* currentCell = sheet_get_cell(state->sheet, state->cursor_row, state->cursor_col);
    
    // Safety check for cellRef
    if (!cellRef) {
        debug_log("WARNING: cellRef is NULL, using default");
        static char defaultRef[] = "A1";
        cellRef = defaultRef;
    }
    
    debug_log("Cell reference: %s", cellRef);
    
    if (state->mode == MODE_NORMAL) {
        if (currentCell && currentCell->type == CELL_FORMULA) {
            sprintf_s(status, sizeof(status), "[%s] %s: %s | %s", 
                    state->sheet->name, cellRef, 
                    currentCell->data.formula.expression,
                    state->status_message);
        } else {
            sprintf_s(status, sizeof(status), "[%s] %s | %s", 
                    state->sheet->name, cellRef, state->status_message);
        }
    } else {
        // Show input buffer with cursor
        char input_with_cursor[300];
        if (state->cursor_visible) {
            // Insert cursor at current input position
            strncpy_s(input_with_cursor, sizeof(input_with_cursor), state->input_buffer, state->input_pos);
            input_with_cursor[state->input_pos] = '\0';
            strcat_s(input_with_cursor, sizeof(input_with_cursor), "_");
            strcat_s(input_with_cursor, sizeof(input_with_cursor), &state->input_buffer[state->input_pos]);
        } else {
            strcpy_s(input_with_cursor, sizeof(input_with_cursor), state->input_buffer);
        }
        
        sprintf_s(status, sizeof(status), "[%s] %s > %s", 
                state->sheet->name, cellRef, input_with_cursor);
    }    
    console_write_string(con, 0, status_y + 1, status, headerColor);
    
    debug_log("Calling console_flip");
    // Update screen
    console_flip(con);
    debug_log("console_flip completed, app_render finished");
}

// Start input mode
void app_start_input(AppState* state, AppMode mode) {
    state->mode = mode;
    state->input_buffer[0] = '\0';
    state->input_pos = 0;
    
    // Use faster blink rate during input
    state->cursor_blink_rate = 300;  // 300ms for faster blinking during input
    state->cursor_visible = TRUE;
    state->cursor_blink_time = GetTickCount();
    
    // Pre-fill with current cell value for editing
    if (mode == MODE_INSERT_FORMULA || mode == MODE_INSERT_NUMBER) {
        Cell* cell = sheet_get_cell(state->sheet, state->cursor_row, state->cursor_col);
        if (cell) {
            if (cell->type == CELL_FORMULA) {
                strcpy_s(state->input_buffer, sizeof(state->input_buffer), cell->data.formula.expression);
            } else if (cell->type == CELL_NUMBER) {
                sprintf_s(state->input_buffer, sizeof(state->input_buffer), "%g", cell->data.number);
            }
            state->input_pos = (int)strlen(state->input_buffer);
        }
    }
}

// Finish input and update cell
void app_finish_input(AppState* state) {
    switch (state->mode) {
        case MODE_INSERT_NUMBER:
        case MODE_INSERT_FORMULA:
            if (state->input_buffer[0] == '=') {
                // Formula
                sheet_set_formula(state->sheet, state->cursor_row, 
                                  state->cursor_col, state->input_buffer);
            } else {
                // Try to parse as number
                char* endptr;
                double value = strtod(state->input_buffer, &endptr);
                if (*endptr == '\0') {
                    sheet_set_number(state->sheet, state->cursor_row, 
                                     state->cursor_col, value);
                } else {
                    // Treat as string if not a valid number
                    sheet_set_string(state->sheet, state->cursor_row, 
                                     state->cursor_col, state->input_buffer);
                }
            }
            sheet_recalculate(state->sheet);
            break;
            
        case MODE_INSERT_STRING:
            sheet_set_string(state->sheet, state->cursor_row, 
                             state->cursor_col, state->input_buffer);
            break;
            
        case MODE_COMMAND:
            app_execute_command(state, state->input_buffer);
            break;
    }
    
    state->mode = MODE_NORMAL;
    // Reset to normal blink rate
    state->cursor_blink_rate = 500;
    state->cursor_visible = TRUE;
    state->cursor_blink_time = GetTickCount();
}

// Cancel input
void app_cancel_input(AppState* state) {
    state->mode = MODE_NORMAL;
    // Reset to normal blink rate
    state->cursor_blink_rate = 500;
    state->cursor_visible = TRUE;
    state->cursor_blink_time = GetTickCount();
    strcpy_s(state->status_message, sizeof(state->status_message), "Cancelled");
}

// Execute command
void app_execute_command(AppState* state, const char* command) {
    if (strcmp(command, "q") == 0 || strcmp(command, "quit") == 0) {
        state->running = FALSE;
    } else if (strncmp(command, "w ", 2) == 0 || strcmp(command, "w") == 0) {
        strcpy_s(state->status_message, sizeof(state->status_message), "Save not implemented yet");
    } else if (strcmp(command, "wq") == 0) {
        strcpy_s(state->status_message, sizeof(state->status_message), "Save not implemented yet");
        state->running = FALSE;
    } else {
        sprintf_s(state->status_message, sizeof(state->status_message), "Unknown command: %s", command);
    }
}

// Copy current cell to clipboard
void app_copy_cell(AppState* state) {
    Cell* cell = sheet_get_cell(state->sheet, state->cursor_row, state->cursor_col);
    sheet_set_clipboard_cell(cell);
    strcpy_s(state->status_message, sizeof(state->status_message), "Cell copied");
}

// Paste clipboard cell to current position
void app_paste_cell(AppState* state) {
    Cell* clipboard = sheet_get_clipboard_cell();
    if (clipboard) {
        sheet_copy_cell(state->sheet, clipboard->row, clipboard->col, 
                       state->cursor_row, state->cursor_col);
        strcpy_s(state->status_message, sizeof(state->status_message), "Cell pasted");
    } else {
        strcpy_s(state->status_message, sizeof(state->status_message), "Nothing to paste");
    }
}

// Copy current cell to system clipboard
void app_copy_to_system_clipboard(AppState* state) {
    Cell* cell = sheet_get_cell(state->sheet, state->cursor_row, state->cursor_col);
    if (cell) {
        char* text = sheet_get_display_value(state->sheet, state->cursor_row, state->cursor_col);
        if (text) {
            set_system_clipboard_text(text);
            strcpy_s(state->status_message, sizeof(state->status_message), "Cell content copied to system clipboard");
        } else {
            strcpy_s(state->status_message, sizeof(state->status_message), "Failed to get cell content");
        }
    } else {
        set_system_clipboard_text(""); // Copy empty string for empty cell
        strcpy_s(state->status_message, sizeof(state->status_message), "Empty cell copied to system clipboard");
    }
}

// Paste from system clipboard to current cell
void app_paste_from_system_clipboard(AppState* state) {
    char* text = get_system_clipboard_text();
    if (text) {
        // Try to determine what type of data this is
        if (strlen(text) == 0) {
            // Empty string - clear cell
            sheet_clear_cell(state->sheet, state->cursor_row, state->cursor_col);
            strcpy_s(state->status_message, sizeof(state->status_message), "Cell cleared from system clipboard");
        } else if (text[0] == '=') {
            // Formula
            sheet_set_formula(state->sheet, state->cursor_row, state->cursor_col, text);
            sheet_recalculate(state->sheet);
            strcpy_s(state->status_message, sizeof(state->status_message), "Formula pasted from system clipboard");
        } else {
            // Try to parse as number
            char* endptr;
            double num = strtod(text, &endptr);
            if (*endptr == '\0' || (*endptr == '\n' && *(endptr+1) == '\0')) {
                // It's a number
                sheet_set_number(state->sheet, state->cursor_row, state->cursor_col, num);
                strcpy_s(state->status_message, sizeof(state->status_message), "Number pasted from system clipboard");
            } else {
                // It's a string - remove trailing newline if present
                size_t len = strlen(text);
                if (len > 0 && text[len-1] == '\n') {
                    text[len-1] = '\0';
                }
                sheet_set_string(state->sheet, state->cursor_row, state->cursor_col, text);
                strcpy_s(state->status_message, sizeof(state->status_message), "Text pasted from system clipboard");
            }
        }
        free(text);
    } else {
        strcpy_s(state->status_message, sizeof(state->status_message), "Failed to get system clipboard content");
    }
}

// System clipboard helper functions
BOOL set_system_clipboard_text(const char* text) {
    if (!OpenClipboard(NULL)) return FALSE;
    
    EmptyClipboard();
    
    size_t len = strlen(text) + 1;
    HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, len);
    if (!hMem) {
        CloseClipboard();
        return FALSE;
    }
    
    char* pMem = (char*)GlobalLock(hMem);
    if (pMem) {
        strcpy_s(pMem, len, text);
        GlobalUnlock(hMem);
        SetClipboardData(CF_TEXT, hMem);
    }
    
    CloseClipboard();
    return TRUE;
}

char* get_system_clipboard_text(void) {
    if (!OpenClipboard(NULL)) return NULL;
    
    HANDLE hData = GetClipboardData(CF_TEXT);
    if (!hData) {
        CloseClipboard();
        return NULL;
    }
    
    char* pData = (char*)GlobalLock(hData);
    if (!pData) {
        CloseClipboard();
        return NULL;
    }
    
    size_t len = strlen(pData) + 1;
    char* result = (char*)malloc(len);
    if (result) {
        strcpy_s(result, len, pData);
    }
    
    GlobalUnlock(hData);
    CloseClipboard();
    return result;
}

// Handle keyboard input
void app_handle_input(AppState* state, KeyEvent* key) {
    if (state->mode == MODE_NORMAL) {
        // Normal mode navigation and commands
        if (key->type == 0) {  // Character key
            switch (key->key.ch) {
                // Navigation
                case 'h':
                    if (state->cursor_col > 0) state->cursor_col--;
                    break;
                case 'l':
                    if (state->cursor_col < state->sheet->cols - 1) state->cursor_col++;
                    break;
                case 'j':
                    if (state->cursor_row < state->sheet->rows - 1) state->cursor_row++;
                    break;
                case 'k':
                    if (state->cursor_row > 0) state->cursor_row--;
                    break;
                    
                // Commands
                case '=':
                    app_start_input(state, MODE_INSERT_FORMULA);
                    break;
                case '"':
                    app_start_input(state, MODE_INSERT_STRING);
                    break;                case ':':
                    app_start_input(state, MODE_COMMAND);
                    break;
                case 'x':
                    sheet_clear_cell(state->sheet, state->cursor_row, state->cursor_col);
                    sheet_recalculate(state->sheet);
                    strcpy_s(state->status_message, sizeof(state->status_message), "Cell cleared");
                    break;
                case 'c':
                    if (key->ctrl && key->shift) {
                        app_copy_to_system_clipboard(state);
                    } else if (key->ctrl) {
                        app_copy_cell(state);
                    }
                    break;
                case 'v':
                    if (key->ctrl && key->shift) {
                        app_paste_from_system_clipboard(state);
                    } else if (key->ctrl) {
                        app_paste_cell(state);
                    }
                    break;
                
                // Quick quit
                case 'q':
                    if (key->ctrl) {
                        state->running = FALSE;
                    }
                    break;
            }
        } else {  // Special key
            switch (key->key.special) {
                case KEY_LEFT:
                    if (state->cursor_col > 0) state->cursor_col--;
                    break;
                case KEY_RIGHT:
                    if (state->cursor_col < state->sheet->cols - 1) state->cursor_col++;
                    break;
                case KEY_UP:
                    if (state->cursor_row > 0) state->cursor_row--;
                    break;
                case KEY_DOWN:
                    if (state->cursor_row < state->sheet->rows - 1) state->cursor_row++;
                    break;
                case KEY_PGUP:
                    state->cursor_row = max(0, state->cursor_row - 10);
                    break;
                case KEY_PGDN:
                    state->cursor_row = min(state->sheet->rows - 1, state->cursor_row + 10);
                    break;
            }
        }
        
        // Adjust view if cursor moved outside
        int visible_rows = state->console->height - 3;
        int visible_cols = (state->console->width - 4) / 10;
        
        if (state->cursor_row < state->view_top) {
            state->view_top = state->cursor_row;
        } else if (state->cursor_row >= state->view_top + visible_rows) {
            state->view_top = state->cursor_row - visible_rows + 1;
        }
        
        if (state->cursor_col < state->view_left) {
            state->view_left = state->cursor_col;
        } else if (state->cursor_col >= state->view_left + visible_cols) {
            state->view_left = state->cursor_col - visible_cols + 1;
        }
    } else {
        // Input mode
        if (key->type == 0) {
            switch (key->key.ch) {
                case KEY_ENTER:
                    app_finish_input(state);
                    break;
                case KEY_ESC:
                    app_cancel_input(state);
                    break;
                case KEY_BACKSPACE:
                    if (state->input_pos > 0) {
                        state->input_pos--;
                        state->input_buffer[state->input_pos] = '\0';
                    }
                    break;
                default:
                    if (isprint(key->key.ch) && state->input_pos < 255) {
                        state->input_buffer[state->input_pos++] = key->key.ch;
                        state->input_buffer[state->input_pos] = '\0';
                    }
                    break;
            }
        }
    }
}

// Main program
int main() {
    debug_init();
    debug_log("=== Starting WinSpread ===");
    
    AppState state;
    
    // Initialize
    debug_log("Calling app_init");
    app_init(&state);
    
    // Check if initialization failed
    if (!state.running) {
        debug_log("ERROR: Initialization failed");
        printf("Failed to initialize application\n");
        debug_cleanup();
        return 1;
    }
    
    debug_log("Entering main loop");
    // Main loop
    while (state.running) {
        debug_log("Main loop iteration - updating cursor blink");
        // Update cursor blinking
        app_update_cursor_blink(&state);
        
        debug_log("Calling app_render");
        app_render(&state);
        debug_log("app_render completed");
        
        debug_log("Checking for key input");
        KeyEvent key;
        if (console_get_key(state.console, &key)) {
            debug_log("Key pressed, handling input");
            app_handle_input(&state, &key);
            
            // Reset cursor to visible when user interacts
            state.cursor_visible = TRUE;
            state.cursor_blink_time = GetTickCount();
            debug_log("Input handled");
        }
        
        Sleep(16);  // ~60 FPS
    }
    
    debug_log("Exiting main loop, cleaning up");
    app_cleanup(&state);
    debug_log("=== WinSpread Ended ===");
    debug_cleanup();
    return 0;
}
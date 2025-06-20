// main.c - Enhanced WinSpread with range selection, formatting, and VLOOKUP
#include <windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>

#include "console.h"
#include "sheet.h"
#include "debug.h"

// Application state
typedef enum {
    MODE_NORMAL,
    MODE_INSERT_NUMBER,
    MODE_INSERT_STRING,
    MODE_INSERT_FORMULA,
    MODE_COMMAND,
    MODE_RANGE_SELECT  // NEW: Range selection mode
} AppMode;

typedef struct {
    Sheet* sheet;
    Console* console;
    AppMode mode;
    int cursor_row;
    int cursor_col;
    int view_top;
    int view_left;
    char input_buffer[256];
    int input_pos;
    char status_message[256];
    BOOL running;
    
    // Cursor blinking state
    DWORD cursor_blink_time;
    BOOL cursor_visible;
    DWORD cursor_blink_rate;
    
    // NEW: Range selection state
    BOOL range_selection_active;
    int range_start_row;
    int range_start_col;
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
void app_update_cursor_blink(AppState* state);

// NEW: Range selection functions
void app_start_range_selection(AppState* state);
void app_extend_range_selection(AppState* state, int row, int col);
void app_finish_range_selection(AppState* state);
void app_cancel_range_selection(AppState* state);

// Copy/paste operations
void app_copy_cell(AppState* state);
void app_paste_cell(AppState* state);
void app_copy_to_system_clipboard(AppState* state);
void app_paste_from_system_clipboard(AppState* state);

// NEW: Range copy/paste operations
void app_copy_range(AppState* state);
void app_paste_range(AppState* state);

// NEW: Formatting functions
void app_set_cell_format(AppState* state, DataFormat format, FormatStyle style);
void app_cycle_date_format(AppState* state);
void app_cycle_datetime_format(AppState* state);

// System clipboard functions
BOOL set_system_clipboard_text(const char* text);
char* get_system_clipboard_text(void);

// Initialize application
void app_init(AppState* state) {
    debug_log("Starting app_init");
    
    state->sheet = sheet_new(1000, 100);
    if (!state->sheet) {
        debug_log("ERROR: Failed to create sheet");
    }
    
    state->console = console_init();
    if (!state->console) {
        debug_log("ERROR: Failed to initialize console");
    }
    
    if (!state->sheet || !state->console) {
        if (state->sheet) sheet_free(state->sheet);
        if (state->console) console_cleanup(state->console);
        state->running = FALSE;
        return;
    }
    
    state->mode = MODE_NORMAL;
    state->cursor_row = 0;
    state->cursor_col = 0;
    state->view_top = 0;
    state->view_left = 0;
    state->input_buffer[0] = '\0';
    state->input_pos = 0;
    strcpy_s(state->status_message, sizeof(state->status_message), "Ready");
    state->running = TRUE;
    
    // NEW: Initialize range selection
    state->range_selection_active = FALSE;
    state->range_start_row = 0;
    state->range_start_col = 0;
    
    state->cursor_blink_time = GetTickCount();
    state->cursor_visible = TRUE;
    state->cursor_blink_rate = 500;
    
    console_hide_cursor(state->console);
    
    // Add enhanced sample data with formatting examples
    sheet_set_string(state->sheet, 0, 0, "Enhanced WinSpread Features");
    
    sheet_set_string(state->sheet, 2, 0, "NEW FEATURES:");
    sheet_set_string(state->sheet, 3, 0, "Range Selection: Shift+arrows");
    sheet_set_string(state->sheet, 4, 0, "Range Copy/Paste: Ctrl+C/V on ranges");
    sheet_set_string(state->sheet, 5, 0, "Cell Formatting: :format commands");
    sheet_set_string(state->sheet, 6, 0, "VLOOKUP function support");
    
    sheet_set_string(state->sheet, 8, 0, "Formatting Examples:");
    
    // Number formatting examples
    sheet_set_string(state->sheet, 9, 0, "Percentage:");
    sheet_set_number(state->sheet, 9, 1, 0.1234);
    Cell* pct_cell = sheet_get_or_create_cell(state->sheet, 9, 1);
    cell_set_format(pct_cell, FORMAT_PERCENTAGE, 0);
    
    sheet_set_string(state->sheet, 10, 0, "Currency:");
    sheet_set_number(state->sheet, 10, 1, 1234.56);
    Cell* curr_cell = sheet_get_or_create_cell(state->sheet, 10, 1);
    cell_set_format(curr_cell, FORMAT_CURRENCY, 0);
    
    sheet_set_string(state->sheet, 11, 0, "Date:");
    sheet_set_number(state->sheet, 11, 1, 45000); // Excel date serial
    Cell* date_cell = sheet_get_or_create_cell(state->sheet, 11, 1);
    cell_set_format(date_cell, FORMAT_DATE, DATE_STYLE_MM_DD_YYYY);
    
    sheet_set_string(state->sheet, 12, 0, "Time:");
    sheet_set_number(state->sheet, 12, 1, 0.5); // 12:00 PM
    Cell* time_cell = sheet_get_or_create_cell(state->sheet, 12, 1);
    cell_set_format(time_cell, FORMAT_TIME, TIME_STYLE_12HR);
    
    // VLOOKUP example
    sheet_set_string(state->sheet, 14, 0, "VLOOKUP Example:");
    sheet_set_string(state->sheet, 15, 0, "Product");
    sheet_set_string(state->sheet, 15, 1, "Price");
    sheet_set_string(state->sheet, 16, 0, "Apple");
    sheet_set_number(state->sheet, 16, 1, 0.50);
    sheet_set_string(state->sheet, 17, 0, "Orange");
    sheet_set_number(state->sheet, 17, 1, 0.75);
    sheet_set_string(state->sheet, 18, 0, "Banana");
    sheet_set_number(state->sheet, 18, 1, 0.30);
    
    sheet_set_string(state->sheet, 20, 0, "Lookup 'Orange':");
    sheet_set_formula(state->sheet, 20, 1, "=VLOOKUP(\"Orange\",A16:B18,2,1)");
      // Commands help
    sheet_set_string(state->sheet, 22, 0, "Format Commands:");
    sheet_set_string(state->sheet, 23, 0, ":format percentage");
    sheet_set_string(state->sheet, 24, 0, ":format currency");
    sheet_set_string(state->sheet, 25, 0, ":format date");
    sheet_set_string(state->sheet, 26, 0, ":format time");
    sheet_set_string(state->sheet, 27, 0, ":format general");
    
    // NEW: Color and resize commands help
    sheet_set_string(state->sheet, 29, 0, "Color Commands:");
    sheet_set_string(state->sheet, 30, 0, ":clrtx red (or #FF0000)");
    sheet_set_string(state->sheet, 31, 0, ":clrbg blue (or #0000FF)");
    
    sheet_set_string(state->sheet, 33, 0, "Resize Commands:");
    sheet_set_string(state->sheet, 34, 0, "Alt+Left/Right: Column width");
    sheet_set_string(state->sheet, 35, 0, "Alt+Up/Down: Row height");
    sheet_set_string(state->sheet, 36, 0, "Works with range selection!");
    
    sheet_set_string(state->sheet, 38, 0, "Colors: black, blue, green, cyan");
    sheet_set_string(state->sheet, 39, 0, "        red, magenta, yellow, white");
    
    sheet_recalculate(state->sheet);
    debug_log("app_init completed successfully");
}

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

void app_update_cursor_blink(AppState* state) {
    DWORD current_time = GetTickCount();
    if (current_time - state->cursor_blink_time > state->cursor_blink_rate) {
        state->cursor_visible = !state->cursor_visible;
        state->cursor_blink_time = current_time;
    }
}

// NEW: Start range selection
void app_start_range_selection(AppState* state) {
    state->range_selection_active = TRUE;
    state->range_start_row = state->cursor_row;
    state->range_start_col = state->cursor_col;
    sheet_start_range_selection(state->sheet, state->cursor_row, state->cursor_col);
    strcpy_s(state->status_message, sizeof(state->status_message), "Range selection started");
}

// NEW: Extend range selection
void app_extend_range_selection(AppState* state, int row, int col) {
    if (state->range_selection_active) {
        sheet_extend_range_selection(state->sheet, row, col);
        
        char* start_ref = cell_reference_to_string(state->range_start_row, state->range_start_col);
        char* end_ref = cell_reference_to_string(row, col);
        sprintf_s(state->status_message, sizeof(state->status_message), 
                 "Selected: %s:%s", start_ref, end_ref);
    }
}

// NEW: Finish range selection
void app_finish_range_selection(AppState* state) {
    state->range_selection_active = FALSE;
    strcpy_s(state->status_message, sizeof(state->status_message), "Range selected");
}

// NEW: Cancel range selection
void app_cancel_range_selection(AppState* state) {
    state->range_selection_active = FALSE;
    sheet_clear_range_selection(state->sheet);
    strcpy_s(state->status_message, sizeof(state->status_message), "Range selection cancelled");
}

// Render the spreadsheet
void app_render(AppState* state) {
    Console* con = state->console;
    
    if (!con || !con->backBuffer) {
        return;
    }
    
    // Colors
    WORD headerColor = MAKE_COLOR(COLOR_BLACK, COLOR_WHITE);
    WORD cellColor = MAKE_COLOR(COLOR_WHITE, COLOR_BLACK);
    WORD selectedColor = MAKE_COLOR(COLOR_BLACK, COLOR_CYAN);
    WORD gridColor = MAKE_COLOR(COLOR_WHITE | COLOR_BRIGHT, COLOR_BLACK);
    WORD rangeColor = MAKE_COLOR(COLOR_BLACK, COLOR_YELLOW);  // NEW: Range selection color
    
    // Clear back buffer
    for (int i = 0; i < con->width * con->height; i++) {
        con->backBuffer[i].Char.AsciiChar = ' ';
        con->backBuffer[i].Attributes = cellColor;
    }
      // Calculate visible area (with dynamic column widths)
    int col_header_width = 4;
    int status_height = 2;
    int visible_cols = 0;
    int visible_rows = con->height - status_height - 1;
    
    // Calculate how many columns fit in the screen width
    int total_width = col_header_width;
    for (int i = 0; i < state->sheet->cols && total_width < con->width; i++) {
        int col_width = sheet_get_column_width(state->sheet, state->view_left + i);
        if (total_width + col_width <= con->width) {
            total_width += col_width;
            visible_cols++;
        } else {
            break;
        }
    }
    
    if (visible_rows < 1 || visible_cols < 1) {
        return;
    }
    
    // Draw column headers (with dynamic widths)
    int current_x = col_header_width;
    for (int i = 0; i < visible_cols && state->view_left + i < state->sheet->cols; i++) {
        char colName[10];
        int col = state->view_left + i;
        int col_width = sheet_get_column_width(state->sheet, col);
        
        if (col < 26) {
            sprintf_s(colName, sizeof(colName), "%c", 'A' + col);
        } else {
            sprintf_s(colName, sizeof(colName), "%c%c", 'A' + (col / 26) - 1, 'A' + (col % 26));
        }
        
        // Center the column name in the column width
        int center_x = current_x + (col_width / 2) - ((int)strlen(colName) / 2);
        console_write_string(con, center_x, 0, colName, headerColor);
        current_x += col_width;
    }
      // Draw row headers (accounting for variable row heights)
    int visual_row = 0;
    for (int sheet_row = state->view_top; sheet_row < state->sheet->rows && visual_row < visible_rows; sheet_row++) {
        int row_height = sheet_get_row_height(state->sheet, sheet_row);
        
        // Draw row number only on the first line of each row
        char rowNum[10];
        sprintf_s(rowNum, sizeof(rowNum), "%3d", sheet_row + 1);
        console_write_string(con, 0, visual_row + 1, rowNum, headerColor);
        
        // Skip visual rows for the height of this row
        visual_row += row_height;
    }
    
    // Draw grid and cell contents (with dynamic sizes)
    current_x = col_header_width;
    for (int row = 0; row < visible_rows && state->view_top + row < state->sheet->rows; row++) {
        int sheet_row = state->view_top + row;
        int row_height = sheet_get_row_height(state->sheet, sheet_row);
        
        // Draw multiple lines for tall rows
        for (int row_line = 0; row_line < row_height && row + row_line < visible_rows; row_line++) {
            int y = row + 1 + row_line;
            current_x = col_header_width;
            
            for (int col = 0; col < visible_cols && state->view_left + col < state->sheet->cols; col++) {
                int sheet_col = state->view_left + col;
                int col_width = sheet_get_column_width(state->sheet, sheet_col);
                
                // Draw vertical grid line
                console_write_char(con, current_x, y, '|', gridColor);
                
                // Only draw content on the first line of each cell
                if (row_line == 0) {
                    // Get cell value
                    char* value = sheet_get_display_value(state->sheet, sheet_row, sheet_col);
                    
                    // Truncate if too long for column width
                    char display[51]; // Max column width
                    int max_len = col_width - 1;
                    if (max_len > 50) max_len = 50;
                    strncpy_s(display, sizeof(display), value, max_len);
                    display[max_len] = '\0';
                    
                    // Determine color
                    WORD color = cellColor;
                    BOOL is_current_cell = (sheet_row == state->cursor_row && sheet_col == state->cursor_col);
                    
                    // NEW: Check if cell is in range selection
                    BOOL is_in_range = sheet_is_in_selection(state->sheet, sheet_row, sheet_col);
                    
                    // NEW: Get cell for color formatting
                    Cell* cell = sheet_get_cell(state->sheet, sheet_row, sheet_col);
                    if (cell && (cell->text_color >= 0 || cell->background_color >= 0)) {
                        // Apply custom colors
                        int fg = (cell->text_color >= 0) ? cell->text_color : COLOR_WHITE;
                        int bg = (cell->background_color >= 0) ? cell->background_color : COLOR_BLACK;
                        color = MAKE_COLOR(fg, bg);
                    }
                      if (is_in_range) {
                        // Range selection takes priority, but distinguish current cell in range
                        if (is_current_cell) {
                            // Current cell in range gets a special highlight
                            color = MAKE_COLOR(COLOR_YELLOW, COLOR_BLUE);
                        } else {
                            color = rangeColor;  // Normal range selection color
                        }
                    } else if (is_current_cell) {
                        if (state->cursor_visible) {
                            color = selectedColor;
                        } else {
                            color = MAKE_COLOR(COLOR_WHITE, COLOR_BLUE);
                        }
                    }
                    
                    // Draw cell content
                    console_write_string(con, current_x + 1, y, display, color);
                    
                    // Draw blinking cursor indicator in the current cell
                    if (is_current_cell && state->cursor_visible) {
                        int cursor_x = current_x + 1 + (int)strlen(display);
                        if (cursor_x < current_x + col_width) {
                            console_write_char(con, cursor_x, y, '_', selectedColor);
                        }
                    }
                }
                
                current_x += col_width;
                if (current_x >= con->width) break;
            }
        }
        
        // Skip additional rows if this row was taller than 1
        if (row_height > 1) {
            row += row_height - 1;
        }
    }
    
    // Draw horizontal line above status
    int status_y = con->height - status_height;
    for (int x = 0; x < con->width; x++) {
        console_write_char(con, x, status_y, '-', headerColor);
    }
    
    // Draw status line
    char status[256];
    char* cellRef = cell_reference_to_string(state->cursor_row, state->cursor_col);
    Cell* currentCell = sheet_get_cell(state->sheet, state->cursor_row, state->cursor_col);
    
    if (!cellRef) {
        static char defaultRef[] = "A1";
        cellRef = defaultRef;
    }
    
    if (state->mode == MODE_NORMAL) {
        if (currentCell && currentCell->type == CELL_FORMULA) {
            sprintf_s(status, sizeof(status), "[%s] %s: %s | %s", 
                    state->sheet->name, cellRef, 
                    currentCell->data.formula.expression,
                    state->status_message);
        } else {
            // NEW: Show cell formatting info
            if (currentCell && currentCell->format != FORMAT_GENERAL) {
                const char* format_name = "General";
                switch (currentCell->format) {
                    case FORMAT_PERCENTAGE: format_name = "Percentage"; break;
                    case FORMAT_CURRENCY: format_name = "Currency"; break;
                    case FORMAT_DATE: format_name = "Date"; break;
                    case FORMAT_TIME: format_name = "Time"; break;
                    case FORMAT_DATETIME: format_name = "DateTime"; break;
                    case FORMAT_NUMBER: format_name = "Number"; break;
                    default: format_name = "General"; break;
                }
                sprintf_s(status, sizeof(status), "[%s] %s (%s) | %s", 
                        state->sheet->name, cellRef, format_name, state->status_message);
            } else {
                sprintf_s(status, sizeof(status), "[%s] %s | %s", 
                        state->sheet->name, cellRef, state->status_message);
            }
        }
    } else if (state->mode == MODE_COMMAND && strlen(state->input_buffer) == 0 && strlen(state->status_message) > 0) {
        sprintf_s(status, sizeof(status), "%s", state->status_message);
    } else {
        // Show input buffer with cursor
        char input_with_cursor[300];
        if (state->cursor_visible) {
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
    
    console_flip(con);
}

// Start input mode
void app_start_input(AppState* state, AppMode mode) {
    state->mode = mode;
    state->input_buffer[0] = '\0';
    state->input_pos = 0;
    
    state->cursor_blink_rate = 300;
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

void app_finish_input(AppState* state) {
    switch (state->mode) {
        case MODE_INSERT_NUMBER:
        case MODE_INSERT_FORMULA:
            if (state->input_buffer[0] == '=') {
                sheet_set_formula(state->sheet, state->cursor_row, 
                                  state->cursor_col, state->input_buffer);
            } else {
                char* endptr;
                double value = strtod(state->input_buffer, &endptr);
                if (*endptr == '\0') {
                    sheet_set_number(state->sheet, state->cursor_row, 
                                     state->cursor_col, value);
                } else {
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
    state->cursor_blink_rate = 500;
    state->cursor_visible = TRUE;
    state->cursor_blink_time = GetTickCount();
}

void app_cancel_input(AppState* state) {
    state->mode = MODE_NORMAL;
    state->cursor_blink_rate = 500;
    state->cursor_visible = TRUE;
    state->cursor_blink_time = GetTickCount();
    strcpy_s(state->status_message, sizeof(state->status_message), "Cancelled");
}

// NEW: Set cell formatting
void app_set_cell_format(AppState* state, DataFormat format, FormatStyle style) {
    Cell* cell = sheet_get_or_create_cell(state->sheet, state->cursor_row, state->cursor_col);
    if (cell) {
        cell_set_format(cell, format, style);
        
        // Update status message based on format type
        switch (format) {
            case FORMAT_PERCENTAGE:
                strcpy_s(state->status_message, sizeof(state->status_message), "Cell formatted as percentage");
                break;
            case FORMAT_CURRENCY:
                strcpy_s(state->status_message, sizeof(state->status_message), "Cell formatted as currency");
                break;
            case FORMAT_DATE:
                strcpy_s(state->status_message, sizeof(state->status_message), "Cell formatted as date");
                break;
            case FORMAT_TIME:
                strcpy_s(state->status_message, sizeof(state->status_message), "Cell formatted as time");
                break;
            case FORMAT_NUMBER:
                strcpy_s(state->status_message, sizeof(state->status_message), "Cell formatted as number");
                break;
            case FORMAT_GENERAL:
            default:
                strcpy_s(state->status_message, sizeof(state->status_message), "Cell formatting reset to general");
                break;
        }    } else {
        strcpy_s(state->status_message, sizeof(state->status_message), "Failed to format cell");
    }
}

// NEW: Enhanced function to cycle through comprehensive date/time formats
void app_cycle_datetime_format(AppState* state) {
    Cell* cell = sheet_get_or_create_cell(state->sheet, state->cursor_row, state->cursor_col);
    if (cell) {
        // Define the cycling order through different formats
        typedef struct {
            DataFormat format;
            FormatStyle style;
            const char* description;
        } FormatOption;
        
        static const FormatOption format_cycle[] = {
            // Date formats
            {FORMAT_DATE, DATE_STYLE_MM_DD_YYYY, "Date format: MM/DD/YYYY"},
            {FORMAT_DATE, DATE_STYLE_DD_MM_YYYY, "Date format: DD/MM/YYYY"},
            {FORMAT_DATE, DATE_STYLE_YYYY_MM_DD, "Date format: YYYY-MM-DD"},
            {FORMAT_DATE, DATE_STYLE_SHORT_DATE, "Date format: MM/DD/YY"},
            {FORMAT_DATE, DATE_STYLE_MON_DD_YYYY, "Date format: Mon DD, YYYY"},
            {FORMAT_DATE, DATE_STYLE_DD_MON_YYYY, "Date format: DD Mon YYYY"},
            // Time formats
            {FORMAT_TIME, TIME_STYLE_12HR, "Time format: 12-hour"},
            {FORMAT_TIME, TIME_STYLE_24HR, "Time format: 24-hour"},
            {FORMAT_TIME, TIME_STYLE_SECONDS, "Time format: with seconds"},
            {FORMAT_TIME, TIME_STYLE_12HR_SECONDS, "Time format: 12-hour with seconds"},
            // DateTime formats
            {FORMAT_DATETIME, DATETIME_STYLE_SHORT, "DateTime format: Short"},
            {FORMAT_DATETIME, DATETIME_STYLE_LONG, "DateTime format: Long"},
            {FORMAT_DATETIME, DATETIME_STYLE_ISO, "DateTime format: ISO 8601"}
        };
        
        const int cycle_length = sizeof(format_cycle) / sizeof(format_cycle[0]);
        int current_index = -1;
        
        // Find current format in the cycle
        for (int i = 0; i < cycle_length; i++) {
            if (cell->format == format_cycle[i].format && 
                cell->format_style == format_cycle[i].style) {
                current_index = i;
                break;
            }
        }
        
        // Move to next format in cycle (or start at beginning)
        int next_index = (current_index + 1) % cycle_length;
        
        // Apply the new format
        cell_set_format(cell, format_cycle[next_index].format, format_cycle[next_index].style);
        
        // Update status message
        strcpy_s(state->status_message, sizeof(state->status_message), format_cycle[next_index].description);
    } else {
        strcpy_s(state->status_message, sizeof(state->status_message), "Failed to format cell");
    }
}

// NEW: Function to cycle through date formats
void app_cycle_date_format(AppState* state) {
    Cell* cell = sheet_get_or_create_cell(state->sheet, state->cursor_row, state->cursor_col);
    if (cell) {
        // Get current format style or start with first style
        FormatStyle current_style = cell->format_style;
        FormatStyle next_style;
        
        // If cell is not currently a date format, start with first date style
        if (cell->format != FORMAT_DATE) {
            next_style = DATE_STYLE_MM_DD_YYYY;
        } else {
            // Cycle through date styles
            switch (current_style) {
                case DATE_STYLE_MM_DD_YYYY:
                    next_style = DATE_STYLE_DD_MM_YYYY;
                    break;
                case DATE_STYLE_DD_MM_YYYY:
                    next_style = DATE_STYLE_YYYY_MM_DD;
                    break;
                case DATE_STYLE_YYYY_MM_DD:
                default:
                    next_style = DATE_STYLE_MM_DD_YYYY;
                    break;
            }
        }
        
        // Apply the new date format
        cell_set_format(cell, FORMAT_DATE, next_style);
        
        // Update status message based on the new style
        switch (next_style) {
            case DATE_STYLE_MM_DD_YYYY:
                strcpy_s(state->status_message, sizeof(state->status_message), "Date format: MM/DD/YYYY");
                break;
            case DATE_STYLE_DD_MM_YYYY:
                strcpy_s(state->status_message, sizeof(state->status_message), "Date format: DD/MM/YYYY");
                break;
            case DATE_STYLE_YYYY_MM_DD:
                strcpy_s(state->status_message, sizeof(state->status_message), "Date format: YYYY-MM-DD");
                break;
            default:
                strcpy_s(state->status_message, sizeof(state->status_message), "Date format applied");
                break;
        }
    } else {
        strcpy_s(state->status_message, sizeof(state->status_message), "Failed to format cell");
    }
}

int ask_preserve_formulas(AppState* state, const char* operation) {
    AppMode old_mode = state->mode;
    char old_status[256];
    strcpy_s(old_status, sizeof(old_status), state->status_message);
    
    state->mode = MODE_COMMAND;
    
    sprintf_s(state->status_message, sizeof(state->status_message), 
              "%s: Type 'f' to flatten (save calculated values) or 'p' to preserve (save formulas as text): ", operation);
    
    state->input_buffer[0] = '\0';
    state->input_pos = 0;
    state->cursor_blink_rate = 300;
    state->cursor_visible = TRUE;
    state->cursor_blink_time = GetTickCount();
    
    app_render(state);
    
    KeyEvent key;
    int result = -1;
    while (state->running) {
        if (console_get_key(state->console, &key)) {
            if (key.type == 1 && key.key.special == VK_ESCAPE) {
                strcpy_s(state->status_message, sizeof(state->status_message), "Cancelled");
                result = -1;
                break;
            } else if (key.type == 0) {
                if (key.key.ch == 'f' || key.key.ch == 'F') {
                    result = 0;
                    break;
                } else if (key.key.ch == 'p' || key.key.ch == 'P') {
                    result = 1;
                    break;
                }
            }
        }
        Sleep(10);
    }
    
    state->mode = old_mode;
    state->input_buffer[0] = '\0';
    state->input_pos = 0;
    state->cursor_blink_rate = 500;
    state->cursor_visible = TRUE;
    state->cursor_blink_time = GetTickCount();
    
    if (result != -1) {
        strcpy_s(state->status_message, sizeof(state->status_message), old_status);
    }
    
    return result;
}

// Execute command
void app_execute_command(AppState* state, const char* command) {
    if (strcmp(command, "q") == 0 || strcmp(command, "quit") == 0) {
        state->running = FALSE;
    } else if (strncmp(command, "savecsv ", 8) == 0) {
        const char* filename = command + 8;
        if (strlen(filename) == 0) {
            strcpy_s(state->status_message, sizeof(state->status_message), "Usage: savecsv <filename>");
            return;
        }
        
        int preserve = ask_preserve_formulas(state, "Save CSV");
        if (preserve == -1) return;
        
        if (sheet_save_csv(state->sheet, filename, preserve)) {
            sprintf_s(state->status_message, sizeof(state->status_message), 
                     "Saved to %s (%s)", filename, preserve ? "formulas preserved" : "values flattened");
        } else {
            sprintf_s(state->status_message, sizeof(state->status_message), 
                     "Failed to save %s", filename);
        }
    } else if (strncmp(command, "loadcsv ", 8) == 0) {
        const char* filename = command + 8;
        if (strlen(filename) == 0) {
            strcpy_s(state->status_message, sizeof(state->status_message), "Usage: loadcsv <filename>");
            return;
        }
        
        int preserve = ask_preserve_formulas(state, "Load CSV");
        if (preserve == -1) return;
        
        if (sheet_load_csv(state->sheet, filename, preserve)) {
            sprintf_s(state->status_message, sizeof(state->status_message), 
                     "Loaded from %s (%s)", filename, preserve ? "formulas preserved" : "values only");
        } else {
            sprintf_s(state->status_message, sizeof(state->status_message), 
                     "Failed to load %s", filename);
        }
    } 
    // NEW: Formatting commands
    else if (strcmp(command, "format percentage") == 0) {
        app_set_cell_format(state, FORMAT_PERCENTAGE, 0);
    } else if (strcmp(command, "format currency") == 0) {
        app_set_cell_format(state, FORMAT_CURRENCY, 0);
    } else if (strcmp(command, "format date") == 0) {
        app_set_cell_format(state, FORMAT_DATE, DATE_STYLE_MM_DD_YYYY);
    } else if (strcmp(command, "format date dd/mm/yyyy") == 0) {
        app_set_cell_format(state, FORMAT_DATE, DATE_STYLE_DD_MM_YYYY);
    } else if (strcmp(command, "format date yyyy-mm-dd") == 0) {
        app_set_cell_format(state, FORMAT_DATE, DATE_STYLE_YYYY_MM_DD);
    } else if (strcmp(command, "format time") == 0) {
        app_set_cell_format(state, FORMAT_TIME, TIME_STYLE_12HR);
    } else if (strcmp(command, "format time 24hr") == 0) {
        app_set_cell_format(state, FORMAT_TIME, TIME_STYLE_24HR);
    } else if (strcmp(command, "format time seconds") == 0) {
        app_set_cell_format(state, FORMAT_TIME, TIME_STYLE_SECONDS);
    } else if (strcmp(command, "format datetime") == 0) {
        app_set_cell_format(state, FORMAT_DATETIME, 0);
    } else if (strcmp(command, "format general") == 0 || strcmp(command, "format number") == 0) {
        app_set_cell_format(state, FORMAT_GENERAL, 0);
    } 
    // NEW: Range formatting commands
    else if (strncmp(command, "range format ", 13) == 0) {
        if (!state->sheet->selection.is_active) {
            strcpy_s(state->status_message, sizeof(state->status_message), "No range selected");
            return;
        }
        
        const char* format_type = command + 13;
        DataFormat format = FORMAT_GENERAL;
        FormatStyle style = 0;
        
        if (strcmp(format_type, "percentage") == 0) {
            format = FORMAT_PERCENTAGE;
        } else if (strcmp(format_type, "currency") == 0) {
            format = FORMAT_CURRENCY;
        } else if (strcmp(format_type, "date") == 0) {
            format = FORMAT_DATE;
            style = DATE_STYLE_MM_DD_YYYY;
        } else if (strcmp(format_type, "time") == 0) {
            format = FORMAT_TIME;
            style = TIME_STYLE_12HR;
        } else if (strcmp(format_type, "general") == 0) {
            format = FORMAT_GENERAL;
        }
        
        // Apply formatting to selected range
        int min_row = state->sheet->selection.start_row < state->sheet->selection.end_row ? 
                      state->sheet->selection.start_row : state->sheet->selection.end_row;
        int max_row = state->sheet->selection.start_row > state->sheet->selection.end_row ? 
                      state->sheet->selection.start_row : state->sheet->selection.end_row;
        int min_col = state->sheet->selection.start_col < state->sheet->selection.end_col ? 
                      state->sheet->selection.start_col : state->sheet->selection.end_col;
        int max_col = state->sheet->selection.start_col > state->sheet->selection.end_col ? 
                      state->sheet->selection.start_col : state->sheet->selection.end_col;
        
        for (int row = min_row; row <= max_row; row++) {
            for (int col = min_col; col <= max_col; col++) {
                Cell* cell = sheet_get_or_create_cell(state->sheet, row, col);
                if (cell) {
                    cell_set_format(cell, format, style);
                }
            }
        }
          sprintf_s(state->status_message, sizeof(state->status_message), 
                 "Range formatted as %s", format_type);
    } 
    // NEW: Color commands for text
    else if (strncmp(command, "clrtx ", 6) == 0) {
        const char* color_str = command + 6;
        int color = parse_color(color_str);
        if (color >= 0) {
            if (state->sheet->selection.is_active) {
                // Apply to range
                int min_row = state->sheet->selection.start_row < state->sheet->selection.end_row ? 
                              state->sheet->selection.start_row : state->sheet->selection.end_row;
                int max_row = state->sheet->selection.start_row > state->sheet->selection.end_row ? 
                              state->sheet->selection.start_row : state->sheet->selection.end_row;
                int min_col = state->sheet->selection.start_col < state->sheet->selection.end_col ? 
                              state->sheet->selection.start_col : state->sheet->selection.end_col;
                int max_col = state->sheet->selection.start_col > state->sheet->selection.end_col ? 
                              state->sheet->selection.start_col : state->sheet->selection.end_col;
                
                for (int row = min_row; row <= max_row; row++) {
                    for (int col = min_col; col <= max_col; col++) {
                        Cell* cell = sheet_get_or_create_cell(state->sheet, row, col);
                        if (cell) {
                            cell_set_text_color(cell, color);
                        }
                    }
                }
                sprintf_s(state->status_message, sizeof(state->status_message), 
                         "Range text color set to %s", color_str);
            } else {
                // Apply to current cell
                Cell* cell = sheet_get_or_create_cell(state->sheet, state->cursor_row, state->cursor_col);
                if (cell) {
                    cell_set_text_color(cell, color);
                    sprintf_s(state->status_message, sizeof(state->status_message), 
                             "Cell text color set to %s", color_str);
                }
            }
        } else {
            sprintf_s(state->status_message, sizeof(state->status_message), 
                     "Invalid color: %s", color_str);
        }
    }
    // NEW: Color commands for background
    else if (strncmp(command, "clrbg ", 6) == 0) {
        const char* color_str = command + 6;
        int color = parse_color(color_str);
        if (color >= 0) {
            if (state->sheet->selection.is_active) {
                // Apply to range
                int min_row = state->sheet->selection.start_row < state->sheet->selection.end_row ? 
                              state->sheet->selection.start_row : state->sheet->selection.end_row;
                int max_row = state->sheet->selection.start_row > state->sheet->selection.end_row ? 
                              state->sheet->selection.start_row : state->sheet->selection.end_row;
                int min_col = state->sheet->selection.start_col < state->sheet->selection.end_col ? 
                              state->sheet->selection.start_col : state->sheet->selection.end_col;
                int max_col = state->sheet->selection.start_col > state->sheet->selection.end_col ? 
                              state->sheet->selection.start_col : state->sheet->selection.end_col;
                
                for (int row = min_row; row <= max_row; row++) {
                    for (int col = min_col; col <= max_col; col++) {
                        Cell* cell = sheet_get_or_create_cell(state->sheet, row, col);
                        if (cell) {
                            cell_set_background_color(cell, color);
                        }
                    }
                }
                sprintf_s(state->status_message, sizeof(state->status_message), 
                         "Range background color set to %s", color_str);
            } else {
                // Apply to current cell
                Cell* cell = sheet_get_or_create_cell(state->sheet, state->cursor_row, state->cursor_col);
                if (cell) {
                    cell_set_background_color(cell, color);
                    sprintf_s(state->status_message, sizeof(state->status_message), 
                             "Cell background color set to %s", color_str);
                }
            }
        } else {
            sprintf_s(state->status_message, sizeof(state->status_message), 
                     "Invalid color: %s", color_str);
        }
    }else {
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

// NEW: Copy range to clipboard
void app_copy_range(AppState* state) {
    if (state->sheet->selection.is_active) {
        sheet_copy_range(state->sheet);
        strcpy_s(state->status_message, sizeof(state->status_message), "Range copied");
    } else {
        app_copy_cell(state);
    }
}

// NEW: Paste range from clipboard
void app_paste_range(AppState* state) {
    if (state->sheet->range_clipboard.is_active) {
        sheet_paste_range(state->sheet, state->cursor_row, state->cursor_col);
        strcpy_s(state->status_message, sizeof(state->status_message), "Range pasted");
    } else {
        app_paste_cell(state);
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
        set_system_clipboard_text("");
        strcpy_s(state->status_message, sizeof(state->status_message), "Empty cell copied to system clipboard");
    }
}

// Paste from system clipboard to current cell
void app_paste_from_system_clipboard(AppState* state) {
    char* text = get_system_clipboard_text();
    if (text) {
        if (strlen(text) == 0) {
            sheet_clear_cell(state->sheet, state->cursor_row, state->cursor_col);
            strcpy_s(state->status_message, sizeof(state->status_message), "Cell cleared from system clipboard");
        } else if (text[0] == '=') {
            sheet_set_formula(state->sheet, state->cursor_row, state->cursor_col, text);
            sheet_recalculate(state->sheet);
            strcpy_s(state->status_message, sizeof(state->status_message), "Formula pasted from system clipboard");
        } else {
            char* endptr;
            double num = strtod(text, &endptr);
            if (*endptr == '\0' || (*endptr == '\n' && *(endptr+1) == '\0')) {
                sheet_set_number(state->sheet, state->cursor_row, state->cursor_col, num);
                strcpy_s(state->status_message, sizeof(state->status_message), "Number pasted from system clipboard");
            } else {
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
        if (key->type == 0) {  // Character key
            switch (key->key.ch) {
                // Navigation                case 'h':
                    if (state->cursor_col > 0) {
                        if (key->shift) {
                            if (!state->range_selection_active) {
                                app_start_range_selection(state);
                            }
                        }
                        state->cursor_col--;
                        if (key->shift) {
                            app_extend_range_selection(state, state->cursor_row, state->cursor_col);
                        } else if (state->range_selection_active) {
                            app_cancel_range_selection(state);
                        }
                    }
                    break;
                case 'l':
                    if (state->cursor_col < state->sheet->cols - 1) {
                        if (key->shift) {
                            if (!state->range_selection_active) {
                                app_start_range_selection(state);
                            }
                        }
                        state->cursor_col++;
                        if (key->shift) {
                            app_extend_range_selection(state, state->cursor_row, state->cursor_col);
                        } else if (state->range_selection_active) {
                            app_cancel_range_selection(state);
                        }
                    }
                    break;
                case 'j':
                    if (state->cursor_row < state->sheet->rows - 1) {
                        if (key->shift) {
                            if (!state->range_selection_active) {
                                app_start_range_selection(state);
                            }
                        }
                        state->cursor_row++;
                        if (key->shift) {
                            app_extend_range_selection(state, state->cursor_row, state->cursor_col);
                        } else if (state->range_selection_active) {
                            app_cancel_range_selection(state);
                        }
                    }
                    break;
                case 'k':
                    if (state->cursor_row > 0) {
                        if (key->shift) {
                            if (!state->range_selection_active) {
                                app_start_range_selection(state);
                            }
                        }
                        state->cursor_row--;
                        if (key->shift) {
                            app_extend_range_selection(state, state->cursor_row, state->cursor_col);
                        } else if (state->range_selection_active) {
                            app_cancel_range_selection(state);
                        }
                    }
                    break;
                    
                // Commands
                case '=':
                    app_start_input(state, MODE_INSERT_FORMULA);
                    break;
                case '"':
                    app_start_input(state, MODE_INSERT_STRING);
                    break;
                case ':':
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
                        app_copy_range(state);  // NEW: Enhanced copy
                    }
                    break;
                case 'v':
                    if (key->ctrl && key->shift) {
                        app_paste_from_system_clipboard(state);
                    } else if (key->ctrl) {
                        app_paste_range(state);  // NEW: Enhanced paste
                    }
                    break;                case 'q':
                    if (key->ctrl) {
                        state->running = FALSE;
                    }
                    break;                // NEW: Date format cycling with Ctrl+#
                case '#':
                    if (key->ctrl) {
                        app_cycle_date_format(state);
                    }
                    break;
                // Additional formatting for numbers 1-9 keys
                case '5':  // Ctrl+5 for percentage (common Excel shortcut)
                    if (key->ctrl && key->shift) {
                        app_set_cell_format(state, FORMAT_PERCENTAGE, 0);
                    }
                    break;
                case '4':  // Ctrl+Shift+4 for currency (common Excel shortcut)
                    if (key->ctrl && key->shift) {
                        app_set_cell_format(state, FORMAT_CURRENCY, 0);
                    }
                    break;                case '3':  // Ctrl+Shift+3 for date/time cycling (enhanced Excel-style)
                    if (key->ctrl && key->shift) {
                        app_cycle_datetime_format(state);
                    }
                    break;
                case '1':  // Ctrl+Shift+1 for number format
                    if (key->ctrl && key->shift) {
                        app_set_cell_format(state, FORMAT_NUMBER, 0);
                    }
                    break;
                // NEW: ESC to cancel range selection
                case KEY_ESC:
                    if (state->range_selection_active) {
                        app_cancel_range_selection(state);
                    }
                    break;
            }
        } else {  // Special key
            switch (key->key.special) {                case KEY_LEFT:
                    if (key->alt) {
                        // Alt+Left: Decrease column width
                        if (state->sheet->selection.is_active) {
                            // Resize selected columns
                            int min_col = state->sheet->selection.start_col < state->sheet->selection.end_col ? 
                                          state->sheet->selection.start_col : state->sheet->selection.end_col;
                            int max_col = state->sheet->selection.start_col > state->sheet->selection.end_col ? 
                                          state->sheet->selection.start_col : state->sheet->selection.end_col;
                            sheet_resize_columns_in_range(state->sheet, min_col, max_col, -1);
                            strcpy_s(state->status_message, sizeof(state->status_message), "Columns resized");
                        } else {
                            // Resize current column
                            sheet_resize_columns_in_range(state->sheet, state->cursor_col, state->cursor_col, -1);
                            strcpy_s(state->status_message, sizeof(state->status_message), "Column resized");
                        }                    } else if (state->cursor_col > 0) {
                        if (key->shift) {
                            if (!state->range_selection_active) {
                                app_start_range_selection(state);
                            }
                        }
                        state->cursor_col--;
                        if (key->shift) {
                            app_extend_range_selection(state, state->cursor_row, state->cursor_col);
                        } else if (state->range_selection_active) {
                            app_cancel_range_selection(state);
                        }
                    }
                    break;
                case KEY_RIGHT:
                    if (key->alt) {
                        // Alt+Right: Increase column width
                        if (state->sheet->selection.is_active) {
                            // Resize selected columns
                            int min_col = state->sheet->selection.start_col < state->sheet->selection.end_col ? 
                                          state->sheet->selection.start_col : state->sheet->selection.end_col;
                            int max_col = state->sheet->selection.start_col > state->sheet->selection.end_col ? 
                                          state->sheet->selection.start_col : state->sheet->selection.end_col;
                            sheet_resize_columns_in_range(state->sheet, min_col, max_col, 1);
                            strcpy_s(state->status_message, sizeof(state->status_message), "Columns resized");
                        } else {
                            // Resize current column
                            sheet_resize_columns_in_range(state->sheet, state->cursor_col, state->cursor_col, 1);
                            strcpy_s(state->status_message, sizeof(state->status_message), "Column resized");
                        }                    } else if (state->cursor_col < state->sheet->cols - 1) {
                        if (key->shift) {
                            if (!state->range_selection_active) {
                                app_start_range_selection(state);
                            }
                        }
                        state->cursor_col++;
                        if (key->shift) {
                            app_extend_range_selection(state, state->cursor_row, state->cursor_col);
                        } else if (state->range_selection_active) {
                            app_cancel_range_selection(state);
                        }
                    }
                    break;
                case KEY_UP:
                    if (key->alt) {
                        // Alt+Up: Decrease row height
                        if (state->sheet->selection.is_active) {
                            // Resize selected rows
                            int min_row = state->sheet->selection.start_row < state->sheet->selection.end_row ? 
                                          state->sheet->selection.start_row : state->sheet->selection.end_row;
                            int max_row = state->sheet->selection.start_row > state->sheet->selection.end_row ? 
                                          state->sheet->selection.start_row : state->sheet->selection.end_row;
                            sheet_resize_rows_in_range(state->sheet, min_row, max_row, -1);
                            strcpy_s(state->status_message, sizeof(state->status_message), "Rows resized");
                        } else {
                            // Resize current row
                            sheet_resize_rows_in_range(state->sheet, state->cursor_row, state->cursor_row, -1);
                            strcpy_s(state->status_message, sizeof(state->status_message), "Row resized");
                        }                    } else if (state->cursor_row > 0) {
                        if (key->shift) {
                            if (!state->range_selection_active) {
                                app_start_range_selection(state);
                            }
                        }
                        state->cursor_row--;
                        if (key->shift) {
                            app_extend_range_selection(state, state->cursor_row, state->cursor_col);
                        } else if (state->range_selection_active) {
                            app_cancel_range_selection(state);
                        }
                    }
                    break;
                case KEY_DOWN:
                    if (key->alt) {
                        // Alt+Down: Increase row height
                        if (state->sheet->selection.is_active) {
                            // Resize selected rows
                            int min_row = state->sheet->selection.start_row < state->sheet->selection.end_row ? 
                                          state->sheet->selection.start_row : state->sheet->selection.end_row;
                            int max_row = state->sheet->selection.start_row > state->sheet->selection.end_row ? 
                                          state->sheet->selection.start_row : state->sheet->selection.end_row;
                            sheet_resize_rows_in_range(state->sheet, min_row, max_row, 1);
                            strcpy_s(state->status_message, sizeof(state->status_message), "Rows resized");
                        } else {
                            // Resize current row
                            sheet_resize_rows_in_range(state->sheet, state->cursor_row, state->cursor_row, 1);
                            strcpy_s(state->status_message, sizeof(state->status_message), "Row resized");
                        }                    } else if (state->cursor_row < state->sheet->rows - 1) {
                        if (key->shift) {
                            if (!state->range_selection_active) {
                                app_start_range_selection(state);
                            }
                        }
                        state->cursor_row++;
                        if (key->shift) {
                            app_extend_range_selection(state, state->cursor_row, state->cursor_col);
                        } else if (state->range_selection_active) {
                            app_cancel_range_selection(state);
                        }
                    }
                    break;
                case KEY_PGUP:
                    state->cursor_row = max(0, state->cursor_row - 10);
                    if (state->range_selection_active) {
                        app_cancel_range_selection(state);
                    }
                    break;
                case KEY_PGDN:
                    state->cursor_row = min(state->sheet->rows - 1, state->cursor_row + 10);
                    if (state->range_selection_active) {
                        app_cancel_range_selection(state);
                    }
                    break;
                case KEY_ESC:
                    if (state->range_selection_active) {
                        app_cancel_range_selection(state);
                    }
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
    debug_log("=== Starting Enhanced WinSpread ===");
    
    AppState state;
    
    app_init(&state);
    
    if (!state.running) {
        printf("Failed to initialize application\n");
        debug_cleanup();
        return 1;
    }
    
    // Main loop
    while (state.running) {
        app_update_cursor_blink(&state);
        app_render(&state);
        
        KeyEvent key;
        if (console_get_key(state.console, &key)) {
            app_handle_input(&state, &key);
            
            state.cursor_visible = TRUE;
            state.cursor_blink_time = GetTickCount();
        }
        
        Sleep(16);  // ~60 FPS
    }
    
    app_cleanup(&state);
    debug_log("=== Enhanced WinSpread Ended ===");
    debug_cleanup();
    return 0;
}
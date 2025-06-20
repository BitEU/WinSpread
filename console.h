// console.h - Windows Console API wrapper for terminal spreadsheet
#ifndef CONSOLE_H
#define CONSOLE_H

#include <windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

// Include debug logging
extern void debug_log(const char* format, ...);
extern FILE* debug_file;

// Console color attributes
#define COLOR_BLACK         0
#define COLOR_BLUE          1
#define COLOR_GREEN         2
#define COLOR_CYAN          3
#define COLOR_RED           4
#define COLOR_MAGENTA       5
#define COLOR_YELLOW        6
#define COLOR_WHITE         7
#define COLOR_BRIGHT        8

// Combine foreground and background colors
#define MAKE_COLOR(fg, bg)  ((bg << 4) | fg)

// Special keys
#define KEY_UP              0x48
#define KEY_DOWN            0x50
#define KEY_LEFT            0x4B
#define KEY_RIGHT           0x4D
#define KEY_PGUP            0x49
#define KEY_PGDN            0x51
#define KEY_HOME            0x47
#define KEY_END             0x4F
#define KEY_F1              0x3B
#define KEY_ESC             0x1B
#define KEY_ENTER           0x0D
#define KEY_BACKSPACE       0x08
#define KEY_TAB             0x09

typedef struct {
    HANDLE hOut;
    HANDLE hIn;
    CONSOLE_SCREEN_BUFFER_INFO originalInfo;
    CHAR_INFO* backBuffer;
    CHAR_INFO* frontBuffer;
    SHORT width;
    SHORT height;
    DWORD originalMode;
} Console;

typedef struct {
    int type;  // 0 = char, 1 = special key
    union {
        char ch;
        int special;
    } key;
    BOOL ctrl;
    BOOL alt;
    BOOL shift;
} KeyEvent;

// Function prototypes
Console* console_init(void);
void console_cleanup(Console* con);
void console_clear(Console* con);
void console_set_cursor(Console* con, SHORT x, SHORT y);
void console_hide_cursor(Console* con);
void console_show_cursor(Console* con);
void console_write_char(Console* con, SHORT x, SHORT y, char ch, WORD attr);
void console_write_string(Console* con, SHORT x, SHORT y, const char* str, WORD attr);
void console_draw_box(Console* con, SHORT x, SHORT y, SHORT w, SHORT h, WORD attr);
void console_flip(Console* con);
BOOL console_get_key(Console* con, KeyEvent* key);
void console_get_size(Console* con, SHORT* width, SHORT* height);

// Implementation
Console* console_init(void) {
    Console* con = (Console*)malloc(sizeof(Console));
    if (!con) return NULL;
    
    // Get console handles
    con->hOut = GetStdHandle(STD_OUTPUT_HANDLE);
    con->hIn = GetStdHandle(STD_INPUT_HANDLE);
    
    // Check if handles are valid
    if (con->hOut == INVALID_HANDLE_VALUE || con->hIn == INVALID_HANDLE_VALUE) {
        free(con);
        return NULL;
    }
    
    // Save original console state
    if (!GetConsoleScreenBufferInfo(con->hOut, &con->originalInfo)) {
        free(con);
        return NULL;
    }
    
    if (!GetConsoleMode(con->hIn, &con->originalMode)) {
        free(con);
        return NULL;
    }
    
    // Set input mode for reading individual keys
    SetConsoleMode(con->hIn, ENABLE_WINDOW_INPUT | ENABLE_MOUSE_INPUT);
    
    // Get console size
    CONSOLE_SCREEN_BUFFER_INFO csbi;
    GetConsoleScreenBufferInfo(con->hOut, &csbi);
    con->width = csbi.srWindow.Right - csbi.srWindow.Left + 1;
    con->height = csbi.srWindow.Bottom - csbi.srWindow.Top + 1;
    
    // Validate console size
    if (con->width <= 0 || con->height <= 0) {
        free(con);
        return NULL;
    }
    
    // Allocate double buffer
    int bufferSize = con->width * con->height;
    con->backBuffer = (CHAR_INFO*)malloc(bufferSize * sizeof(CHAR_INFO));
    con->frontBuffer = (CHAR_INFO*)malloc(bufferSize * sizeof(CHAR_INFO));
    
    // Check if allocation succeeded
    if (!con->backBuffer || !con->frontBuffer) {
        if (con->backBuffer) free(con->backBuffer);
        if (con->frontBuffer) free(con->frontBuffer);
        free(con);
        return NULL;
    }
    
    // Initialize buffers
    for (int i = 0; i < bufferSize; i++) {
        con->backBuffer[i].Char.AsciiChar = ' ';
        con->backBuffer[i].Attributes = MAKE_COLOR(COLOR_WHITE, COLOR_BLACK);
        con->frontBuffer[i] = con->backBuffer[i];
    }
    
    // Clear screen
    console_clear(con);
    
    return con;
}

void console_cleanup(Console* con) {
    if (!con) return;
    
    // Restore original console state
    SetConsoleTextAttribute(con->hOut, con->originalInfo.wAttributes);
    SetConsoleMode(con->hIn, con->originalMode);
    
    // Free buffers
    if (con->backBuffer) {
        free(con->backBuffer);
        con->backBuffer = NULL;
    }
    if (con->frontBuffer) {
        free(con->frontBuffer);
        con->frontBuffer = NULL;
    }
    free(con);
}

void console_clear(Console* con) {
    COORD topLeft = {0, 0};
    DWORD written;
    FillConsoleOutputCharacter(con->hOut, ' ', con->width * con->height, topLeft, &written);
    FillConsoleOutputAttribute(con->hOut, MAKE_COLOR(COLOR_WHITE, COLOR_BLACK), 
                               con->width * con->height, topLeft, &written);
    console_set_cursor(con, 0, 0);
}

void console_set_cursor(Console* con, SHORT x, SHORT y) {
    COORD pos = {x, y};
    SetConsoleCursorPosition(con->hOut, pos);
}

void console_hide_cursor(Console* con) {
    CONSOLE_CURSOR_INFO cursorInfo;
    GetConsoleCursorInfo(con->hOut, &cursorInfo);
    cursorInfo.bVisible = FALSE;
    SetConsoleCursorInfo(con->hOut, &cursorInfo);
}

void console_show_cursor(Console* con) {
    CONSOLE_CURSOR_INFO cursorInfo;
    GetConsoleCursorInfo(con->hOut, &cursorInfo);
    cursorInfo.bVisible = TRUE;
    SetConsoleCursorInfo(con->hOut, &cursorInfo);
}

void console_write_char(Console* con, SHORT x, SHORT y, char ch, WORD attr) {
    if (x >= 0 && x < con->width && y >= 0 && y < con->height) {
        int index = y * con->width + x;
        con->backBuffer[index].Char.AsciiChar = ch;
        con->backBuffer[index].Attributes = attr;
    }
}

void console_write_string(Console* con, SHORT x, SHORT y, const char* str, WORD attr) {
    int len = (int)strlen(str);
    for (int i = 0; i < len && x + i < con->width; i++) {
        console_write_char(con, x + i, y, str[i], attr);
    }
}

void console_draw_box(Console* con, SHORT x, SHORT y, SHORT w, SHORT h, WORD attr) {
    // Draw corners
    console_write_char(con, x, y, '+', attr);
    console_write_char(con, x + w - 1, y, '+', attr);
    console_write_char(con, x, y + h - 1, '+', attr);
    console_write_char(con, x + w - 1, y + h - 1, '+', attr);
    
    // Draw horizontal lines
    for (SHORT i = 1; i < w - 1; i++) {
        console_write_char(con, x + i, y, '-', attr);
        console_write_char(con, x + i, y + h - 1, '-', attr);
    }
    
    // Draw vertical lines
    for (SHORT i = 1; i < h - 1; i++) {
        console_write_char(con, x, y + i, '|', attr);
        console_write_char(con, x + w - 1, y + i, '|', attr);
    }
}

void console_flip(Console* con) {
    if (debug_file) debug_log("console_flip called");
    
    if (!con) {
        if (debug_file) debug_log("ERROR: console_flip - con is NULL");
        return;
    }
    
    if (!con->backBuffer) {
        if (debug_file) debug_log("ERROR: console_flip - backBuffer is NULL");
        return;
    }
    
    if (debug_file) debug_log("console_flip - calling WriteConsoleOutput, size: %dx%d", con->width, con->height);
    
    // Only update changed characters (optimization)
    COORD bufferSize = {con->width, con->height};
    COORD bufferCoord = {0, 0};
    SMALL_RECT writeRegion = {0, 0, con->width - 1, con->height - 1};
    
    BOOL result = WriteConsoleOutput(con->hOut, con->backBuffer, bufferSize, bufferCoord, &writeRegion);
    if (!result) {
        DWORD error = GetLastError();
        if (debug_file) debug_log("ERROR: WriteConsoleOutput failed with error: %lu", error);
        return;
    }
    
    if (debug_file) debug_log("WriteConsoleOutput succeeded, copying buffers");
    // Copy back buffer to front buffer
    memcpy(con->frontBuffer, con->backBuffer, con->width * con->height * sizeof(CHAR_INFO));
    
    if (debug_file) debug_log("console_flip completed successfully");
}

BOOL console_get_key(Console* con, KeyEvent* key) {
    INPUT_RECORD inputRecord;
    DWORD eventsRead;
    
    // Check if input is available
    DWORD numEvents;
    GetNumberOfConsoleInputEvents(con->hIn, &numEvents);
    if (numEvents == 0) return FALSE;
    
    // Read input
    if (ReadConsoleInput(con->hIn, &inputRecord, 1, &eventsRead)) {
        if (inputRecord.EventType == KEY_EVENT && inputRecord.Event.KeyEvent.bKeyDown) {
            KEY_EVENT_RECORD* keyEvent = &inputRecord.Event.KeyEvent;
            
            // Set modifiers
            key->ctrl = (keyEvent->dwControlKeyState & (LEFT_CTRL_PRESSED | RIGHT_CTRL_PRESSED)) != 0;
            key->alt = (keyEvent->dwControlKeyState & (LEFT_ALT_PRESSED | RIGHT_ALT_PRESSED)) != 0;
            key->shift = (keyEvent->dwControlKeyState & SHIFT_PRESSED) != 0;
            
            // Check for special keys
            if (keyEvent->wVirtualKeyCode == VK_UP) {
                key->type = 1;
                key->key.special = KEY_UP;
                return TRUE;
            } else if (keyEvent->wVirtualKeyCode == VK_DOWN) {
                key->type = 1;
                key->key.special = KEY_DOWN;
                return TRUE;
            } else if (keyEvent->wVirtualKeyCode == VK_LEFT) {
                key->type = 1;
                key->key.special = KEY_LEFT;
                return TRUE;
            } else if (keyEvent->wVirtualKeyCode == VK_RIGHT) {
                key->type = 1;
                key->key.special = KEY_RIGHT;
                return TRUE;
            } else if (keyEvent->wVirtualKeyCode == VK_PRIOR) {
                key->type = 1;
                key->key.special = KEY_PGUP;
                return TRUE;
            } else if (keyEvent->wVirtualKeyCode == VK_NEXT) {
                key->type = 1;
                key->key.special = KEY_PGDN;
                return TRUE;
            } else if (keyEvent->wVirtualKeyCode == VK_HOME) {
                key->type = 1;
                key->key.special = KEY_HOME;
                return TRUE;
            } else if (keyEvent->wVirtualKeyCode == VK_END) {
                key->type = 1;
                key->key.special = KEY_END;
                return TRUE;            } else if (keyEvent->wVirtualKeyCode == VK_F1) {
                key->type = 1;
                key->key.special = KEY_F1;
                return TRUE;
            } else if (keyEvent->wVirtualKeyCode >= 'A' && keyEvent->wVirtualKeyCode <= 'Z' && key->ctrl) {
                // Handle Ctrl+letter combinations (with or without Shift)
                key->type = 0;
                key->key.ch = (char)keyEvent->wVirtualKeyCode + ('a' - 'A'); // Convert to lowercase
                return TRUE;
            } else if (key->ctrl && key->shift) {
                // Handle Ctrl+Shift+number combinations for formatting
                switch (keyEvent->wVirtualKeyCode) {
                    case '1':
                        key->type = 0;
                        key->key.ch = '1';
                        return TRUE;
                    case '3':
                        key->type = 0;
                        key->key.ch = '3';
                        return TRUE;
                    case '4':
                        key->type = 0;
                        key->key.ch = '4';
                        return TRUE;
                    case '5':
                        key->type = 0;
                        key->key.ch = '5';
                        return TRUE;
                }
            } else if (key->ctrl) {
                // Handle Ctrl+symbol combinations
                switch (keyEvent->wVirtualKeyCode) {
                    case '5':  // Ctrl+5 for %
                        if (key->shift) {
                            key->type = 0;
                            key->key.ch = '%';
                            return TRUE;
                        }
                        break;
                    case '4':  // Ctrl+4 for $
                        if (key->shift) {
                            key->type = 0;
                            key->key.ch = '$';
                            return TRUE;
                        }
                        break;
                    case '3':  // Ctrl+3 for #
                        if (key->shift) {
                            key->type = 0;
                            key->key.ch = '#';
                            return TRUE;
                        }
                        break;
                }
                // Also handle when the character is already the symbol
                if (keyEvent->uChar.AsciiChar != 0) {
                    key->type = 0;
                    key->key.ch = keyEvent->uChar.AsciiChar;
                    return TRUE;
                }
            } else if (keyEvent->uChar.AsciiChar != 0) {
                // Regular character
                key->type = 0;
                key->key.ch = keyEvent->uChar.AsciiChar;
                return TRUE;
            }
        }
    }
    
    return FALSE;
}

void console_get_size(Console* con, SHORT* width, SHORT* height) {
    CONSOLE_SCREEN_BUFFER_INFO csbi;
    GetConsoleScreenBufferInfo(con->hOut, &csbi);
    *width = csbi.srWindow.Right - csbi.srWindow.Left + 1;
    *height = csbi.srWindow.Bottom - csbi.srWindow.Top + 1;
}

/*
// Demo program showing basic spreadsheet grid
int main() {
    Console* con = console_init();
    if (!con) {
        printf("Failed to initialize console\n");
        return 1;
    }
    
    console_hide_cursor(con);
    
    // Colors
    WORD headerColor = MAKE_COLOR(COLOR_BLACK, COLOR_WHITE);
    WORD cellColor = MAKE_COLOR(COLOR_WHITE, COLOR_BLACK);
    WORD selectedColor = MAKE_COLOR(COLOR_BLACK, COLOR_CYAN);
    WORD formulaColor = MAKE_COLOR(COLOR_GREEN, COLOR_BLACK);
    
    // Grid parameters
    int colWidth = 10;
    int numCols = 7;
    int numRows = 20;
    int currentRow = 0;
    int currentCol = 0;
    
    // Main loop
    BOOL running = TRUE;
    while (running) {
        // Clear back buffer
        for (int i = 0; i < con->width * con->height; i++) {
            con->backBuffer[i].Char.AsciiChar = ' ';
            con->backBuffer[i].Attributes = cellColor;
        }
          // Draw column headers
        for (int col = 0; col < numCols; col++) {
            char colName[10];
            sprintf_s(colName, sizeof(colName), "%c", 'A' + col);
            int x = (col + 1) * colWidth + colWidth/2 - 1;
            console_write_string(con, x, 0, colName, headerColor);
        }
          // Draw row headers
        for (int row = 0; row < numRows && row < con->height - 2; row++) {
            char rowNum[10];
            sprintf_s(rowNum, sizeof(rowNum), "%3d", row + 1);
            console_write_string(con, 0, row + 1, rowNum, headerColor);
        }
        
        // Draw grid lines
        for (int row = 0; row <= numRows && row < con->height - 1; row++) {
            for (int col = 0; col <= numCols; col++) {
                int x = col * colWidth + 4;
                console_write_char(con, x, row, '|', cellColor);
            }
        }
        
        // Draw horizontal lines
        for (int col = 0; col < numCols * colWidth + 5; col++) {
            console_write_char(con, col, 1, '-', cellColor);
        }
        
        // Highlight current cell
        int cellX = (currentCol + 1) * colWidth + 5;
        int cellY = currentRow + 2;
        for (int i = 0; i < colWidth - 1; i++) {
            con->backBuffer[cellY * con->width + cellX + i].Attributes = selectedColor;
        }
          // Draw status line
        char status[100];
        sprintf_s(status, sizeof(status), "[Sheet1] %c%d | F1:Help  :q to quit", 'A' + currentCol, currentRow + 1);
        console_write_string(con, 0, con->height - 1, status, headerColor);
        
        // Update screen
        console_flip(con);
        
        // Handle input
        KeyEvent key;
        if (console_get_key(con, &key)) {
            if (key.type == 0) {  // Character key
                if (key.key.ch == 'q' || key.key.ch == 'Q') {
                    running = FALSE;
                } else if (key.key.ch == 'h' && currentCol > 0) {
                    currentCol--;
                } else if (key.key.ch == 'l' && currentCol < numCols - 1) {
                    currentCol++;
                } else if (key.key.ch == 'j' && currentRow < numRows - 1) {
                    currentRow++;
                } else if (key.key.ch == 'k' && currentRow > 0) {
                    currentRow--;
                }
            } else {  // Special key
                switch (key.key.special) {
                    case KEY_LEFT:
                        if (currentCol > 0) currentCol--;
                        break;
                    case KEY_RIGHT:
                        if (currentCol < numCols - 1) currentCol++;
                        break;
                    case KEY_UP:
                        if (currentRow > 0) currentRow--;
                        break;
                    case KEY_DOWN:
                        if (currentRow < numRows - 1) currentRow++;
                        break;
                    case KEY_ESC:
                        running = FALSE;
                        break;
                }
            }
        }
        
        Sleep(16);  // ~60 FPS
    }
      console_show_cursor(con);
    console_clear(con);
    console_cleanup(con);
    
    return 0;
}
*/

// To compile with MSVC:
// cl console.c /Fe:wtsc_demo.exe user32.lib

// To compile with MinGW:
// gcc console.c -o wtsc_demo.exe

#endif // CONSOLE_H
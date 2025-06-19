// debug.h - Debug logging utilities
#ifndef DEBUG_H
#define DEBUG_H

#include <stdio.h>
#include <time.h>
#include <stdarg.h>

// Debug logging function
void debug_log(const char* format, ...);
void debug_init();
void debug_cleanup();

// Global debug file handle
extern FILE* debug_file;

// Implementation
FILE* debug_file = NULL;

void debug_init() {
    fopen_s(&debug_file, "debug.log", "w");
    if (debug_file) {
        debug_log("=== WinSpread Debug Log Started ===");
    }
}

void debug_cleanup() {
    if (debug_file) {
        debug_log("=== WinSpread Debug Log Ended ===");
        fclose(debug_file);
        debug_file = NULL;
    }
}

void debug_log(const char* format, ...) {
    if (!debug_file) return;
    
    // Get current time
    time_t now = time(NULL);
    struct tm timeinfo;
    localtime_s(&timeinfo, &now);
    
    // Write timestamp
    fprintf(debug_file, "[%02d:%02d:%02d] ", 
            timeinfo.tm_hour, timeinfo.tm_min, timeinfo.tm_sec);
    
    // Write the actual message
    va_list args;
    va_start(args, format);
    vfprintf(debug_file, format, args);
    va_end(args);
    
    fprintf(debug_file, "\n");
    fflush(debug_file);  // Ensure it's written immediately
}

#endif // DEBUG_H

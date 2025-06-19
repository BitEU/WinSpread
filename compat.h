// compat.h - Compatibility layer for different compilers
#ifndef COMPAT_H
#define COMPAT_H

// Handle secure string functions for different compilers
#ifdef _MSC_VER
    // MSVC - use secure versions
    #define STRCPY_S(dest, size, src) strcpy_s(dest, size, src)
    #define STRNCPY_S(dest, size, src, count) strncpy_s(dest, size, src, count)
    #define SPRINTF_S sprintf_s
    #define STRDUP(s) _strdup(s)
#else
    // GCC/MinGW - use standard versions with wrappers
    #include <string.h>
    
    #define STRCPY_S(dest, size, src) (strncpy(dest, src, size), dest[size-1] = '\0')
    #define STRNCPY_S(dest, size, src, count) (strncpy(dest, src, count), dest[size-1] = '\0')
    #define SPRINTF_S(buffer, size, ...) snprintf(buffer, size, __VA_ARGS__)
    #define STRDUP(s) strdup(s)
    
    // Provide secure alternatives if needed
    static inline void strcpy_s(char* dest, size_t size, const char* src) {
        if (dest && src && size > 0) {
            strncpy(dest, src, size - 1);
            dest[size - 1] = '\0';
        }
    }
    
    static inline void strncpy_s(char* dest, size_t size, const char* src, size_t count) {
        if (dest && src && size > 0) {
            size_t copy_size = (count < size - 1) ? count : size - 1;
            strncpy(dest, src, copy_size);
            dest[copy_size] = '\0';
        }
    }
#endif

// Helper macros
#ifndef min
#define min(a,b) ((a) < (b) ? (a) : (b))
#endif

#ifndef max
#define max(a,b) ((a) > (b) ? (a) : (b))
#endif

#endif // COMPAT_H
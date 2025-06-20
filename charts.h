// charts.h - ASCII chart generation for WinSpread terminal spreadsheet
#ifndef CHARTS_H
#define CHARTS_H

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <math.h>
#include <float.h>
#include "sheet.h"
#include "console.h"

// Chart types
typedef enum {
    CHART_LINE,
    CHART_BAR,
    CHART_PIE,
    CHART_SCATTER
} ChartType;

// Chart configuration
typedef struct {
    ChartType type;
    char x_label[64];
    char y_label[64];
    char title[128];
    int width;
    int height;
    int show_grid;
    int show_legend;
} ChartConfig;

// Chart data point
typedef struct {
    double x;
    double y;
    char label[64];
} ChartPoint;

// Chart data series
typedef struct {
    ChartPoint* points;
    int count;
    char name[64];
    char symbol;  // Symbol to use for this series
} ChartSeries;

// Chart structure
typedef struct {
    ChartConfig config;
    ChartSeries* series;
    int series_count;
    double x_min, x_max;
    double y_min, y_max;
    char** canvas;  // ASCII canvas for drawing
    int canvas_width;
    int canvas_height;
} Chart;

// Function prototypes
Chart* chart_create(ChartType type, const char* x_label, const char* y_label);
Chart* chart_create_sized(ChartType type, const char* x_label, const char* y_label, int width, int height);
void chart_free(Chart* chart);
int chart_add_data_from_range(Chart* chart, Sheet* sheet, RangeSelection* range);
void chart_render(Chart* chart);
void chart_display(Chart* chart, Console* console, int x, int y);
char** chart_get_output(Chart* chart, int* line_count);

// Helper functions
void chart_draw_line(Chart* chart, int x1, int y1, int x2, int y2, char symbol);
void chart_draw_axes(Chart* chart);
void chart_plot_line_chart(Chart* chart);
void chart_plot_bar_chart(Chart* chart);
void chart_plot_pie_chart(Chart* chart);
void chart_plot_scatter_chart(Chart* chart);
void chart_set_pixel(Chart* chart, int x, int y, char c);
int chart_scale_x(Chart* chart, double value);
int chart_scale_y(Chart* chart, double value);

// Implementation

Chart* chart_create(ChartType type, const char* x_label, const char* y_label) {
    // Use larger default sizes for better visibility
    return chart_create_sized(type, x_label, y_label, 120, 40);
}

Chart* chart_create_sized(ChartType type, const char* x_label, const char* y_label, int width, int height) {
    Chart* chart = (Chart*)calloc(1, sizeof(Chart));
    if (!chart) return NULL;
    
    chart->config.type = type;
    strncpy_s(chart->config.x_label, sizeof(chart->config.x_label), x_label ? x_label : "X", _TRUNCATE);
    strncpy_s(chart->config.y_label, sizeof(chart->config.y_label), y_label ? y_label : "Y", _TRUNCATE);
    
    // Use provided sizes with more generous bounds checking for full-screen charts
    chart->config.width = width;
    chart->config.height = height;
    if (chart->config.width < 40) chart->config.width = 40;
    if (chart->config.width > 300) chart->config.width = 300;  // Increased from 200
    if (chart->config.height < 15) chart->config.height = 15;
    if (chart->config.height > 100) chart->config.height = 100; // Increased from 50
    
    chart->config.show_grid = 1;
    chart->config.show_legend = 1;
    
    // Initialize bounds
    chart->x_min = DBL_MAX;
    chart->x_max = -DBL_MAX;
    chart->y_min = DBL_MAX;
    chart->y_max = -DBL_MAX;      // Allocate canvas with appropriate margins for labels, axes, and legend
    // Ensure canvas width doesn't exceed what can be displayed
    int legend_space = chart->config.show_legend ? 25 : 5;  // Space for legend or minimal margin
    chart->canvas_width = chart->config.width + legend_space;
    chart->canvas_height = chart->config.height + 12; // Space for X-axis labels and potential legend below
    chart->canvas = (char**)calloc(chart->canvas_height, sizeof(char*));
    
    for (int i = 0; i < chart->canvas_height; i++) {
        chart->canvas[i] = (char*)calloc(chart->canvas_width + 1, sizeof(char));
        memset(chart->canvas[i], ' ', chart->canvas_width);
        chart->canvas[i][chart->canvas_width] = '\0';
    }
    
    return chart;
}

void chart_free(Chart* chart) {
    if (!chart) return;
    
    // Free canvas
    if (chart->canvas) {
        for (int i = 0; i < chart->canvas_height; i++) {
            free(chart->canvas[i]);
        }
        free(chart->canvas);
    }
    
    // Free series data
    if (chart->series) {
        for (int i = 0; i < chart->series_count; i++) {
            if (chart->series[i].points) {
                free(chart->series[i].points);
            }
        }
        free(chart->series);
    }
    
    free(chart);
}

int chart_add_data_from_range(Chart* chart, Sheet* sheet, RangeSelection* range) {
    if (!chart || !sheet || !range || !range->is_active) return 0;
    
    // Calculate range dimensions
    int min_row = range->start_row < range->end_row ? range->start_row : range->end_row;
    int max_row = range->start_row > range->end_row ? range->start_row : range->end_row;
    int min_col = range->start_col < range->end_col ? range->start_col : range->end_col;
    int max_col = range->start_col > range->end_col ? range->start_col : range->end_col;
    
    int rows = max_row - min_row + 1;
    int cols = max_col - min_col + 1;
    
    // For now, assume first column is X values (or labels), rest are Y series
    if (cols < 2) return 0;  // Need at least 2 columns
    
    // Allocate series
    int num_series = cols - 1;
    chart->series = (ChartSeries*)calloc(num_series, sizeof(ChartSeries));
    chart->series_count = num_series;
    
    // Symbols for different series
    const char symbols[] = "*+ox#@$%&";
    
    // Check if first row contains headers
    int has_headers = 0;
    Cell* first_cell = sheet_get_cell(sheet, min_row, min_col);
    if (first_cell && first_cell->type == CELL_STRING) {
        has_headers = 1;
    }
    
    // Process each Y series
    for (int series_idx = 0; series_idx < num_series; series_idx++) {
        ChartSeries* series = &chart->series[series_idx];
        series->symbol = symbols[series_idx % strlen(symbols)];
        
        // Get series name from header if available
        if (has_headers) {
            Cell* header_cell = sheet_get_cell(sheet, min_row, min_col + series_idx + 1);
            if (header_cell && header_cell->type == CELL_STRING) {
                strncpy_s(series->name, sizeof(series->name), header_cell->data.string, _TRUNCATE);
            } else {
                sprintf_s(series->name, sizeof(series->name), "Series %d", series_idx + 1);
            }
        } else {
            sprintf_s(series->name, sizeof(series->name), "Series %d", series_idx + 1);
        }
        
        // Count valid data points
        int data_start = has_headers ? min_row + 1 : min_row;
        int point_count = 0;
        
        for (int row = data_start; row <= max_row; row++) {
            Cell* x_cell = sheet_get_cell(sheet, row, min_col);
            Cell* y_cell = sheet_get_cell(sheet, row, min_col + series_idx + 1);
            
            if (y_cell && (y_cell->type == CELL_NUMBER || 
                         (y_cell->type == CELL_FORMULA && y_cell->data.formula.error == ERROR_NONE))) {
                point_count++;
            }
        }
        
        if (point_count == 0) continue;
        
        // Allocate points
        series->points = (ChartPoint*)calloc(point_count, sizeof(ChartPoint));
        series->count = point_count;
        
        // Fill in data points
        int point_idx = 0;
        for (int row = data_start; row <= max_row; row++) {
            Cell* x_cell = sheet_get_cell(sheet, row, min_col);
            Cell* y_cell = sheet_get_cell(sheet, row, min_col + series_idx + 1);
            
            if (y_cell && (y_cell->type == CELL_NUMBER || 
                         (y_cell->type == CELL_FORMULA && y_cell->data.formula.error == ERROR_NONE))) {
                
                ChartPoint* point = &series->points[point_idx];
                
                // Get X value
                if (x_cell) {
                    switch (x_cell->type) {
                        case CELL_NUMBER:
                            point->x = x_cell->data.number;
                            break;
                        case CELL_FORMULA:
                            if (x_cell->data.formula.error == ERROR_NONE) {
                                point->x = x_cell->data.formula.cached_value;
                            } else {
                                point->x = point_idx;  // Use index as X
                            }
                            break;
                        case CELL_STRING:
                            // For string labels, use index as X value
                            point->x = point_idx;
                            strncpy_s(point->label, sizeof(point->label), x_cell->data.string, _TRUNCATE);
                            break;
                        default:
                            point->x = point_idx;
                            break;
                    }
                } else {
                    point->x = point_idx;
                }
                
                // Get Y value
                if (y_cell->type == CELL_NUMBER) {
                    point->y = y_cell->data.number;
                } else if (y_cell->type == CELL_FORMULA) {
                    point->y = y_cell->data.formula.cached_value;
                }
                
                // Update bounds
                if (point->x < chart->x_min) chart->x_min = point->x;
                if (point->x > chart->x_max) chart->x_max = point->x;
                if (point->y < chart->y_min) chart->y_min = point->y;
                if (point->y > chart->y_max) chart->y_max = point->y;
                
                point_idx++;
            }
        }
    }
    
    // Add some padding to bounds
    double x_range = chart->x_max - chart->x_min;
    double y_range = chart->y_max - chart->y_min;
    
    if (x_range < 0.0001) {
        chart->x_min -= 1.0;
        chart->x_max += 1.0;
    } else {
        chart->x_min -= x_range * 0.1;
        chart->x_max += x_range * 0.1;
    }
    
    if (y_range < 0.0001) {
        chart->y_min -= 1.0;
        chart->y_max += 1.0;
    } else {
        chart->y_min -= y_range * 0.1;
        chart->y_max += y_range * 0.1;
    }
    
    // Ensure Y axis includes zero for bar charts
    if (chart->config.type == CHART_BAR) {
        if (chart->y_min > 0) chart->y_min = 0;
        if (chart->y_max < 0) chart->y_max = 0;
    }
    
    return 1;
}

void chart_set_pixel(Chart* chart, int x, int y, char c) {
    if (x >= 0 && x < chart->canvas_width && y >= 0 && y < chart->canvas_height) {
        chart->canvas[y][x] = c;
    }
}

int chart_scale_x(Chart* chart, double value) {
    if (chart->x_max == chart->x_min) return 10;
    return 10 + (int)((value - chart->x_min) / (chart->x_max - chart->x_min) * (chart->config.width - 1));
}

int chart_scale_y(Chart* chart, double value) {
    if (chart->y_max == chart->y_min) return chart->config.height / 2;
    return chart->config.height - 1 - (int)((value - chart->y_min) / (chart->y_max - chart->y_min) * (chart->config.height - 1));
}

void chart_draw_line(Chart* chart, int x1, int y1, int x2, int y2, char symbol) {
    // Bresenham's line algorithm
    int dx = abs(x2 - x1);
    int dy = abs(y2 - y1);
    int sx = x1 < x2 ? 1 : -1;
    int sy = y1 < y2 ? 1 : -1;
    int err = dx - dy;
    
    while (1) {
        chart_set_pixel(chart, x1, y1, symbol);
        
        if (x1 == x2 && y1 == y2) break;
        
        int e2 = 2 * err;
        if (e2 > -dy) {
            err -= dy;
            x1 += sx;
        }
        if (e2 < dx) {
            err += dx;
            y1 += sy;
        }
    }
}

void chart_draw_axes(Chart* chart) {
    int x_axis_y = chart_scale_y(chart, 0);
    int y_axis_x = 8;  // More space for Y axis labels
    
    // Draw Y axis with double line for visibility
    for (int y = 0; y < chart->config.height; y++) {
        chart_set_pixel(chart, y_axis_x, y, '|');
        chart_set_pixel(chart, y_axis_x + 1, y, '|');
    }
    
    // Draw X axis with double line
    for (int x = y_axis_x; x < y_axis_x + chart->config.width + 2; x++) {
        if (x_axis_y >= 0 && x_axis_y < chart->config.height) {
            chart_set_pixel(chart, x, x_axis_y, '=');
        }
    }
    
    // Draw axis intersection
    if (x_axis_y >= 0 && x_axis_y < chart->config.height) {
        chart_set_pixel(chart, y_axis_x, x_axis_y, '#');
        chart_set_pixel(chart, y_axis_x + 1, x_axis_y, '#');
    }
    
    // Draw Y axis labels with better spacing (10 labels instead of 6)
    for (int i = 0; i <= 10; i++) {
        int y = (chart->config.height - 1) * i / 10;
        double value = chart->y_min + (chart->y_max - chart->y_min) * (10 - i) / 10;
        char label[16];
        
        // Format based on value magnitude
        if (fabs(value) >= 1000) {
            sprintf_s(label, sizeof(label), "%7.0f", value);
        } else if (fabs(value) >= 10) {
            sprintf_s(label, sizeof(label), "%7.1f", value);
        } else {
            sprintf_s(label, sizeof(label), "%7.2f", value);
        }
        
        for (int j = 0; j < 7 && j < (int)strlen(label); j++) {
            if (j < y_axis_x) {
                chart_set_pixel(chart, j, y, label[j]);
            }
        }
    }
      // Draw X axis labels with better spacing and support for string labels
    int num_x_labels = 8;  // More labels for better granularity
    
    // Check if we have string labels from data points
    int has_string_labels = 0;
    if (chart->series_count > 0 && chart->series[0].count > 0) {
        for (int i = 0; i < chart->series[0].count; i++) {
            if (strlen(chart->series[0].points[i].label) > 0) {
                has_string_labels = 1;
                break;
            }
        }
    }
    
    if (has_string_labels && chart->series_count > 0) {
        // Use actual data point labels
        ChartSeries* series = &chart->series[0];
        int max_labels = chart->config.width / 8;  // Limit labels to prevent overlap
        int step = series->count / max_labels;
        if (step < 1) step = 1;
        
        for (int i = 0; i < series->count; i += step) {
            int x = chart_scale_x(chart, series->points[i].x);
            const char* label = series->points[i].label;
            
            if (strlen(label) > 0) {
                int label_y = chart->config.height + 1;
                int label_x = x - (int)strlen(label) / 2;  // Center the label
                
                // Ensure label fits within canvas
                if (label_x < y_axis_x + 2) label_x = y_axis_x + 2;
                if (label_x + (int)strlen(label) >= chart->canvas_width) {
                    label_x = chart->canvas_width - (int)strlen(label) - 1;
                }
                
                for (int j = 0; j < (int)strlen(label) && label_x + j < chart->canvas_width; j++) {
                    chart_set_pixel(chart, label_x + j, label_y, label[j]);
                }
                
                // Draw tick mark
                if (x < chart->canvas_width) {
                    chart_set_pixel(chart, x, chart->config.height, '|');
                }
            }
        }
    } else {
        // Use numerical labels as before
        for (int i = 0; i <= num_x_labels; i++) {
            int x = y_axis_x + 2 + (chart->config.width - 2) * i / num_x_labels;
            double value = chart->x_min + (chart->x_max - chart->x_min) * i / num_x_labels;
            char label[16];
            
            if (fabs(value) >= 1000) {
                sprintf_s(label, sizeof(label), "%.0f", value);
            } else if (fabs(value) >= 10) {
                sprintf_s(label, sizeof(label), "%.1f", value);
            } else {
                sprintf_s(label, sizeof(label), "%.2f", value);
            }
            
            int label_y = chart->config.height + 1;
            int label_x = x - (int)strlen(label) / 2;  // Center the label
            
            for (int j = 0; j < (int)strlen(label); j++) {
                if (label_x + j >= 0 && label_x + j < chart->canvas_width) {
                    chart_set_pixel(chart, label_x + j, label_y, label[j]);
                }
            }
            
            // Draw tick mark
            if (x < chart->canvas_width) {
                chart_set_pixel(chart, x, chart->config.height, '|');
            }
        }
    }
    
    // Draw axis labels with better positioning
    // X axis label (centered)
    int x_label_pos = y_axis_x + 2 + chart->config.width / 2 - (int)strlen(chart->config.x_label) / 2;
    for (int i = 0; i < (int)strlen(chart->config.x_label); i++) {
        chart_set_pixel(chart, x_label_pos + i, chart->config.height + 3, chart->config.x_label[i]);
    }
    
    // Y axis label (vertically, centered)
    int y_label_pos = chart->config.height / 2 - (int)strlen(chart->config.y_label) / 2;
    for (int i = 0; i < (int)strlen(chart->config.y_label); i++) {
        if (y_label_pos + i >= 0 && y_label_pos + i < chart->config.height) {
            chart_set_pixel(chart, 0, y_label_pos + i, chart->config.y_label[i]);
        }
    }
}

void chart_plot_line_chart(Chart* chart) {
    chart_draw_axes(chart);
    
    // Draw grid if enabled with improved visibility
    if (chart->config.show_grid) {
        // Vertical grid lines (more lines for better granularity)
        for (int i = 1; i < 8; i++) {
            int x = 10 + (chart->config.width - 1) * i / 8;
            for (int y = 0; y < chart->config.height; y++) {
                if (chart->canvas[y][x] == ' ') {
                    chart_set_pixel(chart, x, y, '|');
                }
            }
        }
        
        // Horizontal grid lines (more lines)
        for (int i = 1; i < 10; i++) {
            int y = (chart->config.height - 1) * i / 10;
            for (int x = 10; x < 10 + chart->config.width; x++) {
                if (chart->canvas[y][x] == ' ' || chart->canvas[y][x] == '|') {
                    chart_set_pixel(chart, x, y, '-');
                }
            }
        }
    }
    
    // Use better symbols for different series - more visible
    const char symbols[] = "O*+X#@$%&^~!?><";
    
    // Plot each series
    for (int s = 0; s < chart->series_count; s++) {
        ChartSeries* series = &chart->series[s];
        series->symbol = symbols[s % strlen(symbols)];
        
        // Plot points and connect with lines
        for (int i = 0; i < series->count; i++) {
            int x = chart_scale_x(chart, series->points[i].x);
            int y = chart_scale_y(chart, series->points[i].y);
            
            // Draw the point with a larger marker
            if (x >= 0 && x < chart->canvas_width && y >= 0 && y < chart->canvas_height) {
                chart_set_pixel(chart, x, y, series->symbol);
                // Make points more visible with surrounding characters
                if (x > 0) chart_set_pixel(chart, x - 1, y, series->symbol);
                if (x < chart->canvas_width - 1) chart_set_pixel(chart, x + 1, y, series->symbol);
            }
            
            // Connect to previous point with better line drawing
            if (i > 0) {
                int prev_x = chart_scale_x(chart, series->points[i-1].x);
                int prev_y = chart_scale_y(chart, series->points[i-1].y);
                
                // Use different line styles for different series
                char line_char = (s == 0) ? '*' : (s == 1) ? '+' : (s == 2) ? 'x' : '.';
                chart_draw_line(chart, prev_x, prev_y, x, y, line_char);
            }
        }
    }      // Draw legend with improved positioning and formatting
    if (chart->config.show_legend && chart->series_count > 0) {
        // Position legend at the top right, with better bounds checking
        int legend_x = 10 + chart->config.width + 2;
        int legend_y = 2;
        
        // Check if legend fits horizontally, if not move it below the chart
        int max_legend_width = 0;
        for (int s = 0; s < chart->series_count; s++) {
            int width = 6 + (int)strlen(chart->series[s].name);  // symbol + " = " + name
            if (width > max_legend_width) max_legend_width = width;
        }
        
        // Always place legend on the right if there's any space, otherwise below
        if (legend_x + max_legend_width >= chart->canvas_width - 2) {
            // Move legend below the chart
            legend_x = 12;
            legend_y = chart->config.height + 3;
        }
        
        // Draw legend title
        const char* legend_title = "Legend:";
        for (int i = 0; i < (int)strlen(legend_title) && legend_x + i < chart->canvas_width; i++) {
            chart_set_pixel(chart, legend_x + i, legend_y - 1, legend_title[i]);
        }
        
        // Draw legend entries with better formatting
        for (int s = 0; s < chart->series_count && legend_y + s + 1 < chart->canvas_height; s++) {
            int current_y = legend_y + s;
            
            // Draw symbol with double characters for visibility
            chart_set_pixel(chart, legend_x, current_y, chart->series[s].symbol);
            chart_set_pixel(chart, legend_x + 1, current_y, chart->series[s].symbol);
            chart_set_pixel(chart, legend_x + 2, current_y, ' ');
            chart_set_pixel(chart, legend_x + 3, current_y, '=');
            chart_set_pixel(chart, legend_x + 4, current_y, ' ');
            
            // Draw series name with truncation if needed
            const char* name = chart->series[s].name;
            int max_name_len = chart->canvas_width - legend_x - 5 - 1;
            for (int i = 0; i < (int)strlen(name) && i < max_name_len; i++) {
                chart_set_pixel(chart, legend_x + 5 + i, current_y, name[i]);
            }
            
            // Add ellipsis if name was truncated
            if ((int)strlen(name) > max_name_len && max_name_len > 3) {
                chart_set_pixel(chart, legend_x + 5 + max_name_len - 3, current_y, '.');
                chart_set_pixel(chart, legend_x + 5 + max_name_len - 2, current_y, '.');
                chart_set_pixel(chart, legend_x + 5 + max_name_len - 1, current_y, '.');
            }
        }
    }
}

void chart_plot_bar_chart(Chart* chart) {
    chart_draw_axes(chart);
    
    if (chart->series_count == 0 || chart->series[0].count == 0) return;
    
    // For bar charts, we'll use the first series only (for simplicity)
    ChartSeries* series = &chart->series[0];
    
    // Calculate bar width with better spacing
    int total_width = chart->config.width - 4;
    int bar_width = total_width / series->count - 2;  // Leave space between bars
    if (bar_width < 3) bar_width = 3;
    if (bar_width > 12) bar_width = 12;
    
    int spacing = 2;
    if (series->count * (bar_width + spacing) > total_width) {
        spacing = 1;
        bar_width = (total_width / series->count) - spacing;
    }
    
    // Draw bars
    for (int i = 0; i < series->count; i++) {
        int bar_x = 12 + i * (bar_width + spacing);
        int bar_top = chart_scale_y(chart, series->points[i].y);
        int bar_bottom = chart_scale_y(chart, 0);
        
        if (bar_top > bar_bottom) {
            int temp = bar_top;
            bar_top = bar_bottom;
            bar_bottom = temp;
        }
        
        // Draw the bar with better visual appearance
        for (int y = bar_top; y <= bar_bottom && y < chart->config.height; y++) {
            for (int x = 0; x < bar_width; x++) {
                if (bar_x + x < 10 + chart->config.width) {
                    char bar_char = '#';
                    // Create 3D effect with different characters
                    if (x == 0) {
                        bar_char = '[';
                    } else if (x == bar_width - 1) {
                        bar_char = ']';
                    } else if (y == bar_top) {
                        bar_char = '=';
                    } else {
                        bar_char = '#';
                    }
                    chart_set_pixel(chart, bar_x + x, y, bar_char);
                }
            }
        }
        
        // Draw value on top of bar with better formatting
        char value_str[32];
        if (fabs(series->points[i].y) >= 1000) {
            sprintf_s(value_str, sizeof(value_str), "%.0f", series->points[i].y);
        } else {
            sprintf_s(value_str, sizeof(value_str), "%.1f", series->points[i].y);
        }
        
        int value_x = bar_x + (bar_width - (int)strlen(value_str)) / 2;
        int value_y = bar_top - 1;
        
        if (value_y >= 0 && value_y < chart->config.height) {
            for (int j = 0; j < (int)strlen(value_str); j++) {
                if (value_x + j < 10 + chart->config.width) {
                    chart_set_pixel(chart, value_x + j, value_y, value_str[j]);
                }
            }
        }
          // Draw label below bar with better handling for string labels
        const char* label_text = "";
        if (strlen(series->points[i].label) > 0) {
            label_text = series->points[i].label;
        } else {
            // Use index-based label if no string label available
            static char index_label[16];
            sprintf_s(index_label, sizeof(index_label), "Item %d", i + 1);
            label_text = index_label;
        }
        
        int label_y = chart->config.height + 2;
        int label_x = bar_x;
        
        // Handle label length intelligently
        int label_len = (int)strlen(label_text);
        if (label_len <= bar_width) {
            // Center the label under the bar
            int offset = (bar_width - label_len) / 2;
            for (int j = 0; j < label_len; j++) {
                if (label_x + offset + j < chart->canvas_width) {
                    chart_set_pixel(chart, label_x + offset + j, label_y, label_text[j]);
                }
            }
        } else {
            // Label is too long - try different strategies
            if (label_len <= bar_width + 2) {
                // Slightly extend beyond bar if close
                for (int j = 0; j < label_len && label_x + j < chart->canvas_width; j++) {
                    chart_set_pixel(chart, label_x + j, label_y, label_text[j]);
                }
            } else if (bar_width >= 6) {
                // Truncate with ellipsis
                for (int j = 0; j < bar_width - 3; j++) {
                    chart_set_pixel(chart, label_x + j, label_y, label_text[j]);
                }
                chart_set_pixel(chart, label_x + bar_width - 3, label_y, '.');
                chart_set_pixel(chart, label_x + bar_width - 2, label_y, '.');
                chart_set_pixel(chart, label_x + bar_width - 1, label_y, '.');
            } else {
                // Very narrow bars - use vertical text or abbreviation
                for (int j = 0; j < bar_width && j < label_len; j++) {
                    chart_set_pixel(chart, label_x + j, label_y, label_text[j]);
                }
            }
        }
    }
}

void chart_plot_pie_chart(Chart* chart) {
    if (chart->series_count == 0 || chart->series[0].count == 0) return;
    
    // For pie charts, use first series and interpret as categories
    ChartSeries* series = &chart->series[0];
    
    // Calculate total
    double total = 0;
    for (int i = 0; i < series->count; i++) {
        if (series->points[i].y > 0) {
            total += series->points[i].y;
        }
    }
    
    if (total == 0) return;
    
    // Pie chart parameters - much larger
    int center_x = chart->config.width / 2 + 5;
    int center_y = chart->config.height / 2;
    int radius = (chart->config.height / 2) - 3;
    if (radius > 25) radius = 25;  // Increased max radius
    
    // ASCII art pie slices using different characters
    const char* slice_chars = "@#$%&*+=~-:.|ox+";
    int num_chars = (int)strlen(slice_chars);
    
    // Draw the pie with better resolution
    for (int y = -radius; y <= radius; y++) {
        for (int x = -radius * 2; x <= radius * 2; x++) {
            double dx = x / 2.0;  // Correct for aspect ratio
            double dy = y;
            double dist = sqrt(dx * dx + dy * dy);
            
            if (dist <= radius) {
                // Calculate angle
                double angle = atan2(dy, dx);
                if (angle < 0) angle += 2 * 3.14159265359;
                
                // Find which slice this belongs to
                double cumulative = 0;
                int slice_idx = -1;
                
                for (int i = 0; i < series->count; i++) {
                    if (series->points[i].y <= 0) continue;
                    
                    double slice_angle = (series->points[i].y / total) * 2 * 3.14159265359;
                    if (angle >= cumulative && angle < cumulative + slice_angle) {
                        slice_idx = i;
                        break;
                    }
                    cumulative += slice_angle;
                }
                
                if (slice_idx >= 0) {
                    char c = slice_chars[slice_idx % num_chars];
                    // Add border effect
                    if (dist > radius - 1) {
                        c = '*';
                    }
                    chart_set_pixel(chart, center_x + x, center_y + y, c);
                }
            }
        }
    }
    
    // Draw legend with better formatting
    int legend_x = 5;
    int legend_y = 2;
    
    // Draw legend title
    const char* legend_title = "Legend:";
    for (int i = 0; i < (int)strlen(legend_title); i++) {
        chart_set_pixel(chart, legend_x + i, legend_y - 1, legend_title[i]);
    }
    
    for (int i = 0; i < series->count && legend_y + i + 1 < chart->config.height; i++) {
        if (series->points[i].y <= 0) continue;
        
        double percentage = (series->points[i].y / total) * 100;
        
        // Draw slice character with padding
        chart_set_pixel(chart, legend_x, legend_y + i + 1, slice_chars[i % num_chars]);
        chart_set_pixel(chart, legend_x + 1, legend_y + i + 1, slice_chars[i % num_chars]);
        chart_set_pixel(chart, legend_x + 2, legend_y + i + 1, ' ');
        chart_set_pixel(chart, legend_x + 3, legend_y + i + 1, '-');
        chart_set_pixel(chart, legend_x + 4, legend_y + i + 1, ' ');
        
        // Draw label and percentage with better formatting
        char legend_text[80];
        if (strlen(series->points[i].label) > 0) {
            sprintf_s(legend_text, sizeof(legend_text), "%-15s: %8.1f (%5.1f%%)", 
                     series->points[i].label, series->points[i].y, percentage);
        } else {
            sprintf_s(legend_text, sizeof(legend_text), "Item %-10d: %8.1f (%5.1f%%)", 
                     i + 1, series->points[i].y, percentage);
        }
        
        for (int j = 0; j < (int)strlen(legend_text) && legend_x + 5 + j < chart->canvas_width; j++) {
            chart_set_pixel(chart, legend_x + 5 + j, legend_y + i + 1, legend_text[j]);
        }
    }
    
    // Draw title
    const char* pie_title = "Distribution:";
    int title_x = center_x - (int)strlen(pie_title) / 2;
    int title_y = center_y - radius - 2;
    if (title_y >= 0) {
        for (int i = 0; i < (int)strlen(pie_title); i++) {
            chart_set_pixel(chart, title_x + i, title_y, pie_title[i]);
        }
    }
}

void chart_render(Chart* chart) {
    if (!chart) return;
    
    // Clear canvas
    for (int y = 0; y < chart->canvas_height; y++) {
        memset(chart->canvas[y], ' ', chart->canvas_width);
        chart->canvas[y][chart->canvas_width] = '\0';
    }
    
    // Render based on chart type
    switch (chart->config.type) {
        case CHART_LINE:
            chart_plot_line_chart(chart);
            break;
        case CHART_BAR:
            chart_plot_bar_chart(chart);
            break;
        case CHART_PIE:
            chart_plot_pie_chart(chart);
            break;
        case CHART_SCATTER:
            chart_plot_line_chart(chart);  // Use line chart without lines for now
            break;
    }
}

void chart_display(Chart* chart, Console* console, int start_x, int start_y) {
    if (!chart || !console) return;
    
    // Display the chart on the console
    WORD chartColor = MAKE_COLOR(COLOR_WHITE, COLOR_BLACK);
    
    for (int y = 0; y < chart->canvas_height && start_y + y < console->height; y++) {
        console_write_string(console, start_x, start_y + y, chart->canvas[y], chartColor);
    }
}

char** chart_get_output(Chart* chart, int* line_count) {
    if (!chart || !line_count) return NULL;
    
    *line_count = chart->canvas_height;
    return chart->canvas;
}

// Helper function to parse chart command
int parse_chart_command(const char* command, ChartType* type, char* x_label, char* y_label) {
    // Format: "chart_type x_label y_label"
    // Example: "line year revenue"
    
    char type_str[32] = {0};
    int n = sscanf_s(command, "%31s %63s %63s", type_str, (unsigned)sizeof(type_str), 
                     x_label, (unsigned)64, y_label, (unsigned)64);
    
    if (n < 1) return 0;
    
    // Determine chart type
    if (strcmp(type_str, "line") == 0) {
        *type = CHART_LINE;
    } else if (strcmp(type_str, "bar") == 0) {
        *type = CHART_BAR;
    } else if (strcmp(type_str, "pie") == 0) {
        *type = CHART_PIE;
    } else if (strcmp(type_str, "scatter") == 0) {
        *type = CHART_SCATTER;
    } else {
        return 0;  // Unknown chart type
    }
    
    // If labels weren't provided, use defaults
    if (n < 2) strcpy_s(x_label, 64, "X");
    if (n < 3) strcpy_s(y_label, 64, "Y");
    
    return 1;
}

// Helper function to parse chart command with optional size
int parse_chart_command_with_size(const char* command, ChartType* type, char* x_label, char* y_label, int* width, int* height) {
    // Format: "chart_type x_label y_label [width] [height]"
    // Example: "line year revenue 120 40"
    
    char type_str[32] = {0};
    *width = 100;  // Default
    *height = 35;  // Default
    
    int n = sscanf_s(command, "%31s %63s %63s %d %d", 
                     type_str, (unsigned)sizeof(type_str), 
                     x_label, (unsigned)64, 
                     y_label, (unsigned)64,
                     width, height);
    
    if (n < 1) return 0;
    
    // Determine chart type
    if (strcmp(type_str, "line") == 0) {
        *type = CHART_LINE;
    } else if (strcmp(type_str, "bar") == 0) {
        *type = CHART_BAR;
    } else if (strcmp(type_str, "pie") == 0) {
        *type = CHART_PIE;
    } else if (strcmp(type_str, "scatter") == 0) {
        *type = CHART_SCATTER;
    } else {
        return 0;  // Unknown chart type
    }
    
    // If labels weren't provided, use defaults
    if (n < 2) strcpy_s(x_label, 64, "X");
    if (n < 3) strcpy_s(y_label, 64, "Y");
    
    return 1;
}

// Function to display chart in a popup-style window
void display_chart_popup(Console* console, Chart* chart, const char* title) {
    if (!console || !chart) return;
    
    // Save current screen (simplified - in practice you'd save the buffer)
    console_clear(console);
    
    // Use almost the full screen space with minimal borders
    int border_width = console->width - 2;
    int border_height = console->height - 2;
    
    // Center the chart on screen
    int border_x = 1;
    int border_y = 1;
    
    // Draw simple border for full-screen experience
    for (int x = 0; x < console->width; x++) {
        console_write_char(console, x, 0, '=', MAKE_COLOR(COLOR_WHITE | COLOR_BRIGHT, COLOR_BLACK));
        console_write_char(console, x, console->height - 1, '=', MAKE_COLOR(COLOR_WHITE | COLOR_BRIGHT, COLOR_BLACK));
    }
    for (int y = 0; y < console->height; y++) {
        console_write_char(console, 0, y, '|', MAKE_COLOR(COLOR_WHITE | COLOR_BRIGHT, COLOR_BLACK));
        console_write_char(console, console->width - 1, y, '|', MAKE_COLOR(COLOR_WHITE | COLOR_BRIGHT, COLOR_BLACK));
    }
    
    // Draw corners
    console_write_char(console, 0, 0, '#', MAKE_COLOR(COLOR_WHITE | COLOR_BRIGHT, COLOR_BLACK));
    console_write_char(console, console->width - 1, 0, '#', MAKE_COLOR(COLOR_WHITE | COLOR_BRIGHT, COLOR_BLACK));
    console_write_char(console, 0, console->height - 1, '#', MAKE_COLOR(COLOR_WHITE | COLOR_BRIGHT, COLOR_BLACK));
    console_write_char(console, console->width - 1, console->height - 1, '#', MAKE_COLOR(COLOR_WHITE | COLOR_BRIGHT, COLOR_BLACK));
    
    // Draw title with better formatting
    if (title && strlen(title) > 0) {
        int title_x = (console->width - (int)strlen(title)) / 2;
        console_write_string(console, title_x - 1, 0, "[", MAKE_COLOR(COLOR_YELLOW | COLOR_BRIGHT, COLOR_BLACK));
        console_write_string(console, title_x, 0, title, MAKE_COLOR(COLOR_YELLOW | COLOR_BRIGHT, COLOR_BLACK));
        console_write_string(console, title_x + (int)strlen(title), 0, "]", MAKE_COLOR(COLOR_YELLOW | COLOR_BRIGHT, COLOR_BLACK));
    }      // Display chart using the full available space (minus borders)
    int display_start_x = border_x;
    int display_start_y = border_y;
    int display_width = border_width;
    int display_height = border_height;
    
    // Fill the entire available area with the chart
    int chart_rows = chart->canvas_height;
    int chart_cols = chart->canvas_width;
    
    for (int y = 0; y < chart_rows && y < display_height; y++) {
        for (int x = 0; x < chart_cols && x < display_width; x++) {
            if (display_start_x + x < console->width && display_start_y + y < console->height) {
                // Enhanced color coding for much better visibility
                WORD color = MAKE_COLOR(COLOR_WHITE, COLOR_BLACK);
                char c = chart->canvas[y][x];
                
                // Color coding based on character type for better visibility
                if (c == '|' || c == '-' || c == '=' || c == '#') {
                    // Axes and grid
                    color = MAKE_COLOR(COLOR_CYAN | COLOR_BRIGHT, COLOR_BLACK);
                } else if (c == 'O' || c == '*') {
                    // First series - bright yellow
                    color = MAKE_COLOR(COLOR_YELLOW | COLOR_BRIGHT, COLOR_BLACK);
                } else if (c == '+' || c == 'X') {
                    // Second series - bright green
                    color = MAKE_COLOR(COLOR_GREEN | COLOR_BRIGHT, COLOR_BLACK);
                } else if (c == 'x' || c == '@') {
                    // Third series - bright magenta
                    color = MAKE_COLOR(COLOR_MAGENTA | COLOR_BRIGHT, COLOR_BLACK);
                } else if (c == '$' || c == '%') {
                    // Fourth series - bright red
                    color = MAKE_COLOR(COLOR_RED | COLOR_BRIGHT, COLOR_BLACK);
                } else if (c == '[' || c == ']') {
                    // Bar chart borders
                    color = MAKE_COLOR(COLOR_BLUE | COLOR_BRIGHT, COLOR_BLACK);
                } else if (c == '#' && y > 0 && chart->canvas[y-1][x] == '#') {
                    // Bar chart fill
                    color = MAKE_COLOR(COLOR_BLUE | COLOR_BRIGHT, COLOR_BLACK);
                } else if (isdigit(c) || c == '.') {
                    // Numbers and decimal points
                    color = MAKE_COLOR(COLOR_WHITE | COLOR_BRIGHT, COLOR_BLACK);
                } else if (isalpha(c)) {
                    // Letters (labels, legend text)
                    color = MAKE_COLOR(COLOR_WHITE, COLOR_BLACK);
                }
                
                console_write_char(console, display_start_x + x, display_start_y + y, c, color);
            }
        }
    }    
    // Draw instructions with better visibility at the bottom
    const char* instructions = "[ Press any key to close ]";
    int inst_x = (console->width - (int)strlen(instructions)) / 2;
    console_write_string(console, inst_x, console->height - 1, instructions, 
                        MAKE_COLOR(COLOR_CYAN | COLOR_BRIGHT, COLOR_BLACK));
    
    console_flip(console);
}

#endif // CHARTS_H
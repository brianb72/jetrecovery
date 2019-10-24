// 2017 BrianB
/*
Program Description:
    A customer was operating an in-house testing rig that used a pneumatic pump and sensors to measure the long term
    performance of valves. Over one year of data was recorded in a Microsoft Access database.

    The first ~10.5 megabytes of their database file was overwritten with random data. The customer attempted to use
    several commercial recovery tools without success. The only backup of the file was months old. The customer sent
    me a copy of the corrupted database, as well as their old backup.

    The files were MDB Jet4 databases. The damage to the begining of the newer file destroyed the database definition
    page and all table definitions. By looking at the old file I obtained table names and definitions.

    I wrote this utility in C to scan the raw MDB file and recover rows of the customers data. All of the data that
    was needed was located in one table, 'tblResults', which I extracted to a CSV file.

    The corrupted file had an auto-incrementing ID column, and the maximum row value in the data was 2,149,759. The
    first row observed after the corrupted data was 405,378. The older file had row ID's of 62 to ~900,000, and I was
    able to copy data from the older file to replace the missing corrupted data. I spot checked overlapping ID's
    between the old and new files and the data matched.

    'rewrite.py' is a short script that converts the datetime column in the CSV to the proper value.
*/

#define _GNU_SOURCE
#include <stdio.h>
#include <string.h>
#include <ctype.h>
#include <stdint.h>
#include <inttypes.h>

// Globals
int current_page_loaded = -1;   // Which 4096 byte page is currently loaded from the source file
int total_rows_in_csv = 0;

// Jet4 Data Page Header
struct data_page_header {
    uint8_t  page_type;          // Page type   0x01 = data page
    uint8_t  unknown1;           // Always 0x01
    uint16_t free_space;         // Free space in this page
    uint32_t tdef_pg;            // Page pointer to table definition
    uint32_t unknown2;           // Unknown field
    uint16_t num_rows;           // Number of rowss this page
    uint16_t row_offset[];        // Table of offsets for each row starting, first row starts at end of page and moves up.
} __attribute__((packed));

// Fixed Fields for Table Definition page 46
struct fixed_fields {
    uint32_t ID;                     // autonumber integer
    uint64_t DateTime;               // 8 byte DateTime value, a standard 'double'
    uint32_t CalibrationSetpoint1;   // single 3 digit precision, a standard 'float'
    uint32_t CalibrationSetpoint2;
    uint32_t CalibrationSetpoint3;
    uint32_t CalibrationSetpoint4;
    uint32_t CalibrationCheckRiseSetpoint;
    uint32_t CalibrationCheckFallSetpoint;
    uint32_t CalibrationCheckDiffPressure;
    uint16_t CalibrationCount;       // integer
    uint32_t CheckRiseSetpoint;
    uint32_t CheckFallSetpoint;
    uint32_t CheckDiffPressure;
    uint16_t CalibrationSystem;      // integer
    uint16_t CalibrationCell;        // integer
} __attribute__((packed));

// For debugging, print a hex dump on the left and printable characters on the right, with address labels in the row and column headers.
void print_hex(char *hex, int hex_length) {
    const int char_per_line = (8 + 8 + 8);
    char hex_string[1024];
    char char_string[1024];
    char buf[2048];
    int i;
    int printed_chars = 0;
    int printed_lines = 0;
    
    hex_string[0] = char_string[0] = 0;

    // Print the column index at the top
    printf("     ");
    for (i = 0; i < char_per_line; ++i) {
        printf("%02u ", i);
    }
    printf("\n");

    // Print the line under the index    
    printf("     ");
    for (i = 0; i < char_per_line; ++i) {
        printf("---");
    }
    printf("\n");


    // Print each row
    for (i = 0; i < hex_length; ++i) {
        // Add the hex string to the hex string
        sprintf(buf, "%2.2X ", hex[i] & 0xFF);
        strcat(hex_string, buf);       

        // If the character is printable add it to the string buffer, otherwise add '.'
        if (isprint(hex[i])) {
            sprintf(buf, "%c", hex[i]);
            strcat(char_string, buf);
        } else {
            strcat(char_string, ".");
        }
        ++printed_chars;

        // If we've built up enough characters, print them out with a row index
        if ((i + 1) % char_per_line == 0) {
            printed_chars = 0;
            int line_num = char_per_line * printed_lines;
            
            printf("%02u | %s | %s\n", line_num, hex_string, char_string);
            hex_string[0] = char_string[0] = 0;
            ++printed_lines;
        }
    }    
    
    // When hex_length is hit pad the rest of the row with spaces
    if (printed_chars < char_per_line) {
        int n_more = char_per_line - printed_chars;
        for (int t = 0; t < n_more; ++t) {
            strcat(char_string, " ");
            strcat(hex_string, "   ");
        }
        
    }
    // Print the last row with row index
    if (strlen(hex_string) > 0) {
        printf("%02u | %s | %s\n", char_per_line * printed_lines, hex_string, char_string);
    }
}

#define BOUNDS(p_pointer, p_target, target_length) ( \
    (p_pointer >= p_target) \
    && (p_pointer <= &p_target[target_length]) )

// Process a row of data, extract the columns, write to CSV
int process_data_row(int row_number, unsigned char *row_data, unsigned int row_length, FILE *outfp) {
    // Row pointer will point inside row_data
    unsigned char *rp = row_data;

    // The first word is the number of columns in this row
    unsigned int num_columns_in_row = *((uint16_t *)rp);

    // The number of bytes of the null field at the end of the record uses this formula
    unsigned int null_size = (num_columns_in_row + 7) / 8;
    
    // The word before the null field is the number of variable columns.
    rp = &row_data[row_length - null_size - 2];
    if (!BOUNDS(rp, row_data, row_length)) {
        printf("Page [%i] Row [%i] out of bounds fetching num_variable_columns.\n", current_page_loaded, row_number);
        print_hex(row_data, row_length);
        return 0;
    }
    unsigned int num_variable_columns = *((uint16_t *)rp);

    // Find the EOD offset
    rp = &row_data[row_length - null_size - 2 - (num_variable_columns * 2) - 2];
    if (!BOUNDS(rp, row_data, row_length)) {
        printf("Page [%i] Row [%i] out of bounds fetching EOD offset.\n", current_page_loaded, row_number);
        print_hex(row_data, row_length);
        return 0;
    }
    unsigned int eod_value = *((int16_t *)rp);
    unsigned int eod_address = (unsigned int)(rp - row_data);
    
    // Verify the EOD offset matches it's actual address
    if (eod_value != eod_address) {
        printf("ERROR MISMATCH - EOD Address: %Xh, EOD Value: %Xh\n", eod_address, eod_value);
        return 0;
    }

    // Buffer to hold variable length strings
    #define MAX_VARCOL_COUNT         10
    #define MAX_VARCOL_STRING_LENGTH 512
    char string_array[MAX_VARCOL_COUNT][MAX_VARCOL_STRING_LENGTH];

    // General check
    // Check that there are not too many variable columns for our string array
    if (num_variable_columns >= MAX_VARCOL_COUNT) {
        printf("Page [%i] Row [%i] too many variable columns.\n", current_page_loaded, row_number);
        print_hex(row_data, row_length);
        return 0;
    }

    // Check that we have the required number of variable columns for this type of row
    // TODO Specific to table #46
    if (num_variable_columns != 3) {
        printf("Page [%i] Row [%i] does not have 3 variable columns.\n", current_page_loaded, row_number);
        print_hex(row_data, row_length);
        return 0;
    }

    // Advance rp past the EOD word so that it points to the first entry in the variable column offset table
    rp += 2;
    

    unsigned int var_last_offset = 0;

    // Locate the offset and length of each variable column, then copy out it's string
    for (int i = 0; i < num_variable_columns; ++i) {
        // Bounds check rp
        if (!BOUNDS(rp, row_data, row_length)) {
            printf("Page [%i] Row [%i] out of bounds fetching EOD offset.\n", current_page_loaded, row_number);
            print_hex(row_data, row_length);
            return 0;
        }

        // Extract the offset and calculate the length
        unsigned int var_offset = *((int16_t *)rp);
        unsigned int var_length;
        if (i > 0) {
            var_length = var_last_offset - var_offset;
        } else {
            var_length = eod_address - var_offset;
        }

        // Extract the string if length is nonzero, otherwise set as a null string
        if (var_length > 0) {
            // Strip out unicode byte order marker if exists
            if (row_data[var_offset] == 0xFF && row_data[var_offset+1] == 0xFE) {
                memcpy(string_array[i], &row_data[var_offset+2], var_length-2);
                string_array[i][var_length-2] = 0;
            } else {
                memcpy(string_array[i], &row_data[var_offset], var_length);
                string_array[i][var_length] = 0;
            }
        } else {
            // A zero length entry is a null string
            string_array[i][0] = 0;
        }

        // Save the offset 
        var_last_offset = var_offset;

        // Advance rp to the next word in the offset table
        rp += 2;
    }

    // string_array[] is   0 = errorcode,  1 = modelno,   2 = workorder
    
    // Create pointer to fixed fields, which is right after the first number of columns word.
    struct fixed_fields *fixed = (struct fixed_fields *)&row_data[2];

    // Now that all fields have been collected, print out this row to the CSV
    fprintf(outfp, "%i,%lf,%s,%s,%f,%f,%f,%f,%f,%f,%f,%i,%f,%f,%f,%s,%i,%i\n",
            fixed->ID,
            *(double*)(&fixed->DateTime),
            string_array[1],
            string_array[2],
            *(float*)(&fixed->CalibrationSetpoint1),
            *(float*)(&fixed->CalibrationSetpoint2),
            *(float*)(&fixed->CalibrationSetpoint3),
            *(float*)(&fixed->CalibrationSetpoint4),
            *(float*)(&fixed->CalibrationCheckRiseSetpoint),
            *(float*)(&fixed->CalibrationCheckFallSetpoint),
            *(float*)(&fixed->CalibrationCheckDiffPressure),
            fixed->CalibrationCount,
            *(float*)(&fixed->CheckRiseSetpoint),
            *(float*)(&fixed->CheckFallSetpoint),
            *(float*)(&fixed->CheckDiffPressure),
            string_array[0],
            fixed->CalibrationSystem,
            fixed->CalibrationCell
    );

    ++total_rows_in_csv;

    return 1;
}




// Processes each page, scan each row in the row offset table, and process each row
int process_data_page(unsigned char *page_data, FILE *outfp) {
    // Get a pointer to the header for this page
    struct data_page_header *header = (struct data_page_header *)page_data;

    // Walk through the row offset table
    for (int i = 0; i < header->num_rows; ++i) {
        // Calculate row start and end positions within page_data
        unsigned int row_start = header->row_offset[i];
        unsigned int row_end;
        if (i == 0) {
            row_end = 4096 - 1;
        } else {
            row_end = header->row_offset[i-1] - 1;
        }
        unsigned int row_length = row_end - row_start + 1;

        // Bounds check, must be within the 4096 byte page
        if (row_start > 4096 || (row_start + row_length) > 4096) {
            printf("Error - Page [%i] Row [%i] indexes out of bounds  (start: %i length: %i)", 
                current_page_loaded, i, row_start, row_length);
            return 0;
        }

        int r = process_data_row(i, &page_data[row_start], row_length, outfp);
        if (!r) return 0;
    }
    return 1;
}




// Loads the next 4096 byte data page from a Jet4 database that has the specified table definition pointer
// Return 0 if read error and no more data to read
// Return 1 if read success and possibly more data to read
int load_next_data_page(FILE *fp, unsigned char *page_data, int table_page_pointer) {
    do {
        // Read one page of data
        int r = fread(page_data, 1, 1024*4, fp);
        ++current_page_loaded;
        if (r < 1024*4) {
            printf("Page %i Incomplete Read %i\n", current_page_loaded, r);
            long cur_pos = ftell(fp);
            fseek(fp, 0, SEEK_END);
            long end_pos = ftell(fp);
            printf("%lu / %lu\n", cur_pos, end_pos);

            return 0;   // Fail, no more data, do not call again
        }

        // Is this a valid data page?
        struct data_page_header *header = (struct data_page_header *)page_data;
        if (header->page_type == 0x01 && header->unknown1 == 0x01 && header->tdef_pg == table_page_pointer) {
            printf("Data found page %i - Row count %i\n", current_page_loaded, header->num_rows);
            return 1;   // Success, possibly more data
        }
    } while (1);
}



/* **************************************************************************************** */
/* **************************************************************************************** */
/* **************************************************************************************** */


int main() {
    FILE *fp, *outfp;
    unsigned char page_data[1024*1024];      // Big buffer for loading data from file
    const char *csv_header = "ID,DateTime,ModelNo,WorkOrder,CalibrationSetpoint1,CalibrationSetpoint2,CalibrationSetpoint3,CalibrationSetpoint4,"
                "CalibrationCheckRiseSetpoint,CalibrationCheckFallSetpoint,CalibrationCheckDiffPressure,CalibrationCount,CheckRiseSetpoint,"
                "CheckFallSetpoint,CheckDiffPressure,ErrorCode,CalibrationSystem,CalibrationCell";
    
    // Open the input file
    fp = fopen("calbad.mdb", "rb");
    if (!fp) {
        printf("Could not open file.\n");
        return 1;
    }

    // Open the output file
    outfp = fopen("output.csv", "w");
    if (!outfp) {
        printf("Could not open output file.\n");
    }

    fprintf(outfp, "%s\n", csv_header);
    
    while (load_next_data_page(fp, page_data, 46)) {
       int r = process_data_page(page_data, outfp);
       if (!r) return 0;
    }

    // Close files
    fclose(fp);
    fclose(outfp);

    // Exit
    printf("Wrote %i rows to CSV.\n", total_rows_in_csv);
    printf("Done.\n");
    return 0;
}
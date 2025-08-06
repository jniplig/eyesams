"""
Excel File Processor - Merge multiple Excel files with error handling

This module provides functionality to process multiple Excel files,
extracting data from multiple worksheets and combining them into
a single consolidated output file.

Author: Your Name
Date: Today's Date
"""

import pandas as pd
import os
import glob


def process_excel_files(input_directory=None, output_directory=None):
    """
    Process multiple Excel files and merge them into a single output file.
    
    Args:
        input_directory (str): Path to directory containing Excel files
        output_directory (str): Path to directory for output file
    
    Returns:
        str: Path to output file if successful, None if failed
    """
    
    # Set default directories if not provided
    if input_directory is None:
        input_directory = "/content/drive/MyDrive/ISAMS PROJECT/UPLOADS"
    
    if output_directory is None:
        output_directory = "/content/drive/MyDrive/ISAMS PROJECT"
    
    print(f"Starting Excel file processing...")
    print(f"Input directory: {input_directory}")
    print(f"Output directory: {output_directory}")
    
    # Check if input directory exists
    if not os.path.exists(input_directory):
        print(f"Error: Input directory {input_directory} does not exist!")
        return None
    
    # Check if output directory exists
    if not os.path.exists(output_directory):
        print(f"Error: Output directory {output_directory} does not exist!")
        return None
    
    # Find Excel files with error handling
    try:
        files = glob.glob(os.path.join(input_directory, "*.xlsx"))
        if not files:
            print(f"Warning: No Excel files found in {input_directory}")
            return None
        print(f"Found {len(files)} Excel files to process")
    except Exception as e:
        print(f"Error searching for files: {e}")
        return None
    
    # Initialize data collection
    merged_data = []
    processing_stats = {
        'files_found': len(files),
        'files_processed': 0,
        'sheets_processed': 0,
        'errors': []
    }
    
    # Process each Excel file
    for file in files:
        file_name = os.path.basename(file)
        print(f"\nProcessing file: {file_name}")
        
        try:
            excel_file = pd.ExcelFile(file)
            sheets_in_file = 0
        except Exception as e:
            error_msg = f"Error opening {file_name}: {e}"
            print(error_msg)
            processing_stats['errors'].append(error_msg)
            continue  # Skip this file, continue with others
        
        # Process each sheet in the current file
        for sheet_name in excel_file.sheet_names:
            try:
                print(f"  Processing sheet: {sheet_name}")
                
                # Load sheet data
                df = pd.read_excel(file, sheet_name=sheet_name, header=None)
                
                # Validate data structure
                if df.empty:
                    print(f"    Warning: Sheet {sheet_name} is completely empty")
                    continue
                
                if df.shape[0] < 3:
                    print(f"    Warning: Sheet {sheet_name} has insufficient rows ({df.shape[0]})")
                    continue
                
                if df.iloc[2:-2].empty:
                    print(f"    Skipping: Sheet {sheet_name} has no data rows after header/footer removal")
                    continue
                
                # Extract teacher information with validation
                teacher_cell = df.iloc[0, 0]
                if pd.isna(teacher_cell):
                    print(f"    Warning: No teacher data in first cell of {sheet_name}")
                    teacher = "UNKNOWN"
                elif len(str(teacher_cell)) < 5:
                    print(f"    Warning: Teacher data too short in {sheet_name}: '{teacher_cell}'")
                    teacher = str(teacher_cell)  # Use whatever we have
                else:
                    teacher = str(teacher_cell)[-5:]
                
                # Process data structure
                df.columns = df.iloc[1]  # Set column headers from row 2
                df = df.iloc[2:-2].copy()  # Keep only data rows
                
                # Add metadata columns
                df.loc[:, 'Teacher'] = teacher
                df.loc[:, 'Set'] = sheet_name
                
                # Store processed data
                merged_data.append(df)
                processing_stats['sheets_processed'] += 1
                sheets_in_file += 1
                print(f"    ✓ Successfully processed {sheet_name} ({len(df)} rows)")
                
            except Exception as e:
                error_msg = f"Error processing sheet {sheet_name} in {file_name}: {e}"
                print(f"    ✗ {error_msg}")
                processing_stats['errors'].append(error_msg)
                continue
        
        if sheets_in_file > 0:
            processing_stats['files_processed'] += 1
            print(f"  ✓ Completed {file_name}: {sheets_in_file} sheets processed")
        else:
            print(f"  ✗ No sheets processed in {file_name}")
    
    # Check if any data was successfully processed
    if not merged_data:
        print("\nError: No data was successfully processed!")
        print("Check the error messages above for details.")
        return None
    
    # Merge all processed data
    try:
        print(f"\nMerging {len(merged_data)} processed sheets...")
        merged_df = pd.concat(merged_data, ignore_index=True)
        print(f"✓ Successfully merged data: {len(merged_df)} total rows")
    except Exception as e:
        print(f"Error merging data: {e}")
        return None
    
    # Generate output filename with version control
    try:
        base_filename = "merged_sets"
        counter = 1
        
        while True:
            output_file = os.path.join(output_directory, f"{base_filename}_{counter}.xlsx")
            if not os.path.exists(output_file):
                break
            counter += 1
        
        print(f"Saving merged data to: {os.path.basename(output_file)}")
        merged_df.to_excel(output_file, index=False)
        print(f"✓ File saved successfully!")
        
    except PermissionError:
        print("Error: Permission denied. Check if the output file is open in Excel.")
        return None
    except Exception as e:
        print(f"Error saving file: {e}")
        return None
    
    # Print final summary
    print(f"\n" + "="*50)
    print(f"PROCESSING COMPLETE")
    print(f"="*50)
    print(f"Files found: {processing_stats['files_found']}")
    print(f"Files processed: {processing_stats['files_processed']}")
    print(f"Sheets processed: {processing_stats['sheets_processed']}")
    print(f"Total rows in output: {len(merged_df)}")
    print(f"Errors encountered: {len(processing_stats['errors'])}")
    print(f"Output file: {output_file}")
    
    if processing_stats['errors']:
        print(f"\nErrors encountered:")
        for error in processing_stats['errors']:
            print(f"  - {error}")
    
    return output_file


def main():
    """
    Main function to run the Excel processor with default settings.
    """
    print("Excel File Processor")
    print("=" * 30)
    
    # Run with default directories (Colab paths)
    result = process_excel_files()
    
    if result:
        print(f"\n🎉 Success! Output saved to: {result}")
    else:
        print(f"\n❌ Processing failed. Check error messages above.")


if __name__ == "__main__":
    main()

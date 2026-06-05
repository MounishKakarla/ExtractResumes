import pandas as pd
import os
import shutil
from pathlib import Path

# --- Configuration ---
base_destination = 'extracted_resumes' # Base folder for all extracted resumes

def find_resume_column(df):
    """Auto-detect the resume column by looking for common patterns"""
    
    common_patterns = ['resume', 'link', 'path', 'url', 'file']
    
    for col in df.columns:
        col_lower = col.lower()
        # Check if column contains common resume-related keywords
        if any(pattern in col_lower for pattern in common_patterns):
            return col
    
    # If no match, return first column
    return df.columns[0]

def extract_resumes_from_file(excel_file):
    """Extract resumes from a single Excel file"""
    
    # Create a subfolder for this Excel file
    file_name_without_ext = Path(excel_file).stem
    destination_folder = os.path.join(base_destination, file_name_without_ext)
    
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)
        print(f"Created folder: {destination_folder}")

    # Read the Excel data
    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        print(f"Error reading {excel_file}: {e}")
        return 0

    # Auto-detect resume column
    column_name = find_resume_column(df)
    print(f"Using column: '{column_name}'")

    # Iterate through paths and copy files
    count = 0
    for index, row in df.iterrows():
        source_path = row[column_name]
        
        # Ensure path is a string and exists
        if isinstance(source_path, str) and os.path.exists(source_path):
            file_name = os.path.basename(source_path)
            dest_path = os.path.join(destination_folder, file_name)
            
            shutil.copy2(source_path, dest_path)
            print(f"  Copied: {file_name}")
            count += 1
        else:
            print(f"  Skipped/Not found: {source_path}")

    return count

def process_all_excel_files():
    """Process all Excel files in current directory"""
    
    # Find all Excel files
    excel_files = list(Path('.').glob('*.xlsx')) + list(Path('.').glob('*.xls'))
    
    if not excel_files:
        print("No Excel files found in current directory.")
        return
    
    print(f"Found {len(excel_files)} Excel file(s).\n")
    
    total_count = 0
    for excel_file in excel_files:
        print(f"Processing: {excel_file}")
        count = extract_resumes_from_file(str(excel_file))
        print(f"  {count} resumes extracted.\n")
        total_count += count
    
    print(f"=" * 50)
    print(f"All tasks complete. {total_count} total resumes extracted to '{base_destination}'.")

if __name__ == "__main__":
    process_all_excel_files()

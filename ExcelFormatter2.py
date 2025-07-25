import os
import glob
import pandas as pd
from datetime import datetime, timedelta, date
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
import threading
import openpyxl
from openpyxl.styles import Border, Side, Font, Alignment
import calendar
from collections import Counter, defaultdict

def clean_and_count_excel(file_path):
    """Clean and count entries from a single Excel file"""
    try:
        df = pd.read_excel(file_path, header=None, engine='openpyxl', skiprows=2)
        
        # Check if file has enough columns
        if df.shape[1] < 6:
            return 0
            
        # Start from third row (index 2) - already handled by skiprows=2
        df = df.copy()
        
        # Remove rows where column F (index 5) is blank or '-'
        df = df[df.iloc[:, 5].notna()]
        df = df[df.iloc[:, 5] != '-']
        
        # Keep only rows where column F is 'local' or 'imported' (case insensitive)
        df = df[df.iloc[:, 5].astype(str).str.strip().str.lower().isin(['local', 'imported'])]
        
        # Fix date format in column A (index 0)
        df.iloc[:, 0] = df.iloc[:, 0].apply(fix_date)
        
        # Remove rows with invalid dates
        df = df[df.iloc[:, 0].notna()]
        
        return len(df)
    except Exception as e:
        print(f"Error processing {file_path}: {e}")
        return 0

def fix_date(val):
    """Robust date parsing function"""
    if val is None or pd.isna(val):
        return None
        
    # Handle datetime objects first
    if isinstance(val, (pd.Timestamp, datetime)):
        return val.date()
    
    # Handle date objects
    if isinstance(val, date):
        return val
        
    # Handle string dates
    if isinstance(val, str):
        val = val.strip()
        if not val:
            return None
            
        # Try common date formats
        formats = [
            '%Y-%m-%d',      # 2024-03-07
            '%d/%m/%Y',      # 07/03/2024
            '%m/%d/%Y',      # 03/07/2024
            '%Y/%m/%d',      # 2024/03/07
            '%d-%m-%Y',      # 07-03-2024
            '%m-%d-%Y',      # 03-07-2024
            '%d.%m.%Y',      # 07.03.2024
            '%Y.%m.%d',      # 2024.03.07
            '%d %m %Y',      # 07 03 2024
            '%Y %m %d',      # 2024 03 07
        ]
        
        for fmt in formats:
            try:
                return datetime.strptime(val, fmt).date()
            except ValueError:
                continue
        
        # Try pandas parsing as last resort
        try:
            parsed = pd.to_datetime(val, errors='coerce', dayfirst=True)
            if not pd.isna(parsed):
                return parsed.date()
        except:
            pass
            
        return None
        
    # Handle numeric values (Excel serial dates)
    if isinstance(val, (int, float)):
        try:
            # Excel dates start from 1900-01-01 (but Excel thinks 1900 was a leap year)
            if val >= 1:
                # Convert Excel serial date to Python date
                if val < 60:  # Before March 1, 1900
                    base_date = datetime(1899, 12, 30)
                else:  # March 1, 1900 and after (account for fake leap day)
                    base_date = datetime(1899, 12, 31)
                    val -= 1
                
                result_date = base_date + timedelta(days=int(val))
                return result_date.date()
        except:
            pass
            
    return None

def format_date_header(date):
    """Format date as shown in the table (e.g., '7-Mar', '15-May')"""
    return f"{date.day}-{calendar.month_abbr[date.month]}"

def get_all_dates_from_folder(folder, log_callback=None):
    """Scan all Excel files in the folder to collect unique dates"""
    excel_files = get_valid_excel_files(folder)
    all_dates = set()
    
    for file in excel_files:
        try:
            df = pd.read_excel(file, header=None, engine='openpyxl', skiprows=2)
            
            # Check if enough columns
            if df.shape[1] < 6:
                if log_callback:
                    log_callback(f"File {os.path.basename(file)} does not have enough columns, skipping.")
                continue
            
            # Apply same filtering as main processing
            df = df[df.iloc[:, 5].notna()]
            df = df[df.iloc[:, 5] != '-']
            df = df[df.iloc[:, 5].astype(str).str.strip().str.lower().isin(['local', 'imported'])]

            # Fix dates
            df.iloc[:, 0] = df.iloc[:, 0].apply(fix_date)
            df = df[df.iloc[:, 0].notna()]

            # Add unique dates to our set
            file_dates = df.iloc[:, 0].unique()
            valid_dates = [d for d in file_dates if isinstance(d, date)]
            all_dates.update(valid_dates)

        except Exception as e:
            if log_callback:
                log_callback(f"Error scanning dates from {os.path.basename(file)}: {e}")

    # Convert to sorted list
    sorted_dates = sorted(all_dates)

    if log_callback:
        log_callback(f"Found {len(sorted_dates)} unique dates across all files")
        if sorted_dates:
            log_callback(f"Date range: {sorted_dates[0]} to {sorted_dates[-1]}")

    return sorted_dates

def get_valid_excel_files(folder):
    """Get list of valid Excel files, excluding templates and temp files"""
    excel_files = glob.glob(os.path.join(folder, '*.xlsx'))
    
    # Exclude Template.xlsx, summary files, and Excel temp/lock files
    valid_files = []
    for f in excel_files:
        basename = os.path.basename(f).lower()
        if (not basename.startswith('template') and 
            not basename.startswith('summary') and 
            not basename.startswith('~$')):
            valid_files.append(f)
    
    return valid_files

def create_or_update_summary_list(summary_path, region, date_counts, all_dates, log_callback=None):
    """Create or update the summary_list.xlsx file with dynamic date columns"""
    
    regions = [
        "NCR", "Region I", "Region II", "Region III", "Region IV", 
        "MIMAROPA", "Region V", "Region VI", "Region VII", "Region VIII", 
        "Region IX", "Region X", "Region XI", "Region XII", "CAR", 
        "Region XIII", "BARMM"
    ]
    
    # Check if file exists
    if os.path.exists(summary_path):
        # Load existing file
        wb = openpyxl.load_workbook(summary_path)
        ws = wb.active
        if log_callback:
            log_callback(f"Loading existing summary file: {summary_path}")
        
        # Get existing date columns
        existing_dates = []
        for col in range(2, ws.max_column + 1):
            from openpyxl.utils import get_column_letter
            col_letter = get_column_letter(col)
            header_value = ws[f'{col_letter}2'].value
            if header_value:
                # Try to parse the existing header back to a date
                try:
                    day, month_abbr = header_value.split('-')
                    month_num = list(calendar.month_abbr).index(month_abbr)
                    # Use the year from our data
                    year = all_dates[0].year if all_dates else datetime.now().year
                    existing_date = datetime(year, month_num, int(day)).date()
                    existing_dates.append(existing_date)
                except:
                    if log_callback:
                        log_callback(f"Could not parse existing date header: {header_value}")
        
        # Find new dates that need to be added
        new_dates = [d for d in all_dates if d not in existing_dates]
        
        if new_dates:
            # Add new columns for new dates
            start_col = ws.max_column + 1
            for i, date in enumerate(new_dates):
                from openpyxl.utils import get_column_letter
                col_letter = get_column_letter(start_col + i)
                ws[f'{col_letter}2'] = format_date_header(date)
                
                # Apply formatting to new header
                ws[f'{col_letter}2'].font = Font(bold=True)
                ws[f'{col_letter}2'].alignment = Alignment(horizontal='center')
                ws[f'{col_letter}2'].border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
                
                # Apply formatting to data cells in new column
                for row in range(3, len(regions) + 3):
                    cell = ws[f'{col_letter}{row}']
                    cell.border = Border(
                        left=Side(style='thin'),
                        right=Side(style='thin'),
                        top=Side(style='thin'),
                        bottom=Side(style='thin')
                    )
                    cell.alignment = Alignment(horizontal='center')
            
            # Update the merged title cell
            from openpyxl.utils import get_column_letter
            last_col_letter = get_column_letter(ws.max_column)
            ws.merge_cells(f'A1:{last_col_letter}1')
            
            if log_callback:
                log_callback(f"Added {len(new_dates)} new date columns")
    else:
        # Create new workbook
        wb = openpyxl.Workbook()
        ws = wb.active
        
        # Set up headers
        from openpyxl.utils import get_column_letter
        ws['A1'] = "2025 Daily Monitoring - Entry Count"
        last_col_letter = get_column_letter(len(all_dates) + 1)
        ws.merge_cells(f'A1:{last_col_letter}1')

        # Region header
        ws['A2'] = "REGIONS"

        # Date headers
        for i, date in enumerate(all_dates):
            col_letter = get_column_letter(i + 2)  # B, C, D, etc.
            ws[f'{col_letter}2'] = format_date_header(date)

        # Region names
        for i, reg in enumerate(regions):
            ws[f'A{i+3}'] = reg

        # Apply formatting
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        # Format headers
        for row in ws[f'A1:{last_col_letter}2']:
            for cell in row:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center')
                cell.border = thin_border

        # Format region column and data cells
        for row in range(3, len(regions) + 3):
            for col in range(1, len(all_dates) + 2):
                col_letter = get_column_letter(col)
                cell = ws[f'{col_letter}{row}']
                cell.border = thin_border
                if col == 1:  # Region column
                    cell.font = Font(bold=True)
                else:
                    cell.alignment = Alignment(horizontal='center')
        
        if log_callback:
            log_callback(f"Created new summary file with {len(all_dates)} date columns: {summary_path}")
    
    # Update the cells with the counts
    from openpyxl.utils import get_column_letter
    
    # Find the correct row for the region
    region_row = None
    for row in range(3, ws.max_row + 1):
        if ws[f'A{row}'].value == region:
            region_row = row
            break

    if region_row is None:
        if log_callback:
            log_callback(f"Warning: Region '{region}' not found in summary file")
        return

    # Update counts for dates that have data
    for date, count in date_counts.items():
        # Find the correct column for this date
        date_col = None
        for col in range(2, ws.max_column + 1):
            col_letter = get_column_letter(col)
            header_date = ws[f'{col_letter}2'].value
            if header_date == format_date_header(date):
                date_col = col
                break

        if date_col is None:
            if log_callback:
                log_callback(f"Warning: Date column for '{format_date_header(date)}' not found")
            continue

        # Update the cell (ADD to existing value, don't replace)
        col_letter = get_column_letter(date_col)
        current_value = ws[f'{col_letter}{region_row}'].value
        if current_value is None:
            current_value = 0

        new_value = current_value + count
        ws[f'{col_letter}{region_row}'].value = new_value

        if log_callback:
            log_callback(f"Updated {region} for {format_date_header(date)}: {current_value} + {count} = {new_value}")
    
    # Save the file
    wb.save(summary_path)

def process_excels(folder, region, summary_list_path, log_callback=None, export_folder=None):
    """Main processing function"""
    # First, scan all files to get unique dates
    if log_callback:
        log_callback("=" * 50)
        log_callback("Starting Excel processing...")
        log_callback("Scanning all files to determine date columns...")
    
    all_dates = get_all_dates_from_folder(folder, log_callback)
    
    if not all_dates:
        if log_callback:
            log_callback("ERROR: No valid dates found in any files!")
        return
    
    excel_files = get_valid_excel_files(folder)
    
    if not excel_files:
        if log_callback:
            log_callback("ERROR: No valid Excel files found!")
        return
    
    if log_callback:
        log_callback(f"Found {len(excel_files)} valid Excel files to process")
    
    results = []
    merged_df_list = []
    
    # Determine summary file path
    if not summary_list_path or not os.path.exists(summary_list_path):
        summary_path = os.path.join(export_folder if export_folder else folder, 'summary_list.xlsx')
        if log_callback:
            log_callback(f"Will create new summary file at: {summary_path}")
    else:
        summary_path = summary_list_path
        if log_callback:
            log_callback(f"Will update existing summary file: {summary_path}")
    
    # Dictionary to store counts per date
    region_date_counts = defaultdict(int)
    total_count = 0
    file_count = 0
    
    for file in excel_files:
        file_count += 1
        filename = os.path.basename(file)
        
        try:
            if log_callback:
                log_callback(f"Processing file {file_count}/{len(excel_files)}: {filename}")
            
            # Read and process the file
            df = pd.read_excel(file, header=None, engine='openpyxl', skiprows=2)
            
            if df.shape[1] < 6:
                if log_callback:
                    log_callback(f"  SKIPPED: Not enough columns ({df.shape[1]} < 6)")
                continue
            
            # Apply filtering
            original_rows = len(df)
            df = df[df.iloc[:, 5].notna()]
            df = df[df.iloc[:, 5] != '-']
            df = df[df.iloc[:, 5].astype(str).str.strip().str.lower().isin(['local', 'imported'])]
            
            # Fix dates
            df.iloc[:, 0] = df.iloc[:, 0].apply(fix_date)
            df = df[df.iloc[:, 0].notna()]
            
            final_rows = len(df)
            
            if final_rows == 0:
                if log_callback:
                    log_callback(f"  SKIPPED: No valid data after filtering (started with {original_rows} rows)")
                continue
            
            # Count entries per date for this file
            file_date_counts = df.iloc[:, 0].value_counts().to_dict()
            
            # Add to overall region counts
            file_total = 0
            for dte, count in file_date_counts.items():
                if isinstance(dte, date):
                    region_date_counts[dte] += count
                    file_total += count
            
            total_count += file_total
            results.append(f"{filename} - {file_total}")
            merged_df_list.append(df)

            if log_callback:
                date_breakdown = []
                for d in sorted(file_date_counts.keys()):
                    if isinstance(d, date):
                        date_breakdown.append(f"{format_date_header(d)}({file_date_counts[d]})")
                
                date_summary = ", ".join(date_breakdown)
                log_callback(f"  SUCCESS: {file_total} entries | Per date: {date_summary}")
                log_callback(f"  Filtered: {original_rows} → {final_rows} rows")

        except Exception as e:
            if log_callback:
                log_callback(f"  ERROR processing {filename}: {e}")
    
    # Summary of processing
    if log_callback:
        log_callback("=" * 50)
        log_callback(f"PROCESSING SUMMARY:")
        log_callback(f"Files processed successfully: {len(results)}")
        log_callback(f"Total entries found: {total_count}")
        log_callback(f"Date breakdown:")
        for dte in sorted(region_date_counts.keys()):
            if isinstance(dte, date):
                log_callback(f"  {format_date_header(dte)}: {region_date_counts[dte]} entries")
    
    # Update summary list with per-date counts
    if region_date_counts:
        try:
            if log_callback:
                log_callback("Updating summary file...")
            create_or_update_summary_list(summary_path, region, region_date_counts, all_dates, log_callback)
            if log_callback:
                log_callback(f"Successfully updated summary for {region} with {total_count} total entries")
        except Exception as e:
            if log_callback:
                log_callback(f"ERROR updating summary list: {e}")
    
    # Create merged file if template exists
    template_path = os.path.join(folder, 'Template.xlsx')
    if merged_df_list and os.path.exists(template_path):
        try:
            if log_callback:
                log_callback("Creating merged file...")
            
            merged_df = pd.concat(merged_df_list, ignore_index=True)
            values_to_paste = merged_df.values.tolist()
            wb_template = openpyxl.load_workbook(template_path)
            ws_template = wb_template.active
            
            # Paste values starting from row 3
            for i, row_values in enumerate(values_to_paste, start=3):
                for j, value in enumerate(row_values, start=1):
                    if j <= ws_template.max_column:
                        ws_template.cell(row=i, column=j, value=value)
            
            folder_name = os.path.basename(folder)
            merged_path = os.path.join(folder, f'{folder_name}_merged.xlsx')
            wb_template.save(merged_path)
            
            if log_callback:
                log_callback(f"Merged file created: {merged_path}")
        except Exception as e:
            if log_callback:
                log_callback(f"ERROR creating merged file: {e}")
    
    if log_callback:
        log_callback("=" * 50)
        log_callback("PROCESSING COMPLETE!")
        log_callback(f"Final total: {total_count} entries for {region}")

def start_processing(selected_folder, region, summary_list_path, log_widget):
    def log_callback(msg):
        log_widget.config(state='normal')
        log_widget.insert(tk.END, msg + '\n')
        log_widget.see(tk.END)
        log_widget.config(state='disabled')
        log_widget.update_idletasks()
    # Pass export folder to processing
    export_folder = export_folder_var.get() if 'export_folder_var' in locals() else selected_folder
    threading.Thread(
        target=process_excels, 
        args=(selected_folder, region, summary_list_path, log_callback, export_folder), 
        daemon=True
    ).start()

def run_ui():
    root = tk.Tk()
    root.title("Excel Counter & Summary Updater - Enhanced Version")
    root.geometry("700x500")

    folder_var = tk.StringVar()
    folder_var.set(os.path.dirname(os.path.abspath(__file__)))

    region_var = tk.StringVar()
    region_var.set("NCR")
    regions = [
        "NCR", "Region I", "Region II", "Region III", "Region IV", "MIMAROPA", 
        "Region V", "Region VI", "Region VII", "Region VIII", "Region IX", "Region X", 
        "Region XI", "Region XII", "CAR", "Region XIII", "BARMM"
    ]

    def select_folder():
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            folder_var.set(folder_selected)

    # Folder selection
    tk.Label(root, text="Select Folder:", font=('Arial', 10, 'bold')).pack(pady=5)
    folder_frame = tk.Frame(root)
    folder_frame.pack(pady=5)
    tk.Entry(folder_frame, textvariable=folder_var, width=60).pack(side=tk.LEFT)
    tk.Button(folder_frame, text="Browse", command=select_folder).pack(side=tk.LEFT, padx=5)

    # Region selection
    tk.Label(root, text="Select Region:", font=('Arial', 10, 'bold')).pack(pady=5)
    region_dropdown = tk.OptionMenu(root, region_var, *regions)
    region_dropdown.config(width=15)
    region_dropdown.pack(pady=5)

    # Summary list selection and export folder selection
    summary_list_path_var = tk.StringVar()
    summary_list_path_var.set("")
    export_folder_var = tk.StringVar()
    export_folder_var.set(os.path.dirname(os.path.abspath(__file__)))

    def select_summary_list():
        file_selected = filedialog.askopenfilename(
            title="Select existing summary_list.xlsx (optional)", 
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file_selected:
            summary_list_path_var.set(file_selected)

    def select_export_folder():
        folder_selected = filedialog.askdirectory(title="Select folder to export summary_list.xlsx")
        if folder_selected:
            export_folder_var.set(folder_selected)

    tk.Label(root, text="Summary List (optional - will create if not selected):", font=('Arial', 10, 'bold')).pack(pady=(10,5))
    summary_frame = tk.Frame(root)
    summary_frame.pack(pady=5)
    tk.Button(summary_frame, text="Select summary_list.xlsx", command=select_summary_list).pack(side=tk.LEFT)
    summary_entry = tk.Entry(summary_frame, textvariable=summary_list_path_var, width=50, state='readonly')
    summary_entry.pack(side=tk.LEFT, padx=5)
    tk.Button(summary_frame, text="Select Export Folder", command=select_export_folder).pack(side=tk.LEFT, padx=5)
    export_entry = tk.Entry(summary_frame, textvariable=export_folder_var, width=40, state='readonly')
    export_entry.pack(side=tk.LEFT, padx=5)
    def clear_summary():
        summary_list_path_var.set("")
    tk.Button(summary_frame, text="Clear", command=clear_summary).pack(side=tk.LEFT, padx=5)

    # Log widget
    tk.Label(root, text="Processing Log:", font=('Arial', 10, 'bold')).pack(pady=(10,5))
    log_widget = scrolledtext.ScrolledText(root, width=80, height=15, state='disabled')
    log_widget.pack(pady=5, padx=10, fill='both', expand=True)

    def on_start():
        if not folder_var.get():
            messagebox.showerror("Error", "Please select a folder first!")
            return
        
        log_widget.config(state='normal')
        log_widget.delete(1.0, tk.END)
        log_widget.config(state='disabled')
        
        start_processing(folder_var.get(), region_var.get(), summary_list_path_var.get(), log_widget)

    # Start button
    start_btn = tk.Button(root, text="Start Processing", command=on_start, 
                         width=20, height=2, bg='#4CAF50', fg='white', 
                         font=('Arial', 12, 'bold'))
    start_btn.pack(pady=20)

    # Add some instructions
    instructions = tk.Label(root, 
                           text="ENHANCED VERSION: Better error handling, detailed logging, and improved counting logic.\n"
                                "Select folder with Excel files, choose region, optionally select existing summary file.\n"
                                "The tool will show detailed progress and identify any counting issues.",
                           font=('Arial', 9), fg='gray', wraplength=650)
    instructions.pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    run_ui()
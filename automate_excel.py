import os
import datetime
import logging
import openpyxl
from openpyxl.styles import Font
import pandas as pd
from openpyxl.cell.cell import MergedCell

# ------------------------------------------------------------
# Logging Configuration
# ------------------------------------------------------------
logging.basicConfig(
    level=logging.DEBUG,  # Change to INFO to reduce verbosity if desired
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

# ------------------------------------------------------------
# Safe Cell Writing (for merged cells)
# ------------------------------------------------------------
def safe_set_cell(ws, cell_ref, value):
    """
    Safely sets the value of a cell. If the cell is merged, update the top-left cell.
    """
    cell = ws[cell_ref]
    if isinstance(cell, MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                top_left = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
                top_left.value = value
                logging.debug(f"Set merged cell {merged_range} top-left to '{value}'")
                return
    else:
        cell.value = value

# ============================================================
# Clear Data from Copied Worksheet
# ============================================================
def clear_sheet_data(sheet, start_row=7):
    """
    Clears cell values from start_row to the end of the sheet.
    This prevents copying any sample data from the TEMPLATE.
    """
    for row in sheet.iter_rows(min_row=start_row, max_row=sheet.max_row):
        for cell in row:
            cell.value = None

# ============================================================
# Helper Functions for User Input
# ============================================================
def prompt_date_range():
    """
    Prompts for a start and end date (YYYY-MM-DD). These dates are used as sheet names.
    Returns start_date and end_date as date objects.
    """
    print("Enter the date range for your daily sheets:")
    start_date_str = input("  Start date (YYYY-MM-DD): ").strip()
    end_date_str = input("  End date (YYYY-MM-DD): ").strip()
    try:
        start_date = datetime.datetime.strptime(start_date_str, "%Y-%m-%d").date()
        end_date = datetime.datetime.strptime(end_date_str, "%Y-%m-%d").date()
    except ValueError:
        logging.error("Invalid date format. Please use YYYY-MM-DD.")
        exit(1)
    today = datetime.date.today()
    if start_date > today or end_date > today:
        logging.error("Future dates are not allowed.")
        exit(1)
    if start_date > end_date:
        logging.error("Start date must not be later than end date.")
        exit(1)
    logging.info(f"Date range: {start_date} to {end_date}")
    return start_date, end_date

def prompt_file_paths():
    """
    Prompts for one or more file paths (CSV/TXT) and verifies each exists.
    Returns a list of file paths.
    """
    file_paths_str = input("Enter the path(s) to the data file(s) (CSV/TXT, comma-separated): ").strip()
    file_paths = [fp.strip().strip('"') for fp in file_paths_str.split(",")]
    for fp in file_paths:
        if not os.path.exists(fp):
            logging.error(f"File not found: {fp}")
            exit(1)
    return file_paths

def prompt_rate():
    """
    Prompts for an hourly rate.
    """
    rate_str = input("Enter the hourly rate: ").strip()
    try:
        rate = float(rate_str)
    except ValueError:
        logging.error("Invalid rate. Please enter a numeric value.")
        exit(1)
    logging.info(f"Hourly rate: {rate}")
    return rate

# ============================================================
# Data File Parsing Functions
# ============================================================
def read_csv_data(data_file):
    """
    Reads a CSV/TXT file into a DataFrame.
    Uses a tab delimiter if the file extension is '.txt'; otherwise, uses comma.
    """
    try:
        if data_file.lower().endswith(".txt"):
            df = pd.read_csv(data_file, sep='\t')
        else:
            df = pd.read_csv(data_file)
        logging.info(f"Data file '{data_file}' read with {len(df)} rows.")
    except Exception as e:
        logging.error(f"Error reading data file: {e}")
        exit(1)
    return df

def combine_csv_data(file_paths):
    """
    Reads and concatenates multiple CSV/TXT files into a single DataFrame.
    """
    df_list = []
    for fp in file_paths:
        df = read_csv_data(fp)
        df_list.append(df)
    if df_list:
        combined_df = pd.concat(df_list, ignore_index=True)
    else:
        combined_df = pd.DataFrame()
    return combined_df

# ============================================================
# Date Filtering and Daily List Creation
# ============================================================
def filter_df_by_date(df, start_date, end_date):
    """
    If a 'Date' column exists, parses it and filters the DataFrame by the given date range.
    Returns the filtered DataFrame and a flag indicating if a 'Date' column is present.
    """
    if 'Date' in df.columns:
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce').dt.date
        orig_count = len(df)
        df = df[(df['Date'] >= start_date) & (df['Date'] <= end_date)]
        logging.info(f"Filtered tasks from {orig_count} to {len(df)} rows by date.")
        return df, True
    else:
        logging.info("No 'Date' column found; applying tasks to every date.")
        return df, False

def create_date_list(start_date, end_date):
    """
    Creates a list of date objects from start_date to end_date (inclusive).
    """
    date_list = []
    current = start_date
    while current <= end_date:
        date_list.append(current)
        current += datetime.timedelta(days=1)
    return date_list

# ============================================================
# Sheet Population Functions
# ============================================================
def fill_daily_sheet(sheet, date_obj, data_rows, is_dataframe, start_row=7, fallback_date=None):
    """
    Populates a daily sheet:
      - Sets cell B1 to the given date if matching data exists; otherwise uses today's date.
      - Writes headers in row 6 and task data starting at row 7.
      - For single-day ranges, if no data exists, writes fallback_date to B3.
    Returns the last row used.
    """
    if is_dataframe:
        if 'Date' in data_rows.columns:
            day_df = data_rows[data_rows['Date'] == date_obj]
        else:
            day_df = data_rows
        records = day_df.to_dict(orient='records')
    else:
        records = data_rows

    if not records:
        today = datetime.date.today()
        sheet["B1"] = today.strftime("%Y-%m-%d")
        if fallback_date is not None:
            sheet["B3"] = fallback_date.strftime("%Y-%m-%d")
    else:
        sheet["B1"] = date_obj.strftime("%Y-%m-%d")
    
    sheet.freeze_panes = sheet["A7"]
    
    headers = ["Number", "Daily Work Description", "Hr", "Min", "Complete", "Follow up", "Supervisor Comments"]
    for col_idx, header in enumerate(headers, start=1):
        sheet.cell(row=6, column=col_idx).value = header

    current_row = start_row
    for row_dict in records:
        sheet.cell(row=current_row, column=1).value = row_dict.get("Number", "")
        sheet.cell(row=current_row, column=2).value = row_dict.get("Daily Work Description", "")
        sheet.cell(row=current_row, column=3).value = row_dict.get("Hr", "")
        sheet.cell(row=current_row, column=4).value = row_dict.get("Min", "")
        complete_val = row_dict.get("Complete", "")
        complete_cell = sheet.cell(row=current_row, column=5)
        complete_cell.value = complete_val
        if isinstance(complete_val, str):
            if complete_val.lower() == "yes":
                complete_cell.font = Font(color="008000")  # green
            elif complete_val.lower() == "no":
                complete_cell.font = Font(color="FF0000")  # red
        sheet.cell(row=current_row, column=6).value = row_dict.get("Follow up", "")
        sheet.cell(row=current_row, column=7).value = row_dict.get("Supervisor Comments", "")
        current_row += 1
    last_row = current_row - 1
    logging.info(f"{date_obj}: Populated rows {start_row} to {last_row}.")
    return last_row

# ============================================================
# Total Sheet Population Functions
# ============================================================
def update_total_sheet(total_sheet, daily_info, rate):
    """
    Fills the 'Total' sheet with summary information:
      - Headers in cells B3:E3: Date, Hour, Rate, Total Cost.
      - From row 4 onward, each row corresponds to a daily sheet.
      - Column C uses a formula to calculate total hours (hours + minutes/60) from each daily sheet.
    """
    for row in total_sheet.iter_rows(min_row=4, max_row=total_sheet.max_row):
        for cell in row:
            cell.value = None

    row_idx = 4
    for sheet_name, (start_row, last_row) in sorted(daily_info.items()):
        if last_row < start_row:
            continue
        hour_formula = f"=SUM('{sheet_name}'!C7:C{last_row}) + (SUM('{sheet_name}'!D7:D{last_row})/60)"
        safe_set_cell(total_sheet, f"B{row_idx}", sheet_name)
        safe_set_cell(total_sheet, f"C{row_idx}", hour_formula)
        safe_set_cell(total_sheet, f"D{row_idx}", rate)
        safe_set_cell(total_sheet, f"E{row_idx}", f"=C{row_idx}*D{row_idx}")
        row_idx += 1

def create_or_update_total_sheet(wb, daily_info, rate):
    """
    Ensures a 'Total' sheet exists, writes summary headers in B3:E3,
    populates summary rows, and freezes rows 1-3.
    """
    if "Total" in wb.sheetnames:
        total_sheet = wb["Total"]
    else:
        total_sheet = wb.create_sheet("Total")
    safe_set_cell(total_sheet, "B3", "Date")
    safe_set_cell(total_sheet, "C3", "Hour")
    safe_set_cell(total_sheet, "D3", "Rate")
    safe_set_cell(total_sheet, "E3", "Total Cost")
    update_total_sheet(total_sheet, daily_info, rate)
    total_sheet.freeze_panes = total_sheet["A4"]
    return total_sheet

# ============================================================
# Main Workflow
# ============================================================
def main():
    """
    Main workflow:
      1. Prompt for a date range, data file(s), and hourly rate.
      2. Load the template workbook (which must have 'Template' and 'Total' sheets).
      3. Process data from CSV/TXT files (with a 'Date' column).
         Create one daily sheet per date in the prompted range using the imported data.
      4. Create a 'Total' sheet that summarizes the daily sheets.
      5. Hide the TEMPLATE sheet.
      6. Save the resulting workbook (for a single day, expect 2 sheets: one daily and one Total).
    """
    # 1. User Inputs
    start_date, end_date = prompt_date_range()
    file_paths = prompt_file_paths()
    rate = prompt_rate()
    
    # 2. Load Template Workbook
    template_path = r"C:\Users\ellie\Downloads\excel_automation\template_daily_recap.xlsx"
    if not os.path.exists(template_path):
        logging.error(f"Template file not found: {template_path}")
        exit(1)
    wb = openpyxl.load_workbook(template_path)
    template_sheet_name = "Template"
    if template_sheet_name not in wb.sheetnames:
        logging.error(f"No sheet named '{template_sheet_name}' in {template_path}")
        exit(1)
    
    daily_info = {}
    
    # 3. Process Data (CSV/TXT with Date column)
    combined_df = combine_csv_data(file_paths)
    combined_df, has_date = filter_df_by_date(combined_df, start_date, end_date)
    
    if has_date and not combined_df.empty:
        unique_dates = sorted(set(combined_df['Date'].dropna()))
    else:
        unique_dates = create_date_list(start_date, end_date)
    
    for d in unique_dates:
        sheet_name = d.strftime("%Y-%m-%d")
        new_sheet = wb.copy_worksheet(wb[template_sheet_name])
        new_sheet.title = sheet_name
        # Clear any data copied from the template
        clear_sheet_data(new_sheet, start_row=7)
        # For single-day range, pass fallback date to render in B3 if needed
        fallback = d if start_date == end_date else None
        last_row = fill_daily_sheet(new_sheet, d, combined_df, is_dataframe=True, start_row=7, fallback_date=fallback)
        daily_info[sheet_name] = (7, last_row)
    
    # 4. Create/Update Total Sheet
    create_or_update_total_sheet(wb, daily_info, rate)
    
    # 5. Hide the TEMPLATE Sheet
    wb[template_sheet_name].sheet_state = "hidden"
    
    # 6. Save the Workbook
    if start_date == end_date:
        output_filename = f"{start_date.strftime('%Y-%m-%d')}.xlsx"
    else:
        output_filename = f"{start_date.strftime('%Y-%m-%d')}_to_{end_date.strftime('%Y-%m-%d')}.xlsx"
    try:
        wb.save(output_filename)
    except PermissionError as e:
        logging.error(f"Permission error saving '{output_filename}': {e}")
        exit(1)
    logging.info(f"Workbook '{output_filename}' created successfully.")

if __name__ == "__main__":
    main()

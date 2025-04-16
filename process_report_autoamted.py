# Filename: process_report_automated.py
# (Correct Sunday Calc, Friday Name, No Unblinded Sheet)

import os
import re
import traceback  # For detailed error logging
from datetime import datetime, timedelta

import pandas as pd
import xlwings as xw

# --- Configuration ---
try:
    # Best practice: Use __file__ when available
    SCRIPT_DIRECTORY = os.path.dirname(os.path.abspath(__file__))
except NameError:
    # Fallback for environments where __file__ is not defined (e.g., interactive)
    SCRIPT_DIRECTORY = os.getcwd()
    print("Warning: __file__ not defined.")

print(f"SCRIPT_DIRECTORY set to: {SCRIPT_DIRECTORY}")

TEMPLATE_NAME = 'Report_Template.xlsm'
SOURCE_VEEVA_NAME = 'DMA Report - Documents in Available and Assigned tasks for RQC.xlsx'
SOURCE_STUDY_ALLOC_NAME = 'RQC Studies & POC list.xlsx'
OUTPUT_FOLDER = 'output_reports'  # Subfolder for results

# --- Helper Function: Find Latest File ---

def find_latest_file(base_filename_pattern, directory=SCRIPT_DIRECTORY):
    """
    Finds the file matching the pattern with the highest number in
    parentheses, or the base file if no numbered versions exist,
    within the specified directory.

    Uses a robust regex and the passed directory argument.
    """
    print(
        f"Searching for latest file matching pattern: "
        f"'{base_filename_pattern}' in directory: '{directory}'"
    )
    base_name, ext = os.path.splitext(base_filename_pattern)
    # Regex to find base name optionally followed by space/(number)/extension
    # Added re.IGNORECASE for flexibility (e.g. .xlsx vs .XLSX)
    pattern = re.compile(
        rf"^{re.escape(base_name)}(?:\s*\((\d+)\))?{re.escape(ext)}$",
        re.IGNORECASE
    )

    latest_file = None  # Start with None
    max_number = -1

    try:
        files_in_directory = os.listdir(directory)
    except FileNotFoundError:
        print(f"ERROR: Directory not found when searching for files: {directory}")
        # Return the original pattern so the main script can report
        # the file not found error clearly
        return base_filename_pattern
    except Exception as e:
        print(f"ERROR listing files in directory '{directory}': {e}")
        return base_filename_pattern  # Return pattern on other errors

    found_match = False
    for filename in files_in_directory:
        match = pattern.match(filename)
        if match:
            found_match = True
            number_str = match.group(1)  # Captured digits, or None

            if number_str:
                number = int(number_str)
                if number > max_number:
                    max_number = number
                    latest_file = filename
            elif max_number == -1:
                # Match is the base file (no number), and we haven't
                # found a numbered one yet.
                latest_file = filename

    if not found_match:
        print(
            f"WARNING: No files found matching the pattern "
            f"'{base_filename_pattern}' in '{directory}'."
        )
        # Return the original pattern name; os.path.exists check will fail clearly.
        return base_filename_pattern
    elif latest_file:
        print(f"Found latest matching file: '{latest_file}'")
        return latest_file
    else:
        # Fallback, should ideally not happen if found_match is True
        print(
            f"Warning: Matched pattern but couldn't determine latest file. "
            f"Using base pattern: {base_filename_pattern}"
        )
        return base_filename_pattern

# --- Find the specific latest files ---
latest_veeva_filename = find_latest_file(SOURCE_VEEVA_NAME)
latest_study_alloc_name = find_latest_file(SOURCE_STUDY_ALLOC_NAME)

# --- Helper Function: Add RQC User ---

def add_rqc_user_to_sheet(sheet, study_rqc_map):
    """Adds the 'RQC User' column directly to the xlwings sheet object."""
    print("Attempting to add RQC User column...")
    try:
        # Read header reliably, handling single cell header
        header_range = sheet.range('A1').expand('right')
        header_range_values = header_range.value
        if not isinstance(header_range_values, (list, tuple)):
            header_range_values = [header_range_values] # Wrap single value
        header_range_values = header_range_values or [] # Ensure list if None

        # Find 'Study' column index
        study_col_name = 'Study'
        study_col_index = -1
        try:
            # Using original index search method (xlwings is 1-based)
            study_col_index = header_range_values.index(study_col_name) + 1
            print(f"Found '{study_col_name}' column at index: {study_col_index}")
        except ValueError:
            print(
                f"ERROR: Column '{study_col_name}' not found in header: "
                f"{header_range_values}"
            )
            raise ValueError(
                f"Column '{study_col_name}' not found in RawData sheet header."
            )

        # Find the last row with data in the 'Study' column
        # Using sheet.api for potentially better performance/reliability
        last_row = sheet.range(
            sheet.api.Rows.Count, study_col_index
        ).end('up').row

        if last_row < 2:  # If only header or empty
            print(
                "Warning: No data rows found based on Study column. "
                "Stopping RQC User addition."
            )
            return  # Nothing to process

        print(f"Processing rows up to: {last_row}")

        # Find or Add 'RQC User' column
        rqc_user_col_name = 'RQC User'
        rqc_user_col_index = -1
        try:
            # Using original index search method (xlwings is 1-based)
            rqc_user_col_index = header_range_values.index(rqc_user_col_name) + 1
            print(
                f"'{rqc_user_col_name}' column already exists at index "
                f"{rqc_user_col_index}. Overwriting."
            )
        except ValueError:
            # Add header if column doesn't exist
            rqc_user_col_index = len(header_range_values) + 1
            sheet.cells(1, rqc_user_col_index).value = rqc_user_col_name
            print(
                f"Adding '{rqc_user_col_name}' column at index "
                f"{rqc_user_col_index}."
            )

        # Batch read study values for efficiency
        study_range = sheet.range(
            (2, study_col_index), (last_row, study_col_index)
        )
        # options(ndim=1) simplifies to a flat list if single column
        study_values = study_range.options(ndim=1).value

        # Handle case where only one row of data exists
        if not isinstance(study_values, list) and last_row >= 2:
            study_values = [study_values]
        elif last_row < 2:
            study_values = []  # No data rows

        # Prepare RQC user values based on the map
        rqc_users_to_write = []
        for study_cell_value in study_values:
            rqc_user_list = []
            if study_cell_value not in [None, '']:
                # Handle multiple studies separated by comma in one cell
                studies_in_cell = [
                    s.strip() for s in str(study_cell_value).split(',')
                ]
                for study in studies_in_cell:
                    # Map study to RQC user, default to 'NA'
                    rqc_user = study_rqc_map.get(str(study).strip(), 'NA')
                    # Avoid duplicates if multiple studies map to the same user
                    if str(rqc_user) not in rqc_user_list:
                        rqc_user_list.append(str(rqc_user))
                # Join multiple users or use 'NA' if list is empty after lookup
                rqc_users_to_write.append(
                    ', '.join(rqc_user_list) if rqc_user_list else 'NA'
                )
            else:
                # If study cell is empty, assign 'NA'
                rqc_users_to_write.append('NA')

        # Batch write RQC user values
        if rqc_users_to_write:
            target_range = sheet.range(
                (2, rqc_user_col_index),
                (len(rqc_users_to_write) + 1, rqc_user_col_index)
            )
            # Write as a column: use options(transpose=False) and nested list
            target_range.options(transpose=False).value = [
                [val] for val in rqc_users_to_write
            ]
            print(f"Successfully wrote {len(rqc_users_to_write)} RQC User values.")
        else:
            print("No data rows found to write RQC User values.")

        print("'RQC User' column processing complete.")

    except Exception as e:
        print(f"Error in add_rqc_user_to_sheet: {e}")
        traceback.print_exc()
        raise  # Re-raise the exception to stop the script if critical

# --- Main Processing Function ---

def main():
    """
    Main function to process Veeva reports using a template.
    Steps include: loading data, adding RQC users, running VBA,
    creating a pivot summary, and saving the final report.
    """
    template_path = os.path.join(SCRIPT_DIRECTORY, TEMPLATE_NAME)
    source_veeva_path = os.path.join(SCRIPT_DIRECTORY, latest_veeva_filename)
    source_study_alloc_path = os.path.join(
        SCRIPT_DIRECTORY, latest_study_alloc_name
    )
    output_dir = os.path.join(SCRIPT_DIRECTORY, OUTPUT_FOLDER)
    os.makedirs(output_dir, exist_ok=True)

    output_filename = (
        f"Processed_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsm"
    )
    output_path = os.path.join(output_dir, output_filename)

    # --- Input File Validation ---
    if not os.path.exists(source_veeva_path):
        print(f"ERROR: Source Veeva file not found at: {source_veeva_path}")
        return
    if not os.path.exists(source_study_alloc_path):
        print(
            "ERROR: Source Study Allocation file not found at: "
            f"{source_study_alloc_path}"
        )
        return
    if not os.path.exists(template_path):
        print(f"ERROR: Template file not found at: {template_path}")
        return

    excel_app = None
    wb_template = None
    wb_source = None

    try:
        print("Starting Excel Process...")
        # Start Excel invisibly and configure settings
        excel_app = xw.App(visible=False)
        excel_app.display_alerts = False
        excel_app.calculation = 'manual'  # Speed up operations
        excel_app.screen_updating = False # Speed up operations

        # --- 1. Load Study Allocation ---
        print(f"Loading Study Allocation: {latest_study_alloc_name}")
        # Read only first two columns, no header assumed
        df_study_alloc = pd.read_excel(
            source_study_alloc_path, header=None, usecols=[0, 1]
        )
        # Simple validation
        if df_study_alloc.shape[1] < 2:
            print(
                f"ERROR: The '{source_study_alloc_path}' file needs to have at "
                "least 2 columns (A for Study, B for RQC User)."
            )
            return  # Exit if format is wrong

        # Create mapping dictionary: Study (col 0) -> RQC User (col 1)
        # Ensure keys and values are strings, handle potential NaN in RQC User
        study_rqc_map = pd.Series(
            df_study_alloc[1].fillna('NA').astype(str).values,
            index=df_study_alloc[0].astype(str)
        ).to_dict()
        print("Study Allocation map created.")

        # --- 2. Open Source Veeva Report and Template ---
        print(f"Opening Source Veeva Report: {latest_veeva_filename}")
        wb_source = excel_app.books.open(source_veeva_path)
        try:
            sht_source = wb_source.sheets['Sheet0']
        except Exception as e:
            print(
                f"ERROR: Cannot find 'Sheet0' in {latest_veeva_filename}. "
                f"Error: {e}"
            )
            # Clean up before exiting
            if wb_source: wb_source.close()
            if excel_app: excel_app.quit()
            return

        print(f"Opening Template: {TEMPLATE_NAME}")
        wb_template = excel_app.books.open(template_path)
        try:
            sht_raw_data = wb_template.sheets['RawData']
        except Exception as e:
            print(
                f"ERROR: Cannot find 'RawData' sheet in {TEMPLATE_NAME}. "
                f"Error: {e}"
            )
            # Clean up before exiting
            if wb_source: wb_source.close()
            if wb_template: wb_template.close()
            if excel_app: excel_app.quit()
            return

        # --- 3. Copy Data with Hyperlinks ---
        print("Copying data from Veeva Report to Template RawData sheet...")
        sht_raw_data.clear()  # Clear existing content first
        sht_source.used_range.copy(sht_raw_data.range('A1'))
        excel_app.cut_copy_mode = False  # Clear clipboard flashing
        print("Data copied, preserving hyperlinks.")

        wb_source.close()  # Close source file now that data is copied
        wb_source = None # Prevent trying to close again in finally
        print(f"Closed {latest_veeva_filename}")

        # --- 4. Add RQC User Column ---
        # This function now handles adding the column header if needed
        add_rqc_user_to_sheet(sht_raw_data, study_rqc_map)

        # --- 5. Create 'Unblinded' Sheet (DISABLED) ---
        print("Step 5: Python creation of separate 'Unblinded' sheet is DISABLED.")
        # --- No action taken here ---

        # --- 6. Run VBA Macro for Today/Friday Filters ---
        # (Using Sunday date logic implicitly if VBA uses cell values updated by Python)
        print(
            "Running VBA macro 'CreateFilteredSheetsWithHyperlinks' "
            "(for Today/Friday)..."
        )
        try:
            # Ensure correct macro path (Module.MacroName)
            vba_macro = wb_template.macro(
                'Module1.CreateFilteredSheetsWithHyperlinks'
            )
            vba_macro()  # Execute the macro
            print("VBA macro execution finished.")
        except Exception as e:
            # Provide more specific error guidance
            macro_path = 'Module1.CreateFilteredSheetsWithHyperlinks'
            if "Cannot run the macro" in str(e) or "Macro not found" in str(e):
                print(
                    f"ERROR running VBA macro: Macro '{macro_path}' "
                    "not found or disabled."
                )
                print(
                    "Check VBA code (module/macro name), macro security "
                    "settings, and ensure the template is macro-enabled (.xlsm)."
                )
            else:
                print(f"ERROR running VBA macro '{macro_path}': {e}")
                traceback.print_exc()
            raise # Stop script if VBA fails, as subsequent steps might depend on it

        # --- 7. Create Pivot Table ---
        print("Creating Pivot Table using Python...")
        try:
            # Read data from RawData sheet AFTER VBA might have run (if needed)
            # Read directly into DataFrame for easier manipulation
            df_for_pivot = sht_raw_data.used_range.options(
                pd.DataFrame, header=True, index=False
            ).value

            # Define column names for clarity and easier modification
            rqc_user_col = 'RQC User'
            due_date_col = 'Task Due Date'
            content_col = 'Content'  # Needed for 'Unblinded' check

            # Validate necessary columns exist
            if rqc_user_col not in df_for_pivot.columns:
                raise ValueError(
                    f"'{rqc_user_col}' column missing for pivot table."
                )
            if due_date_col not in df_for_pivot.columns:
                raise ValueError(
                    f"'{due_date_col}' column missing for pivot calculations."
                )
            if content_col not in df_for_pivot.columns:
                print(
                    f"WARNING: '{content_col}' column missing for pivot "
                    "calculation ('Unblinded' count will be 0)."
                )
                # Add a dummy column to avoid errors later if missing
                df_for_pivot[content_col] = ''

            # --- Pivot Calculations ---
            # Convert Due Date column to date objects, coerce errors to NaT
            df_for_pivot[due_date_col] = pd.to_datetime(
                df_for_pivot[due_date_col], errors='coerce'
            ).dt.date

            today_date = datetime.now().date()

            # Calculate upcoming Sunday (includes today if today is Sunday)
            # weekday(): Monday is 0, Sunday is 6
            days_until_sunday = (6 - today_date.weekday() + 7) % 7
            next_sunday_date = today_date + timedelta(days=days_until_sunday)

            print(
                f"Calculating 'Due By Friday' pivot column based on tasks "
                f"due on or before coming Sunday: {next_sunday_date}"
            )

            # Calculate helper columns for pivot aggregation
            # Count_Unblinded: Check if 'Content' contains 'Unblinded' (case-insensitive)
            df_for_pivot['Count_Unblinded'] = df_for_pivot[content_col].astype(str).str.contains(
                'Unblinded', na=False, case=False
            ).astype(int)

            # Count_DueToday: Check if Due Date is valid and <= today
            df_for_pivot['Count_DueToday'] = (
                ~pd.isna(df_for_pivot[due_date_col]) &
                (df_for_pivot[due_date_col] <= today_date)
            ).astype(int)

            # Count_DueByCriteriaDate: Check if Due Date is valid and <= Sunday
            df_for_pivot['Count_DueByCriteriaDate'] = (
                ~pd.isna(df_for_pivot[due_date_col]) &
                (df_for_pivot[due_date_col] <= next_sunday_date)
            ).astype(int)

            # Check if there's any data to pivot after potential NaNs
            if df_for_pivot.empty or df_for_pivot[rqc_user_col].isna().all():
                print("WARNING: No valid data found to create Pivot Table.")
                # Create an empty pivot table with expected columns
                pivot_table = pd.DataFrame(
                    columns=['Unblinded', 'Due By Today', 'Due By Friday']
                )
            else:
                # Create the pivot table
                pivot_table = pd.pivot_table(
                    df_for_pivot.dropna(subset=[rqc_user_col]), # Ignore rows with no RQC user
                    index=rqc_user_col,
                    values=[
                        'Count_Unblinded',
                        'Count_DueToday',
                        'Count_DueByCriteriaDate'
                    ],
                    aggfunc='sum',
                    fill_value=0  # Replace NaN in results with 0
                )

                # Rename columns for the final report
                pivot_table.rename(
                    columns={
                        'Count_Unblinded': 'Unblinded',
                        'Count_DueToday': 'Due By Today',
                        'Count_DueByCriteriaDate': 'Due By Friday' # Final name
                    },
                    inplace=True
                )

            print("Pivot Table calculated:")
            print(pivot_table)

            # --- Write Pivot Table to New Sheet ---
            pivot_sheet_name = 'Pivot_Summary'
            try:
                # Check if sheet exists, clear it; otherwise add it
                try:
                    sht_pivot = wb_template.sheets[pivot_sheet_name]
                    sht_pivot.clear_contents()
                    print(f"Cleared existing '{pivot_sheet_name}' sheet.")
                except Exception:  # Catch error if sheet doesn't exist
                    print(f"Adding '{pivot_sheet_name}' sheet for pivot table.")
                    # Add sheet after the last existing sheet
                    sht_pivot = wb_template.sheets.add(
                        pivot_sheet_name, after=wb_template.sheets[-1]
                    )

                # Write DataFrame to the sheet
                if not pivot_table.empty:
                    # options(index=True, header=True) writes index and columns
                    sht_pivot.range('A1').options(
                        index=True, header=True
                    ).value = pivot_table
                else:
                    # Write message if pivot is empty
                    sht_pivot.range('A1').value = "No data for Pivot Table"

                print(f"Pivot Table written to '{pivot_sheet_name}' sheet.")

            except Exception as pivot_write_err:
                print(
                    f"!!! ERROR writing Pivot Table to sheet: {pivot_write_err}"
                )
                traceback.print_exc()
                # Continue to saving attempt even if pivot writing fails

        except Exception as e:
            print(f"Error creating or writing Pivot Table: {e}")
            traceback.print_exc()
            # Continue to saving attempt

        # --- 8. Save Final Report ---
        # Restore Excel settings before saving
        excel_app.calculation = 'automatic'
        excel_app.screen_updating = True
        excel_app.display_alerts = True

        print(f"Saving final report to: {output_path}")
        try:
            # Save the modified template workbook to the new output path
            wb_template.save(output_path)
            print("Workbook saved successfully.")
        except Exception as save_err:
            print(f"!!! ERROR saving workbook: {save_err}")
            # Attempt to save with a different name as a recovery measure
            try:
                alt_output_path = output_path.replace(
                    ".xlsm",
                    f"_SAVE_ERROR_{datetime.now().strftime('%H%M%S')}.xlsm"
                )
                print(f"Attempting to save to alternate path: {alt_output_path}")
                wb_template.save(alt_output_path)
                print(f"Workbook saved to alternate path: {alt_output_path}")
            except Exception as alt_save_err:
                print(f"!!! FAILED to save to alternate path: {alt_save_err}")
                traceback.print_exc()
                # Indicate that saving failed completely

    except Exception as e:
        # Catch any unexpected errors during the main process
        print("\n--- An Unexpected Error Occurred During Main Processing ---")
        print(f"Error: {e}")
        traceback.print_exc()
        print("Attempting cleanup...")

    finally:
        # --- Cleanup: Ensure Excel instances are closed ---
        print("--- Running Cleanup ---")
        # Restore Excel settings in case of error before closing
        if excel_app is not None:
            try:
                excel_app.calculation = 'automatic'
                excel_app.screen_updating = True
                excel_app.display_alerts = True
            except Exception as setting_err:
                print(f"Warning: Error restoring Excel settings: {setting_err}")

        # Close workbooks gracefully
        if wb_source is not None: # Check if source was opened and not closed yet
             try:
                 wb_source.close()
                 print("Closed source workbook (in finally block).")
             except Exception as e:
                 print(f"Error closing source workbook: {e}")

        if wb_template is not None:
            was_saved = False
            # Check if the primary output path exists to infer successful save
            if 'output_path' in locals() and os.path.exists(output_path):
                was_saved = True
            if not was_saved:
                # Check for the alternate save path as well
                 if 'alt_output_path' in locals() and os.path.exists(alt_output_path):
                     was_saved = True

            if not was_saved:
                 print("Warning: Template workbook may not have saved correctly.")

            try:
                wb_template.close(save_changes=False) # Don't save again on close
                print("Closed template workbook.")
            except Exception as e:
                print(f"Error closing template workbook: {e}")

        # Quit Excel application
        if excel_app is not None:
            try:
                # Only quit if it's visible or has no open books left
                # (Avoids quitting a user's existing Excel instance if script attached)
                # Note: This check might not be foolproof depending on how xw manages instances
                if excel_app.visible or len(excel_app.books) == 0:
                     excel_app.quit()
                     print("Quit Excel application.")
                else:
                     print("Excel application not quit (possibly other books open).")
            except Exception as e:
                 # Ignore common "RPC server is unavailable" error if Excel already closed
                if "RPC server is unavailable" not in str(e):
                    print(f"Error quitting Excel application: {e}")

        print("Script finished.")

# --- Run the script ---
if __name__ == "__main__":
    main()
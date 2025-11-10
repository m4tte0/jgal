import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border
from openpyxl.utils import get_column_letter
import re
import os
from pathlib import Path
from copy import copy

def extract_date_from_filename(filename):
    """Extract date from Planning filename in format Planning_yy_mm_dd.xlsx"""
    match = re.search(r'Planning_(\d{2})_(\d{2})_(\d{2})\.xlsx', filename)
    if match:
        yy, mm, dd = match.groups()
        return f"20{yy}-{mm}-{dd}"
    return None

def copy_cell_style(source_cell, target_cell):
    """Copy all style attributes from source to target cell"""
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)

def main():
    # Define paths
    base_path = Path("_ref/usbilli")
    source_file = base_path / "Avanzamento schede 3° trimestre 2025.xlsx"
    planning_folder = base_path / "Planning"
    output_file = "Avanzamento_schede_automated.xlsx"

    # Get all Planning files and extract dates
    planning_files = sorted(planning_folder.glob("Planning_*.xlsx"))
    planning_dates = []
    for pf in planning_files:
        date = extract_date_from_filename(pf.name)
        if date:
            planning_dates.append((date, pf))

    print(f"Found {len(planning_dates)} Planning files")
    for date, pf in planning_dates:
        print(f"  - {pf.name}: {date}")

    # Load source workbook
    print(f"\nLoading source file: {source_file}")
    source_wb = openpyxl.load_workbook(source_file)
    source_ws = source_wb.active

    # Find columns to exclude
    columns_to_exclude = []
    header_row = 1  # Assuming headers are in row 1

    for col_idx in range(1, source_ws.max_column + 1):
        cell_value = source_ws.cell(header_row, col_idx).value
        if cell_value and ("Data effettiva avanzamento" in str(cell_value) or
                          "Data prevista avanzamento" in str(cell_value)):
            columns_to_exclude.append(col_idx)
            print(f"Column to exclude: {get_column_letter(col_idx)} - {cell_value}")

    # Create new workbook
    new_wb = openpyxl.Workbook()
    new_ws = new_wb.active
    new_ws.title = source_ws.title

    # Copy all data except excluded columns
    print("\nCopying data and formatting...")
    col_mapping = {}  # Maps old column index to new column index
    new_col_idx = 1

    for old_col_idx in range(1, source_ws.max_column + 1):
        if old_col_idx in columns_to_exclude:
            continue

        col_mapping[old_col_idx] = new_col_idx

        # Copy column width
        old_col_letter = get_column_letter(old_col_idx)
        new_col_letter = get_column_letter(new_col_idx)
        if old_col_letter in source_ws.column_dimensions:
            new_ws.column_dimensions[new_col_letter].width = source_ws.column_dimensions[old_col_letter].width

        # Copy all cells in this column
        for row_idx in range(1, source_ws.max_row + 1):
            source_cell = source_ws.cell(row_idx, old_col_idx)
            target_cell = new_ws.cell(row_idx, new_col_idx)

            # Copy value (but skip formulas that reference excluded columns)
            cell_value = source_cell.value
            if cell_value and isinstance(cell_value, str) and cell_value.startswith('='):
                # This is a formula - check if it references excluded columns
                skip_formula = False
                for excluded_col in columns_to_exclude:
                    excluded_letter = get_column_letter(excluded_col)
                    if excluded_letter in cell_value:
                        skip_formula = True
                        break

                if skip_formula:
                    # Clear the formula to avoid circular references
                    target_cell.value = None
                else:
                    target_cell.value = cell_value
            else:
                target_cell.value = cell_value

            # Copy style
            copy_cell_style(source_cell, target_cell)

        new_col_idx += 1

    # Copy row heights
    for row_idx in range(1, source_ws.max_row + 1):
        if row_idx in source_ws.row_dimensions:
            new_ws.row_dimensions[row_idx].height = source_ws.row_dimensions[row_idx].height

    # Add "Data prevista avanzamento" column for each Planning file
    print(f"\nAdding {len(planning_dates)} 'Data prevista avanzamento' columns...")

    # First, find the Matricola and Articolo columns in the new worksheet
    matricola_col_idx = None
    articolo_col_idx = None
    for col_idx in range(1, new_ws.max_column + 1):
        header_value = new_ws.cell(header_row, col_idx).value
        if header_value == "Matricola":
            matricola_col_idx = col_idx
        elif header_value == "Articolo":
            articolo_col_idx = col_idx

    if not matricola_col_idx or not articolo_col_idx:
        print("ERROR: Could not find 'Matricola' or 'Articolo' column in source file!")
        return None

    print(f"Found 'Matricola' column at index {matricola_col_idx}")
    print(f"Found 'Articolo' column at index {articolo_col_idx}")

    for idx, (date, planning_file) in enumerate(planning_dates):
        col_idx = new_col_idx + idx
        col_letter = get_column_letter(col_idx)

        # Set header
        header_cell = new_ws.cell(header_row, col_idx)
        if len(planning_dates) == 1:
            header_cell.value = "Data prevista avanzamento"
        else:
            header_cell.value = f"Data prevista avanzamento ({date})"

        # Copy header style from first column
        first_header = new_ws.cell(header_row, 1)
        copy_cell_style(first_header, header_cell)

        # Set column width
        new_ws.column_dimensions[col_letter].width = 20

        print(f"  Processing column {col_letter}: {header_cell.value}")

        # Load Planning file and extract dates
        print(f"    Loading Planning file: {planning_file}")
        planning_wb = openpyxl.load_workbook(planning_file, data_only=True)
        planning_ws = planning_wb.active

        # Build mappings from Planning file
        # Column 2 = Matricola, Column 4 = Articolo, Column 31 = Rilascio DiBa/Disegni (Mecc. + Idr.)
        matricola_to_date = {}
        articolo_to_date = {}
        planning_data_start_row = 5  # Data starts at row 5

        for row_idx in range(planning_data_start_row, planning_ws.max_row + 1):
            matricola_cell = planning_ws.cell(row_idx, 2)  # Column 2 = Matricola
            articolo_cell = planning_ws.cell(row_idx, 4)   # Column 4 = Articolo
            date_cell = planning_ws.cell(row_idx, 31)      # Column 31 = Rilascio DiBa/Disegni

            if matricola_cell.value:
                matricola = str(matricola_cell.value).strip()
                date_value = date_cell.value
                matricola_to_date[matricola] = date_value

            if articolo_cell.value:
                articolo = str(articolo_cell.value).strip()
                date_value = date_cell.value
                articolo_to_date[articolo] = date_value

        print(f"    Found {len(matricola_to_date)} matricola and {len(articolo_to_date)} articolo entries in Planning file")

        # Now populate the new worksheet by matching Matricola first, then Articolo as fallback
        matches_by_matricola = 0
        matches_by_articolo = 0
        for row_idx in range(2, new_ws.max_row + 1):  # Start from row 2 (skip header)
            target_matricola_cell = new_ws.cell(row_idx, matricola_col_idx)
            target_articolo_cell = new_ws.cell(row_idx, articolo_col_idx)
            target_date_cell = new_ws.cell(row_idx, col_idx)

            date_value = None

            # Try matching by Matricola first
            if target_matricola_cell.value:
                target_matricola = str(target_matricola_cell.value).strip()
                if target_matricola in matricola_to_date:
                    date_value = matricola_to_date[target_matricola]
                    matches_by_matricola += 1

            # If no match by Matricola, try Articolo
            if not date_value and target_articolo_cell.value:
                target_articolo = str(target_articolo_cell.value).strip()
                if target_articolo in articolo_to_date:
                    date_value = articolo_to_date[target_articolo]
                    matches_by_articolo += 1

            # Populate the cell if we found a match
            if date_value:
                target_date_cell.value = date_value
                # Copy number format for dates
                if date_value:
                    target_date_cell.number_format = 'YYYY-MM-DD'

        print(f"    Matched {matches_by_matricola} rows by Matricola, {matches_by_articolo} rows by Articolo")

        planning_wb.close()

    # Add consolidated "Data prevista avanzamento" column (no date in label)
    consolidated_col_idx = new_col_idx + len(planning_dates)
    consolidated_col_letter = get_column_letter(consolidated_col_idx)

    consolidated_header = new_ws.cell(header_row, consolidated_col_idx)
    consolidated_header.value = "Data prevista avanzamento"

    first_header = new_ws.cell(header_row, 1)
    copy_cell_style(first_header, consolidated_header)
    new_ws.column_dimensions[consolidated_col_letter].width = 20

    print(f"\nAdding consolidated column {consolidated_col_letter}: Data prevista avanzamento")

    # Populate consolidated column using the last Planning file date, ignoring 'KOM' values
    consolidated_count = 0
    for row_idx in range(2, new_ws.max_row + 1):  # Start from row 2 (skip header)
        # Collect valid dates from the planning columns (ignore 'KOM' and non-date values)
        # Start from the last planning file and work backwards
        last_valid_date = None

        for col_offset in range(len(planning_dates) - 1, -1, -1):  # Iterate backwards from last to first
            date_col_idx = new_col_idx + col_offset
            date_cell = new_ws.cell(row_idx, date_col_idx)
            cell_value = date_cell.value

            # Check if it's a valid date (not 'KOM' and not None)
            if cell_value and str(cell_value).strip().upper() != 'KOM':
                # Check if it's a datetime object (actual date)
                if hasattr(cell_value, 'year'):  # datetime object
                    last_valid_date = cell_value
                    break

        # Populate consolidated column if we found a valid date
        if last_valid_date:
            consolidated_cell = new_ws.cell(row_idx, consolidated_col_idx)
            consolidated_cell.value = last_valid_date
            consolidated_cell.number_format = 'YYYY-MM-DD'
            consolidated_count += 1

    print(f"  Populated {consolidated_count} rows with consolidated dates (using last valid Planning date)")

    # Add "Data effettiva avanzamento" column (empty for now, waiting for instructions)
    final_col_idx = consolidated_col_idx + 1
    final_col_letter = get_column_letter(final_col_idx)

    final_header = new_ws.cell(header_row, final_col_idx)
    final_header.value = "Data effettiva avanzamento"

    first_header = new_ws.cell(header_row, 1)
    copy_cell_style(first_header, final_header)
    new_ws.column_dimensions[final_col_letter].width = 20

    print(f"  Added column {final_col_letter}: Data effettiva avanzamento (empty)")

    # Save the new workbook
    print(f"\nSaving output file: {output_file}")
    new_wb.save(output_file)
    print("Done!")

    return output_file

if __name__ == "__main__":
    main()

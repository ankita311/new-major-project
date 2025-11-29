from collections import defaultdict
import math

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter


def _clean_value(value):
    """Convert NaN, None, or pandas NaN to empty string."""
    if value is None:
        return ""
    try:
        # Check for pandas/numpy NaN
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    try:
        # Check for float NaN
        if isinstance(value, float) and math.isnan(value):
            return ""
    except (TypeError, ValueError):
        pass
    # Check for string "nan" or "NaN"
    if isinstance(value, str) and value.lower() in ('nan', 'none', ''):
        return ""
    return value

def _split_roll_and_branch(raw_value: str):
    """Return (roll_no, branch) tuple from a raw string."""
    # Handle NaN, None, or empty values
    if raw_value is None or pd.isna(raw_value):
        return "", ""

    text = str(raw_value).strip()
    if not text or text.lower() == 'nan':
        return "", ""

    if "\n" in text:
        roll, branch = text.split("\n", 1)
    else:
        parts = text.split()
        roll = parts[0]
        branch = " ".join(parts[1:]) if len(parts) > 1 else ""

    return roll.strip(), branch.strip()

def upload_students(file):
    df = pd.read_excel(file,
                    sheet_name="main",
                    usecols = ['Roll No. Series-1', 'Roll No. Series-2'],
                    )

    pairs = df.to_dict(orient="records") #[{'Roll No. Series-1': str or nan , 'Roll No. Series-2': '2200970700064\n MBA-II'}]
    return pairs

def upload_rooms(file):
    df = pd.read_excel(file,
                    sheet_name="main",
                    usecols = ['Room No.', 'Row', 'Column'],
                    )

    rooms = df.dropna(how="all").to_dict(orient="records") #[{'Room No.': 'D-104' or nan, 'Row': 8.0 or nan , 'Column': 4.0 or nan}...]
    return rooms

def upload_college_sem(file):
    """Read college name and exam name from Excel file.
    Returns (college_name, exam_name) tuple from the first non-empty record."""
    df = pd.read_excel(file,
                    sheet_name="main",
                    usecols = ['College Name', 'Exam Name'])
    info = df.dropna(how="all").to_dict(orient="records")
    
    if not info:
        return "", ""
    
    first_record = info[0]
    college_name = str(first_record.get("College Name", "")).strip() if first_record.get("College Name") else ""
    exam_name = str(first_record.get("Exam Name", "")).strip() if first_record.get("Exam Name") else ""
    
    return college_name, exam_name

def find_capacity_per_room(rooms: dict):
    room_capacity = {}
    for room in rooms:
        room_no = room['Room No.']
        rows = int(room['Row'] or 0)
        cols = int(room['Column'] or 0)

        room_capacity[room_no] = {
            "rows": rows,
            "cols": cols,
            "capacity": rows * cols
        }
    return room_capacity        #room_capacity = {'D-104': {'rows':8,'cols':4,'capacity':32}...}

def fill_room(pairs: list, room_capacity: dict):
    room_layout = defaultdict(list)  # creates an empty dictionary with values as lists {some_key: []}
    branch_counts_per_room = defaultdict(lambda: defaultdict(int))  # {room_no: {branch: count}}
    pair_idx = 0

    for room_no, spec in room_capacity.items():
        rows = int(spec.get("rows", 0) or 0)
        cols = int(spec.get("cols", 0) or 0)

        # build a grid [row][col], but fill column by column so
        # students in the same "current_row" list end up in one column
        grid = [[None for _ in range(cols)] for _ in range(rows)]

        for c in range(cols):
            for r in range(rows):
                if pair_idx >= len(pairs):
                    break
                grid[r][c] = pairs[pair_idx]
                
                # Count branches for this pair
                pair = pairs[pair_idx]
                s1_raw = pair.get("Roll No. Series-1", pair.get("s1", ""))
                s2_raw = pair.get("Roll No. Series-2", pair.get("s2", ""))
                
                _, branch1 = _split_roll_and_branch(s1_raw)
                _, branch2 = _split_roll_and_branch(s2_raw)
                
                if branch1:
                    branch_counts_per_room[room_no][branch1] += 1
                if branch2:
                    branch_counts_per_room[room_no][branch2] += 1
                
                pair_idx += 1
            if pair_idx >= len(pairs):
                break

        # convert grid back to list-of-lists (rows), dropping empty seats
        cleaned_rows = []
        for r in range(rows):
            row_seats = [seat for seat in grid[r] if seat is not None]
            if row_seats:
                cleaned_rows.append(row_seats)

        room_layout[room_no] = cleaned_rows

        if pair_idx >= len(pairs):
            break  # no more students to allocate

    # Convert defaultdict to regular dict for return
    branch_counts_dict = {room: dict(branches) for room, branches in branch_counts_per_room.items()}
    
    return room_layout, branch_counts_dict      # ({'D-104': [[{pair1}, {pair2}, ...], [{pairN}, ...]]}, {'D-104': {'branch1': count, 'branch2': count}})

def generate_qpd(branch_counts_per_room: dict, sem: str, branch: str, subject_code: str):
    pass

def fill_room_row_gap(pairs: list, room_capacity: dict):
    room_layout = defaultdict(list)
    branch_counts_per_room = defaultdict(lambda: defaultdict(int))  # {room_no: {branch: count}}
    pair_idx = 0

    for room_no, spec in room_capacity.items():
        rows = int(spec.get("rows", 0) or 0)
        cols = int(spec.get("cols", 0) or 0)
        grid = [[None for _ in range(cols)] for _ in range(rows)]

        for c in range(cols):
            for r in range(rows):
                if r % 2 != 0:  # skip odd rows to keep alternate rows empty
                    continue
                if pair_idx >= len(pairs):
                    break
                grid[r][c] = pairs[pair_idx]
                
                # Count branches for this pair
                pair = pairs[pair_idx]
                s1_raw = pair.get("Roll No. Series-1", pair.get("s1", ""))
                s2_raw = pair.get("Roll No. Series-2", pair.get("s2", ""))
                
                _, branch1 = _split_roll_and_branch(s1_raw)
                _, branch2 = _split_roll_and_branch(s2_raw)
                
                if branch1:
                    branch_counts_per_room[room_no][branch1] += 1
                if branch2:
                    branch_counts_per_room[room_no][branch2] += 1
                
                pair_idx += 1
            if pair_idx >= len(pairs):
                break

        cleaned_rows = []
        for r in range(rows):
            row_seats = [seat for seat in grid[r] if seat is not None]
            # Include all rows (filled or empty) to show skipped alternate rows
            if row_seats:
                cleaned_rows.append(row_seats)
            else:
                # Insert empty row to indicate this row was skipped
                cleaned_rows.append([])

        room_layout[room_no] = cleaned_rows

        if pair_idx >= len(pairs):
            break

    unallocated = len(pairs) - pair_idx
    # Convert defaultdict to regular dict for return
    branch_counts_dict = {room: dict(branches) for room, branches in branch_counts_per_room.items()}
    
    return room_layout, unallocated, branch_counts_dict

def fill_room_col_gap(pairs: list, room_capacity: dict):
    room_layout = defaultdict(list)
    branch_counts_per_room = defaultdict(lambda: defaultdict(int))  # {room_no: {branch: count}}
    pair_idx = 0

    for room_no, spec in room_capacity.items():
        rows = int(spec.get("rows", 0) or 0)
        cols = int(spec.get("cols", 0) or 0)
        grid = [[None for _ in range(cols)] for _ in range(rows)]

        for c in range(cols):
            if c % 2 != 0:  # skip odd columns to leave a gap between columns
                continue
            for r in range(rows):
                if pair_idx >= len(pairs):
                    break
                grid[r][c] = pairs[pair_idx]
                
                # Count branches for this pair
                pair = pairs[pair_idx]
                s1_raw = pair.get("Roll No. Series-1", pair.get("s1", ""))
                s2_raw = pair.get("Roll No. Series-2", pair.get("s2", ""))
                
                _, branch1 = _split_roll_and_branch(s1_raw)
                _, branch2 = _split_roll_and_branch(s2_raw)
                
                if branch1:
                    branch_counts_per_room[room_no][branch1] += 1
                if branch2:
                    branch_counts_per_room[room_no][branch2] += 1
                
                pair_idx += 1
            if pair_idx >= len(pairs):
                break

        cleaned_rows = []
        for r in range(rows):
            # Include all columns (filled or empty) to show skipped alternate columns
            row_seats = grid[r]  # Keep all columns, including None for skipped columns
            # Only include rows that have at least one filled seat
            if any(seat is not None for seat in row_seats):
                cleaned_rows.append(row_seats)

        room_layout[room_no] = cleaned_rows

        if pair_idx >= len(pairs):
            break

    unallocated = len(pairs) - pair_idx
    # Convert defaultdict to regular dict for return
    branch_counts_dict = {room: dict(branches) for room, branches in branch_counts_per_room.items()}
    
    return room_layout, unallocated, branch_counts_dict


# def fill_room_one_student_per_bench(pairs: list, room_capacity: dict):
#     """
#     Fill room with one student per bench - students from the same pair are separated by a column gap.
#     Pattern: s1 in col 0, gap in col 1, s2 in col 2, s1 (next pair) in col 3, gap in col 4, s2 (next pair) in col 5, etc.
#     """
#     room_layout = defaultdict(list)
#     branch_counts_per_room = defaultdict(lambda: defaultdict(int))  # {room_no: {branch: count}}
#     pair_idx = 0

#     for room_no, spec in room_capacity.items():
#         rows = int(spec.get("rows", 0) or 0)
#         cols = int(spec.get("cols", 0) or 0)

#         # Build grid - need more columns to accommodate gaps (every 3rd column pattern: s1, gap, s2)
#         # Each pair takes 3 columns: s1, gap, s2
#         grid = [[None for _ in range(cols * 3)] for _ in range(rows)]  # Expand columns to fit gaps
        
#         for c in range(0, cols * 3, 3):  # Step by 3: s1 column, gap column, s2 column
#             for r in range(rows):
#                 if pair_idx >= len(pairs):
#                     break
                
#                 pair = pairs[pair_idx]
#                 s1_raw = pair.get("Roll No. Series-1", pair.get("s1", ""))
#                 s2_raw = pair.get("Roll No. Series-2", pair.get("s2", ""))
                
#                 # Count branches
#                 _, branch1 = _split_roll_and_branch(s1_raw)
#                 _, branch2 = _split_roll_and_branch(s2_raw)
                
#                 if branch1:
#                     branch_counts_per_room[room_no][branch1] += 1
#                 if branch2:
#                     branch_counts_per_room[room_no][branch2] += 1
                
#                 # Create single-student pairs: s1 only in first column, s2 only in third column
#                 # Column c: s1 (with empty s2)
#                 grid[r][c] = {
#                     "Roll No. Series-1": s1_raw,
#                     "Roll No. Series-2": ""
#                 }
                
#                 # Column c+1: gap (None, already set)
#                 # Column c+2: s2 (with empty s1)
#                 grid[r][c + 2] = {
#                     "Roll No. Series-1": "",
#                     "Roll No. Series-2": s2_raw
#                 }
                
#                 pair_idx += 1
#             if pair_idx >= len(pairs):
#                 break

#         # Convert grid back to list-of-lists (rows)
#         # Keep the structure with gaps (None values) to show the column separation
#         cleaned_rows = []
#         for r in range(rows):
#             row_seats = []
#             for c in range(cols * 3):
#                 seat = grid[r][c]
#                 if seat is not None:
#                     row_seats.append(seat)
#                 elif c % 3 == 1:  # This is a gap column, keep it as None to show gap
#                     row_seats.append(None)
#             # Only include rows that have at least one filled seat
#             if any(seat is not None for seat in row_seats):
#                 cleaned_rows.append(row_seats)

#         room_layout[room_no] = cleaned_rows

#         if pair_idx >= len(pairs):
#             break
    
#     unallocated = len(pairs) - pair_idx
#     # Convert defaultdict to regular dict for return
#     branch_counts_dict = {room: dict(branches) for room, branches in branch_counts_per_room.items()}
    
#     return room_layout, unallocated, branch_counts_dict


def build_room_sheet(ws, room_name: str, rows: list, college_name: str = "", exam_name: str = "", branch_counts: dict = None):
    if not rows:
        return

    max_seats = max(len(row) for row in rows)
    total_columns = max(1, max_seats * 2)  # s1 and s2 occupy separate columns
    arrow_banner = "^" * (max(5, total_columns * 2))
    
    # Use provided branch_counts or empty dict if not provided
    if branch_counts is None:
        branch_counts = {}

    # Calculate required width for college name (font size 20, bold)
    # Excel column width: 1 unit ≈ 1 character at default font size
    # For font size 20, we need approximately: len(text) * (20/11) * 1.2 (for bold)
    base_column_width = 18  # Default width for seating columns
    column_width = base_column_width  # Will be adjusted if needed
    
    if college_name:
        college_name_clean = _clean_value(college_name) or ""
        if college_name_clean:
            # Estimate width needed: account for larger font (20pt vs 11pt default) and bold
            # Font size 20 is ~1.8x larger than default 11pt, plus 1.3 factor for bold and spacing
            required_width = len(college_name_clean) * (20 / 11) * 1.3
            current_total_width = total_columns * base_column_width
            
            # If college name needs more width, adjust column widths
            if required_width > current_total_width:
                # Calculate new column width to accommodate college name
                column_width = max(base_column_width, required_width / total_columns)
    
    # Set all columns to the calculated width (before displaying college name)
    for col in range(1, total_columns + 1):
        ws.column_dimensions[get_column_letter(col)].width = column_width

    def merge_and_set(row_idx, value, font_size=14, bold=True):
        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=total_columns)
        cell = ws.cell(row=row_idx, column=1, value=_clean_value(value))
        cell.font = Font(size=font_size, bold=bold)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    current_row = 1
    
    # Display college name in big font (no blank line after)
    if college_name:
        merge_and_set(current_row, college_name, font_size=20, bold=True)
        current_row += 1
    
    # Display exam name in big font (no blank line after)
    if exam_name:
        merge_and_set(current_row, exam_name, font_size=20, bold=True)
        current_row += 1
    
    # Display 'Seating Plan' heading (no blank line after)
    merge_and_set(current_row, "Seating Plan", font_size=18)
    current_row += 1
    
    # Display room name (no blank line after)
    merge_and_set(current_row, room_name, font_size=16)
    current_row += 1

    thin = Side(border_style="thin", color="000000")
    border = Border(top=thin, bottom=thin, left=thin, right=thin)
    # data_start_row starts right after the room name (blackboard will be first row of table)
    data_start_row = current_row + 1
    
    # Add blackboard heading as first row of the table
    blackboard_row = data_start_row
    ws.row_dimensions[blackboard_row].height = 36
    ws.merge_cells(start_row=blackboard_row, start_column=1, end_row=blackboard_row, end_column=total_columns)
    blackboard_cell = ws.cell(row=blackboard_row, column=1, value=f"{arrow_banner}  Black Board  {arrow_banner}")
    blackboard_cell.font = Font(size=11, bold=False)
    blackboard_cell.alignment = Alignment(horizontal="center", vertical="center")
    # Apply border to the merged cell (apply to all cells in merged range for proper display)
    for col in range(1, total_columns + 1):
        ws.cell(row=blackboard_row, column=col).border = border
    
    # Adjust data_start_row to start after blackboard row
    data_start_row = blackboard_row + 1

    for row_offset, row in enumerate(rows):
        excel_row = data_start_row + row_offset
        ws.row_dimensions[excel_row].height = 36

        for seat_idx in range(1, max_seats + 1):
            s1_col = (seat_idx - 1) * 2 + 1
            s2_col = s1_col + 1

            for col in (s1_col, s2_col):
                cell = ws.cell(row=excel_row, column=col)
                # Don't overwrite column width - use the width already set for college name
                # ws.column_dimensions[get_column_letter(col)].width is already set above
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = border

            if seat_idx <= len(row):
                student = row[seat_idx - 1] or {}
                s1_raw = student.get("Roll No. Series-1", student.get("s1", ""))
                s2_raw = student.get("Roll No. Series-2", student.get("s2", ""))
                
                # Clean NaN values before processing
                s1_raw = _clean_value(s1_raw) if s1_raw else ""
                s2_raw = _clean_value(s2_raw) if s2_raw else ""

                roll1, branch1 = _split_roll_and_branch(s1_raw)
                roll2, branch2 = _split_roll_and_branch(s2_raw)

                s1_value = "\n".join(filter(None, [roll1, branch1]))
                s2_value = "\n".join(filter(None, [roll2, branch2]))
                ws.cell(row=excel_row, column=s1_col).value = s1_value if s1_value else ""
                ws.cell(row=excel_row, column=s2_col).value = s2_value if s2_value else ""

    if branch_counts:
        summary_start = data_start_row + len(rows) + 2
        # Use columns after the seating table to avoid conflicts
        # Place summary in a dedicated area (columns 1-2, but ensure proper width)
        summary_col1 = 1
        summary_col2 = 2
        
        # Header row
        header_row = summary_start
        name_header = ws.cell(header_row, summary_col1, "Branch Name")
        name_header.font = Font(bold=True, size=12)
        name_header.alignment = Alignment(horizontal="left", vertical="center")
        
        count_header = ws.cell(header_row, summary_col2, "No. of Students")
        count_header.font = Font(bold=True, size=12)
        count_header.alignment = Alignment(horizontal="left", vertical="center")
        
        # Ensure column widths are adequate for summary (only if not already set wider)
        if summary_col1 <= total_columns:
            current_width = ws.column_dimensions[get_column_letter(summary_col1)].width or 18
            ws.column_dimensions[get_column_letter(summary_col1)].width = max(current_width, 20)
        else:
            ws.column_dimensions[get_column_letter(summary_col1)].width = 20
            
        if summary_col2 <= total_columns:
            current_width = ws.column_dimensions[get_column_letter(summary_col2)].width or 18
            ws.column_dimensions[get_column_letter(summary_col2)].width = max(current_width, 18)
        else:
            ws.column_dimensions[get_column_letter(summary_col2)].width = 18
        
        # Data rows - each branch gets its own row
        for idx, (branch, count) in enumerate(branch_counts.items(), start=1):
            data_row = summary_start + idx
            name_cell = ws.cell(data_row, summary_col1, _clean_value(branch))
            name_cell.alignment = Alignment(horizontal="left", vertical="center")
            name_cell.font = Font(size=11)
            
            count_cell = ws.cell(data_row, summary_col2, _clean_value(count))
            count_cell.alignment = Alignment(horizontal="center", vertical="center")
            count_cell.font = Font(size=11)


def build_workbook(room_layout: dict, output_path: str = "seating_plan.xlsx", college_name: str = "", exam_name: str = "", branch_counts_per_room: dict = None):
    wb = Workbook()
    ws = wb.active
    first_sheet = True

    for room_name, rows in room_layout.items():
        if first_sheet:
            ws.title = room_name
            first_sheet = False
        else:
            ws = wb.create_sheet(title=room_name)
        
        # Get branch counts for this room if provided
        branch_counts = branch_counts_per_room.get(room_name, {}) if branch_counts_per_room else {}
        build_room_sheet(ws, room_name, rows, college_name, exam_name, branch_counts)

    wb.save(output_path)

def build_workbook_in_memory(room_layout: dict, college_name: str = "", exam_name: str = "", branch_counts_per_room: dict = None):
    """Build workbook in memory and return the workbook object (doesn't save to disk)."""
    wb = Workbook()
    ws = wb.active
    first_sheet = True

    # Process all rooms to ensure all sheets are created
    for room_name, rows in room_layout.items():
        if first_sheet:
            ws.title = room_name
            first_sheet = False
        else:
            ws = wb.create_sheet(title=room_name)
        
        # Get branch counts for this room if provided
        branch_counts = branch_counts_per_room.get(room_name, {}) if branch_counts_per_room else {}
        build_room_sheet(ws, room_name, rows, college_name, exam_name, branch_counts)
    
    # Ensure workbook is properly finalized
    # Remove default empty sheet if we created other sheets
    if len(wb.worksheets) > 1 and wb.worksheets[0].title == "Sheet":
        wb.remove(wb.worksheets[0])

    return wb



if __name__ == "__main__":
    with open("C:/Users/Ankita/OneDrive/Desktop/CAE-II_JULY_2023_MS.xlsx", "rb") as f:
        pairs = upload_students(f)
        rooms = upload_rooms(f)
        college_name, exam_name = upload_college_sem(f)
        room_capacity = find_capacity_per_room(rooms)

        ## Uncomment the below functions to generate different formats
        room_layout, branch_counts_per_room = fill_room(pairs, room_capacity)
        # room_layout, unallocated, branch_counts_per_room = fill_room_row_gap(pairs, room_capacity)
        # room_layout, unallocated, branch_counts_per_room = fill_room_col_gap(pairs, room_capacity)
        # room_layout, unallocated, branch_counts_per_room = fill_room_one_student_per_bench(pairs, room_capacity)
        # print(branch_counts_per_room)
        build_workbook(room_layout, "seating_plan.xlsx", college_name, exam_name, branch_counts_per_room)


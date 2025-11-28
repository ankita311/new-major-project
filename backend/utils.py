from collections import defaultdict

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter


def _split_roll_and_branch(raw_value: str):
    """Return (roll_no, branch) tuple from a raw string."""
    if raw_value is None:
        return "", ""

    text = str(raw_value).strip()
    if not text:
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

    return room_layout      #{'D-104': [[{pair1}, {pair2}, ...], [{pairN}, ...]]}

def fill_room_row_gap(pairs: list, room_capacity: dict):
    room_layout = defaultdict(list)
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
    return room_layout, unallocated

def fill_room_col_gap(pairs: list, room_capacity: dict):
    room_layout = defaultdict(list)
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
    return room_layout, unallocated

def build_room_sheet(ws, room_name: str, rows: list, college_name: str = "", exam_name: str = ""):
    if not rows:
        return

    max_seats = max(len(row) for row in rows)
    total_columns = max(1, max_seats * 2)  # s1 and s2 occupy separate columns
    arrow_banner = "^" * (max(5, total_columns * 2))
    branch_counts = defaultdict(int)

    def merge_and_set(row_idx, value, font_size=14, bold=True):
        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=total_columns)
        cell = ws.cell(row=row_idx, column=1, value=value)
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
                ws.column_dimensions[get_column_letter(col)].width = 18
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = border

            if seat_idx <= len(row):
                student = row[seat_idx - 1] or {}
                s1_raw = student.get("Roll No. Series-1", student.get("s1", ""))
                s2_raw = student.get("Roll No. Series-2", student.get("s2", ""))

                roll1, branch1 = _split_roll_and_branch(s1_raw)
                roll2, branch2 = _split_roll_and_branch(s2_raw)

                ws.cell(row=excel_row, column=s1_col).value = "\n".join(filter(None, [roll1, branch1]))
                ws.cell(row=excel_row, column=s2_col).value = "\n".join(filter(None, [roll2, branch2]))

                if branch1:
                    branch_counts[branch1] += 1
                if branch2:
                    branch_counts[branch2] += 1

    if branch_counts:
        summary_start = data_start_row + len(rows) + 2
        name_header = ws.cell(summary_start, 1, "Branch Name")
        count_header = ws.cell(summary_start + 1, 1, "No. of Students")
        name_header.font = Font(bold=True)
        count_header.font = Font(bold=True)
        name_header.alignment = Alignment(horizontal="left")
        count_header.alignment = Alignment(horizontal="left")

        for idx, (branch, count) in enumerate(branch_counts.items(), start=2):
            name_cell = ws.cell(summary_start, idx, branch)
            count_cell = ws.cell(summary_start + 1, idx, count)
            name_cell.alignment = Alignment(horizontal="center")
            count_cell.alignment = Alignment(horizontal="center")


def build_workbook(room_layout: dict, output_path: str = "seating_plan.xlsx", college_name: str = "", exam_name: str = ""):
    wb = Workbook()
    ws = wb.active
    first_sheet = True

    for room_name, rows in room_layout.items():
        if first_sheet:
            ws.title = room_name
            first_sheet = False
        else:
            ws = wb.create_sheet(title=room_name)
        build_room_sheet(ws, room_name, rows, college_name, exam_name)

    wb.save(output_path)


if __name__ == "__main__":
    with open("C:/Users/Ankita/OneDrive/Desktop/CAE-II_JULY_2023_MS.xlsx", "rb") as f:
        pairs = upload_students(f)
        rooms = upload_rooms(f)
        college_name, exam_name = upload_college_sem(f)
        room_capacity = find_capacity_per_room(rooms)

        ## Uncomment the below functions to generate different formats
        room_layout = fill_room(pairs, room_capacity)
        # room_layout, unallocated = fill_room_row_gap(pairs, room_capacity)
        # room_layout, unallocated = fill_room_col_gap(pairs, room_capacity)
        build_workbook(room_layout, "seating_plan.xlsx", college_name, exam_name)



    

from fastapi import FastAPI, File, UploadFile
from fastapi.responses import StreamingResponse
from io import BytesIO
import tempfile
import os
import time
from backend import utils, schemas
from backend.schemas import UploadInfo

app = FastAPI()

@app.get('/root')
def root():
    return {"message": "Exam Hall Seat Allocation System"}

@app.post('/upload-file', response_model= schemas.UploadInfo)
async def upload_file(file: UploadFile = File(...)):
    f = file.file
    
    pairs = utils.upload_students(f)
    f.seek(0)

    rooms = utils.upload_rooms(f)
    f.seek(0)

    college_name, exam_name = utils.upload_college_sem(f)
    f.seek(0)

    room_capacity = utils.find_capacity_per_room(rooms)
    f.seek(0)

    return {
        "pairs": pairs,
        "rooms": rooms,
        "college_name": college_name,
        "exam_name": exam_name,
        "room_capacity": room_capacity
    }

@app.post('/generate-plan')
def generate_plan(info: schemas.UploadInfo):
    room_layout, branch_counts_per_room = utils.fill_room(info.pairs, info.room_capacity)
    
    # Generate workbook in memory
    wb = utils.build_workbook_in_memory(room_layout, info.college_name, info.exam_name, branch_counts_per_room)
    
    # Save to temporary file to ensure complete write
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            tmp_path = tmp_file.name
        
        # Save workbook and ensure it's closed
        wb.save(tmp_path)
        wb.close()
        
        # Ensure file is fully written by checking file size is stable
        prev_size = 0
        for _ in range(10):  # Check up to 10 times
            if os.path.exists(tmp_path):
                current_size = os.path.getsize(tmp_path)
                if current_size == prev_size and current_size > 0:
                    break
                prev_size = current_size
            time.sleep(0.1)
        
        # Read the complete file into BytesIO
        with open(tmp_path, 'rb') as f:
            file_bytes = f.read()
        
        # Create BytesIO from complete bytes
        file_stream = BytesIO(file_bytes)
        
        # Return as streaming response
        return StreamingResponse(
            file_stream,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=seating_plan.xlsx"}
        )
    finally:
        # Clean up temporary file
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)

@app.post('/generate-plan-row-gap')
def generate_plan_row_gap(info: schemas.UploadInfo):
    room_layout, unallocated, branch_counts_per_room = utils.fill_room_row_gap(info.pairs, info.room_capacity)
    
    # Generate workbook in memory
    wb = utils.build_workbook_in_memory(room_layout, info.college_name, info.exam_name, branch_counts_per_room)
    
    # Save to temporary file to ensure complete write
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            tmp_path = tmp_file.name
        
        # Save workbook and ensure it's closed
        wb.save(tmp_path)
        wb.close()
        
        # Ensure file is fully written by checking file size is stable
        import time
        prev_size = 0
        for _ in range(10):  # Check up to 10 times
            if os.path.exists(tmp_path):
                current_size = os.path.getsize(tmp_path)
                if current_size == prev_size and current_size > 0:
                    break
                prev_size = current_size
            time.sleep(0.1)
        
        # Read the complete file into BytesIO
        with open(tmp_path, 'rb') as f:
            file_bytes = f.read()
        
        # Create BytesIO from complete bytes
        file_stream = BytesIO(file_bytes)
        
        # Return as streaming response with unallocated count in headers
        return StreamingResponse(
            file_stream,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": "attachment; filename=seating_plan_row_gap.xlsx",
                "Unallocated-Seats": str(unallocated)
            }
        )
    finally:
        # Clean up temporary file
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)

@app.post('/generate-plan-col-gap')
def generate_plan_col_gap(info: schemas.UploadInfo):
    room_layout, unallocated, branch_counts_per_room = utils.fill_room_col_gap(info.pairs, info.room_capacity)
    
    # Generate workbook in memory
    wb = utils.build_workbook_in_memory(room_layout, info.college_name, info.exam_name, branch_counts_per_room)
    
    # Save to temporary file to ensure complete write
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            tmp_path = tmp_file.name
        
        # Save workbook and ensure it's closed
        wb.save(tmp_path)
        wb.close()
        
        # Ensure file is fully written by checking file size is stable
        import time
        prev_size = 0
        for _ in range(10):  # Check up to 10 times
            if os.path.exists(tmp_path):
                current_size = os.path.getsize(tmp_path)
                if current_size == prev_size and current_size > 0:
                    break
                prev_size = current_size
            time.sleep(0.1)
        
        # Read the complete file into BytesIO
        with open(tmp_path, 'rb') as f:
            file_bytes = f.read()
        
        # Create BytesIO from complete bytes
        file_stream = BytesIO(file_bytes)
        
        # Return as streaming response
        return StreamingResponse(
            file_stream,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": "attachment; filename=seating_plan_col_gap.xlsx",
                "Unallocated-Seats": str(unallocated)
            }
        )
    finally:
        # Clean up temporary file
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("backend.main:app", host="127.0.0.1", port=8000, reload=True)




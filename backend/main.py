from fastapi import FastAPI, File, UploadFile
from fastapi.responses import StreamingResponse
from io import BytesIO
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
    room_layout = utils.fill_room(info.pairs, info.room_capacity)
    
    # Generate workbook in memory
    wb = utils.build_workbook_in_memory(room_layout, info.college_name, info.exam_name)
    
    # Save to BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    
    # Return as streaming response
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=seating_plan.xlsx"}
    )

@app.post('/generate-plan-row-gap')
def generate_plan_row_gap(info: schemas.UploadInfo):
    room_layout, unallocated = utils.fill_room_row_gap(info.pairs, info.room_capacity)
    
    # Generate workbook in memory
    wb = utils.build_workbook_in_memory(room_layout, info.college_name, info.exam_name)
    
    # Save to BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    
    # Return as streaming response
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=seating_plan_row_gap.xlsx"}
    )

@app.post('/generate-plan-col-gap')
def generate_plan_col_gap(info: schemas.UploadInfo):
    room_layout, unallocated= utils.fill_room_col_gap(info.pairs, info.room_capacity)
    
    # Generate workbook in memory
    wb = utils.build_workbook_in_memory(room_layout, info.college_name, info.exam_name)
    
    # Save to BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    
    # Return as streaming response
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=seating_plan_col_gap.xlsx"}
    )

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("backend.main:app", host="127.0.0.1", port=8000, reload=True)




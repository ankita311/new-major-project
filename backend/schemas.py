from pydantic import BaseModel

class UploadInfo(BaseModel):
    pairs: list
    rooms: list
    college_name: str
    exam_name: str
    room_capacity: dict
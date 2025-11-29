#Ignore the main.py file as it contains the api routes.
#All main functions are in utils.py 
##To run main.py

1. cd backend
2. uv sync
3. cd ..
4. python -m backend.main

##To run utils.py
1. cd backend
2. uv sync
3. python utils.py
4. seating_plan.xlsx will be generated

## utils.py

Utility module for exam hall seat allocation system. Handles:
- **Data Upload**: Reads student pairs, room configurations, and college/exam info from Excel
- **Room Allocation**: Fills rooms with students using different patterns (normal, row gap, column gap)
- **Branch Counting**: Tracks student counts per branch for each room
- **Excel Generation**: Creates formatted seating plan workbooks with room layouts, headers, and branch summaries
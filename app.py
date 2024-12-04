from fastapi import FastAPI
from fastapi.responses import StreamingResponse
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Or use specific origins like ["http://localhost:3000"]
    allow_credentials=True,
    allow_methods=["*"],  # Allows all methods (GET, POST, etc.)
    allow_headers=["*"],  # Allows all headers
)

@app.post("/generate-excel/")
async def generate_excel(data: dict):
    # Extract columns and rows from the input JSON
    columns = data.get("columns", [])
    rows = data.get("rows", [])
    sheetname = data.get("sheetname", str)
    filename = data.get("filename", str)

    # Create a DataFrame
    df = pd.DataFrame(rows, columns=columns)

    # Save the DataFrame to an Excel file in memory
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheetname)
        workbook = writer.book
        worksheet = writer.sheets[sheetname]

        # Apply styles to the header row
        header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        header_font = Font(bold=True)

        for cell in worksheet[1]:  # 1st row for the header
            cell.fill = header_fill
            cell.font = header_font
            cell.border = None
    
    buffer.seek(0)

    # Return the Excel file as a streaming response
    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}.xlsx"}
    )

# Run the server with: uvicorn filename:app --reload

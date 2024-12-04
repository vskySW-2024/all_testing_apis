from fastapi import FastAPI
from fastapi.responses import StreamingResponse
import pandas as pd
import io

app = FastAPI()

@app.post("/generate-excel/")
async def generate_excel(data: dict):
    # Extract columns and rows from the input JSON
    columns = data.get("columns", [])
    rows = data.get("rows", [])

    # Create a DataFrame
    df = pd.DataFrame(rows, columns=columns)

    # Save the DataFrame to an Excel file in memory
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    
    buffer.seek(0)

    # Return the Excel file as a streaming response
    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=data.xlsx"}
    )

# Run the server with: uvicorn filename:app --reload

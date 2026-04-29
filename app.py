import os
import shutil
import uuid
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse
from parse_docx import process_document
from dotenv import load_dotenv
load_dotenv()

host = os.getenv("host")
port = int(os.getenv("port"))

app = FastAPI(title="UPSC Question Parser API")

# Ensure temporary directory exists
TEMP_DIR = "temp_uploads"
os.makedirs(TEMP_DIR, exist_ok=True)

@app.post("/process")
async def process_docx_endpoint(file: UploadFile = File(...)):
    # Validate file extension
    if not file.filename.endswith(".docx"):
        raise HTTPException(status_code=400, detail="Only .docx files are supported.")

    # Create a unique filename to avoid collisions
    file_id = str(uuid.uuid4())
    temp_file_path = os.path.join(TEMP_DIR, f"{file_id}_{file.filename}")
    output_json_path = os.path.join(TEMP_DIR, f"{file_id}_result.json")

    try:
        # Save the uploaded file
        with open(temp_file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Process the document
        # Returns a dict with "processed", "parsing_failures", and "translation_failures"
        result = process_document(temp_file_path, output_file=output_json_path)

        return JSONResponse(content=result)

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

    finally:
        # Delete the uploaded .docx file but keep the .json result
        if os.path.exists(temp_file_path):
            os.remove(temp_file_path)

@app.get("/")
async def root():
    return {"message": "hello world"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app",host=host, port=port,reload=True)

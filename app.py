import os
import shutil
import uuid
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse
from fastapi.concurrency import run_in_threadpool
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
async def process_docx_endpoint(
    english_file: UploadFile = File(...),
    hindi_file: UploadFile = File(None)
):
    # Validate file extensions
    if not english_file.filename.endswith(".docx"):
        raise HTTPException(status_code=400, detail="Only .docx files are supported for English.")
    
    if hindi_file and not hindi_file.filename.endswith(".docx"):
        raise HTTPException(status_code=400, detail="Only .docx files are supported for Hindi.")

    # Create unique filenames and directories to avoid collisions
    file_id = str(uuid.uuid4())
    en_temp_path = os.path.join(TEMP_DIR, f"{file_id}_en_{english_file.filename}")
    hi_temp_path = os.path.join(TEMP_DIR, f"{file_id}_hi_{hindi_file.filename}") if hindi_file else None
    output_json_path = os.path.join(TEMP_DIR, f"{file_id}_result.json")
    
    # Isolated image directory for this specific request
    img_dir = os.path.join("data", "images", file_id)

    try:
        # Save the English file
        with open(en_temp_path, "wb") as buffer:
            shutil.copyfileobj(english_file.file, buffer)
            
        # Save the Hindi file if provided
        if hi_temp_path:
            with open(hi_temp_path, "wb") as buffer:
                shutil.copyfileobj(hindi_file.file, buffer)

        # Process the documents using run_in_threadpool to avoid blocking the main event loop
        # Returns a dict with "processed", "parsing_failures", and "translation_failures"
        result = await run_in_threadpool(
            process_document, 
            en_temp_path, 
            hi_file_path=hi_temp_path, 
            output_file=output_json_path,
            img_dir=img_dir
        )

        return JSONResponse(content=result)

    except Exception as e:
        import traceback
        print(traceback.format_exc())
        raise HTTPException(status_code=500, detail=str(e))

    finally:
        # Cleanup temporary files
        if os.path.exists(en_temp_path):
            os.remove(en_temp_path)
        if hi_temp_path and os.path.exists(hi_temp_path):
            os.remove(hi_temp_path)

@app.get("/")
async def root():
    return {"message": "hello world"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app",host=host, port=port,reload=True)

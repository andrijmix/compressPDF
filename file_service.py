from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
from pathlib import Path
import os
import shutil
from datetime import datetime, timedelta
import asyncio

app = FastAPI(title="File Storage Service")

# Storage configuration
STORAGE_DIR = Path(os.getenv("STORAGE_DIR", "./storage"))
STORAGE_DIR.mkdir(exist_ok=True)

# File retention (clean up files older than 1 hour)
FILE_RETENTION_HOURS = int(os.getenv("FILE_RETENTION_HOURS", "1"))


@app.post("/upload/{job_id}")
async def upload_file(job_id: str, file: UploadFile = File(...)):
    """
    Upload a file for processing
    """
    try:
        # Create job directory
        job_dir = STORAGE_DIR / job_id
        job_dir.mkdir(exist_ok=True)

        # Save original file
        original_path = job_dir / f"original_{file.filename}"
        with open(original_path, "wb") as buffer:
            content = await file.read()
            buffer.write(content)

        # Store metadata
        metadata = {
            "filename": file.filename,
            "size": len(content),
            "upload_time": datetime.now().isoformat(),
            "content_type": file.content_type
        }

        import json
        with open(job_dir / "metadata.json", "w") as f:
            json.dump(metadata, f)

        return {
            "job_id": job_id,
            "filename": file.filename,
            "size": len(content),
            "status": "uploaded"
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Upload failed: {str(e)}")


@app.get("/file/{job_id}/original")
async def get_original_file(job_id: str):
    """
    Get the original uploaded file path
    """
    job_dir = STORAGE_DIR / job_id
    if not job_dir.exists():
        raise HTTPException(status_code=404, detail="Job not found")

    # Find original file
    original_files = list(job_dir.glob("original_*"))
    if not original_files:
        raise HTTPException(status_code=404, detail="Original file not found")

    return {"file_path": str(original_files[0])}


@app.post("/file/{job_id}/compressed")
async def save_compressed_file(job_id: str, file: UploadFile = File(...)):
    """
    Save the compressed file
    """
    try:
        job_dir = STORAGE_DIR / job_id
        if not job_dir.exists():
            raise HTTPException(status_code=404, detail="Job not found")

        # Save compressed file
        compressed_path = job_dir / f"compressed_{file.filename}"
        with open(compressed_path, "wb") as buffer:
            content = await file.read()
            buffer.write(content)

        return {
            "job_id": job_id,
            "compressed_size": len(content),
            "status": "saved"
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Save failed: {str(e)}")


@app.get("/download/{job_id}")
async def download_compressed_file(job_id: str):
    """
    Download the compressed file
    """
    job_dir = STORAGE_DIR / job_id
    if not job_dir.exists():
        raise HTTPException(status_code=404, detail="Job not found")

    # Find compressed file
    compressed_files = list(job_dir.glob("compressed_*"))
    if not compressed_files:
        raise HTTPException(status_code=404, detail="Compressed file not found")

    compressed_file = compressed_files[0]
    return FileResponse(
        path=str(compressed_file),
        filename=compressed_file.name,
        media_type="application/pdf"
    )


@app.get("/info/{job_id}")
async def get_file_info(job_id: str):
    """
    Get file information
    """
    job_dir = STORAGE_DIR / job_id
    if not job_dir.exists():
        raise HTTPException(status_code=404, detail="Job not found")

    # Read metadata
    metadata_path = job_dir / "metadata.json"
    if not metadata_path.exists():
        raise HTTPException(status_code=404, detail="Metadata not found")

    import json
    with open(metadata_path, "r") as f:
        metadata = json.load(f)

    # Check for compressed file
    compressed_files = list(job_dir.glob("compressed_*"))
    if compressed_files:
        compressed_size = compressed_files[0].stat().st_size
        metadata["compressed_size"] = compressed_size
        metadata["compression_ratio"] = round((1 - compressed_size / metadata["size"]) * 100, 2)

    return metadata


@app.delete("/cleanup/{job_id}")
async def cleanup_job(job_id: str):
    """
    Clean up job files
    """
    job_dir = STORAGE_DIR / job_id
    if job_dir.exists():
        shutil.rmtree(job_dir)
        return {"status": "cleaned up"}
    return {"status": "not found"}


@app.get("/health")
async def health_check():
    """
    Health check
    """
    return {
        "status": "healthy",
        "storage_dir": str(STORAGE_DIR),
        "storage_exists": STORAGE_DIR.exists()
    }


async def cleanup_old_files():
    """
    Background task to clean up old files
    """
    while True:
        try:
            cutoff_time = datetime.now() - timedelta(hours=FILE_RETENTION_HOURS)

            for job_dir in STORAGE_DIR.iterdir():
                if not job_dir.is_dir():
                    continue

                metadata_path = job_dir / "metadata.json"
                if not metadata_path.exists():
                    continue

                import json
                try:
                    with open(metadata_path, "r") as f:
                        metadata = json.load(f)

                    upload_time = datetime.fromisoformat(metadata["upload_time"])
                    if upload_time < cutoff_time:
                        shutil.rmtree(job_dir)
                        print(f"Cleaned up old job: {job_dir.name}")
                except:
                    pass

        except Exception as e:
            print(f"Cleanup error: {e}")

        # Wait 30 minutes before next cleanup
        await asyncio.sleep(1800)


@app.on_event("startup")
async def startup_event():
    """
    Start background cleanup task
    """
    asyncio.create_task(cleanup_old_files())


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8001)
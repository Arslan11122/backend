from fastapi import APIRouter, File, UploadFile, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from typing import List
import os
import time
import asyncio
from pathlib import Path
import aiofiles

from models.conversion import (
    ConversionJob, 
    ConversionResponse, 
    ConversionError, 
    ConversionStatus,
    ConversionFormat
)
from services.conversion_service import ConversionService

router = APIRouter(prefix="/api/convert", tags=["conversion"])

# In-memory storage for conversion jobs (in production, use database)
conversion_jobs = {}

ALLOWED_FORMATS = {
    ConversionFormat.WORD_TO_PDF: ['doc', 'docx'],
    ConversionFormat.PDF_TO_WORD: ['pdf'],
    ConversionFormat.TXT_TO_PDF: ['txt'],
    ConversionFormat.IMAGE_TO_PDF: ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'tiff'],
    ConversionFormat.EXCEL_TO_PDF: ['xls', 'xlsx', 'csv']
}

async def process_conversion(job: ConversionJob):
    """Background task to process file conversion"""
    start_time = time.time()
    
    try:
        job.status = ConversionStatus.PROCESSING
        conversion_jobs[job.id] = job
        
        success = False
        
        # Perform conversion based on type
        if job.from_format == "word" and job.to_format == "pdf":
            success = await ConversionService.word_to_pdf(job.file_path, job.converted_file_path)
        elif job.from_format == "pdf" and job.to_format == "word":
            success = await ConversionService.pdf_to_word(job.file_path, job.converted_file_path)
        elif job.from_format == "txt" and job.to_format == "pdf":
            success = await ConversionService.text_to_pdf(job.file_path, job.converted_file_path)
        elif job.from_format == "image" and job.to_format == "pdf":
            success = await ConversionService.image_to_pdf([job.file_path], job.converted_file_path)
        elif job.from_format == "excel" and job.to_format == "pdf":
            success = await ConversionService.excel_to_pdf(job.file_path, job.converted_file_path)
        
        # Update job status
        if success and os.path.exists(job.converted_file_path):
            job.status = ConversionStatus.COMPLETED
            job.converted_file_size = os.path.getsize(job.converted_file_path)
            job.conversion_time = round(time.time() - start_time, 2)
            job.download_url = f"/api/download/{job.id}"
        else:
            job.status = ConversionStatus.FAILED
            job.error_message = "Conversion failed"
        
    except Exception as e:
        job.status = ConversionStatus.FAILED
        job.error_message = f"Conversion error: {str(e)}"
    
    finally:
        job.completed_at = time.time()
        conversion_jobs[job.id] = job

@router.post("/word-to-pdf", response_model=ConversionResponse)
async def convert_word_to_pdf(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...)
):
    """Convert Word document to PDF"""
    return await handle_conversion(
        file, background_tasks, 
        ConversionFormat.WORD_TO_PDF,
        "word", "pdf"
    )

@router.post("/pdf-to-word", response_model=ConversionResponse)
async def convert_pdf_to_word(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...)
):
    """Convert PDF to Word document"""
    return await handle_conversion(
        file, background_tasks,
        ConversionFormat.PDF_TO_WORD,
        "pdf", "word"
    )

@router.post("/txt-to-pdf", response_model=ConversionResponse)
async def convert_txt_to_pdf(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...)
):
    """Convert text file to PDF"""
    return await handle_conversion(
        file, background_tasks,
        ConversionFormat.TXT_TO_PDF,
        "txt", "pdf"
    )

@router.post("/image-to-pdf", response_model=ConversionResponse)
async def convert_image_to_pdf(
    background_tasks: BackgroundTasks,
    files: List[UploadFile] = File(...)
):
    """Convert image(s) to PDF"""
    if not files:
        raise HTTPException(status_code=400, detail="No files provided")
    
    # For simplicity, handle single image for now
    # Multiple images would need batch processing
    return await handle_conversion(
        files[0], background_tasks,
        ConversionFormat.IMAGE_TO_PDF,
        "image", "pdf"
    )

@router.post("/excel-to-pdf", response_model=ConversionResponse)
async def convert_excel_to_pdf(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...)
):
    """Convert Excel file to PDF"""
    return await handle_conversion(
        file, background_tasks,
        ConversionFormat.EXCEL_TO_PDF,
        "excel", "pdf"
    )

async def handle_conversion(
    file: UploadFile,
    background_tasks: BackgroundTasks,
    conversion_type: ConversionFormat,
    from_format: str,
    to_format: str
) -> ConversionResponse:
    """Handle file conversion process"""
    
    # Validate file type
    file_ext = Path(file.filename).suffix.lower().lstrip('.')
    allowed_extensions = ALLOWED_FORMATS[conversion_type]
    
    if file_ext not in allowed_extensions:
        raise HTTPException(
            status_code=400,
            detail=f"Invalid file type. Supported formats: {', '.join(allowed_extensions)}"
        )
    
    # Create conversion job
    job = ConversionJob(
        original_filename=file.filename,
        converted_filename=ConversionService.generate_unique_filename(
            file.filename, 
            "docx" if to_format == "word" else to_format
        ),
        from_format=from_format,
        to_format=to_format,
        file_size=0  # Will be updated after saving
    )
    
    try:
        # Save uploaded file
        job.file_path = await ConversionService.save_uploaded_file(file, f"{job.id}_{job.original_filename}")
        job.file_size = os.path.getsize(job.file_path)
        
        # Validate file
        is_valid, validation_message = ConversionService.validate_file(job.file_path, allowed_extensions)
        if not is_valid:
            ConversionService.cleanup_file(job.file_path)
            raise HTTPException(status_code=400, detail=validation_message)
        
        # Set output file path
        conversion_dir = Path("/tmp/conversions")
        job.converted_file_path = str(conversion_dir / job.converted_filename)
        
        # Store job and start background conversion
        conversion_jobs[job.id] = job
        background_tasks.add_task(process_conversion, job)
        
        # Return immediate response with job info
        return ConversionResponse(
            success=True,
            job_id=job.id,
            original_filename=job.original_filename,
            converted_filename=job.converted_filename,
            file_size=job.file_size,
            download_url=f"/api/download/{job.id}",
            message="Conversion started successfully"
        )
        
    except HTTPException:
        raise
    except Exception as e:
        if 'job' in locals() and hasattr(job, 'file_path') and job.file_path:
            ConversionService.cleanup_file(job.file_path)
        raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")

@router.get("/status/{job_id}")
async def get_conversion_status(job_id: str):
    """Get conversion job status"""
    if job_id not in conversion_jobs:
        raise HTTPException(status_code=404, detail="Conversion job not found")
    
    job = conversion_jobs[job_id]
    
    return {
        "job_id": job.id,
        "status": job.status.value,
        "progress": 100 if job.status == ConversionStatus.COMPLETED else 
                   50 if job.status == ConversionStatus.PROCESSING else 0,
        "original_filename": job.original_filename,
        "converted_filename": job.converted_filename,
        "conversion_time": job.conversion_time,
        "error_message": job.error_message,
        "download_url": job.download_url if job.status == ConversionStatus.COMPLETED else None
    }

@router.get("/download/{job_id}")
async def download_converted_file(job_id: str):
    """Download converted file"""
    if job_id not in conversion_jobs:
        raise HTTPException(status_code=404, detail="Conversion job not found")
    
    job = conversion_jobs[job_id]
    
    if job.status != ConversionStatus.COMPLETED:
        raise HTTPException(status_code=400, detail="Conversion not completed yet")
    
    if not os.path.exists(job.converted_file_path):
        raise HTTPException(status_code=404, detail="Converted file not found")
    
    return FileResponse(
        path=job.converted_file_path,
        filename=job.converted_filename,
        media_type='application/octet-stream'
    )

@router.delete("/cleanup/{job_id}")
async def cleanup_conversion_files(job_id: str):
    """Cleanup conversion files"""
    if job_id not in conversion_jobs:
        raise HTTPException(status_code=404, detail="Conversion job not found")
    
    job = conversion_jobs[job_id]
    
    # Cleanup files
    if job.file_path:
        ConversionService.cleanup_file(job.file_path)
    if job.converted_file_path:
        ConversionService.cleanup_file(job.converted_file_path)
    
    # Remove from memory
    del conversion_jobs[job_id]
    
    return {"message": "Files cleaned up successfully"}
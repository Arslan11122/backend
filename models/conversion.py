from pydantic import BaseModel, Field
from typing import Optional, List
from datetime import datetime
from enum import Enum
import uuid

class ConversionStatus(str, Enum):
    PENDING = "pending"
    PROCESSING = "processing"
    COMPLETED = "completed"
    FAILED = "failed"

class ConversionFormat(str, Enum):
    WORD_TO_PDF = "word_to_pdf"
    PDF_TO_WORD = "pdf_to_word"
    TXT_TO_PDF = "txt_to_pdf"
    IMAGE_TO_PDF = "image_to_pdf"
    EXCEL_TO_PDF = "excel_to_pdf"

class ConversionJob(BaseModel):
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    original_filename: str
    converted_filename: str
    from_format: str
    to_format: str
    status: ConversionStatus = ConversionStatus.PENDING
    file_size: int
    converted_file_size: Optional[int] = None
    conversion_time: Optional[float] = None
    created_at: datetime = Field(default_factory=datetime.utcnow)
    completed_at: Optional[datetime] = None
    error_message: Optional[str] = None
    download_url: Optional[str] = None
    file_path: Optional[str] = None
    converted_file_path: Optional[str] = None

class ConversionRequest(BaseModel):
    conversion_type: ConversionFormat

class ConversionResponse(BaseModel):
    success: bool
    job_id: str
    original_filename: str
    converted_filename: str
    file_size: int
    converted_file_size: Optional[int] = None
    conversion_time: Optional[float] = None
    download_url: str
    message: str

class ConversionError(BaseModel):
    success: bool = False
    error: str
    details: Optional[str] = None
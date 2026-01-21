#!/usr/bin/env python3
"""
DOCX to XML Conversion Pipeline REST API

This module provides a FastAPI-based REST API for converting DOCX files
to RittDoc-compliant DocBook XML. The API structure mirrors the PDF
conversion pipeline for consistency.

Usage:
    uvicorn api:app --host 0.0.0.0 --port 8000
"""

from __future__ import annotations

import asyncio
import io
import json
import os
import shutil
import tempfile
import uuid
from concurrent.futures import ThreadPoolExecutor
from dataclasses import asdict, dataclass, field
from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

from fastapi import BackgroundTasks, FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from pydantic import BaseModel, Field

from config import get_config, PipelineConfig
from docx_orchestrator import DocxOrchestrator, ConversionResult


# ============================================================================
# CONFIGURATION
# ============================================================================

class APIConfig:
    """API Configuration settings."""
    UPLOAD_DIR: Path = Path(os.environ.get("DOCXTOXML_UPLOAD_DIR", "./uploads"))
    OUTPUT_DIR: Path = Path(os.environ.get("DOCXTOXML_OUTPUT_DIR", "./output"))
    TEMP_DIR: Path = Path(os.environ.get("DOCXTOXML_TEMP_DIR", tempfile.gettempdir()))
    
    MAX_CONCURRENT_JOBS: int = int(os.environ.get("DOCXTOXML_MAX_CONCURRENT", "5"))
    RESULT_RETENTION_HOURS: int = int(os.environ.get("DOCXTOXML_RETENTION_HOURS", "24"))
    
    @classmethod
    def ensure_directories(cls):
        """Ensure all required directories exist."""
        cls.UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
        cls.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


# ============================================================================
# MODELS
# ============================================================================

class JobStatus(str, Enum):
    """Conversion job status."""
    PENDING = "pending"
    PROCESSING = "processing"
    EXTRACTING = "extracting"
    CONVERTING = "converting"
    PACKAGING = "packaging"
    COMPLETED = "completed"
    FAILED = "failed"
    CANCELLED = "cancelled"


class ConversionOptions(BaseModel):
    """Options for DOCX conversion."""
    extract_images: bool = Field(default=True, description="Extract images from DOCX")
    extract_tables: bool = Field(default=True, description="Extract tables from DOCX")
    create_package: bool = Field(default=True, description="Create RittDoc ZIP package")
    preserve_formatting: bool = Field(default=True, description="Preserve text formatting")


class JobInfo(BaseModel):
    """Information about a conversion job."""
    job_id: str
    status: JobStatus
    progress: float = Field(ge=0, le=100)
    filename: str
    created_at: str
    updated_at: str
    error: Optional[str] = None
    output_files: List[str] = Field(default_factory=list)
    metrics: Dict[str, Any] = Field(default_factory=dict)


class DashboardStats(BaseModel):
    """Dashboard statistics."""
    total_conversions: int = 0
    successful: int = 0
    failed: int = 0
    in_progress: int = 0
    total_images_extracted: int = 0
    total_tables_extracted: int = 0
    average_duration_seconds: float = 0.0
    recent_conversions: List[Dict[str, Any]] = Field(default_factory=list)


# ============================================================================
# JOB MANAGEMENT
# ============================================================================

@dataclass
class ConversionJob:
    """Internal representation of a conversion job."""
    job_id: str
    filename: str
    docx_path: Path
    output_dir: Path
    options: ConversionOptions
    status: JobStatus = JobStatus.PENDING
    progress: float = 0.0
    created_at: datetime = field(default_factory=datetime.now)
    updated_at: datetime = field(default_factory=datetime.now)
    completed_at: Optional[datetime] = None
    error: Optional[str] = None
    output_files: List[str] = field(default_factory=list)
    metrics: Dict[str, Any] = field(default_factory=dict)
    
    def to_info(self) -> JobInfo:
        """Convert to API model."""
        return JobInfo(
            job_id=self.job_id,
            status=self.status,
            progress=self.progress,
            filename=self.filename,
            created_at=self.created_at.isoformat(),
            updated_at=self.updated_at.isoformat(),
            error=self.error,
            output_files=self.output_files,
            metrics=self.metrics,
        )


class JobManager:
    """Manages conversion jobs."""
    
    def __init__(self):
        self.jobs: Dict[str, ConversionJob] = {}
        self.executor = ThreadPoolExecutor(max_workers=APIConfig.MAX_CONCURRENT_JOBS)
    
    def create_job(
        self,
        filename: str,
        docx_path: Path,
        output_dir: Path,
        options: ConversionOptions
    ) -> ConversionJob:
        """Create a new conversion job."""
        job_id = str(uuid.uuid4())[:8]
        
        job = ConversionJob(
            job_id=job_id,
            filename=filename,
            docx_path=docx_path,
            output_dir=output_dir,
            options=options,
        )
        self.jobs[job_id] = job
        return job
    
    def get_job(self, job_id: str) -> Optional[ConversionJob]:
        """Get a job by ID."""
        return self.jobs.get(job_id)
    
    def update_job(
        self,
        job_id: str,
        status: Optional[JobStatus] = None,
        progress: Optional[float] = None,
        error: Optional[str] = None,
        output_files: Optional[List[str]] = None,
        metrics: Optional[Dict[str, Any]] = None,
    ):
        """Update job status."""
        job = self.jobs.get(job_id)
        if job:
            if status:
                job.status = status
            if progress is not None:
                job.progress = progress
            if error:
                job.error = error
            if output_files:
                job.output_files = output_files
            if metrics:
                job.metrics.update(metrics)
            job.updated_at = datetime.now()
            if status in (JobStatus.COMPLETED, JobStatus.FAILED, JobStatus.CANCELLED):
                job.completed_at = datetime.now()
    
    def list_jobs(self, status: Optional[JobStatus] = None, limit: int = 50) -> List[ConversionJob]:
        """List jobs, optionally filtered by status."""
        jobs = list(self.jobs.values())
        if status:
            jobs = [j for j in jobs if j.status == status]
        jobs.sort(key=lambda j: j.created_at, reverse=True)
        return jobs[:limit]
    
    def get_dashboard_stats(self) -> DashboardStats:
        """Get dashboard statistics."""
        jobs = list(self.jobs.values())
        
        completed = [j for j in jobs if j.status == JobStatus.COMPLETED]
        failed = [j for j in jobs if j.status == JobStatus.FAILED]
        in_progress = [j for j in jobs if j.status in (
            JobStatus.PROCESSING, JobStatus.EXTRACTING,
            JobStatus.CONVERTING, JobStatus.PACKAGING
        )]
        
        # Calculate metrics
        total_images = sum(j.metrics.get("images", 0) for j in jobs)
        total_tables = sum(j.metrics.get("tables", 0) for j in jobs)
        
        durations = []
        for j in completed:
            if j.completed_at and j.created_at:
                durations.append((j.completed_at - j.created_at).total_seconds())
        
        avg_duration = sum(durations) / len(durations) if durations else 0.0
        
        # Recent conversions
        recent = [
            {
                "job_id": j.job_id,
                "filename": j.filename,
                "status": j.status.value,
                "created_at": j.created_at.isoformat(),
                "duration": (j.completed_at - j.created_at).total_seconds() if j.completed_at else None,
            }
            for j in sorted(jobs, key=lambda x: x.created_at, reverse=True)[:10]
        ]
        
        return DashboardStats(
            total_conversions=len(jobs),
            successful=len(completed),
            failed=len(failed),
            in_progress=len(in_progress),
            total_images_extracted=total_images,
            total_tables_extracted=total_tables,
            average_duration_seconds=avg_duration,
            recent_conversions=recent,
        )


# Global job manager
job_manager = JobManager()


# ============================================================================
# CONVERSION WORKER
# ============================================================================

def run_conversion(job: ConversionJob):
    """Run the DOCX to XML conversion."""
    try:
        job_manager.update_job(job.job_id, status=JobStatus.PROCESSING, progress=10)
        
        # Configure extraction
        config = get_config()
        config.extraction.extract_images = job.options.extract_images
        config.extraction.extract_tables = job.options.extract_tables
        config.extraction.preserve_formatting = job.options.preserve_formatting
        
        # Create orchestrator
        orchestrator = DocxOrchestrator(config=config, verbose=False)
        
        # Update status
        job_manager.update_job(job.job_id, status=JobStatus.EXTRACTING, progress=30)
        
        # Run conversion
        result = orchestrator.convert(
            docx_path=job.docx_path,
            output_dir=job.output_dir,
            create_package=job.options.create_package
        )
        
        job_manager.update_job(job.job_id, status=JobStatus.CONVERTING, progress=70)
        
        if result.success:
            # Collect output files
            output_files = []
            if result.xml_path:
                output_files.append(Path(result.xml_path).name)
            if result.package_path:
                output_files.append(Path(result.package_path).name)
            
            # Update job with results
            job_manager.update_job(
                job.job_id,
                status=JobStatus.COMPLETED,
                progress=100,
                output_files=output_files,
                metrics={
                    "text_blocks": result.text_blocks,
                    "images": result.images,
                    "tables": result.tables,
                    "chapters": result.chapters,
                    "duration_seconds": result.duration_seconds,
                }
            )
        else:
            error_msg = "; ".join(result.errors) if result.errors else "Unknown error"
            job_manager.update_job(
                job.job_id,
                status=JobStatus.FAILED,
                error=error_msg
            )
    
    except Exception as e:
        job_manager.update_job(
            job.job_id,
            status=JobStatus.FAILED,
            error=str(e)
        )


# ============================================================================
# API APPLICATION
# ============================================================================

def create_app() -> FastAPI:
    """Create and configure the FastAPI application."""
    
    app = FastAPI(
        title="DOCX to XML Conversion API",
        description="""
REST API for converting DOCX documents to RittDoc DTD-compliant DocBook XML.

## Features

- Fast text extraction (no AI required for basic content)
- Image extraction with metadata
- Table extraction with structure preservation
- RittDoc ZIP package generation
- Compatible with PDF pipeline output format

## Workflow

1. **Upload & Convert**: `POST /api/v1/convert` - Upload DOCX, returns job_id
2. **Poll Status**: `GET /api/v1/jobs/{job_id}` - Wait for `completed` status
3. **Download Files**: `GET /api/v1/jobs/{job_id}/files/{filename}` - Get outputs
        """,
        version="1.0.0",
        docs_url="/docs",
        redoc_url="/redoc",
    )
    
    # CORS middleware
    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )
    
    @app.on_event("startup")
    async def startup_event():
        APIConfig.ensure_directories()
    
    # ========================================================================
    # CONVERSION ENDPOINTS
    # ========================================================================
    
    @app.post("/api/v1/convert", response_model=JobInfo, tags=["Conversion"])
    async def start_conversion(
        background_tasks: BackgroundTasks,
        file: UploadFile = File(..., description="DOCX file to convert"),
        extract_images: bool = Form(default=True),
        extract_tables: bool = Form(default=True),
        create_package: bool = Form(default=True),
        preserve_formatting: bool = Form(default=True),
    ):
        """
        Upload a DOCX file and start conversion.
        
        The conversion runs in the background. Poll the job status
        until it reaches `completed`.
        """
        # Validate file
        if not file.filename:
            raise HTTPException(status_code=400, detail="No filename provided")
        
        if not file.filename.lower().endswith(('.docx', '.doc')):
            raise HTTPException(status_code=400, detail="File must be a DOCX document")
        
        # Create job directory
        job_id = str(uuid.uuid4())[:8]
        job_dir = APIConfig.OUTPUT_DIR / job_id
        job_dir.mkdir(parents=True, exist_ok=True)
        
        # Save uploaded file
        docx_path = job_dir / file.filename
        content = await file.read()
        docx_path.write_bytes(content)
        
        # Create options
        options = ConversionOptions(
            extract_images=extract_images,
            extract_tables=extract_tables,
            create_package=create_package,
            preserve_formatting=preserve_formatting,
        )
        
        # Create job
        job = job_manager.create_job(
            filename=file.filename,
            docx_path=docx_path,
            output_dir=job_dir,
            options=options,
        )
        
        # Start conversion in background
        background_tasks.add_task(run_conversion, job)
        
        return job.to_info()
    
    @app.get("/api/v1/jobs/{job_id}", response_model=JobInfo, tags=["Conversion"])
    async def get_job_status(job_id: str):
        """Get the status of a conversion job."""
        job = job_manager.get_job(job_id)
        if not job:
            raise HTTPException(status_code=404, detail="Job not found")
        return job.to_info()
    
    @app.get("/api/v1/jobs", response_model=List[JobInfo], tags=["Conversion"])
    async def list_jobs(status: Optional[str] = None, limit: int = 50):
        """List all conversion jobs."""
        status_filter = None
        if status:
            try:
                status_filter = JobStatus(status)
            except ValueError:
                raise HTTPException(status_code=400, detail=f"Invalid status: {status}")
        
        jobs = job_manager.list_jobs(status=status_filter, limit=limit)
        return [j.to_info() for j in jobs]
    
    @app.delete("/api/v1/jobs/{job_id}", tags=["Conversion"])
    async def cancel_job(job_id: str):
        """Cancel a pending or in-progress job."""
        job = job_manager.get_job(job_id)
        if not job:
            raise HTTPException(status_code=404, detail="Job not found")
        
        if job.status in (JobStatus.COMPLETED, JobStatus.FAILED, JobStatus.CANCELLED):
            raise HTTPException(status_code=400, detail="Job already finished")
        
        job_manager.update_job(job_id, status=JobStatus.CANCELLED)
        return {"message": "Job cancelled"}
    
    # ========================================================================
    # FILE ENDPOINTS
    # ========================================================================
    
    @app.get("/api/v1/jobs/{job_id}/files", tags=["Files"])
    async def list_output_files(job_id: str):
        """List output files for a completed job."""
        job = job_manager.get_job(job_id)
        if not job:
            raise HTTPException(status_code=404, detail="Job not found")
        
        if job.status not in (JobStatus.COMPLETED,):
            raise HTTPException(status_code=400, detail="Job not completed yet")
        
        files = []
        if job.output_dir.exists():
            for f in job.output_dir.iterdir():
                if f.is_file() and f.name != job.filename:
                    files.append({
                        "name": f.name,
                        "size": f.stat().st_size,
                        "download_url": f"/api/v1/jobs/{job_id}/files/{f.name}",
                    })
        
        return {"files": files}
    
    @app.get("/api/v1/jobs/{job_id}/files/{filename}", tags=["Files"])
    async def download_file(job_id: str, filename: str):
        """Download an output file."""
        job = job_manager.get_job(job_id)
        if not job:
            raise HTTPException(status_code=404, detail="Job not found")
        
        file_path = job.output_dir / filename
        if not file_path.exists() or not file_path.is_file():
            raise HTTPException(status_code=404, detail="File not found")
        
        return FileResponse(
            path=file_path,
            filename=filename,
            media_type="application/octet-stream",
        )
    
    # ========================================================================
    # DASHBOARD ENDPOINTS
    # ========================================================================
    
    @app.get("/api/v1/dashboard", response_model=DashboardStats, tags=["Dashboard"])
    async def get_dashboard():
        """Get dashboard statistics."""
        return job_manager.get_dashboard_stats()
    
    @app.get("/api/v1/dashboard/export", tags=["Dashboard"])
    async def export_dashboard():
        """Export dashboard data as JSON."""
        stats = job_manager.get_dashboard_stats()
        jobs = job_manager.list_jobs(limit=1000)
        
        export_data = {
            "exported_at": datetime.now().isoformat(),
            "statistics": stats.dict(),
            "jobs": [j.to_info().dict() for j in jobs],
        }
        
        return JSONResponse(content=export_data)
    
    # ========================================================================
    # HEALTH & INFO ENDPOINTS
    # ========================================================================
    
    @app.get("/api/v1/health", tags=["System"])
    async def health_check():
        """Health check endpoint."""
        return {
            "status": "healthy",
            "timestamp": datetime.now().isoformat(),
            "service": "docx-to-xml-pipeline",
        }
    
    @app.get("/api/v1/info", tags=["System"])
    async def get_info():
        """Get API configuration and capabilities."""
        return {
            "version": "1.0.0",
            "service": "DOCX to XML Conversion Pipeline",
            "config": {
                "max_concurrent_jobs": APIConfig.MAX_CONCURRENT_JOBS,
            },
            "capabilities": {
                "docx_parsing": True,
                "image_extraction": True,
                "table_extraction": True,
                "rittdoc_packaging": True,
                "ai_enhancement": False,  # Optional, not required for DOCX
            },
        }
    
    return app


# Create default app instance
app = create_app()


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)

"""
FastAPI backend for PowerPoint to Video converter.
Provides REST API endpoints for the conversion service.
"""

import os
import uuid
import asyncio
import json
from datetime import datetime
from typing import Dict, List, Optional
from pathlib import Path

from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel

# Import our existing conversion logic
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from auto_presenter import (
    configure_gemini_vision_model,
    extract_slides_as_images_linux,
    generate_script_for_slide,
    synthesize_speech_with_coqui,
    create_video_with_moviepy,
    save_script_to_file,
    load_script_from_file,
    should_regenerate_audio
)

# Load environment variables
from dotenv import load_dotenv
load_dotenv()

app = FastAPI(
    title="PowerPoint to Video API",
    description="Convert PowerPoint presentations to narrated videos using AI",
    version="1.0.0"
)

# Add CORS middleware for React frontend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:3000", "http://localhost:5173"],  # React dev servers
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Global variables for services
vision_model = None
tts_engine = None

# Job storage (in production, use a proper database)
jobs: Dict[str, Dict] = {}

# Data models
class JobStatus(BaseModel):
    job_id: str
    status: str  # "pending", "processing", "completed", "failed"
    progress: int  # 0-100
    message: str
    created_at: datetime
    slides_total: Optional[int] = None
    slides_processed: Optional[int] = None
    video_url: Optional[str] = None

class ScriptUpdate(BaseModel):
    scripts: Dict[int, str]  # slide_number -> script_text

class SlideScript(BaseModel):
    slide_number: int
    script: str
    image_url: str

# Initialize AI services on startup
@app.on_event("startup")
async def startup_event():  # noqa: deprecation - lifespan events preferred for newer FastAPI
    global vision_model, tts_engine
    
    print("Initializing AI services...")
    
    # Initialize Gemini Vision Model
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        print("Warning: GEMINI_API_KEY not found. Some features may not work.")
    else:
        try:
            vision_model = configure_gemini_vision_model(api_key)
            print("✓ Gemini Vision Model initialized")
        except Exception as e:
            print(f"Failed to initialize Gemini: {e}")
    
    # Initialize TTS Engine
    try:
        from TTS.api import TTS
        tts_engine = TTS("tts_models/en/ljspeech/vits")
        print("✓ Coqui TTS Engine initialized")
    except ImportError:
        print("⚠️  TTS library not available - install TTS for audio generation")
        tts_engine = None
    except Exception as e:
        print(f"Failed to initialize TTS: {e}")
        tts_engine = None

@app.get("/")
async def root():
    return {
        "message": "PowerPoint to Video API",
        "version": "1.0.0",
        "status": "running"
    }

@app.get("/health")
async def health_check():
    return {
        "status": "healthy",
        "gemini_available": vision_model is not None,
        "tts_available": tts_engine is not None
    }

@app.post("/upload", response_model=JobStatus)
async def upload_presentation(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...)
):
    """Upload a PowerPoint presentation and start conversion."""
    
    # Validate file type
    if not file.filename.lower().endswith('.pptx'):
        raise HTTPException(
            status_code=400,
            detail="Only PowerPoint (.pptx) files are supported"
        )
    
    # Create job ID and directory
    job_id = str(uuid.uuid4())
    job_dir = Path(f"uploads/{job_id}")
    job_dir.mkdir(parents=True, exist_ok=True)
    
    # Save uploaded file
    file_path = job_dir / file.filename
    try:
        with open(file_path, "wb") as buffer:
            content = await file.read()
            buffer.write(content)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to save file: {e}")
    
    # Create job record
    job = {
        "job_id": job_id,
        "status": "pending",
        "progress": 0,
        "message": "File uploaded, starting conversion...",
        "created_at": datetime.now(),
        "filename": file.filename,
        "file_path": str(file_path),
        "slides_total": None,
        "slides_processed": 0,
        "video_url": None
    }
    
    jobs[job_id] = job
    
    # Start background conversion
    background_tasks.add_task(process_presentation, job_id)
    
    return JobStatus(**job)

@app.get("/status/{job_id}", response_model=JobStatus)
async def get_job_status(job_id: str):
    """Get the current status of a conversion job."""
    
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found")
    
    job = jobs[job_id]
    return JobStatus(**job)

@app.get("/jobs", response_model=List[JobStatus])
async def list_jobs():
    """List all conversion jobs."""
    return [JobStatus(**job) for job in jobs.values()]

@app.get("/scripts/{job_id}")
async def get_scripts(job_id: str):
    """Get the generated scripts for all slides."""
    
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found")
    
    job = jobs[job_id]
    if job["status"] not in ["processing", "completed"]:
        raise HTTPException(status_code=400, detail="Scripts not yet available")
    
    # Find the temp directory
    file_path = Path(job["file_path"])
    base_name = file_path.stem
    temp_dir = file_path.parent / f"{base_name}_temp_files"
    
    scripts = []
    slide_num = 1
    
    while True:
        script_path = temp_dir / f"script_{slide_num}.txt"
        image_path = temp_dir / f"slide_{slide_num}.png"
        
        if not script_path.exists():
            break
            
        script_text = load_script_from_file(str(script_path)) or ""
        
        scripts.append(SlideScript(
            slide_number=slide_num,
            script=script_text,
            image_url=f"/slides/{job_id}/{slide_num}"
        ))
        
        slide_num += 1
    
    return scripts

@app.put("/scripts/{job_id}")
async def update_scripts(
    job_id: str,
    script_update: ScriptUpdate,
    background_tasks: BackgroundTasks
):
    """Update scripts and regenerate affected audio/video."""
    
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found")
    
    job = jobs[job_id]
    if job["status"] not in ["completed"]:
        raise HTTPException(status_code=400, detail="Job must be completed before editing scripts")
    
    # Find the temp directory
    file_path = Path(job["file_path"])
    base_name = file_path.stem
    temp_dir = file_path.parent / f"{base_name}_temp_files"
    
    # Update script files
    updated_scripts = []
    for slide_num, script_text in script_update.scripts.items():
        script_path = temp_dir / f"script_{slide_num}.txt"
        if save_script_to_file(script_text, str(script_path), slide_num):
            updated_scripts.append(slide_num)
    
    if updated_scripts:
        # Update job status
        job["status"] = "processing"
        job["message"] = f"Regenerating audio for {len(updated_scripts)} updated scripts..."
        job["progress"] = 0
        
        # Start background regeneration
        background_tasks.add_task(regenerate_audio_and_video, job_id, updated_scripts)
        
        return {"message": f"Started regeneration for slides: {updated_scripts}"}
    else:
        return {"message": "No scripts were updated"}

@app.get("/download/{job_id}")
async def download_video(job_id: str):
    """Download the generated video."""
    
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found")
    
    job = jobs[job_id]
    if job["status"] != "completed":
        raise HTTPException(status_code=400, detail="Video not yet ready")
    
    # Find the video file
    file_path = Path(job["file_path"])
    base_name = file_path.stem
    video_path = file_path.parent / f"{base_name}_presentation.mp4"
    
    if not video_path.exists():
        raise HTTPException(status_code=404, detail="Video file not found")
    
    return FileResponse(
        path=str(video_path),
        filename=f"{base_name}_presentation.mp4",
        media_type="video/mp4"
    )

@app.get("/slides/{job_id}/{slide_num}")
async def get_slide_image(job_id: str, slide_num: int):
    """Get slide image for preview."""
    
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found")
    
    job = jobs[job_id]
    file_path = Path(job["file_path"])
    base_name = file_path.stem
    temp_dir = file_path.parent / f"{base_name}_temp_files"
    image_path = temp_dir / f"slide_{slide_num}.png"
    
    if not image_path.exists():
        raise HTTPException(status_code=404, detail="Slide image not found")
    
    return FileResponse(path=str(image_path), media_type="image/png")

# Background task functions
async def process_presentation(job_id: str):
    """Background task to process the presentation."""
    
    try:
        job = jobs[job_id]
        file_path = job["file_path"]
        
        # Update status
        job["status"] = "processing"
        job["message"] = "Extracting slides..."
        job["progress"] = 10
        
        # Extract slides
        base_name = Path(file_path).stem
        temp_dir = Path(file_path).parent / f"{base_name}_temp_files"
        
        slide_images = extract_slides_as_images_linux(file_path, str(temp_dir))
        if not slide_images:
            job["status"] = "failed"
            job["message"] = "Failed to extract slides"
            return
        
        job["slides_total"] = len(slide_images)
        job["message"] = f"Processing {len(slide_images)} slides..."
        job["progress"] = 20
        
        # Generate scripts and audio
        audio_files = []
        for i, img_path in enumerate(slide_images):
            slide_num = i + 1
            
            # Update progress
            progress = 20 + (60 * i // len(slide_images))
            job["progress"] = progress
            job["message"] = f"Processing slide {slide_num} of {len(slide_images)}..."
            job["slides_processed"] = slide_num
            
            script_path = temp_dir / f"script_{slide_num}.txt"
            audio_path = temp_dir / f"audio_{slide_num}.wav"
            
            # Generate script if not exists
            script = None
            if script_path.exists():
                script = load_script_from_file(str(script_path))
            
            if not script and vision_model:
                script = generate_script_for_slide(
                    vision_model, img_path, slide_num, len(slide_images)
                )
                if script:
                    save_script_to_file(script, str(script_path), slide_num)
            
            # Generate audio
            if script and tts_engine:
                if should_regenerate_audio(str(script_path), str(audio_path)):
                    audio_file = synthesize_speech_with_coqui(
                        tts_engine, script, str(audio_path), slide_num
                    )
                    audio_files.append(audio_file)
                else:
                    audio_files.append(str(audio_path) if audio_path.exists() else None)
            else:
                audio_files.append(None)
        
        # Create video
        job["message"] = "Creating video..."
        job["progress"] = 90
        
        video_path = Path(file_path).parent / f"{base_name}_presentation.mp4"
        create_video_with_moviepy(slide_images, audio_files, str(video_path))
        
        # Complete
        job["status"] = "completed"
        job["message"] = "Video creation completed successfully!"
        job["progress"] = 100
        job["video_url"] = f"/download/{job_id}"
        
    except Exception as e:
        job = jobs[job_id]
        job["status"] = "failed"
        job["message"] = f"Error: {str(e)}"
        print(f"Error processing job {job_id}: {e}")

async def regenerate_audio_and_video(job_id: str, updated_slides: List[int]):
    """Regenerate audio and video for updated scripts."""
    
    try:
        job = jobs[job_id]
        file_path = job["file_path"]
        
        base_name = Path(file_path).stem
        temp_dir = Path(file_path).parent / f"{base_name}_temp_files"
        
        # Find all slides
        slide_images = []
        audio_files = []
        slide_num = 1
        
        while True:
            image_path = temp_dir / f"slide_{slide_num}.png"
            if not image_path.exists():
                break
            slide_images.append(str(image_path))
            slide_num += 1
        
        total_slides = len(slide_images)
        
        # Regenerate audio for updated slides
        for i in range(total_slides):
            slide_num = i + 1
            
            # Update progress
            progress = 10 + (70 * i // total_slides)
            job["progress"] = progress
            job["message"] = f"Checking slide {slide_num} of {total_slides}..."
            
            script_path = temp_dir / f"script_{slide_num}.txt"
            audio_path = temp_dir / f"audio_{slide_num}.wav"
            
            if slide_num in updated_slides:
                # Regenerate this slide's audio
                script = load_script_from_file(str(script_path))
                if script and tts_engine:
                    audio_file = synthesize_speech_with_coqui(
                        tts_engine, script, str(audio_path), slide_num
                    )
                    audio_files.append(audio_file)
                else:
                    audio_files.append(None)
            else:
                # Use existing audio
                audio_files.append(str(audio_path) if audio_path.exists() else None)
        
        # Recreate video
        job["message"] = "Recreating video with updated audio..."
        job["progress"] = 90
        
        video_path = Path(file_path).parent / f"{base_name}_presentation.mp4"
        create_video_with_moviepy(slide_images, audio_files, str(video_path))
        
        # Complete
        job["status"] = "completed"
        job["message"] = "Video regenerated successfully!"
        job["progress"] = 100
        
    except Exception as e:
        job = jobs[job_id]
        job["status"] = "failed"
        job["message"] = f"Error during regeneration: {str(e)}"
        print(f"Error regenerating job {job_id}: {e}")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000, reload=True)
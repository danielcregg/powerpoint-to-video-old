# PowerPoint to Video Converter (Legacy)

![Python](https://img.shields.io/badge/Python-3776AB?style=flat-square&logo=python&logoColor=white)
![React](https://img.shields.io/badge/React-61DAFB?style=flat-square&logo=react&logoColor=black)
![TypeScript](https://img.shields.io/badge/TypeScript-3178C6?style=flat-square&logo=typescript&logoColor=white)
![FastAPI](https://img.shields.io/badge/FastAPI-009688?style=flat-square&logo=fastapi&logoColor=white)
![FFmpeg](https://img.shields.io/badge/FFmpeg-007808?style=flat-square&logo=ffmpeg&logoColor=white)
![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=flat-square)

A full-stack web application that converts PowerPoint presentations into narrated videos using AI-powered script generation and text-to-speech synthesis.

## Overview

This application accepts a `.pptx` file, uses Google Gemini to generate a natural presenter script for each slide, synthesizes speech with Coqui TTS, and assembles the final narrated video using MoviePy and FFmpeg. It includes both a modern React web interface and a standalone command-line tool.

## Features

- **AI-powered script generation** using Google Gemini vision models
- **High-quality text-to-speech** synthesis via Coqui TTS (offline, no API fees)
- **Drag-and-drop web interface** built with React, TypeScript, and Tailwind CSS
- **Real-time progress tracking** with detailed status updates
- **Script editing and regeneration** without full reprocessing
- **Multi-job management** with persistent state
- **Cross-platform slide conversion** using LibreOffice headless mode
- **Multiple video codec fallbacks** (H.264, MP4V) for broad compatibility
- **CLI support** via the standalone `auto_presenter.py` script

## Prerequisites

- Python 3.11 or higher
- Node.js 18 or higher
- LibreOffice (headless mode for PPTX-to-PDF conversion)
- FFmpeg (video encoding)
- A Google Gemini API key

## Getting Started

### Installation

**GitHub Codespaces (Recommended):**

1. Create a new Codespace from this repository. The devcontainer will automatically install all dependencies.
2. Set up your Gemini API key:
   ```bash
   echo "GEMINI_API_KEY=your_api_key_here" > .env
   ```
3. Run the startup script:
   ```bash
   bash start-dev.sh
   ```

**Local Development:**

1. Clone the repository:
   ```bash
   git clone https://github.com/danielcregg/powerpoint-to-video-old.git
   cd powerpoint-to-video-old
   ```

2. Install backend dependencies:
   ```bash
   cd backend
   pip install -r requirements.txt
   echo "GEMINI_API_KEY=your_api_key_here" > ../.env
   ```

3. Install frontend dependencies:
   ```bash
   cd ../frontend
   npm install
   ```

### Usage

**Web Interface:**

1. Start the backend:
   ```bash
   cd backend && python app.py
   ```
2. Start the frontend (in a new terminal):
   ```bash
   cd frontend && npm run dev
   ```
3. Open `http://localhost:3000` in your browser and upload a `.pptx` file.

**Command-Line Interface:**

```bash
python auto_presenter.py presentation.pptx
```

**API Endpoints:**

| Method | Endpoint | Description |
|--------|----------|-------------|
| `POST` | `/upload` | Upload a PowerPoint file and start conversion |
| `GET` | `/status/{job_id}` | Get conversion progress and status |
| `GET` | `/download/{job_id}` | Download completed video |
| `GET` | `/scripts/{job_id}` | Get generated scripts for editing |
| `PUT` | `/scripts/{job_id}` | Update scripts and regenerate audio |
| `GET` | `/jobs` | List all conversion jobs |
| `GET` | `/health` | Check service availability |

## Tech Stack

- **Python** -- Backend logic and AI orchestration
- **FastAPI** -- REST API framework with async support
- **React 18** -- Frontend UI with TypeScript
- **Tailwind CSS** -- Utility-first styling
- **Google Gemini** -- AI vision model for slide script generation
- **Coqui TTS** -- Offline text-to-speech synthesis
- **MoviePy** -- Video assembly from images and audio
- **PyMuPDF** -- PDF-to-image extraction
- **LibreOffice** -- Headless PPTX-to-PDF conversion
- **FFmpeg** -- Video encoding and processing
- **Vite** -- Frontend build tooling

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

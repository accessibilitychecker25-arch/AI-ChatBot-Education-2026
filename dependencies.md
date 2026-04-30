# Python Backend Deployment Guide

## Installation

To install all required dependencies:

```bash
cd Accessibility-Checker-BE/python-server
pip install -r requirements.txt
```

## Running Locally

```bash
cd Accessibility-Checker-BE/python-server
uvicorn server2:app --reload
```

The API will be available at `http://localhost:8000`

## Deployment

For Hugging Face Spaces or other Python hosting:

1. All dependencies are listed in `Accessibility-Checker-BE/python-server/requirements.txt`
2. The main application is in `Accessibility-Checker-BE/python-server/server2.py`
3. Models will auto-download on first use (~2GB for BLIP)

## Dependencies

See `Accessibility-Checker-BE/python-server/requirements.txt` for the complete list of required packages.

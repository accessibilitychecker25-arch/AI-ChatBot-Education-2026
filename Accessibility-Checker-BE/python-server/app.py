#!/usr/bin/env python3
"""
Entry point for Hugging Face Spaces deployment.
This file launches the FastAPI application from server2.py
"""

from server2 import app

# The app variable is automatically detected by HF Spaces
# HF Spaces will run: uvicorn app:app --host 0.0.0.0 --port 7860

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=7860)

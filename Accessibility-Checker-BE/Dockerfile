FROM python:3.11-slim

WORKDIR /app

# Copy requirements and install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy the entire backend
COPY python-server/ ./python-server/

# Set working directory to python-server
WORKDIR /app/python-server

# Expose port (HF Spaces uses 7860)
EXPOSE 7860

# Start the app
CMD ["uvicorn", "server2:app", "--host", "0.0.0.0", "--port", "7860"]

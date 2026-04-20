FROM python:3.11-slim

# System deps
RUN apt-get update && apt-get install -y --no-install-recommends \
        gcc \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python deps first (cached layer)
COPY src/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy source
COPY src/ ./src/

# /data is the mount point for input/ and output/ folders
VOLUME ["/data"]

# Run pipeline from src directory so relative imports work
WORKDIR /app/src

CMD ["python", "pipeline.py"]

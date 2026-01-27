FROM python:3.11-slim-bookworm

# Instalar Tesseract (español + inglés) y dependencias
RUN apt-get update && apt-get install -y --no-install-recommends \
    tesseract-ocr tesseract-ocr-spa tesseract-ocr-eng \
    poppler-utils libgl1 libglib2.0-0 libsm6 libxext6 libxrender1 \
&& rm -rf /var/lib/apt/lists/*

# Instalar Python dependencies
RUN pip install --no-cache-dir \
    Flask==3.0.0 \
    PyMuPDF==1.24.4 \
    pytesseract==0.3.13 \
    Pillow==10.2.0 \
    numpy==1.26.4 \
    scikit-learn==1.4.0 \
    mistralai>=0.1.2 \
    openpyxl==3.1.2 \
    pandas==2.2.1 \
    opencv-python-headless==4.9.0.80\
    python-dotenv==1.0.0 \
    toml==0.10.2

WORKDIR /app

EXPOSE 8501
CMD ["python", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]

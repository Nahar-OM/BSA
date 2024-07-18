# Use an official Python runtime as a parent image
FROM python:3.11.8

# Set the working directory in the container
WORKDIR /app

# Copy the current directory contents into the container at /app
COPY . /app

# Install system dependencies
RUN apt-get update && apt-get install -y \
    build-essential \
    libtesseract-dev \
    libleptonica-dev \
    tesseract-ocr \
    poppler-utils \
    wget \
    unzip \
    clang \
    llvm \
    g++ \
    gdb \
    make \
    cmake \
    && apt-get clean

# Install any needed Python packages specified in requirements.txt
RUN pip install --no-cache-dir -r requirements1.txt
RUN pip install --no-cache-dir -r requirements2.txt

# Set environment variable for TensorFlow optimization
ENV TF_ENABLE_ONEDNN_OPTS=0

# Set Tesseract path environment variable
ENV TESSERACT_PATH="/usr/bin/tesseract"

# Make port 80 available to the world outside this container
EXPOSE 80

# Run main.py when the container launches
#CMD ["python", "app.py"]

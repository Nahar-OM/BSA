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
    libgl1-mesa-glx \
    libglib2.0-0 \
    libhdf5-dev \
    && apt-get clean

# Upgrade pip
RUN pip install --upgrade pip

# Installing h5py separately
RUN pip install --no-cache-dir h5py

# Installing any needed Python packages specified in requirements.txt
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt || \
    (pip install --no-cache-dir $(grep -v h5py requirements.txt) && \
     pip install --no-cache-dir h5py)

# Downloading and installing spacy model
RUN wget https://github.com/explosion/spacy-models/releases/download/en_core_web_sm-3.7.1/en_core_web_sm-3.7.1-py3-none-any.whl \
    && pip install en_core_web_sm-3.7.1-py3-none-any.whl \
    && rm en_core_web_sm-3.7.1-py3-none-any.whl

# Set environment variable for TensorFlow optimization
ENV TF_ENABLE_ONEDNN_OPTS=0

# Set Tesseract path environment variable
ENV TESSERACT_PATH="/usr/bin/tesseract"

# Install Bun
RUN curl -fsSL https://bun.sh/install | bash
ENV PATH="/root/.bun/bin:${PATH}"

# Install TypeScript
RUN bun add -g typescript

# Make port 3000 available to the world outside this container
EXPOSE 3000

# Copy the TypeScript server file
COPY server.ts .

# Run the Bun server when the container launches
CMD ["/root/.bun/bin/bun", "run", "server.ts"]

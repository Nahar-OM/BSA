FROM python:3.11.8

WORKDIR /app

COPY . /app

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

RUN pip install --no-cache-dir -r requirements1.txt
RUN pip install --no-cache-dir -r requirements2.txt

ENV TF_ENABLE_ONEDNN_OPTS=0

# Add poppler-utils bin directory to PATH
ENV PATH="/usr/bin:${PATH}"

ENV TESSERACT_PATH="/usr/bin/tesseract"

# Install Bun (A environment to run js/ts backend)
RUN curl -fsSL https://bun.sh/install | bash

# Adding Bun to PATH
ENV PATH="/root/.bun/bin:${PATH}"

# Installing TypeScript
RUN bun add -g typescript

# Make port 3000 available to outside this container
EXPOSE 3000

# Copy the TypeScript server file
COPY server.ts .

CMD bun run server.ts

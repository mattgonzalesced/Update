# Local OCR with PaddleOCR

This repository includes a helper script to run OCR locally using [PaddleOCR](https://github.com/PaddlePaddle/PaddleOCR). The model files are fetched automatically and stored in your cache, so an internet connection is required the first time you run the script.

## Setup

1. Install Python 3.8 or later.
2. Install dependencies:
   ```bash
   pip install paddleocr
   ```
   Optional: install `huggingface_hub` if you plan to download models via HF APIs.

## Usage

```bash
python local_ocr.py [path/to/image.jpg]
```

If no path is provided, the script opens a Tkinter file dialog so you can choose an image. The recognized text is printed with confidence scores and the bounding boxes as `(xmin, ymin)` and `(xmax, ymax)` coordinates.

## Example output

```
Text: Hello, Score: 0.99, Box: (100, 120), (200, 150)
```

This demonstrates how each piece of text is returned with its coordinates, allowing you to locate it on the original page.

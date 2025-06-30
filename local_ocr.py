#!/usr/bin/env python3
"""Run OCR locally using PaddleOCR.
Requires: pip install paddleocr
Usage: python local_ocr.py [image_path]
Outputs bounding boxes with text. If no path is given, a file dialog will open
to choose an image.
"""
import sys
import tkinter as tk
from tkinter import filedialog
from paddleocr import PaddleOCR

if len(sys.argv) >= 2:
    image_path = sys.argv[1]
else:
    root = tk.Tk()
    root.withdraw()
    image_path = filedialog.askopenfilename(
        title="Select image file",
        filetypes=[
            ("Image Files", "*.jpg *.jpeg *.png *.bmp *.tiff"),
            ("All Files", "*.*"),
        ],
    )
    if not image_path:
        print("No file selected. Exiting.")
        sys.exit(0)
ocr = PaddleOCR(use_angle_cls=True, lang='en')

result = ocr.ocr(image_path, cls=True)

for line in result:
    for box, (text, score) in line:
        # box is [[x1,y1],[x2,y2],[x3,y3],[x4,y4]]
        xs = [int(b[0]) for b in box]
        ys = [int(b[1]) for b in box]
        xmin, xmax = min(xs), max(xs)
        ymin, ymax = min(ys), max(ys)
        print(f'Text: {text}, Score: {score:.2f}, Box: ({xmin}, {ymin}), ({xmax}, {ymax})')

#!/usr/bin/env python3
"""Run OCR locally using PaddleOCR.
Requires: pip install paddleocr

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

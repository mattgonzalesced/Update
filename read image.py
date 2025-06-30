from PIL import Image
from transformers import pipeline, TrOCRProcessor, VisionEncoderDecoderModel
from craft_hw_ocr import OCR

def ocr_with_coordinates(image_path):
    # Load and preprocess image
    image = Image.open(image_path).convert("RGB")
    
    # Load detectors and models
    detector  = pipeline("object-detection", model="hezarai/CRAFT")
    tac_proc  = TrOCRProcessor.from_pretrained("microsoft/trocr-large-printed")
    tac_model = VisionEncoderDecoderModel.from_pretrained("microsoft/trocr-large-printed")
    
    # Detect text regions
    detections = detector(image)
    boxes      = [det["box"] for det in detections]
    
    # Recognize text
    results = []
    for xmin, ymin, w, h in boxes:
        crop   = image.crop((xmin, ymin, xmin + w, ymin + h))
        inputs = tac_proc(crop, return_tensors="pt").to(tac_model.device)
        out    = tac_model.generate(**inputs)
        text   = tac_proc.batch_decode(out, skip_special_tokens=True)[0]
        results.append({"box": (xmin, ymin, w, h), "text": text})
    
    return results

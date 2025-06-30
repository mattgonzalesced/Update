from transformers import pipeline
from PIL import Image
import cv2

# 1.1 Initialize the CRAFT-based detector
detector = pipeline(
    "object-detection",
    model="hezarai/CRAFT",        # CRAFT text detector on Hugging Face
    device=-1                      # CPU (use 0 for GPU if available)
)

# 1.2 Quick smoke test
image = Image.open(r"C:\Users\m.gonzales\OneDrive - CoolSys Inc\Desktop\CHL-A.pdf").convert("RGB")
outputs = detector(image)

print(f"Detected {len(outputs)} text regions")
print("Sample output:", outputs[0])

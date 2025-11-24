from PIL import Image, ImageEnhance
import json
import io
import os

def process_bg_image(bg_path, config_path):
    # Load config
    with open(config_path, 'r', encoding='utf-8') as f:
        cfg = json.load(f)
    contrast = cfg.get('contrast', 100) / 100.0
    brightness = cfg.get('brightness', 100) / 100.0
    opacity = cfg.get('opacity', 100) / 100.0
    # Load image
    img = Image.open(bg_path).convert('RGBA')
    # Apply contrast
    img = ImageEnhance.Contrast(img).enhance(contrast)
    # Apply brightness
    img = ImageEnhance.Brightness(img).enhance(brightness)
    # Apply opacity
    alpha = img.split()[-1]
    alpha = alpha.point(lambda p: int(p * opacity))
    img.putalpha(alpha)
    # Save to bytes
    buf = io.BytesIO()
    img.save(buf, format='PNG')
    buf.seek(0)
    return buf.read()

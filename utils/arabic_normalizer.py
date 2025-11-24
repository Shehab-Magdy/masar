# utils/arabic_normalizer.py
def normalize_arabic(text: str) -> str:
    if not text:
        return text
    replacements = {
        "أ": "ا",
        "إ": "ا",
        "آ": "ا",
        "ة": "ه",
        "ى": "ي",
    }
    text = text.replace("ـ", "")
    for src, target in replacements.items():
        text = text.replace(src, target)
    return text.strip()
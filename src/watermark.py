import os
from pathlib import Path
from PIL import Image, ImageEnhance


def create_transparent_watermark(image_path: str, transparency_percent: float = 93.0, output_path: str = None) -> str:
    """
    Creates a copy of the input logo image with specified transparency percentage.
    For example, 93% transparent means 7% opacity.
    """
    img_path = Path(image_path)
    if not img_path.exists():
        raise FileNotFoundError(f"Logo file not found at: {image_path}")

    img = Image.open(img_path).convert("RGBA")
    
    # Calculate opacity ratio (e.g. 93% transparent -> 0.07 opacity)
    opacity_factor = max(0.0, min(1.0, (100.0 - transparency_percent) / 100.0))

    # Extract channels
    r, g, b, alpha = img.split()
    
    # Apply opacity factor to existing alpha channel
    alpha = alpha.point(lambda p: int(p * opacity_factor))
    
    transparent_img = Image.merge("RGBA", (r, g, b, alpha))

    if output_path is None:
        cache_dir = img_path.parent / ".cache"
        cache_dir.mkdir(parents=True, exist_ok=True)
        output_path = cache_dir / f"watermark_{int(transparency_percent)}pct_{img_path.name}"

    transparent_img.save(output_path, format="PNG")
    return str(output_path)

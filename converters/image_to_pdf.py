from PIL import Image
import os

def convert_image_to_pdf(input_path: str, output_filename: str = None) -> str:
    image = Image.open(input_path)
    if image.mode != "RGB":
        image = image.convert("RGB")
    if output_filename:
        output_path = os.path.join(os.path.dirname(input_path), output_filename)
    else:
        output_path = os.path.splitext(input_path)[0] + ".pdf"
    image.save(output_path, "PDF", resolution=100.0)
    return output_path

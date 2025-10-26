from pdf2image import convert_from_path
from PIL import Image
import os
import zipfile

def pdf_to_images_zip(pdf_path, output_dir, image_format='png', dpi=200, cleanup=True):
    os.makedirs(output_dir, exist_ok=True)

    # Convert all pages to images
    pages = convert_from_path(pdf_path, dpi=dpi)
    image_paths = []

    for i, page in enumerate(pages):
        img_name = f"page_{i + 1:03d}.{image_format.lower()}"  # e.g., page_001.png
        img_path = os.path.join(output_dir, img_name)
        page.save(img_path, image_format.upper())
        image_paths.append(img_path)

    # Create ZIP file containing all images (page-wise order)
    zip_filename = os.path.join(output_dir, "converted_images.zip")
    with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for img_path in image_paths:
            arcname = os.path.basename(img_path)  # keep only filename in ZIP
            zipf.write(img_path, arcname=arcname)

    # Optionally clean up image files after zipping
    if cleanup:
        for img_path in image_paths:
            try:
                os.remove(img_path)
            except Exception as e:
                print(f"Warning: could not delete {img_path}: {e}")

    return {
        "zip_file": zip_filename,
        "deleted_temp_images": cleanup
    }

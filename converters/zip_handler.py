import zipfile, os, logging
from PIL import Image
from reportlab.lib.pagesizes import letter, A4
from reportlab.pdfgen import canvas
from .docx_to_pdf import convert_docx_to_pdf
from .txt_to_pdf import convert_txt_to_pdf

logger = logging.getLogger(__name__)

def handle_zip_file(input_path: str) -> str:
    """
    Extract ZIP and convert contents:
    - TXT/DOCX files → separate PDFs
    - All images → single merged PDF (in order)
    Returns path to ZIP file containing all PDFs
    """
    import tempfile, shutil
    # Create a unique temp directory for this conversion
    temp_dir = tempfile.mkdtemp(prefix="convert_zip_")
    extract_dir = os.path.join(temp_dir, "extracted")
    pdf_output_dir = os.path.join(temp_dir, "pdfs")
    os.makedirs(extract_dir, exist_ok=True)
    os.makedirs(pdf_output_dir, exist_ok=True)
    
    logger.info(f"Extracting ZIP: {input_path}")
    with zipfile.ZipFile(input_path, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)
    
    # Collect files by type
    document_files = []
    image_files = []
    
    for root, dirs, files in os.walk(extract_dir):
        for filename in files:
            file_path = os.path.join(root, filename)
            ext = os.path.splitext(filename)[1].lower()
            
            if ext in ['.txt', '.docx']:
                document_files.append((filename, file_path, ext))
            elif ext in ['.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff']:
                image_files.append((filename, file_path))
    
    logger.info(f"Found {len(document_files)} documents and {len(image_files)} images")
    
    # Convert each document to separate PDF
    for filename, file_path, ext in document_files:
        try:
            base_name = os.path.splitext(filename)[0]
            output_filename = f"{base_name}.pdf"
            
            if ext == '.docx':
                logger.info(f"Converting DOCX: {filename}")
                pdf_path = convert_docx_to_pdf(file_path, output_filename)
            elif ext == '.txt':
                logger.info(f"Converting TXT: {filename}")
                pdf_path = convert_txt_to_pdf(file_path, output_filename)
            
            # Move PDF to output directory
            if pdf_path and os.path.exists(pdf_path):
                dest_path = os.path.join(pdf_output_dir, output_filename)
                if os.path.abspath(pdf_path) != os.path.abspath(dest_path):
                    os.rename(pdf_path, dest_path)
                logger.info(f"Created PDF: {output_filename}")
        except Exception as e:
            logger.error(f"Failed to convert {filename}: {str(e)}")
    
    # Merge all images into single PDF
    if image_files:
        try:
            logger.info(f"Merging {len(image_files)} images into single PDF")
            images_pdf_path = os.path.join(pdf_output_dir, "images_merged.pdf")
            merge_images_to_pdf(image_files, images_pdf_path)
            logger.info("Images merged successfully")
        except Exception as e:
            logger.error(f"Failed to merge images: {str(e)}")
    
    # Create output ZIP containing all PDFs
    output_zip_path = os.path.join(temp_dir, "converted_files.zip")
    with zipfile.ZipFile(output_zip_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
        for filename in os.listdir(pdf_output_dir):
            file_path = os.path.join(pdf_output_dir, filename)
            zip_out.write(file_path, filename)
    
    logger.info(f"Created output ZIP: {output_zip_path}")
    # Optionally, clean up temp_dir after returning zip (if you want to delete temp files)
    # shutil.rmtree(temp_dir)
    return output_zip_path


def merge_images_to_pdf(image_files, output_path):
    """
    Merge multiple images into a single PDF file, preserving order.
    Each image is placed on a separate page, scaled to fit the page.
    """
    # Keep images in the order they appear in the ZIP (don't sort)
    # image_files is already in ZIP order from os.walk
    
    # Open first image to get dimensions and create canvas
    first_img = Image.open(image_files[0][1])
    img_width, img_height = first_img.size
    first_img.close()
    
    # Use A4 if image is portrait-ish, otherwise use custom size
    if img_height > img_width:
        pagesize = A4
    else:
        pagesize = (img_width * 72 / 96, img_height * 72 / 96)  # Convert pixels to points
    
    c = canvas.Canvas(output_path, pagesize=pagesize)
    page_width, page_height = pagesize
    
    for filename, image_path in image_files:
        try:
            img = Image.open(image_path)
            
            # Convert to RGB if necessary
            if img.mode not in ('RGB', 'L'):
                img = img.convert('RGB')
            
            img_width, img_height = img.size
            
            # Calculate scaling to fit page while maintaining aspect ratio
            width_ratio = page_width / img_width
            height_ratio = page_height / img_height
            scale = min(width_ratio, height_ratio)
            
            new_width = img_width * scale
            new_height = img_height * scale
            
            # Center image on page
            x = (page_width - new_width) / 2
            y = (page_height - new_height) / 2
            
            # Save temp file for reportlab (it needs file path)
            temp_img_path = image_path + "_temp.jpg"
            img.save(temp_img_path, "JPEG", quality=95)
            img.close()
            
            # Draw image on canvas
            c.drawImage(temp_img_path, x, y, new_width, new_height)
            c.showPage()
            
            # Clean up temp file
            if os.path.exists(temp_img_path):
                os.remove(temp_img_path)
                
        except Exception as e:
            logger.error(f"Failed to add image {filename} to PDF: {str(e)}")
            continue
    
    c.save()
    logger.info(f"Saved merged PDF with {len(image_files)} images")

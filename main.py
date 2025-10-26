from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
import os, shutil, uuid, logging
from converters.docx_to_pdf import convert_docx_to_pdf
from converters.txt_to_pdf import convert_txt_to_pdf
from converters.image_to_pdf import convert_image_to_pdf
from converters.zip_handler import handle_zip_file
from converters.pdf_to_images import pdf_to_images_zip
from utils.file_utils import get_file_extension, create_temp_dir

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

app = FastAPI(title="File Converter Hub API")

UPLOAD_DIR = "temp_uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

@app.post("/convert-to-pdf")
async def convert_file(file: UploadFile = File(...)):
    ext = get_file_extension(file.filename)
    file_id = str(uuid.uuid4())
    input_path = os.path.join(UPLOAD_DIR, f"{file_id}_{file.filename}")
    
    # Generate output filename based on original filename
    original_name = os.path.splitext(file.filename)[0]
    output_filename = f"{original_name}.pdf"
    
    logger.info(f"Converting {file.filename} (type: {ext})")

    # Save uploaded file temporarily
    with open(input_path, "wb") as f:
        shutil.copyfileobj(file.file, f)
    
    logger.info(f"Saved input file: {input_path} ({os.path.getsize(input_path)} bytes)")

    try:
        if ext == ".docx":
            output_path = convert_docx_to_pdf(input_path, output_filename)
        elif ext == ".txt":
            output_path = convert_txt_to_pdf(input_path, output_filename)
        elif ext in [".jpg", ".jpeg", ".png"]:
            output_path = convert_image_to_pdf(input_path, output_filename)
        elif ext == ".zip":
            output_path = handle_zip_file(input_path)
            output_filename = os.path.basename(output_path)
        else:
            raise HTTPException(status_code=400, detail=f"Unsupported file type: {ext}")
        
        if not os.path.exists(output_path):
            logger.error(f"Conversion failed: output file not created")
            raise HTTPException(status_code=500, detail="Conversion failed: output file not created")
        
        output_size = os.path.getsize(output_path)
        logger.info(f"Conversion successful: {output_path} ({output_size} bytes)")
        
        if output_size < 1000:
            logger.warning(f"Output file is suspiciously small: {output_size} bytes")

        return FileResponse(output_path, filename=output_filename)
    except Exception as e:
        logger.error(f"Conversion error: {str(e)}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")
    finally:
        # Clean up temporary input file
        if os.path.exists(input_path):
            os.remove(input_path)
            logger.info(f"Cleaned up input file: {input_path}")

@app.post("/pdf-to-images")
async def pdf_to_images_endpoint(
    file: UploadFile = File(...),
    image_format: str = "png"
):
    """
    Convert PDF to images (PNG, JPEG, JPG) and return ZIP of images.
    """
    allowed_formats = ["png", "jpeg", "jpg"]
    fmt = image_format.lower()
    if fmt not in allowed_formats:
        raise HTTPException(status_code=400, detail=f"Invalid format. Choose from: {allowed_formats}")

    file_id = str(uuid.uuid4())
    input_path = os.path.join(UPLOAD_DIR, f"{file_id}_{file.filename}")
    output_dir = os.path.join(UPLOAD_DIR, f"{file_id}_images")

    # Save uploaded PDF
    with open(input_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    logger.info(f"Converting PDF to images: {input_path} as {fmt}")
    result = pdf_to_images_zip(input_path, output_dir, image_format=fmt)
    zip_path = result["zip_file"]

    if not os.path.exists(zip_path):
        raise HTTPException(status_code=500, detail="Image conversion failed.")

    return FileResponse(zip_path, filename="converted_images.zip", media_type="application/zip")

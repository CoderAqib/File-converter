import os, tempfile

def get_file_extension(filename: str) -> str:
    return os.path.splitext(filename)[1].lower()

def create_temp_dir() -> str:
    temp_dir = tempfile.mkdtemp()
    return temp_dir

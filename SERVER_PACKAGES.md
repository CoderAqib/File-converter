first install python 3.10.11

## install pandoc on your system using
choco install pandoc -y

## install wkhtmltopdf on your system using
choco install wkhtmltopdf -y

## install poppler on your system 
- go to: https://github.com/oschwartz10612/poppler-windows/releases/

- Extract the ZIP to a folder, e.g., C:\poppler.

- Add C:\poppler\Library\bin (or the bin folder inside your extracted Poppler directory) to your Windows PATH environment variable.

- restart your PC

### install requirements
pip install -r requirements.txt
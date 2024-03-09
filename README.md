# PDF Data Extractor

Application is to extract text from pdf using OCR and data fields using regex

## OCR
OCR (Optional character recognition) is used to scan text from images.
This application uses tesseract-ocr and pytesseract package.

Install tesseract from [here](https://tesseract-ocr.github.io/tessdoc/Installation.html).


## Running application
- Create and activate python environment.

```sh
python -m venv pdfenv

#Windows
pdfenv\Scripts\activate.bat

#Unix or Mac
pdfenv/bin/activate
```

- Install packages

```sh
pip install -r requirements.txt
```

- Run

```sh
python main.py "<PDF file or dir path>" "<Output excel file path>
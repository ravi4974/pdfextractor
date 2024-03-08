import pytesseract
import fitz
from PIL import Image
import pandas as pd
import argparse, os, re

regexp={
    '_text':'(To[,.;]*(.|\n)+?)Sub:'
}

pytesseract.pytesseract.tesseract_cmd=r'C:\Users\outlo\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

def get_value_or_na(pattern,index=0):
    def inner(text):
        value=re.findall(pattern, text)
        if len(value):
            if type(value[0]) is tuple:
                return value[0][index]
            else:
                return value[index]
        else:
            raise ValueError(f'${pattern} not found in ${text}')
    return inner

def get_images_from_pdf(path, mat=fitz.Matrix(2,2)):
    with fitz.open(path) as pdf:
        for page in pdf:
            clip=fitz.Rect(0,120,page.rect.width,page.rect.height*0.35)
            pixmap=page.get_pixmap(matrix=mat,clip=clip)
            image=Image.frombytes('RGB',[pixmap.width,pixmap.height],pixmap.samples).convert('L')
            yield image

def get_text_from_image(image):
    text=pytesseract.image_to_string(image)
    return '\n'.join(line.strip() for line in text.split('\n') if line.strip()) 


def get_data_from_text(text,fields):
    data={}
    subtext=get_value_or_na(regexp['_text'])(text)
    data=dict((field,func(subtext)) for field, func in fields.items())
   
    address='\n'.join([line for line in subtext.split('\n')[1:] if line and not any(line.find(v)>-1 for _,v in data.items())])

    data['address']=address

    return data

def get_rows_from_pdf(path):
    rows=[]

    fields={
        'name':get_value_or_na('To[,.;]*[\n\s]+(.+)'),
        'phone':get_value_or_na('(Mob|Tel)[^\d]*(\d+)',1),
        'email':get_value_or_na('Email.*?:\s*(.+)')
    }

    images=get_images_from_pdf(path)
    for image in images:
        text=get_text_from_image(image)
        data=get_data_from_text(text,fields)
        rows.append(data)
    
    return rows

def main(path,output_path):
    path=os.path.abspath(path)
    if os.path.isdir(path):
        base,_,files=os.walk(path)
    elif os.path.isfile(path):
        base,files=os.path.dirname(path), [os.path.basename(path)]

    rows=[]
    for file in files:
        if not file.lower().endswith('.pdf'): continue
        rows+=get_rows_from_pdf(os.path.join(base,file))
    
    #save into xls
    pd.DataFrame(rows).to_excel(output_path)


if __name__ == '__main__':
    parser = argparse.ArgumentParser('PDFExtractor','To extract text from pdf')
    parser.add_argument('path')
    parser.add_argument('output')

    args=parser.parse_args()

    if not os.path.exists(args.path):
        raise ValueError('${args.path} is not a valid path')
    
    main(args.path,args.output)
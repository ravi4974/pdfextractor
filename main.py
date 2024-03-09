import pytesseract
import fitz
from PIL import Image
import pandas as pd
import argparse, os, re

subtext_regex=[
    '(M/s(.|\n)+?E[-]*mail.*?:.+)',
    'To[,.;]?((.|\n)+?E[-]*mail.*?:.+)',
    'Dated:.*((.|\n)+)Sub:',
]

pytesseract.pytesseract.tesseract_cmd=r'<path of tesseract ocr exe>'

def get_value_or_na(pattern,index=0,optional=False):
    def inner(text):
        if type(pattern) is list:
            for p in pattern:
                try:
                    v=get_value_or_na(p,index,optional)(text)
                    if not v: continue
                    return v
                except ValueError as e:
                    if p==pattern[-1] and not optional:
                        raise e
                    else:
                        pass
            return None

        value=re.findall(pattern, text)
        if len(value):
            if type(value[0]) is tuple:
                return value[0][index].strip()
            else:
                return value[index].strip()
        elif optional:
            return None
        else:
            raise ValueError(f'${pattern} not found in ${text}')
    return inner

def get_images_from_pdf(path, mat=fitz.Matrix(2,2)):
    with fitz.open(path) as pdf:
        for page in pdf:
            clip=fitz.Rect(0,120,page.rect.width,page.rect.height*0.40)
            pixmap=page.get_pixmap(matrix=mat,clip=clip)
            image=Image.frombytes('RGB',[pixmap.width,pixmap.height],pixmap.samples).convert('L')
            yield image

def get_text_from_image(image):
    text=pytesseract.image_to_string(image)
    return '\n'.join(line.strip() for line in text.split('\n') if line.strip()) 


def get_data_from_text(text,fields):
    subtext=get_value_or_na(subtext_regex,0,True)(text)
    
    data={'text':text}

    data.update(dict((field,func(subtext or text) or 'NA') for field, func in fields.items()))

    if not subtext: return data
   
    address='\n'.join([line for line in subtext.split('\n')[1:] if line and not any(line.find(v)>-1 for _,v in data.items())])

    data['address']=address

    return data

def get_rows_from_pdf(path):
    rows=[]

    fields={
        'name':get_value_or_na(['(M/s.+)','To[,.;]*[\n\s]+(.+)','(.+)']),
        'phone':get_value_or_na('(Mob|Tel|Ph)[^\d]*(.+)',1,True),
        'email':get_value_or_na('E[-]*mail.*?:\W*(.+)',0,True)
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
    pd.DataFrame(rows).to_excel(output_path, index=False,columns=["name","address","phone","email","text"])


if __name__ == '__main__':
    parser = argparse.ArgumentParser('PDFExtractor','To extract data from pdf')
    parser.add_argument('path')
    parser.add_argument('output')

    args=parser.parse_args()

    if not os.path.exists(args.path):
        raise ValueError('${args.path} is not a valid path')
    
    main(args.path,args.output)
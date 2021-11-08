#Python 3 program to divide PDF on pages depend on information on this pages
#autor: Michał Franczak
#license: MIT
#version: 0.0.3
from PIL import Image
import pytesseract
import os, errno
import tkinter as tk
from tkinter import filedialog
from PyPDF2 import PdfFileWriter, PdfFileReader
import fitz
from fuzzywuzzy import process
import logging
import re
import settings
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("debug.log"),
        logging.StreamHandler()
    ]
)

wDir=''
actDocPages=0
struct=[]
dict=[]
pytesseract.pytesseract.tesseract_cmd=r'Tesseract-OCR\tesseract.exe'
def prepareSubFolder(f):
    logging.info('Preparing song: '+f)
    global struct
    global wDir
    struct=[]
    actDocPages=0
    path=wDir+'/'+f
    tmppath=wDir+'/tmp/'+f
    try:
        #os.mkdir(path)
        os.mkdir(tmppath)
    except OSError as e:
        if e.errno != errno.EEXIST:
            logging.error('Cannot create direcotry for this file')
            exit(1)
        else:
            logging.warning('Directory already exsit - content will be overwritten!')
            #for file in os.listdir(path):
            #    os.remove(path+'/'+file)
    else:
        logging.info('Subfolders created')

def extract_images(s,f):
    logging.info('Extracting images from file')
    global wDir
    global actDocPages
    global struct
    in_file=fitz.open(wDir+'/'+s+'/'+f)
    actDocPages=len(in_file)
    for i in range(actDocPages):
        page=in_file.load_page(i)
        mat = fitz.Matrix(settings.zoom,settings.zoom)
        clip = fitz.Rect(settings.minx,settings.miny,settings.maxx,settings.maxy)
        pix=page.get_pixmap(matrix=mat,clip=clip)
        output=wDir+'/tmp/'+f[0:-4]+'/page'+str(i)+'.png'
        struct.append('unknown')
        pix.save(output)

def find_instruments(text):
    lines=text.splitlines()
    name=''
    for line in lines:
        if len(line)>3:
            line = line.lower()
            #if any(inst in line for inst in dict):
            #    print(line)
            find=process.extractOne(line,dict)
            if find[1]>settings.minscore:
                logging.info('Found new instrument: '+line+' based on match: '+str(find))
                name=name+' '+line
    if name=='': 
        name='-'
    return name

def define_pages(f):
    global wDir
    global struct
    logging.info('Finding instruments names')
    #print(struct)
    actinst='unknown'
    for i in range(actDocPages):
        imfile=wDir+"/tmp/"+f[0:-4]+"/page"+str(i)+".png"
        text=pytesseract.image_to_string(imfile)
        find=find_instruments(text)
        if find != '-':
            actinst=find
        struct[i]=actinst
        logging.info('Intrument'+actinst+' added to structure with index: '+str(i))

def divide_pages(s,f):
    global wDir
    global struct
    input_pdf = PdfFileReader(wDir+'/'+s+'/'+f)
    resfile= open(wDir+'/'+s+'/description.txt', 'a')
    #actfile=struct[0]
    output = PdfFileWriter()
    for i, inst in enumerate(struct):
        if i==0: actfile=inst
        actPage=input_pdf.getPage(i)
        if actfile==inst:
            output.addPage(actPage)
        else:
            logging.info('Saving file: '+actfile)
            inst_safe=re.sub('[^a-zA-Z0-9 \n\.]',"_",actfile)
            dest_path=wDir+'/'+s+'/'+s+'_'+inst_safe+'.pdf'
            with open(dest_path,"wb") as output_stream:
                output.write(output_stream)
            output = PdfFileWriter()
            actfile=inst
            output.addPage(actPage)
        if i==len(struct)-1:
            logging.info('Saving file: '+actfile)
            inst_safe=re.sub('[^a-zA-Z0-9 \n\.]',"_",actfile)
            dest_path=wDir+'/'+s+'/'+s+'_'+inst_safe+'.pdf'
            with open(dest_path,"wb") as output_stream:
                output.write(output_stream)
        resfile.write(s+'; '+f+'; '+str(i+1)+'; '+re.sub('[^a-zA-Z0-9 \n\.]',"_",inst)+'\r\n')
    
def main():
    global wDir
    global dict
    logging.info('START')
    root = tk.Tk()
    root.withdraw()
    wDir=filedialog.askdirectory()
    dict_file=open('instruments.txt','r')
    dict=dict_file.read().splitlines()
    for i in range(len(dict)):
        dict[i]=dict[i].lower()
    #print(dict)
    if not wDir:
        logging.error("You didn't choose working directory!")
        exit(1)
    try:
        os.mkdir(wDir+'/tmp')
    except OSError as e:
        if e.errno != errno.EEXIST:
            logging.error('Cannot create temorary directory')
            exit(1)
        else:
            logging.info('Temporary directory exist - OK')
    for songname in os.listdir(wDir):
        if (songname=='tmp'): continue
        for filename in os.listdir(wDir+'/'+songname):
            prepareSubFolder(songname)
            if filename.endswith(".pdf"):
                extract_images(songname,filename)
                define_pages(filename)
                divide_pages(songname,filename)


if __name__ == '__main__':
    main()

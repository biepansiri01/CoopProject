from pathlib import Path
import subprocess
import os
from pdf2image import convert_from_path, convert_from_bytes
from PyPDF2 import PdfFileWriter, PdfFileReader
import shutil 
from PIL import Image


class ConverterAndSplit:

    
    libreoffice  = r"C:\Program Files\LibreOffice\program\soffice.exe"
    outputJPG = r"multipagedocclassify\DocDirectory\TMPfile\convertedJPG"
    outputPDF = r"multipagedocclassify\DocDirectory\TMPfile\convertedPDF"
    outputsplited = r"multipagedocclassify\DocDirectory\TMPfile\splitedPDF"


    def splitPDFPage(self, inputfile):
    
        with open(inputfile,"rb") as pdf_file:
            inputpdf = PdfFileReader(pdf_file)
        
            
            file_path = os.path.basename(inputfile)
            file_name = os.path.splitext(file_path)[0]

            numpage = len(range(inputpdf.numPages))

            if numpage == 1: #ถ้ามี 1 หน้าจะ copy ไฟล์ไปจัดเก็บที่ \splitedPDF\
                shutil.copy2(inputfile, self.outputsplited)
            else:
                """
                แบ่งหน้าของไฟล์ pdf โดยมีเลขกำกับต่อท้ายไฟล์ไว้ e.g. ไฟล์[เลขหน้า]
                """    

                for i in range(inputpdf.numPages): #หลายหน้า
                    output = PdfFileWriter()
                    output.addPage(inputpdf.getPage(i))
                    with open(self.outputsplited+"\\"+file_name+"[%s].pdf" % (i+1), "wb") as outputStream:
                        output.write(outputStream)
                    


    def convertPDF(self, inputfile):
    
        images = convert_from_bytes(open(inputfile, 'rb').read())
        image_no = 1
        
        if range(len(images))==range(0, 1): #เช็คว่าไฟล์ PDF มี 1 หน้าไหม ถ้ามีก็จะแปลงหน้านั้นเป็นไฟล์ JPG เลย
                    savepath = self.outputJPG+"\\"+ Path(inputfile).stem + ".jpg"
                    images[0].save(savepath, 'JPEG')
                    image_no+=1
        else:
            for i in range(len(images)): #ถ้ามีหลายหน้าก็จะแปลงแต่ละหน้าเป็น JPG โดยมีเลขกำกับต่อท้ายไฟล์ไว้ e.g. ไฟล์[เลขหน้า]
                savepath = self.outputJPG+"\\"+ Path(inputfile).stem + "["+ str(image_no)+ "]" + ".jpg"
                images[i].save(savepath, 'JPEG')
                image_no+=1
                
        self.splitPDFPage(inputfile) #จากนั้นทำการแบ่งหน้าของไฟล์ PDF


    def convertDOC(self, inputfile):
        r"""

        ต้อง install โปรแกรม LibreOffice จากนั้น
        กำหนด path ของตัวโปรแกรม LibreOffice ถ้าเป็น window ใช้ soffice.exe อยู่ในโฟลเดอร์ LibreOffice\program

        ref : https://stackoverflow.com/questions/50982064/converting-docx-to-pdf-with-pure-python-on-linux-without-libreoffice
              https://michalzalecki.com/converting-docx-to-pdf-using-python

        """
        soffice = self.libreoffice 
        converted2PDFfilepath = self.outputPDF+"\\"+Path(inputfile).stem+".pdf"  #ตำแหน่ง outputfile ที่แปลง doc to pdf

        #กำหนดคำสั่งของ commandline ในการแปลงไฟล์ doc/docx เป็น PDF
        cmdconv2pdf = f'"{soffice}"' + " --headless --convert-to pdf " + f'"{inputfile}"' +" --outdir " + f'"{self.outputPDF}"'
        
        #เรียกใช้คำสั่ง 
        subprocess.call(cmdconv2pdf) #call command to convert doc/docx file to pdf
        
        #ทำการแปลงไฟล์เป็น JPG แล้วทำการแบ่งหน้าไฟล์ PDF ที่ได้
        self.convertPDF(converted2PDFfilepath)

        #ลบไฟล์ที่แปลงจาก doc to pdf
        os.remove(converted2PDFfilepath) 


    def convertIMG2JPG(self, inputfile):

        #แปลงไฟล์รูปเป็น jpg เพื่อนำไปเข้า predict 
        #รองรับ file_extension == ".gif" ,".jfif",".jpeg",".jpg",".BMP",".png"
        #แปลงไฟล์ jpg to jpg เพราะ บางครั้งไฟล์ต้นฉบับเสียวิธีนี้แก้ปัญหาได้
        im = Image.open(inputfile)
        rgb_im = im.convert('RGB')
        rgb_im.save(self.outputJPG+"\\"+ Path(inputfile).stem + ".jpg")
        

    
    def convertIMG2pdf(self, inputfile): 
        
        #แปลงไฟล์รูปเป็น pdf 
        im = Image.open(inputfile)
        rgb_im = im.convert('RGB')
        rgb_im.save(self.outputsplited+"\\"+ Path(inputfile).stem + ".pdf")    

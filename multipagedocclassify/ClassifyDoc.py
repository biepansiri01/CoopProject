#from multipagedocclassify.ConverterAndsplit import Converter
#import sys
#sys.path.append("multipagedocclassify")

from ConverterAndSplit import ConverterAndSplit


from glob import glob
import os
from PIL import Image

import torch
import torchvision
from torchvision import models, transforms

from PyPDF2 import PdfFileMerger
import shutil 



class ClassifyDoc:

    converter = ConverterAndSplit()
    document_class = {0: 'certificate', 1: 'other', 2: 'resume', 3: 'transcript'}
    model = torch.load(r"multipagedocclassify\PredictionModel.pt") #โมเดลจากการทำ image classification
    predicteddir = r"multipagedocclassify\DocDirectory\predictedfile" #โฟลเดอร์ที่จัดเก็บไฟล์ต้นฉบับ
    outputdir = r"multipagedocclassify\DocDirectory\TMPfile\docclass" #ที่เก็บไฟล์ชั่วคราวที่ทำนายผลแล้วเป็น PDF
    
    doc_extension = [".doc", "docx"]
    image_extension = [".gif", ".jfif", ".jpeg", ".jpg", ".BMP", ".png"]



    def listAllFile(self, dirpath): #ลิสต์ชื่อไฟล์ในโฟลเดอร์ให้เป็น List
        listOfFiles = list()
        for (dirpath, dirnames, filenames) in os.walk(dirpath):
            listOfFiles += [os.path.join(dirpath, file) for file in filenames]

        return listOfFiles



    def predictIMG(self, inputpath): #ทำนายผลจากโมเดลที่ได้มาจาก image classification

        image_transforms = { 
            'test': transforms.Compose([
                transforms.Resize(size=(256,256)),
                transforms.CenterCrop(size=224),
                transforms.Grayscale(3),
                transforms.ToTensor(),
            ])}
            
            
        transform = image_transforms['test']
        
        image = Image.open(inputpath)
        
        image_tensor = transform(image)
    
        if torch.cuda.is_available():
            image_tensor = image_tensor.view(1, 3, 224, 224).cuda()
            #print("use cuda")
        else:
            image_tensor = image_tensor.view(1, 3, 224, 224)
            #print("use cuda")
        
        with torch.no_grad():
            self.model.eval()
            # Model outputs log probabilities
            out = self.model(image_tensor)

            confident = torch.exp(out)
            #print("confident of all class : ",confident[0].tolist())
            topk, topclass = confident.topk(1, dim=1)
            fileclass = self.document_class[topclass.cpu().numpy()[0][0]]
            #print("Output class :  ", document_class[topclass.cpu().numpy()[0][0]])
            os.remove(inputpath) #ลบไฟล์รูปที่นำมาทำนายผล
            return confident,fileclass #รีเทิร์นค่าออกมาเป็น ค่า confident และ ผลของคลาสเอกสารที่ทำนายออกมา



    def convertDocument(self, inputfile): #เช็คนามสกุลไฟล์ เมื่อตรงตามเงื่อนไขก็ทำการแปลงไฟล์
        file_extension = os.path.splitext(inputfile)[1]
        
        if(file_extension == ".pdf"):
            
            self.converter.convertPDF(inputfile)
            return True
            
        elif bool([ele for ele in self.doc_extension if(ele in file_extension)]) == True:
            
            self.converter.convertDOC(inputfile)
            return True
        
        elif bool([ele for ele in self.image_extension if(ele in file_extension)]) == True:
            
            self.converter.convertIMG2JPG(inputfile)
            self.converter.convertIMG2pdf(inputfile)
            return True
            
        else:
            return False

    @property
    def listDocClassDir(self): #ลิสต์โฟลเดอร์คลาสทั้งหมดใน outputdir
        path = self.outputdir
        return glob(path+"//*")



    def mergeFile(self, inputfile, outputpath): #ทำการรวมไฟล์ที่ถูกทำนายมาใน clas เดียวกันในเป็นไฟล์เดียว
        docclassdir = self.listDocClassDir
    
        file_path = os.path.basename(inputfile)
        file_name = os.path.splitext(file_path)[0]
        os.mkdir(outputpath)
        
        for dirpath in docclassdir:
            if not os.listdir(dirpath) :
                continue
            else:
                listfile = self.listAllFile(dirpath)
                merger = PdfFileMerger()
                for pdf in listfile:
                    merger.append(pdf)

                
                os.mkdir(outputpath+"//"+dirpath.split(os.sep)[-1])
                docpath = outputpath+"//"+dirpath.split(os.sep)[-1]
                with open(docpath+"\\"+file_name+".pdf", "wb") as fout:
                    merger.write(fout)
                    merger.close()
                    
                for pdf in listfile:
                    os.remove(pdf)


    def checkDocDir(self, inputfile): #เช็คว่ามีโฟลเดอรืซ้ำไหมใน predictedfile ถ้าไม่มีจะ return เป็นชื่อไฟล์แทน ถ้ามีก็จะเพิ่มเลขกำกับต่อท้ายไว้
        filename = os.path.basename(inputfile)
        
        docdir = self.predicteddir+"//"+filename
        
        i = 1
        while os.path.exists(docdir):
            docdir = self.predicteddir+"//"+filename+"(" + str(i) + ")"
            i=i+1
        
        return docdir

        
    def saveFileToItsClass(self, inputfile): #สั่งแปลงไฟล์ ทำนายผล และรวมไฟล์

        convert = self.convertDocument(inputfile)
        outputpath = self.checkDocDir(inputfile)
        
        if convert is True:
        
            listofpdf = self.listAllFile(self.converter.outputsplited)
            listofjpg = self.listAllFile(self.converter.outputJPG)

            for i in range(len(listofjpg)):
                prediction = self.predictIMG(listofjpg[i])
                shutil.move(listofpdf[i],self.outputdir+"\\"+prediction[1])
            
            self.mergeFile(inputfile,outputpath)
        


    def classifyDocument(self, inputfile): 
        file_extension = os.path.splitext(inputfile)[1]
        outputfiledir = self.checkDocDir(inputfile)

        self.saveFileToItsClass(inputfile)

        if ((bool([ele for ele in self.doc_extension if(ele in file_extension)]) == True) or
            (bool([ele for ele in self.image_extension if(ele in file_extension)]) == True) or
            (file_extension == ".pdf")):
            
            shutil.copy2(inputfile, outputfiledir) #จัดเก็บไฟล์ต้นฉบับในโฟลเดอร์ที่ทำนายผลแล้ว
        else:
            print("file not support")








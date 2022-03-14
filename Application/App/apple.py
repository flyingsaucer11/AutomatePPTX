from openpyxl import load_workbook
from pathlib import Path
import os

from Application.App import banana


class Apple():

    def __init__(self,flname):

        self.flname=flname
        self.startnum=4
        self.myworkbook = load_workbook(filename=self.flname, read_only=True, data_only=True)
        self.sheet_ranges = self.myworkbook.worksheets[0]
        #print(self.flname)

    def countTotal(self):
        counter=0
        for i in range(self.startnum,1000):
            str='B'+ f'{i}'
            if self.sheet_ranges[str].value == None:
                break
            else:
                counter=counter+1
        return counter

    def getImageFromRaw(self,x):
        name=self.sheet_ranges['R'+str(x)].value
        name=str(name)
        if name==str('First Person'):
            #print(name)
            signatureImage='production_data/signature1.png'
            extnum=1
        elif name==str('Second Person'):
            #print(name)
            signatureImage='production_data/signature2.png'
            extnum=2
        else:
            signatureImage='production_data/signature2.png'
            extnum=3
        return signatureImage, extnum

    def getOneRawValue(self,x):

        emailAddress=self.sheet_ranges['B'+str(x)].value
        firstName=self.sheet_ranges['C'+str(x)].value
        lastName=self.getImageFromRaw(x)
        idType=self.sheet_ranges['K'+str(x)].value
        idNumber=self.sheet_ranges['L'+str(x)].value
        expiration=self.sheet_ranges['O'+str(x)].value
        dateIssued=self.sheet_ranges['N'+str(x)].value
        dateExpires=self.sheet_ranges['O'+str(x)].value
        img_path , extnum = self.getImageFromRaw(x)

        return emailAddress,firstName,lastName,idType,idNumber,expiration,dateIssued,dateExpires,img_path, extnum

    def getAndSetData(self,prbar):
        total=self.countTotal()
        #print(total)
        self.stopnum=self.startnum+total

        for id in range(self.startnum,self.stopnum):

            emailAddress, firstName, lastName, idType, idNumber, expiration, dateIssued, dateExpires, img_path, extnum= self.getOneRawValue(id)
            idNumber=str(idNumber).encode("utf-8").decode("utf-8")
            expiration=expiration.strftime("%m/%d/%Y").encode("utf-8").decode("utf-8")
            dateIssued=dateIssued.strftime("%m/%d/%Y").encode("utf-8").decode("utf-8")
            dateExpires=dateExpires.strftime("%m/%d/%Y").encode("utf-8").decode("utf-8")


            banana.generatePPTX(firstName, lastName, idType, idNumber, expiration, dateIssued, dateExpires, img_path, extnum)
            percentage=((total-(self.stopnum-id)+1)/total)*100
            #print('Completed %.2f percent' %percentage)
            prbar.setValue(percentage)

        self.packup()

    def packup(self):
        for File in os.listdir("."):
            if File.endswith(".pptx"):
                fl=File.title()
                Path(fl).rename('output/'+fl)

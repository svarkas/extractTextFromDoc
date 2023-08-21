#!/usr/local/bin/python3.7

import sys
import binascii
import compoundfiles

class OldDoc:
 
    def getFIB(self, f_wd):
        f_wd.seek(0)
        fib= f_wd.read(1472)
        return fib

    def getTable(self, tblLabel):
        TABLE_MASK = b'\x02\x00'
        TABLE_MASK = int.from_bytes(TABLE_MASK,'big')
        tblLabel=int.from_bytes(tblLabel,'little')
        TABLE_FLAG = tblLabel & TABLE_MASK
        if TABLE_FLAG == 512:
            return "1Table"
        else:
            return "0Table"
    
    def getCLX(self, f_wd):
        f_wd.seek(418)
        fcClx=f_wd.read(4)
        lcbClx=f_wd.read(4)
        return fcClx, lcbClx

    def getPieceTable(self, CLX):
        pos=0
        goOn=True
        while goOn:
            if (CLX[pos]  == 2):
                goOn=False
                pieceLength=CLX[pos+1:pos+5]
                pieceLength=int.from_bytes(pieceLength,'little')
                pieceTable=CLX[pos+5:(pos+5)+pieceLength]
            elif (CLX[pos]  == 1):
                pos= pos + 1 + 1 + CLX[pos+1]
            else:
                goOn=false

        return pieceLength, pieceTable

    def getPieceDescriptors(self, pieceLength, pieceTable):
        pieceDescriptorsList = []
        pieceCount=(pieceLength-4)/12
        pieceCount=int(pieceCount)
        
        for i in range(pieceCount):
            cpStart = int.from_bytes(pieceTable[i*4:(i*4)+4],'little')
            cpEnd = int.from_bytes(pieceTable[(i+1)*4:((i+1)*4)+4], 'little')
            pieceDescriptor = pieceTable[((pieceCount + 1)*4) + (i*8):(((pieceCount + 1)*4) + (i*8))+8]
            pieceDescriptorDic = dict(pieceDescriptor = pieceDescriptor, cpStart=cpStart, cpEnd=cpEnd)
            pieceDescriptorsList.append(pieceDescriptorDic)
        
        return pieceDescriptorsList
    
    def isANSI(self, pieceDescriptor):
        fcValue = int.from_bytes(pieceDescriptor[2:6],'little')
        ENCFLAG_MASK = b'\x40\x00\x00\x00'
        ENCFLAG_MASK = int.from_bytes(ENCFLAG_MASK,'big')
        isANSI= False

        if ( fcValue & ENCFLAG_MASK == ENCFLAG_MASK ):
            isANSI = True 
        
        return isANSI
        
    def getFC(self, pieceDescriptor):
        fcValue = int.from_bytes(pieceDescriptor[2:6],'little')
        FCFLAG_MASK=b'\xBF\xFF\xFF\xFF'
        FCFLAG_MASK = int.from_bytes(FCFLAG_MASK,'big')
        fc = fcValue & FCFLAG_MASK
        fc= fc/2
        
        return int(fc)
       
    def extractText(self, inputFile):
        doc = compoundfiles.CompoundFileReader(inputFile)
        od = OldDoc()
        f_wd=doc.open(doc.root['WordDocument'])
        FIB= od.getFIB(f_wd)
        fcClx= FIB[418:422]
        lcbClx = FIB[422:426]
        fcClx=int.from_bytes(fcClx,'little')
        lcbClx=int.from_bytes(lcbClx,'little')
        table=od.getTable(FIB[10:12])
        f_wd.close()
        f_tbl=doc.open(doc.root[table])
        f_tbl.seek(fcClx)
        CLX=bytearray(f_tbl.read(lcbClx))
        pieceLength, pieceTable = od.getPieceTable(CLX)


        text=''
        for pieceDescriptorDic in od.getPieceDescriptors(pieceLength, pieceTable):
            cpStart = pieceDescriptorDic["cpStart"]
            cpEnd = pieceDescriptorDic["cpEnd"]
            isANSI = od.isANSI(pieceDescriptorDic["pieceDescriptor"])
            fc = od.getFC(pieceDescriptorDic["pieceDescriptor"])

            encoding = 'cp1252'
            cp = cpEnd - cpStart
            if ( not isANSI ):
                encoding = 'utf-8'
                cp=cp*2
            f_od=doc.open(doc.root['WordDocument'])
            f_od.seek(fc)
            textPart=f_od.read(cp)
            f_od.close()
            textPart = textPart.replace(b'\r',b'\n')
            text = text + str(textPart, encoding, errors='ignore')
        return text 

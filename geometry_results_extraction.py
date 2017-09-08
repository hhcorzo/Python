
"""extraction for propagator files"""
import os
import re
import openpyxl
import sys
workbook = openpyxl.Workbook()  #creates openpyxl workbook

#logFilesFolder is name of folder containing log files
logFilesFolder='/Logs'

'''path to this file'''
path=os.path.dirname(os.path.realpath(__file__))
pathorigin=path     #used to save workbook in this location
excelFilePathName='/geometry_results.xlsx'

#columns for each variable in workbook
colFileInformation='A'
colMolecule='B'
colCharge='C'
colMultiplicity='D'
colBasis='E'
colNineFive='F'
colFullPointGroup='G'
colLargestAbelianSubgroup='H'
colLargestConciseAbelianSubgroup='I'
colCCSDT = 'J'
colHF='K'
colCORR='L'
colMP2='M'
colMP3='N'
colMP4D='O'
colMP4DQ='P'
colMP4SDQ='Q'
colCCSD = 'R'


#def writeDataToExcel(worksheet, row, fileInformation, molecule, charge, multiplicity, basis, nineFive,\
#            orbital, eigenValue, ps, cff):
def writeDataToExcel(worksheet, row, fileInformation, molecule, charge, multiplicity, basis, nineFive,fullPointGroup,largestAbelianSubgroup,largestConciseAbelianSubgroup, \
        hf, ccsdt,mp2,mp3,mp4d,mp4dq,mp4sdq,ccsd,difference):
    '''writesDataToExcel takes is called by dataExtract. It takes in the variables found in
    data extraction and writes it into the openpyxl workbook'''

    worksheet[colFileInformation+str(row)]=fileInformation
    worksheet[colMolecule+str(row)]=molecule
    worksheet[colCharge+str(row)]=charge
    worksheet[colMultiplicity+str(row)]=multiplicity
    worksheet[colBasis+str(row)]=basis
    worksheet[colNineFive+str(row)]=nineFive
    worksheet[colFullPointGroup+str(row)]=fullPointGroup
    worksheet[colLargestAbelianSubgroup+str(row)]=largestAbelianSubgroup
    worksheet[colLargestConciseAbelianSubgroup+str(row)]=largestConciseAbelianSubgroup
    worksheet[colCCSDT+str(row)]=float(ccsdt)
    worksheet[colCORR+str(row)]=difference

    worksheet[colMP2+str(row)]=float(mp2)
    worksheet[colMP3+str(row)]=float(mp3)
    worksheet[colMP4D+str(row)]=float(mp4d)
    worksheet[colMP4DQ+str(row)]=float(mp4dq)

    worksheet[colMP4SDQ+str(row)]=float(mp4sdq)
    worksheet[colCCSD+str(row)]=float(ccsd)
    worksheet[colHF+str(row)]=float(hf)


def numberOfBasisSets(logarray):
    '''returns a list of the split log arrays by basis set. length is number of basis sets'''
    commandLocation=[]
    logsToReturn=[]
    x=0
    while x < len(logarray):
        if logarray[x] =='Final':
            commandLocation.append(x)
        x+=1
    commandLocation.append(len(logarray))
    x=0
    logsToReturn.append(logarray[:commandLocation[0]])
    #the first log in the array is from the start of the file to the first keyword
    while x< len(commandLocation)-1:
        b=logarray[commandLocation[x]:commandLocation[x+1]]
        logsToReturn.append(b)
        x+=1
    return logsToReturn


def dataExtract(path):
    '''Main function in script. Calls other functions. Takes in path of this file and extracts
    data from the log files folder. Then calls functions above to add data to the openpyxl file'''

    #prepare openpyxl first
    worksheet=workbook.active
    worksheet.title="Data"
    #creates worksheet Data

    #add headings to each column
    worksheet[colFileInformation+'1']='File'
    worksheet[colMolecule+'1']='Molecule'
    worksheet[colCharge+'1']='Charge'
    worksheet[colMultiplicity+'1']='Multiplicity'
    worksheet[colBasis+'1']='Basis'
    worksheet[colNineFive+'1']='9/5'
    worksheet[colCCSDT+'1']='CCSD(T)'
    worksheet[colHF+'1']='HF'
    worksheet[colCORR+'1']='CCSD(T)-HF'
    worksheet[colMP2+'1']='MP2'
    worksheet[colMP3+'1']='MP3'
    worksheet[colMP4D+'1']='MP4D'
    worksheet[colMP4DQ+'1']='MP4DQ'
    worksheet[colMP4SDQ+'1']='MP4SDQ'
    worksheet[colCCSD+'1']='CCSD'
    worksheet[colFullPointGroup+'1']='Full Point Group'
    worksheet[colLargestAbelianSubgroup+'1']='Largest Abelian Subgroup'
    worksheet[colLargestConciseAbelianSubgroup+'1']='Largest Concise Abelian Subgroup'

    row=2
    #extraction from log files starts here
    logFiles=[]
    w=0
    for path, subdirs, files in os.walk(path+logFilesFolder):
        for name in files:
            if os.path.join(path, name)[len(os.path.join(path, name))-4:len(os.path.join(path, name))]=='.log':
                for line in open(os.path.join(path, name),'r'):
                    if re.search("Normal",line):
                        logFiles.append(os.path.join(path, name))


    for currentFile in logFiles:
        log = open(currentFile, 'r').read()
        splitLog = re.split(r'[\\\s]\s*', log)  #splits string with \ (\\), empty space (\s) and = and ,
        fileInformation=currentFile
        nineFiveFound=False
        x=0
        while x<len(splitLog):
            if splitLog[x]=='Stoichiometry':
                molecule=splitLog[x+1]
            if splitLog[x]=='Charge' and splitLog[x-1]=='Z-matrix:':
                charge=splitLog[x+2]
            if splitLog[x]=='Multiplicity':
                multiplicity=splitLog[x+2]
            if splitLog[x]=='Standard' and splitLog[x+1]=='basis:':
                basis=splitLog[x+2] +' '+splitLog[x+3]+splitLog[x+4]
            if splitLog[x][0:4]=='9/5=' and nineFiveFound==False:
                nineFive=splitLog[x][4]
                s=0
                f=0
                n=0
                while n<len(splitLog[x]):
                    if splitLog[x][n]=='=' and s==0:
                        s=n+1
                    if splitLog[x][n]==',' and f==0:
                        f=n
                    n+=1
                nineFive=splitLog[x][s:f]
                nineFiveFound=True
            if splitLog[x]=='Full':
                fullPointGroup=splitLog[x+3]
            if splitLog[x]=='Largest' and splitLog[x+1]=='Abelian':
                largestAbelianSubgroup=splitLog[x+3]
            if splitLog[x]=='concise':
                largestConciseAbelianSubgroup=splitLog[x+3]


            if splitLog[x]=='SP':

                y=0
                while splitLog[x+y]!='@':
                    y+=1
                valuesBlock=''.join(splitLog[x:x+y])    #block of text containing values needs to be isolated
                l=0
                while l < len(valuesBlock):
                                                    #find HF
                    if valuesBlock[l:l+3]=='HF=' and valuesBlock[l-2:l]!='PU':
                        start=l+3
                        end=l+5
                        numberDone=False
                        while numberDone==False:
                            end+=1
                            try:
                                float(valuesBlock[start:end])
                            except:
                                numberDone=True
                        hf=valuesBlock[start:end-1]
                                                        #find CCSD(T)
                                                        #calculate difference
                    if valuesBlock[l:l+8]=='CCSD(T)=':
                        start=l+8
                        end=l+11
                        numberDone=False
                        while numberDone==False:
                            end+=1
                            try:
                                float(valuesBlock[start:end])
                            except:
                                numberDone=True
                        ccsdt=valuesBlock[start:end-1]
                        difference = float(ccsdt) - float(hf)
                                                        #find MP2
                    if valuesBlock[l:l+4]=='MP2=':
                        start=l+4
                        end=l+6
                        numberDone=False
                        while numberDone==False:
                            end+=1
                            try:
                                float(valuesBlock[start:end])
                            except:
                                numberDone=True
                        mp2=valuesBlock[start:end-1]
                                                    #find MP3
                    if valuesBlock[l:l+4]=='MP3=':
                        start=l+4
                        end=l+6
                        numberDone=False
                        while numberDone==False:
                            end+=1
                            try:
                                float(valuesBlock[start:end])
                            except:
                                numberDone=True
                        mp3=valuesBlock[start:end-1]
                                                    #find MP4D
                    if valuesBlock[l:l+5]=='MP4D=':
                        start=l+5
                        end=l+6
                        numberDone=False
                        while numberDone==False:
                            end+=1
                            try:
                                float(valuesBlock[start:end])
                            except:
                                numberDone=True
                        mp4d=valuesBlock[start:end-1]
                                                    #find MP4DQ
                    if valuesBlock[l:l+6]=='MP4DQ=':
                        start=l+6
                        end=l+7
                        numberDone=False
                        while numberDone==False:
                            end+=1
                            try:
                                float(valuesBlock[start:end])
                            except:
                                numberDone=True
                        mp4dq=valuesBlock[start:end-1]
                                                        #find MP4SDQ
                    if valuesBlock[l:l+7]=='MP4SDQ=':
                        start=l+7
                        end=l+8
                        numberDone=False
                        while numberDone==False:
                            end+=1
                            try:
                                float(valuesBlock[start:end])
                            except:
                                numberDone=True
                        mp4sdq=valuesBlock[start:end-1]
                                                        #find CCSD
                    if valuesBlock[l:l+5]=='CCSD=':
                        start=l+5
                        end=l+6
                        numberDone=False
                        while numberDone==False:
                            end+=1
                            try:
                                float(valuesBlock[start:end])
                            except:
                                numberDone=True
                        ccsd=valuesBlock[start:end-1]
                    l+=1

            x+=1
            #send variables from data extraction to writeDataToExcel
            #writeDataToExcel(worksheet, row, fileInformation, molecule, charge, multiplicity, basis, nineFive,\
            #orbital, eigenValue, ps, cff)
        #print(worksheet, row, fileInformation, molecule, charge, multiplicity, basis, nineFive)
        difference = float(ccsdt) - float(hf)
        writeDataToExcel(worksheet, row, fileInformation, molecule, charge, multiplicity, basis, nineFive,fullPointGroup,largestAbelianSubgroup,largestConciseAbelianSubgroup, \
        hf, ccsdt,mp2,mp3,mp4d,mp4dq,mp4sdq,ccsd, difference)


        row+=1

    workbook.save(pathorigin + excelFilePathName)     #saves file

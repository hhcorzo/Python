import os
import re
import openpyxl
import stat      #implemented to make files executable
workbook = openpyxl.Workbook()  #creates openpyxl workbook

#logFilesFolder is name of folder containing log files
logFilesFolder='/Final_Geom_Extra'

'''path to this file'''
path=os.path.dirname(os.path.realpath(__file__))
pathorigin=path     #used to save workbook in this location
excelFilePathName='/geometry_logFile_data.xlsx'
gjfFileFolder='/gjf_files'

#columns for each variable in workbook
colFileInformation='A'
colMolecule='B'
colCharge='C'
colMultiplicity='D'
colBasis='E'
colSymmetry='F'
colHarmonicFrequency='G'
colLastLine='H'

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
worksheet[colSymmetry+'1']='Symmetry'
worksheet[colHarmonicFrequency+'1']='Harmonic Frequency'
worksheet[colLastLine+'1']='Last Line'

parametersFileText='large\n8\n\n8gb'

gaussCommand='rung09'

userCharge=0
userMultiplicity=1

def run():
    dataExtract(path)

def gjfFile(name,charge,multiplicity,geometry):
    basisSets=['Aug-cc-pvDz','Aug-cc-pvTz','Aug-cc-pvQz','Aug-cc-pv5z']
    if not os.path.exists(path+gjfFileFolder+'/'+str(userCharge)+str(userMultiplicity)):
            os.makedirs(path+gjfFileFolder+'/'+str(userCharge)+str(userMultiplicity))
            if not os.path.exists(path+gjfFileFolder+'/'+str(userCharge)+str(userMultiplicity)+'/parameters'):
                file=open(path+gjfFileFolder+'/'+str(userCharge)+str(userMultiplicity)+'/parameters', "w")
                file.write(parametersFileText)
                file.close()
                st=os.stat(path+gjfFileFolder+'/'+str(userCharge)+str(userMultiplicity)+'/parameters')
                os.chmod(path+gjfFileFolder+'/'+str(userCharge)+str(userMultiplicity)+'/parameters', st.st_mode | stat.S_IEXEC)   #makes file exectubale
            if not os.path.exists(path+gjfFileFolder+'/'+str(userCharge)+str(userMultiplicity)+'/run'):
                file=open(path+gjfFileFolder+'/'+str(userCharge)+str(userMultiplicity)+'/run', "w")
                file.close()
                st=os.stat(path+gjfFileFolder+'/'+str(userCharge)+str(userMultiplicity)+'/run')
                os.chmod(path+gjfFileFolder+'/'+str(userCharge)+str(userMultiplicity)+'/run', st.st_mode | stat.S_IEXEC)   #makes file exectubale


    geometryText=''
    for g in geometry:
        geometryText+=g
        geometryText+='\n'


    for b in basisSets:
        file=open(path+gjfFileFolder+'/'+str(userCharge)+str(userMultiplicity)+'/'+str(name)+'_'+b[-2]+".gjf","w")
        file.write('%nprocshared=8 \n#CCSD(T)/ '+b+' tran=abcd\n\n'\
        #+str(name)+'\n\n'+str(charge)+' '+str(multiplicity)+'\n'+geometryText+'\n\n')
        +str(name)+'\n\n'+str(userCharge)+' '+str(userMultiplicity)+'\n'+geometryText+'\n\n')
        file.close()

        runfile = open(path+gjfFileFolder+'/'+str(userCharge)+str(userMultiplicity)+'/run', "a")     #a lets you append file
        runfile.write(gaussCommand+' '+str(name)+'_'+b[-2]+".gjf < parameters"+'\n')
        print(gaussCommand+' '+str(name)+'_'+b[-2]+".gjf < parameters"+'\n')
        #runfile.close()

def writeDataToExcel(row,fileInformation,molecule,charge,multiplicity,basis,symmetry,harmonicFrequency,lastLine):
    '''writesDataToExcel takes is called by dataExtract. It takes in the variables found in
    data extraction and writes it into the openpyxl workbook'''

    worksheet[colFileInformation+str(row)]=fileInformation
    worksheet[colMolecule+str(row)]=molecule
    worksheet[colCharge+str(row)]=charge
    worksheet[colMultiplicity+str(row)]=multiplicity
    worksheet[colBasis+str(row)]=basis
    worksheet[colSymmetry+str(row)]=symmetry
    worksheet[colHarmonicFrequency+str(row)]=harmonicFrequency
    worksheet[colLastLine+str(row)]=lastLine


def dataExtract(path):

    row=2
    #extraction from log files starts here
    logFiles=[]

    for path, subdirs, files in os.walk(path+logFilesFolder):
        for name in files:
            if os.path.join(path, name)[len(os.path.join(path, name))-4:len(os.path.join(path, name))]=='.log':
                logFiles.append(os.path.join(path, name))

    for currentFile in logFiles:
        log = open(currentFile, 'r').read()
        splitLog = re.split(r'[\\\s]\s*', log)  #splits string with \ (\\), empty space (\s) and = and ,
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
            if splitLog[x]=='Full' and splitLog[x+1]=='point' and splitLog[x+2]=='group':
                symmetry=splitLog[x+3]
            if splitLog[x]=='normal' and splitLog[x+1]=='coordinates:':
                y=0
                while splitLog[x+y]!='Frequencies':
                    y+=1
                harmonicFrequency=splitLog[x+y+2]
            if splitLog[x]=='Redundant':
                y=0
                while splitLog[x+y]!='Recover':
                    y+=1
                geometry=splitLog[x+6:x+y]

            x+=1
        y=-1
        lastLine=None
        while y>-15:
            if splitLog[y]=='Normal':
                lastLine=splitLog[y:]
                lastLine=' '.join(lastLine)
                break
            y-=1
        if lastLine==None:
            lastLine=splitLog[-15:]
            lastLine=' '.join(lastLine)
        fileInformation=currentFile
        #find the name of the file using fileInformation
        n=-1
        lastLetterName=None
        firstLetterName=None
        while n*-1<len(fileInformation):
            if fileInformation[n]=='.' and fileInformation[n+1]=='c':
                lastLetterName=n
            if fileInformation[n]=='/':
                firstLetterName=n+1
            if lastLetterName!=None and firstLetterName!=None:
                break
            n-=1
        name=fileInformation[firstLetterName:lastLetterName]

        gjfFile(name,charge,multiplicity,geometry)
        writeDataToExcel(row,fileInformation,molecule,charge,multiplicity,basis,symmetry,harmonicFrequency,lastLine)
        row+=1
    workbook.save(pathorigin + excelFilePathName)     #saves file

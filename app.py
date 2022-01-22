from modules import pylightxl as xl
import csv
import time
import glob


timeAsString = time.strftime("%Y%m%d_%H.%M.%S")
inputFilesPath = './input'
gradeSheetName = 'Grade Export'
outputFileName = 'aggregated_'+ timeAsString + '.csv'


inputFiles = list(glob.iglob(inputFilesPath + '/*.xlsx'))

#Remove Excel's temporary owner files from list
for file in inputFiles:
    if file.startswith(inputFilesPath + '/~$'):
        inputFiles.remove(file)


#grab header from the first row of a random file:

mainWorkBookPath = inputFiles[0]
mainWorkBook = xl.readxl(fn=mainWorkBookPath)
headerData = mainWorkBook.ws(ws=gradeSheetName).row(row=1)



#Create new CSV file and write header to it

with open(outputFileName, 'w', newline='') as myfile:
    wr = csv.writer(myfile, dialect='excel')
    wr.writerow(headerData)

#append grades to CSV file
for filepath in inputFiles:

    workBook = xl.readxl(fn=filepath)

    if workBook.ws(ws=gradeSheetName).row(row=1) != headerData:
        raise Exception("E: Header (First Row) Mismatch between", mainWorkBookPath, " and ", filepath, "\nThe output file is not guaranteed to be valid.")

    gradeData = workBook.ws(ws=gradeSheetName).row(row=2)

    with open(outputFileName, 'a', newline='') as myfile:
        wr = csv.writer(myfile, dialect='excel')
        wr.writerow(gradeData)

print("Aggregated grades saved as " + outputFileName)

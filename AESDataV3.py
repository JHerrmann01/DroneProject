###American Environmental Solutions Data Manipulation###
##             Created by Jeremy Herrmann             ##


##Import Libraries##
from __future__ import print_function
from os.path import join, dirname, abspath
import xlrd
from xlrd.sheet import ctype_text
import xlsxwriter
####################

def loadSpreadsheet():
    fname = join(dirname(dirname(abspath(__file__))), 'AES/First Spreadsheet', 'GBZ65745 Excel SE855 GLENWOOD RD-1 (copy).xls')
    xl_workbook = xlrd.open_workbook(fname)
    xl_sheet = xl_workbook.sheet_by_name("Results")
    return xl_workbook, xl_sheet

def grabSimpleInformation(xl_workbook, xl_sheet):
    numSpaces = 0
    generalAreas = {}
    num_cols = xl_sheet.ncols
    for row_idx in range(8, xl_sheet.nrows-7):
        if(xl_sheet.cell(row_idx,0).value == "Mercury"):
            Mercury_Values_Raw  = (xl_sheet.row(row_idx))
        if(xl_sheet.cell(row_idx,0).value == "pH at 25C - Soil"):
            Corrosivity_Values_Raw  = (xl_sheet.row(row_idx))
        if(xl_sheet.cell(row_idx,0).value == "Flash Point"):
            Flashpoint_Values_Raw  = (xl_sheet.row(row_idx))
        if(xl_sheet.cell(row_idx,0).value == "Ignitability"):
            Ignitability_Values_Raw  = (xl_sheet.row(row_idx))
        if(xl_sheet.cell(row_idx,0).value == "Reactivity  Cyanide"):
            Reactivity_Values_Cyanide_Raw  = (xl_sheet.row(row_idx))
        if(xl_sheet.cell(row_idx,0).value == "Reactivity Sulfide"):
            Reactivity_Values_Sulfide_Raw  = (xl_sheet.row(row_idx))
        if(xl_sheet.cell(row_idx,0).value == "Total Cyanide (SW9010C Distill.)"):
            Cyanide_Values_Raw  = (xl_sheet.row(row_idx))
        if(numSpaces%3 == 0):
            generalAreas[int(row_idx)] = str(xl_sheet.cell(row_idx,0).value)
            numSpaces +=1
        if(xl_sheet.cell(row_idx,0).value == ""):
            numSpaces += 1
    return Mercury_Values_Raw, Corrosivity_Values_Raw, Flashpoint_Values_Raw, Ignitability_Values_Raw, Reactivity_Values_Cyanide_Raw, Reactivity_Values_Sulfide_Raw, Cyanide_Values_Raw, generalAreas
    
def sortGeneralAreas(generalAreas):
    keys = generalAreas.keys()
    sortedGenAreas = [[0 for i in range(2)]for i in range(len(keys))]
    for x in range(0,len(keys)):
        smallestKey = 100000
        for key in generalAreas.keys():
            if(key < smallestKey):
                smallestKey = key
        sortedGenAreas[x][0] = int(smallestKey)
        sortedGenAreas[x][1] = str(generalAreas.pop(smallestKey))
    return sortedGenAreas

def insertRowsIntoAreas(xl_sheet, sortedGenAreas):
    rowsInArea = [[""]for i in range(len(sortedGenAreas))]
    for x in range(0,len(sortedGenAreas)):
        rowsInArea[x][0] = sortedGenAreas[x][1]
    numAreas = len(sortedGenAreas)
    for x in range(0 , numAreas):
        if(x < numAreas-1):
            for y in range(sortedGenAreas[x][0]+1, sortedGenAreas[x+1][0]-2):
                rowsInArea[x].append(xl_sheet.row(y))            
        else:
            for y in range(sortedGenAreas[x][0]+1, xl_sheet.nrows-7):
                rowsInArea[x].append(xl_sheet.row(y))
    return rowsInArea

print("Beginning program...")
#Loading the file to be parsed
xl_workbook, xl_sheet = loadSpreadsheet()

#Grabbing basic information
Company_Name = xl_sheet.cell(0, 0).value
Type_Samples_Collected_Raw = xl_sheet.row(4)

global firstIndex
firstIndex = 6

#Begin parsing to find simple useful information
Mercury_Values_Raw, Corrosivity_Values_Raw, Flashpoint_Values_Raw, Ignitability_Values_Raw, Reactivity_Values_Cyanide_Raw, Reactivity_Values_Sulfide_Raw, Cyanide_Values_Raw, generalAreas = grabSimpleInformation(xl_workbook, xl_sheet)        

#Sort the general areas in increasing order(Row number)
sortedGenAreas = sortGeneralAreas(generalAreas)
    
#Insert the rows that belong to each respective area
rowsInArea = insertRowsIntoAreas(xl_sheet, sortedGenAreas)
print("Done Parsing")
print()

########################################################################################################################
def startWritingFinalFile():
    workbook = xlsxwriter.Workbook('/home/jeremy/Desktop/AES/Excel_Reformatting.xlsx')
    worksheet = workbook.add_worksheet()
    return workbook, worksheet

#Refining a given row
def valueRefinerMetals(inputArrayRaw):
    outputArray = []
    pos = 0
    units = str(inputArrayRaw[2].value)
    divisor = 1
    if(units[0:2] == "ug"):
        divisor = 1000
    for value in inputArrayRaw:
        if((pos >= firstIndex and pos%2 == firstIndex%2) or (pos == 0) or (pos == 2)):
            if(pos == 0):
                outputArray.append(str(value.value))
            elif(pos == 2):
                outputArray.append("ppm")
                outputArray.append("")
            elif(str(value.value).find("<") == -1):
                outputArray.append(str(round((float(value.value)/divisor), 5)))
            else:
                outputArray.append("N.D.")
        pos+=1
    return(outputArray)

def isDetected(compound):
    hasFloat = False
    for x in compound:
        try:
            val = float(x)
            hasFloat = True
            break
        except Exception as e:
            val = ""        
    return hasFloat

def isNumber(value):
    try:
        val = float(value)
        return True
    except Exception as e:
        return False

def removeUselessRows(rowsInArea, index):
    y = 1
    lenRow = (len(rowsInArea[index][1]))
    while(y < len(rowsInArea[index])):
        if not isDetected(rowsInArea[index][y]):
            rowsInArea[index].remove(rowsInArea[index][y])
            y -= 1
        y += 1
    if(len(rowsInArea[index]) == 1):
        emptyArray = ["None Detected", "_", "_"]
        for x in range(len(emptyArray), lenRow):
            emptyArray.append("N.D.")
        rowsInArea[index].append(emptyArray)
    return rowsInArea[index]

def createBeginning(worksheet, currLine):
    line = 1

    x = len(Type_Samples_Collected)
    offset = 4

    finalLetter=""
    
    if 64+x+offset > 90:
        firstLetter = chr(int(65+(((x+offset)-26)/26)))
        secondLetter = chr(64+((x+offset)%26))
        finalLetter = firstLetter+secondLetter
    else:
        finalLetter = chr(64+x+offset)
    
    for x in range(0, 5):
        worksheet.merge_range("B"+str(line)+":"+finalLetter+str(line), "")
        line += 1
    return worksheet, currLine

def createHeading(worksheet, currLine, Type_Samples_Collected_Raw, formatOne):
    formatOne.set_text_wrap(True)
    Type_Samples_Collected = []
    pos = 0
    for value in Type_Samples_Collected_Raw:
        if((pos >= firstIndex and pos%2 == firstIndex%2) or (pos ==0)):
            Type_Samples_Collected.append(value.value)
        pos+=1
    worksheet.write('B'+str(currLine), 'Parameter', formatOne)
    worksheet.write('C'+str(currLine), 'Compounds Detected', formatOne)
    worksheet.write('D'+str(currLine), 'Units', formatOne)
    worksheet.write('E'+str(currLine), 'NYSDEC Part 375 Unrestricted Use Criteria', formatOne)

    offset = 4
    
    for x in range(1,len(Type_Samples_Collected)):
        if(64+x+offset < 90):
            worksheet.write(str(chr(65+x+offset))+str(currLine), str(Type_Samples_Collected[x]), formatOne)
        else:
            firstLetter = chr(int(65+(((x+offset)-26)/26)))
            secondLetter = chr(65+((x+offset)%26))
            col = firstLetter + secondLetter
            worksheet.write(col+str(currLine), str(Type_Samples_Collected[x]), formatOne)
    currLine += 1
    return worksheet, currLine, Type_Samples_Collected

def addMercuryValues(worksheet, currLine, Mercury_Values_Raw, formatOne, formatTwo):
    Mercury_Values = valueRefinerMetals(Mercury_Values_Raw)
    offset = 2
    worksheet.write('B'+str(currLine), 'Mercury 7471', formatOne)
    for x in range(0, len(Mercury_Values)):
        if(isNumber(Mercury_Values[x])):
            if(64+x+offset < 90):
                worksheet.write(str(chr(65+x+offset))+str(currLine), str(Mercury_Values[x]), formatTwo)
            else:
                firstLetter = chr(int(65+(((x+offset)-26)/26)))
                secondLetter = chr(65+((x+offset)%26))
                col = firstLetter + secondLetter
                worksheet.write(col+str(currLine), str(Mercury_Values[x]), formatTwo)
        else:
            if(64+x+offset < 90):
                worksheet.write(str(chr(65+x+offset))+str(currLine), str(Mercury_Values[x]), formatOne)
            else:
                firstLetter = chr(int(65+(((x+offset)-26)/26)))
                secondLetter = chr(65+((x+offset)%26))
                col = firstLetter + secondLetter
                worksheet.write(col+str(currLine), str(Mercury_Values[x]), formatOne)
    currLine += 1
    return worksheet, currLine, Mercury_Values

def addPCBValues(worksheet, currLine, sortedGenAreas, rowsInArea, formatOne, formatTwo):
    indexOfPCBS = 0
    for x in range(0, len(sortedGenAreas)):
        if sortedGenAreas[x][1] == "PCBs By SW8082A":
            indexOfPCBS = x
    for x in range(1, len(rowsInArea[indexOfPCBS])):
        rowsInArea[indexOfPCBS][x] = valueRefinerMetals(rowsInArea[indexOfPCBS][x])
        
    rowsInArea[indexOfPCBS] = removeUselessRows(rowsInArea, indexOfPCBS)
    firstLine = currLine
    offset = 2
    
    for x in range(1, len(rowsInArea[indexOfPCBS])):
        for y in range(0, len(rowsInArea[indexOfPCBS][x])):
            if(isNumber(rowsInArea[indexOfPCBS][x][y])):
                if(64+y+offset < 90):
                    worksheet.write(str(chr(65+y+offset))+str(currLine), str(rowsInArea[indexOfPCBS][x][y]), formatTwo)
                else:
                    firstLetter = chr(int(65+(((y+offset)-26)/26)))
                    secondLetter = chr(65+((y+offset)%26))
                    col = firstLetter + secondLetter
                    worksheet.write(col+str(currLine), str(rowsInArea[indexOfPCBS][x][y]), formatTwo)
            else:
                if(64+y+offset < 90):
                    worksheet.write(str(chr(65+y+offset))+str(currLine), rowsInArea[indexOfPCBS][x][y], formatOne)
                else:
                    firstLetter = chr(int(65+(((y+offset)-26)/26)))
                    secondLetter = chr(65+((y+offset)%26))
                    col = firstLetter + secondLetter
                    worksheet.write(col+str(currLine), str(rowsInArea[indexOfPCBS][x][y]), formatOne)
        currLine += 1
    lastLine = currLine - 1
    if(lastLine - firstLine != 0):
        worksheet.merge_range('B'+str(firstLine)+":B"+str(lastLine), 'PCBS', formatOne)
    else:
        worksheet.write('B'+str(firstLine), 'PCBS',formatOne)
    return worksheet, currLine

def addPesticideValues(worksheet, currLine, sortedGenAreas, rowsInArea, formatOne, formatTwo):
    indexOfPesticides = 0
    for x in range(0, len(sortedGenAreas)):
        if sortedGenAreas[x][1] == "Pesticides - Soil By SW8081B":
            indexOfPesticides = x
    for x in range(1, len(rowsInArea[indexOfPesticides])):
        rowsInArea[indexOfPesticides][x] = valueRefinerMetals(rowsInArea[indexOfPesticides][x])

    rowsInArea[indexOfPesticides] = removeUselessRows(rowsInArea, indexOfPesticides)  
    firstLine = currLine
    offset = 2
    for x in range(1, len(rowsInArea[indexOfPesticides])):
        for y in range(0, len(rowsInArea[indexOfPesticides][x])):
            if(isNumber(rowsInArea[indexOfPesticides][x][y])):
                if(64+y+offset < 90):
                    worksheet.write(str(chr(65+y+offset))+str(currLine), str(rowsInArea[indexOfPesticides][x][y]), formatTwo)
                else:
                    firstLetter = chr(int(65+(((y+offset)-26)/26)))
                    secondLetter = chr(65+((y+offset)%26))
                    col = firstLetter + secondLetter
                    worksheet.write(col+str(currLine), str(rowsInArea[indexOfPesticides][x][y]), formatTwo)
            else:
                if(64+y+offset < 90):
                    worksheet.write(str(chr(65+y+offset))+str(currLine), rowsInArea[indexOfPesticides][x][y], formatOne)
                else:
                    firstLetter = chr(int(65+(((y+offset)-26)/26)))
                    secondLetter = chr(65+((y+offset)%26))
                    col = firstLetter + secondLetter
                    worksheet.write(col+str(currLine), str(rowsInArea[indexOfPesticides][x][y]), formatOne)
        currLine += 1
    lastLine = currLine - 1
    if(lastLine - firstLine != 0):
        worksheet.merge_range('B'+str(firstLine)+":B"+str(lastLine), 'Pesticides', formatOne)
    else:
        worksheet.write('B'+str(firstLine), 'Pesticides', formatOne)
    return worksheet, currLine

def addMetalValues(worksheet, currLine, sortedGenAreas, rowsInArea, formatOne, formatTwo):
    indexOfMetals = 0
    for x in range(0, len(sortedGenAreas)):
        if sortedGenAreas[x][1] == "Metals, Total":
            indexOfMetals = x
    for x in range(1, len(rowsInArea[indexOfMetals])):
        rowsInArea[indexOfMetals][x] = valueRefinerMetals(rowsInArea[indexOfMetals][x])

    rowsInArea[indexOfMetals] = removeUselessRows(rowsInArea, indexOfMetals)
    firstLine = currLine
    offset = 2
    worksheet.write('B'+str(currLine), 'Metals, Total')
    for x in range(1, len(rowsInArea[indexOfMetals])):
        if(rowsInArea[indexOfMetals][x][0] != "Mercury"):
            for y in range(0, len(rowsInArea[indexOfMetals][x])):
                if(isNumber(rowsInArea[indexOfMetals][x][y])):
                    if(64+y+offset < 90):
                        worksheet.write(str(chr(65+offset+y))+str(currLine), str(rowsInArea[indexOfMetals][x][y]), formatTwo)
                    else:
                        firstLetter = chr(int(65+(((y+offset)-26)/26)))
                        secondLetter = chr(65+((y+offset)%26))
                        col = firstLetter + secondLetter
                        worksheet.write(col+str(currLine), str(rowsInArea[indexOfMetals][x][y]), formatTwo)
                else:
                    if(64+y+offset < 90):
                        worksheet.write(str(chr(65+y+offset))+str(currLine), rowsInArea[indexOfMetals][x][y], formatOne)
                    else:
                        firstLetter = chr(int(65+(((y+offset)-26)/26)))
                        secondLetter = chr(65+((y+offset)%26))
                        col = firstLetter + secondLetter
                        worksheet.write(col+str(currLine), str(rowsInArea[indexOfMetals][x][y]), formatOne)
            currLine += 1
    lastLine = currLine - 1
    if(lastLine - firstLine != 0):
        worksheet.merge_range('B'+str(firstLine)+":B"+str(lastLine), 'Metals', formatOne)
    else:
        worksheet.write('B'+str(firstLine), 'Metals', formatOne)
            
    return worksheet, currLine

def addCyanideValues(worksheet, currLine, Cyanide_Values_Raw, formatOne, formatTwo):
    Cyanide_Values = valueRefinerMetals(Cyanide_Values_Raw)
    worksheet.write('B'+str(currLine), 'Cyanide', formatOne)
    offset = 2
    for x in range(0, len(Cyanide_Values)):
        if(isNumber(Cyanide_Values[x])):
            if(64+x+offset < 90):
                worksheet.write(str(chr(65+x+offset))+str(currLine), str(Cyanide_Values[x]), formatTwo)
            else:
                firstLetter = chr(int(65+(((x+offset)-26)/26)))
                secondLetter = chr(65+((x+offset)%26))
                col = firstLetter + secondLetter
                worksheet.write(col+str(currLine), str(Cyanide_Values[x]), formatTwo)
        else:
            if(64+x+offset < 90):
                worksheet.write(str(chr(65+x+offset))+str(currLine), str(Cyanide_Values[x]), formatOne)
            else:
                firstLetter = chr(int(65+(((x+offset)-26)/26)))
                secondLetter = chr(65+((x+offset)%26))
                col = firstLetter + secondLetter
                worksheet.write(col+str(currLine), str(Cyanide_Values[x]), formatOne)
    currLine += 1
    return worksheet, currLine, Cyanide_Values

def addSemiVolatileValues(worksheet, currLine, sortedGenAreas, rowsInArea, formatOne, formatTwo):
    indexOfSemiVolatiles = 0
    for x in range(0, len(sortedGenAreas)):
        if sortedGenAreas[x][1] == "Semivolatiles By SW8270D":
            indexOfSemiVolatiles = x
    for x in range(1, len(rowsInArea[indexOfSemiVolatiles])):
        rowsInArea[indexOfSemiVolatiles][x] = valueRefinerMetals(rowsInArea[indexOfSemiVolatiles][x])
        
    rowsInArea[indexOfSemiVolatiles] = removeUselessRows(rowsInArea, indexOfSemiVolatiles)
    firstLine = currLine
    offset = 2
    for x in range(1, len(rowsInArea[indexOfSemiVolatiles])):
        for y in range(0, len(rowsInArea[indexOfSemiVolatiles][x])):
            if(isNumber(rowsInArea[indexOfSemiVolatiles][x][y])):
                if(64+y+offset < 90):
                    worksheet.write(str(chr(65+y+offset))+str(currLine), str(rowsInArea[indexOfSemiVolatiles][x][y]), formatTwo)
                else:
                    firstLetter = chr(int(65+(((y+offset)-26)/26)))
                    secondLetter = chr(65+((y+offset)%26))
                    col = firstLetter + secondLetter
                    worksheet.write(col+str(currLine), str(rowsInArea[indexOfSemiVolatiles][x][y]), formatTwo)
            else:
                if(64+y+offset < 90):
                    worksheet.write(str(chr(65+y+offset))+str(currLine), rowsInArea[indexOfSemiVolatiles][x][y], formatOne)
                else:
                    firstLetter = chr(int(65+(((y+offset)-26)/26)))
                    secondLetter = chr(65+((y+offset)%26))
                    col = firstLetter + secondLetter
                    worksheet.write(col+str(currLine), str(rowsInArea[indexOfSemiVolatiles][x][y]), formatOne)
        currLine += 1
    lastLine = currLine - 1
    if(lastLine - firstLine != 0):
        worksheet.merge_range('B'+str(firstLine)+":B"+str(lastLine), 'SemiVolatiles', formatOne)
    else:
        worksheet.write('B'+str(firstLine), 'SemiVolatiles', formatOne)
    return worksheet, currLine

def addVolatileValues(worksheet, currLine, sortedGenAreas, rowsInArea, formatOne, formatTwo):
    indexOfVolatiles = 0
    for x in range(0, len(sortedGenAreas)):
        if sortedGenAreas[x][1] == "Volatiles (TCL) By SW8260C":
            indexOfVolatiles = x
    for x in range(1, len(rowsInArea[indexOfVolatiles])):
        rowsInArea[indexOfVolatiles][x] = valueRefinerMetals(rowsInArea[indexOfVolatiles][x])

    rowsInArea[indexOfVolatiles] = removeUselessRows(rowsInArea, indexOfVolatiles)
    firstLine = currLine
    offset = 2
    worksheet.write('B'+str(currLine), 'Volatiles (TCL) By SW8260C')
    for x in range(1, len(rowsInArea[indexOfVolatiles])):
        for y in range(0, len(rowsInArea[indexOfVolatiles][x])):
            if(isNumber(rowsInArea[indexOfVolatiles][x][y])):
                if(64+y+offset < 90):
                    worksheet.write(str(chr(65+y+offset))+str(currLine), str(rowsInArea[indexOfVolatiles][x][y]), formatTwo)
                else:
                    firstLetter = chr(int(65+(((y+offset)-26)/26)))
                    secondLetter = chr(65+((y+offset)%26))
                    col = firstLetter + secondLetter
                    worksheet.write(col+str(currLine), str(rowsInArea[indexOfVolatiles][x][y]), formatTwo)
            else:
                if(64+y+offset < 90):
                    worksheet.write(str(chr(65+y+offset))+str(currLine), rowsInArea[indexOfVolatiles][x][y], formatOne)
                else:
                    firstLetter = chr(int(65+(((y+offset)-26)/26)))
                    secondLetter = chr(65+((y+offset)%26))
                    col = firstLetter + secondLetter
                    worksheet.write(col+str(currLine), str(rowsInArea[indexOfVolatiles][x][y]), formatOne)
        currLine += 1
    lastLine = currLine - 1
    if(lastLine - firstLine != 0):
        worksheet.merge_range('B'+str(firstLine)+":B"+str(lastLine), 'Volatiles', formatOne)
    else:
        worksheet.write('B'+str(firstLine), 'Volatiles', formatOne)
    return worksheet, currLine

def createSecondHeading(worksheet, currLine, Type_Samples_Collected, formatOne):
    worksheet.set_row(currLine-1,50)
    worksheet.write('B'+str(currLine), 'RCRA Characteristics ', formatOne)
    worksheet.merge_range('C'+str(currLine)+':E'+str(currLine), 'Regulatory Criteria', formatOne)
    
    offset = 4
    
    for x in range(1,len(Type_Samples_Collected)):
        if(64+x+offset < 90):
            worksheet.write(str(chr(65+x+offset))+str(currLine), str(Type_Samples_Collected[x]), formatOne)
        else:
            firstLetter = chr(int(65+(((x+offset)-26)/26)))
            secondLetter = chr(65+((x+offset)%26))
            col = firstLetter + secondLetter
            worksheet.write(col+str(currLine), str(Type_Samples_Collected[x]), formatOne)
    currLine += 1
    return worksheet, currLine

def addCorrosivityValues(worksheet, currLine, Corrosivity_Values_Raw, formatOne):
    Corrosivity_Values = valueRefinerMetals(Corrosivity_Values_Raw)
    worksheet.write('B'+str(currLine), 'Corrosivity', formatOne)
    offset = 2
    for x in range(0,len(Corrosivity_Values)):
        if(64+x+offset < 90):
            worksheet.write(str(chr(65+x+offset))+str(currLine), str(Corrosivity_Values[x]), formatOne)
        else:
            firstLetter = chr(int(65+(((x+offset)-26)/26)))
            secondLetter = chr(65+((x+offset)%26))
            col = firstLetter + secondLetter
            worksheet.write(col+str(currLine), str(Corrosivity_Values[x]), formatOne)
    currLine += 1
    return worksheet, currLine, Corrosivity_Values

def addFlashpointValues(worksheet, currLine, Flashpoint_Values_Raw, formastOne):
    Flashpoint_Values = []
    pos = 0
    for value in Flashpoint_Values_Raw:
        if(pos == 0):
            Flashpoint_Values.append(value.value)
            Flashpoint_Values.append(" ")
            Flashpoint_Values.append("Degree F")
            Flashpoint_Values.append(">200 Degree F")
        if((pos >= firstIndex and pos%2 == firstIndex%2)):
            Flashpoint_Values.append(value.value)
        pos+=1

    offset = 1

    for x in range(0,len(Flashpoint_Values)):
        if(64+x+offset < 90):
            worksheet.write(str(chr(65+x+offset))+str(currLine), str(Flashpoint_Values[x]), formatOne)
        else:
            firstLetter = chr(int(65+(((x+offset)-26)/26)))
            secondLetter = chr(65+((x+offset)%26))
            col = firstLetter + secondLetter
            worksheet.write(col+str(currLine), str(Flashpoint_Values[x]), formatOne)
    currLine += 1
    return worksheet, currLine, Flashpoint_Values

def addIgnitabilityValues(worksheet, currLine, Ignitability_Values_Raw, formatOne):
    Ignitability_Values = []
    pos = 0
    for value in Ignitability_Values_Raw:
        if(pos == 0):
            Ignitability_Values.append(value.value)
            Ignitability_Values.append(" ")
            Ignitability_Values.append("Degree F")
            Ignitability_Values.append("<140 Degree F")
        if((pos >= firstIndex and pos%2 == firstIndex%2)):
            Ignitability_Values.append(value.value)
        pos+=1

    offset = 1
    for x in range(0,len(Ignitability_Values)):
        if(64+x+offset < 90):
            worksheet.write(str(chr(65+x+offset))+str(currLine), str(Ignitability_Values[x]), formatOne)
        else:
            firstLetter = chr(int(65+(((x+offset)-26)/26)))
            secondLetter = chr(65+((x+offset)%26))
            col = firstLetter + secondLetter
            worksheet.write(col+str(currLine), str(Ignitability_Values[x]), formatOne)
    currLine += 1
    return worksheet, currLine, Ignitability_Values

def addReactivityValues(worksheet, currLine, Reactivity_Values_Cyanide_Raw, Reactivity_Values_Sulfide_Raw, formatOne):
    Reactivity_Values_Cyanide = valueRefinerMetals(Reactivity_Values_Cyanide_Raw)
    worksheet.merge_range('B'+str(currLine)+":B"+str(currLine+1), 'Reactivity', formatOne)
    worksheet.write('C'+str(currLine), 'Cyanide', formatOne)

    offset = 2
    for x in range(1,len(Reactivity_Values_Cyanide)):
        if(64+x+offset < 90):
            worksheet.write(str(chr(65+x+offset))+str(currLine), str(Reactivity_Values_Cyanide[x]), formatOne)
        else:
            firstLetter = chr(int(65+(((x+offset)-26)/26)))
            secondLetter = chr(65+((x+offset)%26))
            col = firstLetter + secondLetter
            worksheet.write(col+str(currLine), str(Reactivity_Values_Cyanide[x]), formatOne)
    currLine += 1
    
    Reactivity_Values_Sulfide = valueRefinerMetals(Reactivity_Values_Sulfide_Raw)
    worksheet.write('C'+str(currLine), 'Sulfide', formatOne)
    for x in range(1,len(Reactivity_Values_Sulfide)):
        if(64+x+offset < 90):
            worksheet.write(str(chr(65+x+offset))+str(currLine), str(Reactivity_Values_Sulfide[x]), formatOne)
        else:
            firstLetter = chr(int(65+(((x+offset)-26)/26)))
            secondLetter = chr(65+((x+offset)%26))
            col = firstLetter + secondLetter
            worksheet.write(col+str(currLine), str(Reactivity_Values_Sulfide[x]), formatOne)
    currLine += 1
    return worksheet, currLine, Reactivity_Values_Cyanide, Reactivity_Values_Sulfide

def createThirdHeading(worksheet, currLine, Type_Samples_Collected, formatOne):
    worksheet.set_row(currLine-1,50)
    worksheet.write('B'+str(currLine), 'Toxicity ', formatOne)
    worksheet.merge_range('C'+str(currLine)+':E'+str(currLine), 'TCLP Regulatory Criteria', formatOne)

    x = len(Type_Samples_Collected)
    offset = 4

    finalLetter=""
    
    if 64+x+offset > 90:
        firstLetter = chr(int(65+(((x+offset)-26)/26)))
        print(firstLetter)
        secondLetter = chr(64+((x+offset)%26))
        print(secondLetter)
        finalLetter = firstLetter+secondLetter
    else:
        finalLetter = chr(64+x+offset)
        
    worksheet.merge_range("F"+str(currLine)+":"+finalLetter+str(currLine), "", formatOne)
    currLine += 1
    return worksheet, currLine

def addTCLPMetalValues(worksheet, currLine, sortedGenAreas, rowsInArea, formatOne, formatTwo):
    indexOfTCLPMetals = 0
    for x in range(0, len(sortedGenAreas)):
        if sortedGenAreas[x][1] == "Metals, TCLP":
            indexOfTCLPMetals = x
    for x in range(1, len(rowsInArea[indexOfTCLPMetals])):
        rowsInArea[indexOfTCLPMetals][x] = valueRefinerMetals(rowsInArea[indexOfTCLPMetals][x])
    rowsInArea[indexOfTCLPMetals] = removeUselessRows(rowsInArea, indexOfTCLPMetals)
    firstLine = currLine
    offset = 2
    for x in range(1, len(rowsInArea[indexOfTCLPMetals])):
        for y in range(0, len(rowsInArea[indexOfTCLPMetals][x])):
            if(isNumber(rowsInArea[indexOfTCLPMetals][x][y])):
                if(64+y+offset < 90):
                    worksheet.write(str(chr(65+y+offset))+str(currLine), str(rowsInArea[indexOfTCLPMetals][x][y]), formatTwo)
                else:
                    firstLetter = chr(int(65+(((y+offset)-26)/26)))
                    secondLetter = chr(65+((y+offset)%26))
                    col = firstLetter + secondLetter
                    worksheet.write(col+str(currLine), str(rowsInArea[indexOfTCLPMetals][x][y]), formatTwo)
            else:
                if(64+y+offset < 90):
                    worksheet.write(str(chr(65+y+offset))+str(currLine), rowsInArea[indexOfTCLPMetals][x][y], formatOne)
                else:
                    firstLetter = chr(int(65+(((y+offset)-26)/26)))
                    secondLetter = chr(65+((y+offset)%26))
                    col = firstLetter + secondLetter
                    worksheet.write(col+str(currLine), str(rowsInArea[indexOfTCLPMetals][x][y]), formatOne)
        currLine += 1
    lastLine = currLine - 1
    if(lastLine - firstLine != 0):
        worksheet.merge_range('B'+str(firstLine)+":B"+str(lastLine), 'TCLP Metals', formatOne)
    else:
        worksheet.write('B'+str(firstLine), 'TCLP Metals', formatOne)
    return worksheet, currLine

def addVOCSValues(worksheet, currLine, sortedGenAreas, rowsInArea, formatOne, formatTwo):
    indexOfVOCS = 0
    for x in range(0, len(sortedGenAreas)):
        if sortedGenAreas[x][1] == "TCLP Volatiles By SW8260C":
            indexOfVOCS = x
    for x in range(1, len(rowsInArea[indexOfVOCS])):
        rowsInArea[indexOfVOCS][x] = valueRefinerMetals(rowsInArea[indexOfVOCS][x])

    rowsInArea[indexOfVOCS] = removeUselessRows(rowsInArea, indexOfVOCS)
    firstLine = currLine
    offset = 2
    for x in range(1, len(rowsInArea[indexOfVOCS])):
        for y in range(0, len(rowsInArea[indexOfVOCS][x])):
            if(isNumber(rowsInArea[indexOfVOCS][x][y])):
                if(64+y+offset < 90):
                    worksheet.write(str(chr(65+y+offset))+str(currLine), str(rowsInArea[indexOfVOCS][x][y]), formatTwo)
                else:
                    firstLetter = chr(int(65+(((y+offset)-26)/26)))
                    secondLetter = chr(65+((y+offset)%26))
                    col = firstLetter + secondLetter
                    worksheet.write(col+str(currLine), str(rowsInArea[indexOfVOCS][x][y]), formatTwo)
            else:
                if(64+y+offset < 90):
                    worksheet.write(str(chr(65+y+offset))+str(currLine), rowsInArea[indexOfVOCS][x][y], formatOne)
                else:
                    firstLetter = chr(int(65+(((y+offset)-26)/26)))
                    secondLetter = chr(65+((y+offset)%26))
                    col = firstLetter + secondLetter
                    worksheet.write(col+str(currLine), str(rowsInArea[indexOfVOCS][x][y]), formatOne)
        currLine += 1
    lastLine = currLine - 1
    if(lastLine - firstLine != 0):
        worksheet.merge_range('B'+str(firstLine)+":B"+str(lastLine), 'TCLP Vocs', formatOne)
    else:
        worksheet.write('B'+str(firstLine), 'TCLP Vocs', formatOne)
    return worksheet, currLine

def addSVOCSValues(worksheet, currLine, sortedGenAreas, rowsInArea, formatOne):
    indexOfSVOCS = 0
    for x in range(0, len(sortedGenAreas)):
        if sortedGenAreas[x][1] == "TCLP Acid/Base-Neutral By SW8270D":
            indexOfSVOCS = x
    for x in range(1, len(rowsInArea[indexOfSVOCS])):
        rowsInArea[indexOfSVOCS][x] = valueRefinerMetals(rowsInArea[indexOfSVOCS][x])

    rowsInArea[indexOfSVOCS] = removeUselessRows(rowsInArea, indexOfSVOCS)
    firstLine = currLine
    offset = 2
    for x in range(1, len(rowsInArea[indexOfSVOCS])):
        for y in range(0, len(rowsInArea[indexOfSVOCS][x])):
            if(isNumber(rowsInArea[indexOfSVOCS][x][y])):
                if(64+y+offset < 90):
                    worksheet.write(str(chr(65+y+offset))+str(currLine), str(rowsInArea[indexOfSVOCS][x][y]), formatTwo)
                else:
                    firstLetter = chr(int(65+(((y+offset)-26)/26)))
                    secondLetter = chr(65+((y+offset)%26))
                    col = firstLetter + secondLetter
                    worksheet.write(col+str(currLine), str(rowsInArea[indexOfSVOCS][x][y]), formatTwo)
            else:
                if(64+y+offset < 90):
                    worksheet.write(str(chr(65+y+offset))+str(currLine), rowsInArea[indexOfSVOCS][x][y], formatOne)
                else:
                    firstLetter = chr(int(65+(((y+offset)-26)/26)))
                    secondLetter = chr(65+((y+offset)%26))
                    col = firstLetter + secondLetter
                    worksheet.write(col+str(currLine), str(rowsInArea[indexOfSVOCS][x][y]), formatOne)
        currLine += 1
    lastLine = currLine - 1
    if(lastLine - firstLine != 0):
        worksheet.merge_range('B'+str(firstLine)+":B"+str(lastLine), 'TCLP SVocs', formatOne)
    else:
        worksheet.write('B'+str(firstLine), 'TCLP SVocs', formatOne)
    return worksheet, currLine

def addTCLPPesticideValues(worksheet, currLine, sortedGenAreas, rowsInArea, formatOne):
    indexOfTCLPPesticides = 0
    for x in range(0, len(sortedGenAreas)):
        if sortedGenAreas[x][1] == "TCLP Pesticides By SW8081B":
            indexOfTCLPPesticides = x
    for x in range(1, len(rowsInArea[indexOfTCLPPesticides])):
        rowsInArea[indexOfTCLPPesticides][x] = valueRefinerMetals(rowsInArea[indexOfTCLPPesticides][x])

    rowsInArea[indexOfTCLPPesticides] = removeUselessRows(rowsInArea, indexOfTCLPPesticides)
    firstLine = currLine
    offset = 2

    for x in range(1, len(rowsInArea[indexOfTCLPPesticides])):
        for y in range(0, len(rowsInArea[indexOfTCLPPesticides][x])):
            if(isNumber(rowsInArea[indexOfTCLPPesticides][x][y])):
                if(64+y+offset < 90):
                    worksheet.write(str(chr(65+y+offset))+str(currLine), str(rowsInArea[indexOfTCLPPesticides][x][y]), formatTwo)
                else:
                    firstLetter = chr(int(65+(((y+offset)-26)/26)))
                    secondLetter = chr(65+((y+offset)%26))
                    col = firstLetter + secondLetter
                    worksheet.write(col+str(currLine), str(rowsInArea[indexOfTCLPPesticides][x][y]), formatTwo)
            else:
                if(64+y+offset < 90):
                    worksheet.write(str(chr(65+y+offset))+str(currLine), rowsInArea[indexOfTCLPPesticides][x][y], formatOne)
                else:
                    firstLetter = chr(int(65+(((y+offset)-26)/26)))
                    secondLetter = chr(65+((y+offset)%26))
                    col = firstLetter + secondLetter
                    worksheet.write(col+str(currLine), str(rowsInArea[indexOfTCLPPesticides][x][y]), formatOne)
        currLine += 1
    lastLine = currLine - 1
    if(lastLine - firstLine != 0):
        worksheet.merge_range('B'+str(firstLine)+":B"+str(lastLine), 'TCLP Pesticides', formatOne)
    else:
        worksheet.write('B'+str(firstLine), 'TCLP Pesticides', formatOne)
    return worksheet, currLine

def addTCLPHerbicideValues(worksheet, currLine, sortedGenAreas, rowsInArea, formatOne):
    indexOfTCLPHerbicides = 0
    for x in range(0, len(sortedGenAreas)):
        if sortedGenAreas[x][1] == "TCLP Herbicides By SW8151A":
            indexOfHerbicides = x
    for x in range(1, len(rowsInArea[indexOfHerbicides])):
        rowsInArea[indexOfHerbicides][x] = valueRefinerMetals(rowsInArea[indexOfHerbicides][x])

    rowsInArea[indexOfTCLPHerbicides] = removeUselessRows(rowsInArea, indexOfTCLPHerbicides)
    firstLine = currLine
    offset = 2
    for x in range(1, len(rowsInArea[indexOfHerbicides])):
        for y in range(0, len(rowsInArea[indexOfHerbicides][x])):
            if(isNumber(rowsInArea[indexOfHerbicides][x][y])):
                if(64+y+offset < 90):
                    worksheet.write(str(chr(65+y+offset))+str(currLine), str(rowsInArea[indexOfHerbicides][x][y]), formatTwo)
                else:
                    firstLetter = chr(int(65+(((y+offset)-26)/26)))
                    secondLetter = chr(65+((y+offset)%26))
                    col = firstLetter + secondLetter
                    worksheet.write(col+str(currLine), str(rowsInArea[indexOfHerbicides][x][y]), formatTwo)
            else:
                if(64+y+offset < 90):
                    worksheet.write(str(chr(65+y+offset))+str(currLine), rowsInArea[indexOfHerbicides][x][y], formatOne)
                else:
                    firstLetter = chr(int(65+(((y+offset)-26)/26)))
                    secondLetter = chr(65+((y+offset)%26))
                    col = firstLetter + secondLetter
                    worksheet.write(col+str(currLine), str(rowsInArea[indexOfHerbicides][x][y]), formatOne)
        currLine += 1
    lastLine = currLine - 1
    if(lastLine - firstLine != 0):
        worksheet.merge_range('B'+str(firstLine)+":B"+str(lastLine), 'TCLP Pesticides / Herbicides', formatOne)
    else:
        worksheet.write('B'+str(firstLine), 'TCLP Pesticides / Herbicides', formatOne)
    return worksheet, currLine

def addTPHValues(worksheet, currLine, sortedGenAreas, rowsInArea, formatOne):
    indexOfGasolineHydrocarbons = 0
    for x in range(0, len(sortedGenAreas)):
        if sortedGenAreas[x][1] == "Gasoline Range Hydrocarbons (C6-C10) By SW8015D":
            indexOfGasolineHydrocarbons = x
    for x in range(1, len(rowsInArea[indexOfGasolineHydrocarbons])):
        rowsInArea[indexOfGasolineHydrocarbons][x] = valueRefinerMetals(rowsInArea[indexOfGasolineHydrocarbons][x])

    indexOfDieselHydrocarbons = 0
    for x in range(0, len(sortedGenAreas)):
        if sortedGenAreas[x][1] == "TPH By SW8015D DRO":
            indexOfDieselHydrocarbons = x
    for x in range(1, len(rowsInArea[indexOfDieselHydrocarbons])):
        rowsInArea[indexOfDieselHydrocarbons][x] = valueRefinerMetals(rowsInArea[indexOfDieselHydrocarbons][x])

    offset = 2

    worksheet.merge_range('B'+str(currLine)+":B"+str(currLine+1), 'Total Petroleum Hydrocarbons', formatOne)
    for x in range(1, len(rowsInArea[indexOfGasolineHydrocarbons])):
        for y in range(0, len(rowsInArea[indexOfGasolineHydrocarbons][x])):
            if(isNumber(rowsInArea[indexOfGasolineHydrocarbons][x][y])):
                if(64+y+offset < 90):
                    worksheet.write(str(chr(65+y+offset))+str(currLine), str(rowsInArea[indexOfGasolineHydrocarbons][x][y]), formatTwo)
                else:
                    firstLetter = chr(int(65+(((y+offset)-26)/26)))
                    secondLetter = chr(65+((y+offset)%26))
                    col = firstLetter + secondLetter
                    worksheet.write(col+str(currLine), str(rowsInArea[indexOfGasolineHydrocarbons][x][y]), formatTwo)
            else:
                if(64+y+offset < 90):
                    worksheet.write(str(chr(65+y+offset))+str(currLine), rowsInArea[indexOfGasolineHydrocarbons][x][y], formatOne)
                else:
                    firstLetter = chr(int(65+(((y+offset)-26)/26)))
                    secondLetter = chr(65+((y+offset)%26))
                    col = firstLetter + secondLetter
                    worksheet.write(col+str(currLine), str(rowsInArea[indexOfGasolineHydrocarbons][x][y]), formatOne)
        currLine += 1
    for x in range(1, len(rowsInArea[indexOfDieselHydrocarbons])):
        for y in range(0, len(rowsInArea[indexOfDieselHydrocarbons][x])):
            if(isNumber(rowsInArea[indexOfDieselHydrocarbons][x][y])):
                if(64+y+offset < 90):
                    worksheet.write(str(chr(65+y+offset))+str(currLine), str(rowsInArea[indexOfDieselHydrocarbons][x][y]), formatTwo)
                else:
                    firstLetter = chr(int(65+(((y+offset)-26)/26)))
                    secondLetter = chr(65+((y+offset)%26))
                    col = firstLetter + secondLetter
                    worksheet.write(col+str(currLine), str(rowsInArea[indexOfDieselHydrocarbons][x][y]), formatTwo)
            else:
                if(64+y+offset < 90):
                    worksheet.write(str(chr(65+y+offset))+str(currLine), rowsInArea[indexOfDieselHydrocarbons][x][y], formatOne)
                else:
                    firstLetter = chr(int(65+(((y+offset)-26)/26)))
                    secondLetter = chr(65+((y+offset)%26))
                    col = firstLetter + secondLetter
                    worksheet.write(col+str(currLine), str(rowsInArea[indexOfDieselHydrocarbons][x][y]), formatOne)
        currLine += 1
    return worksheet, currLine


print("Writing to Excel File...")
workbook, worksheet = startWritingFinalFile()

worksheet.set_column('B:B', 25)
worksheet.set_column('C:C', 30)
worksheet.set_column('E:E', 15)
worksheet.set_row(5,50)

#Important Information - Titles, etc..
formatOne = workbook.add_format()
formatOne.set_align('center')
formatOne.set_align('vcenter')
formatOne.set_font_name('Arial')
formatOne.set_font_size('12')
formatOne.set_border(6)

#Numbers within the text
formatTwo = workbook.add_format()
formatTwo.set_align('center')
formatTwo.set_align('vcenter')
formatTwo.set_font_name('Arial')
formatTwo.set_font_size('12')
formatTwo.set_border(6)
formatTwo.set_bg_color('#87CEFF')
formatTwo.set_bold()


#Current Line to overwrite each process
currLine = 6

#Heading for each column
worksheet, currLine, Type_Samples_Collected = createHeading(worksheet, currLine, Type_Samples_Collected_Raw, formatOne)

#Adding Mercury Values
worksheet, currLine, Mercury_Values = addMercuryValues(worksheet, currLine, Mercury_Values_Raw, formatOne, formatTwo)

#Adding PCB Values
worksheet, currLine = addPCBValues(worksheet, currLine, sortedGenAreas, rowsInArea, formatOne, formatTwo)
    
#Adding Pesticide Values
worksheet, currLine = addPesticideValues(worksheet, currLine, sortedGenAreas, rowsInArea, formatOne, formatTwo)

#Adding Metal Values
worksheet, currLine = addMetalValues(worksheet, currLine, sortedGenAreas, rowsInArea, formatOne, formatTwo)

#Adding Cyanide Values
worksheet, currLine, Cyanide_Values = addCyanideValues(worksheet, currLine, Cyanide_Values_Raw, formatOne, formatTwo)
    
#Adding Semi Volatile Organic Compounds
worksheet, currLine = addSemiVolatileValues(worksheet, currLine, sortedGenAreas, rowsInArea, formatOne, formatTwo)

#Adding Volatile Organic Compounds
worksheet, currLine = addVolatileValues(worksheet, currLine, sortedGenAreas, rowsInArea, formatOne, formatTwo)

#RCRA Second Heading
worksheet, currLine = createSecondHeading(worksheet, currLine, Type_Samples_Collected, formatOne)

#Adding Corrosivity(pH) Values
worksheet, currLine, Corrosivity_Values = addCorrosivityValues(worksheet, currLine, Corrosivity_Values_Raw, formatOne)

#Adding Flashpoint Values
worksheet, currLine, Flashpoint_Values = addFlashpointValues(worksheet, currLine, Flashpoint_Values_Raw, formatOne)

#Adding Ignitability Values
worksheet, currLine, Ignitability_Values = addIgnitabilityValues(worksheet, currLine, Ignitability_Values_Raw, formatOne)

#Adding Reactivity Values
worksheet, currLine, Reactivity_Values_Cyanide, Reactivity_Values_Sulfide = addReactivityValues(worksheet, currLine, Reactivity_Values_Cyanide_Raw, Reactivity_Values_Sulfide_Raw, formatOne)

#Toxicity Third Heading   
worksheet, currLine = createThirdHeading(worksheet, currLine, Type_Samples_Collected, formatOne)

#Adding TCLP Metals(Barium / Lead)
worksheet, currLine = addTCLPMetalValues(worksheet, currLine, sortedGenAreas, rowsInArea, formatOne, formatTwo)

#Adding TCLP VOCS
worksheet, currLine = addVOCSValues(worksheet, currLine, sortedGenAreas, rowsInArea, formatOne, formatTwo)

#Adding TCLP SVOCS
worksheet, currLine = addSVOCSValues(worksheet, currLine, sortedGenAreas, rowsInArea, formatOne)

#Adding TCLP Pesticides
worksheet, currLine = addTCLPPesticideValues(worksheet, currLine, sortedGenAreas, rowsInArea, formatOne)

#Adding TCLP Herbicides
worksheet, currLine = addTCLPHerbicideValues(worksheet, currLine, sortedGenAreas, rowsInArea, formatOne)

#Adding Total Petroleum Hydrocarbons
worksheet, currLine = addTPHValues(worksheet, currLine, sortedGenAreas, rowsInArea, formatOne)

#Beginning information(Company Name, Address, Dates Samples were collected)
worksheet, currLine = createBeginning(worksheet, currLine)

workbook.close()
print("Done Writing")





import xlrd
from xlrd import open_workbook
import xlsxwriter
from plantManagement import *
from excelHelpers import *

# Plants in Excel is a resource for any non-specific file interaction.
# This file is responsible for all of the save and load mechanics in GC SAM.
# This file should not be changed unless the plant management system has new values added to it.

#Only the selectedLines are saved using SaveToFile
#selectedLines should be a vector of strings
#Each entry in selectedLines should be a valid line name within lineManager
def SaveToFile(lineManager,xlsxName,selectedLines):
    selectedLinesManager = LineManager()
    for l in selectedLines:
        selectedLines.AddLine(lineManager.GetLine(l))
    SaveToFile(selectedLinesManager,xlsxName)

#Write an xlsx file with all data within lineManger.
def SaveToFile(lineManager,xlsxName):
    xlFile = xlsxwriter.Workbook(xlsxName)
    xlOut = xlFile.add_worksheet('Saved Lines')
    
    #row/column pos
    r = 0
    startC = 0
    c = startC

    lineFields = 4
    plantFields = 10
    FAMEFields = 11

    xlOut.write(r,c,"Line Name")
    xlOut.write(r,c+1,"Line Number")
    xlOut.write(r,c+2,"Line Color")
    xlOut.write(r,c+3,"Line Unique ID")
    xlOut.write(r,c+lineFields,"Plant Name")
    xlOut.write(r,c+lineFields+1,"Plant Number")
    xlOut.write(r,c+lineFields+2,"Plant Color")
    xlOut.write(r,c+lineFields+3,"Plant Unique ID")
    xlOut.write(r,c+lineFields+4,"Western Rank")
    xlOut.write(r,c+lineFields+5,"Western Value")
    xlOut.write(r,c+lineFields+6,"Zygosity")
    xlOut.write(r,c+lineFields+7,"mgFA Dry Weight")
    xlOut.write(r,c+lineFields+8,"mg FAMEs Std")
    xlOut.write(r,c+lineFields+9,"Weight Fraction")
    xlOut.write(r,c+lineFields+plantFields,"FAME Name")
    xlOut.write(r,c+lineFields+plantFields+1,"FAME Number")
    xlOut.write(r,c+lineFields+plantFields+2,"FAME Color")
    xlOut.write(r,c+lineFields+plantFields+3,"FAME Unique ID")
    xlOut.write(r,c+lineFields+plantFields+4,"Retention Time")
    xlOut.write(r,c+lineFields+plantFields+5,"Peak Area")
    xlOut.write(r,c+lineFields+plantFields+6,"Percent FA")
    xlOut.write(r,c+lineFields+plantFields+7,"Percent of Total FA")
    xlOut.write(r,c+lineFields+plantFields+8,"ugFA")
    xlOut.write(r,c+lineFields+plantFields+9,"Name Match Discrepancy")
    xlOut.write(r,c+lineFields+plantFields+10,"Best Match")

    for l in lineManager.m_lines:
        r+= 1
        c = startC
        xlOut.write(r,c,l.m_name)
        xlOut.write(r,c+1,l.m_number)
        xlOut.write(r,c+2,l.m_colorID)
        xlOut.write(r,c+3,l.m_uniqueID)
        for p in l.m_plants:
            r+= 1
            xlOut.write(r,c+lineFields,p.m_name)
            xlOut.write(r,c+lineFields+1,p.m_number)
            xlOut.write(r,c+lineFields+2,p.m_colorID)
            xlOut.write(r,c+lineFields+3,p.m_uniqueID)
            xlOut.write(r,c+lineFields+4,p.m_expressionLevel.m_westernRank)
            xlOut.write(r,c+lineFields+5,p.m_expressionLevel.m_westernValue)
            xlOut.write(r,c+lineFields+6,p.m_zygosity)
            xlOut.write(r,c+lineFields+7,p.m_mgFADW)
            xlOut.write(r,c+lineFields+8,p.m_mgStd)
            xlOut.write(r,c+lineFields+9,p.m_totalWeightFraction)
            for f in p.m_fames:
                r+= 1
                xlOut.write(r,c+lineFields+plantFields, f.m_name)
                xlOut.write(r,c+lineFields+plantFields+1, f.m_number)
                xlOut.write(r,c+lineFields+plantFields+2, f.m_colorID)
                xlOut.write(r,c+lineFields+plantFields+3, f.m_uniqueID)
                xlOut.write(r,c+lineFields+plantFields+4, f.m_retentionTime)
                xlOut.write(r,c+lineFields+plantFields+5, f.m_peakArea)
                xlOut.write(r,c+lineFields+plantFields+6, f.m_percentFA)
                xlOut.write(r,c+lineFields+plantFields+7, f.m_percentOfTotalFA)
                xlOut.write(r,c+lineFields+plantFields+8, f.m_ugFA)
                xlOut.write(r,c+lineFields+plantFields+9, f.m_nameMatchDiscrepancy)
                if(f.m_bestMatch):
                    xlOut.write(r,c+lineFields+plantFields+10, "T")
                else:
                    xlOut.write(r,c+lineFields+plantFields+10, "F")

    xlFile.close()

    print("Saved data to",xlsxName)

#Read in the data structure of a file generated with SaveToFile
#RETURNS: a LineManager with all available data in the given file.
def LoadFromFile(xlsxName):

    ret = LineManager()

    #FIXME: duplicate code
    lineFields = 4
    plantFields = 10
    FAMEFields = 11

    r = 0
    startC = 0
    c = startC

    #TODO: figure out forward declarations in python
    currentLine = Line()
    currentPlant = Plant()
    currentFAME = FAME()

    print("Loading",xlsxName)
    # print("Minimum rows for a line:",lineFields)
    # print("Minimum rows for a plant:",lineFields+plantFields)
    # print("Minimum rows for a FAME:",lineFields+plantFields+FAMEFields)
    wb = open_workbook(xlsxName)

    for sheet in wb.sheets():
        print("Reading in",sheet.name,sheet.nrows,"x",sheet.ncols)
        for r in range(sheet.nrows):
            #Read a line
            if(sheet.cell_type(r,c) != xlrd.XL_CELL_EMPTY and (c+lineFields == sheet.ncols or sheet.cell_type(r,c+lineFields) == xlrd.XL_CELL_EMPTY) ):
                currentLine = Line()
                currentLine.m_name = sheet.cell(r,c).value
                currentLine.m_number = sheet.cell(r,c+1).value
                currentLine.m_colorID = sheet.cell(r,c+2).value
                currentLine.m_uniqueID = sheet.cell(r,c+3).value
                ret.AddLine(currentLine)

            #Read a plant
            elif(sheet.cell_type(r,c+lineFields) != xlrd.XL_CELL_EMPTY and (c+lineFields+plantFields == sheet.ncols or sheet.cell_type(r,c+lineFields+plantFields) == xlrd.XL_CELL_EMPTY) ):
                currentPlant = Plant()
                currentPlant.m_name = sheet.cell(r,c+lineFields).value
                currentPlant.m_number = sheet.cell(r,c+lineFields+1).value
                currentPlant.m_colorID = sheet.cell(r,c+lineFields+2).value
                currentPlant.m_uniqueID = sheet.cell(r,c+lineFields+3).value
                currentPlant.m_expressionLevel.m_westernRank = sheet.cell(r,c+lineFields+4).value
                currentPlant.m_expressionLevel.m_westernValue = sheet.cell(r,c+lineFields+5).value
                currentPlant.m_zygosity = sheet.cell(r,c+lineFields+6).value
                currentPlant.m_mgFADW = sheet.cell(r,c+lineFields+7).value
                currentPlant.m_mgStd = sheet.cell(r,c+lineFields+8).value
                currentPlant.m_totalWeightFraction = sheet.cell(r,c+lineFields+9).value
                currentLine.AddPlant(currentPlant)

            #Read a FAME
            elif(sheet.cell_type(r,c+lineFields+plantFields) != xlrd.XL_CELL_EMPTY and (c+lineFields+plantFields+FAMEFields == sheet.ncols or sheet.cell_type(r,c+lineFields+plantFields+FAMEFields) == xlrd.XL_CELL_EMPTY) ):
                currentFAME = FAME()
                currentFAME.m_name = sheet.cell(r,c+lineFields+plantFields).value
                currentFAME.m_number = sheet.cell(r,c+lineFields+plantFields+1).value
                currentFAME.m_colorID = sheet.cell(r,c+lineFields+plantFields+2).value
                currentFAME.m_uniqueID = sheet.cell(r,c+lineFields+plantFields+3).value
                currentFAME.m_retentionTime = sheet.cell(r,c+lineFields+plantFields+4).value
                currentFAME.m_peakArea = sheet.cell(r,c+lineFields+plantFields+5).value
                currentFAME.m_percentFA = sheet.cell(r,c+lineFields+plantFields+6).value
                currentFAME.m_percentOfTotalFA = sheet.cell(r,c+lineFields+plantFields+7).value
                currentFAME.m_ugFA = sheet.cell(r,c+lineFields+plantFields+8).value
                currentFAME.m_nameMatchDiscrepancy = sheet.cell(r,c+lineFields+plantFields+9).value
                bestMatch = sheet.cell(r,c+lineFields+plantFields+10).value
                if(bestMatch == "T"):
                    currentFAME.m_bestMatch = True
                if(bestMatch == "F"):
                    currentFAME.m_bestMatch = False
                currentPlant.AddFAME(currentFAME)

            else:
                print("Ignoring row",r,": data format not accepted.")

    return ret

#Write an xlsx file with all data within lineManger.
def Export(lineManager,xlsxName):
    xlFile = xlsxwriter.Workbook(xlsxName)
    xlOut = xlFile.add_worksheet('Exported Lines')

    r = 0

    for index, val in enumerate([
        "Line Name",
        "Plant Name",
        "Total FA",
        '18:2/3 Ratio',
        "C16:0",
        "C16:1",
        "C18:0",
        "C18:1",
        "C18:2",
        "C18:3",
        "C20:0",
        "C22:0",
        "C22:1",
        "C24:0"]
        ):
        xlOut.write(r, index, val)
    r += 1

    for l in lineManager.m_lines:
        for p in l.m_plants:
            for index, val in enumerate([
                l.m_name,
                p.m_name,
                p.m_totalWeightFraction * 100,
                Get18_2_to_3(p),
                p.GetFAME("C16:0").m_percentFA,
                p.GetFAME("C16:1").m_percentFA,
                p.GetFAME("C18:0").m_percentFA,
                p.GetFAME("C18:1").m_percentFA,
                p.GetFAME("C18:2").m_percentFA,
                p.GetFAME("C18:3").m_percentFA,
                p.GetFAME("C20:0").m_percentFA,
                p.GetFAME("C22:0").m_percentFA,
                p.GetFAME("C22:1").m_percentFA,
                p.GetFAME("C24:0").m_percentFA]
                ):
                xlOut.write(r, index, val)
            r += 1

    xlFile.close()

    print("Exported data to",xlsxName)
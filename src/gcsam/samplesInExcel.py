import xlrd
from xlrd import open_workbook
import xlsxwriter
from sampleManagement import *
from excelHelpers import *

# Samples in Excel is a resource for any non-specific file interaction.
# This file is responsible for all of the save and load mechanics in GC SAM.
# This file should not be changed unless the sample management system has new values added to it.

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
    sampleFields = 10
    FAMEFields = 11

    xlOut.write(r,c,"Line Name")
    xlOut.write(r,c+1,"Line Number")
    xlOut.write(r,c+2,"Line Color")
    xlOut.write(r,c+3,"Line Unique ID")
    xlOut.write(r,c+lineFields,"Sample Name")
    xlOut.write(r,c+lineFields+1,"Sample Number")
    xlOut.write(r,c+lineFields+2,"Sample Color")
    xlOut.write(r,c+lineFields+3,"Sample Unique ID")
    xlOut.write(r,c+lineFields+4,"Western Rank")
    xlOut.write(r,c+lineFields+5,"Western Value")
    xlOut.write(r,c+lineFields+6,"Zygosity")
    xlOut.write(r,c+lineFields+7,"mgFA Dry Weight")
    xlOut.write(r,c+lineFields+8,"mg FAMEs Std")
    xlOut.write(r,c+lineFields+9,"Weight Fraction")
    xlOut.write(r,c+lineFields+sampleFields,"FAME Name")
    xlOut.write(r,c+lineFields+sampleFields+1,"FAME Number")
    xlOut.write(r,c+lineFields+sampleFields+2,"FAME Color")
    xlOut.write(r,c+lineFields+sampleFields+3,"FAME Unique ID")
    xlOut.write(r,c+lineFields+sampleFields+4,"Retention Time")
    xlOut.write(r,c+lineFields+sampleFields+5,"Peak Area")
    xlOut.write(r,c+lineFields+sampleFields+6,"Percent FA")
    xlOut.write(r,c+lineFields+sampleFields+7,"Percent of Total FA")
    xlOut.write(r,c+lineFields+sampleFields+8,"ugFA")
    xlOut.write(r,c+lineFields+sampleFields+9,"Name Match Discrepancy")
    xlOut.write(r,c+lineFields+sampleFields+10,"Best Match")

    for l in lineManager.m_lines:
        r+= 1
        c = startC
        xlOut.write(r,c,l.m_name)
        xlOut.write(r,c+1,l.m_number)
        xlOut.write(r,c+2,l.m_colorID)
        xlOut.write(r,c+3,l.m_uniqueID)
        for s in l.m_samples:
            r+= 1
            xlOut.write(r,c+lineFields,s.m_name)
            xlOut.write(r,c+lineFields+1,s.m_number)
            xlOut.write(r,c+lineFields+2,s.m_colorID)
            xlOut.write(r,c+lineFields+3,s.m_uniqueID)
            xlOut.write(r,c+lineFields+4,s.m_expressionLevel.m_westernRank)
            xlOut.write(r,c+lineFields+5,s.m_expressionLevel.m_westernValue)
            xlOut.write(r,c+lineFields+6,s.m_zygosity)
            xlOut.write(r,c+lineFields+7,s.m_mgFADW)
            xlOut.write(r,c+lineFields+8,s.m_mgStd)
            xlOut.write(r,c+lineFields+9,s.m_totalWeightFraction)
            for f in s.m_fames:
                r+= 1
                xlOut.write(r,c+lineFields+sampleFields, f.m_name)
                xlOut.write(r,c+lineFields+sampleFields+1, f.m_number)
                xlOut.write(r,c+lineFields+sampleFields+2, f.m_colorID)
                xlOut.write(r,c+lineFields+sampleFields+3, f.m_uniqueID)
                xlOut.write(r,c+lineFields+sampleFields+4, f.m_retentionTime)
                xlOut.write(r,c+lineFields+sampleFields+5, f.m_peakArea)
                xlOut.write(r,c+lineFields+sampleFields+6, f.m_percentFA)
                xlOut.write(r,c+lineFields+sampleFields+7, f.m_percentOfTotalFA)
                xlOut.write(r,c+lineFields+sampleFields+8, f.m_ugFA)
                xlOut.write(r,c+lineFields+sampleFields+9, f.m_nameMatchDiscrepancy)
                if(f.m_bestMatch):
                    xlOut.write(r,c+lineFields+sampleFields+10, "T")
                else:
                    xlOut.write(r,c+lineFields+sampleFields+10, "F")

    xlFile.close()

    print("Saved data to",xlsxName)

#Read in the data structure of a file generated with SaveToFile
#RETURNS: a LineManager with all available data in the given file.
def LoadFromFile(xlsxName):

    ret = LineManager()

    #FIXME: duplicate code
    lineFields = 4
    sampleFields = 10
    FAMEFields = 11

    r = 0
    startC = 0
    c = startC

    #TODO: figure out forward declarations in python
    currentLine = Line()
    currentSample = Sample()
    currentFAME = FAME()

    print("Loading",xlsxName)
    # print("Minimum rows for a line:",lineFields)
    # print("Minimum rows for a sample:",lineFields+sampleFields)
    # print("Minimum rows for a FAME:",lineFields+sampleFields+FAMEFields)
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

            #Read a sample
            elif(sheet.cell_type(r,c+lineFields) != xlrd.XL_CELL_EMPTY and (c+lineFields+sampleFields == sheet.ncols or sheet.cell_type(r,c+lineFields+sampleFields) == xlrd.XL_CELL_EMPTY) ):
                currentSample = Sample()
                currentSample.m_name = sheet.cell(r,c+lineFields).value
                currentSample.m_number = sheet.cell(r,c+lineFields+1).value
                currentSample.m_colorID = sheet.cell(r,c+lineFields+2).value
                currentSample.m_uniqueID = sheet.cell(r,c+lineFields+3).value
                currentSample.m_expressionLevel.m_westernRank = sheet.cell(r,c+lineFields+4).value
                currentSample.m_expressionLevel.m_westernValue = sheet.cell(r,c+lineFields+5).value
                currentSample.m_zygosity = sheet.cell(r,c+lineFields+6).value
                currentSample.m_mgFADW = sheet.cell(r,c+lineFields+7).value
                currentSample.m_mgStd = sheet.cell(r,c+lineFields+8).value
                currentSample.m_totalWeightFraction = sheet.cell(r,c+lineFields+9).value
                currentLine.AddSample(currentSample)

            #Read a FAME
            elif(sheet.cell_type(r,c+lineFields+sampleFields) != xlrd.XL_CELL_EMPTY and (c+lineFields+sampleFields+FAMEFields == sheet.ncols or sheet.cell_type(r,c+lineFields+sampleFields+FAMEFields) == xlrd.XL_CELL_EMPTY) ):
                currentFAME = FAME()
                currentFAME.m_name = sheet.cell(r,c+lineFields+sampleFields).value
                currentFAME.m_number = sheet.cell(r,c+lineFields+sampleFields+1).value
                currentFAME.m_colorID = sheet.cell(r,c+lineFields+sampleFields+2).value
                currentFAME.m_uniqueID = sheet.cell(r,c+lineFields+sampleFields+3).value
                currentFAME.m_retentionTime = sheet.cell(r,c+lineFields+sampleFields+4).value
                currentFAME.m_peakArea = sheet.cell(r,c+lineFields+sampleFields+5).value
                currentFAME.m_percentFA = sheet.cell(r,c+lineFields+sampleFields+6).value
                currentFAME.m_percentOfTotalFA = sheet.cell(r,c+lineFields+sampleFields+7).value
                currentFAME.m_ugFA = sheet.cell(r,c+lineFields+sampleFields+8).value
                currentFAME.m_nameMatchDiscrepancy = sheet.cell(r,c+lineFields+sampleFields+9).value
                bestMatch = sheet.cell(r,c+lineFields+sampleFields+10).value
                if(bestMatch == "T"):
                    currentFAME.m_bestMatch = True
                if(bestMatch == "F"):
                    currentFAME.m_bestMatch = False
                currentSample.AddFAME(currentFAME)

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
        "Sample Name",
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
        for s in l.m_samples:
            for index, val in enumerate([
                l.m_name,
                s.m_name,
                s.m_totalWeightFraction * 100,
                Get18_2_to_3(s),
                s.GetFAME("C16:0").m_percentFA,
                s.GetFAME("C16:1").m_percentFA,
                s.GetFAME("C18:0").m_percentFA,
                s.GetFAME("C18:1").m_percentFA,
                s.GetFAME("C18:2").m_percentFA,
                s.GetFAME("C18:3").m_percentFA,
                s.GetFAME("C20:0").m_percentFA,
                s.GetFAME("C22:0").m_percentFA,
                s.GetFAME("C22:1").m_percentFA,
                s.GetFAME("C24:0").m_percentFA]
                ):
                xlOut.write(r, index, val)
            r += 1

    xlFile.close()

    print("Exported data to",xlsxName)
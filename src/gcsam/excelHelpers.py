import xlrd
from xlrd import open_workbook
import xlsxwriter
from .sampleManagement import *

# Excel Helpers is a collection of helper functions that interact with the specific excel documents the user will have.
# This file is highly malleable and should be adjusted to suit the needs of the user.
# For example, if the user has a new file format to read in, the code to parse the file would go here. Or if the user wants the config sheet data to be transposed, that change would be made here as well.
# Another good example is changing the analysis output. 
# Ultimately, users should not be afraid to change this file to suit their needs.



#RETURNS: a LineManager with all of the Lines from the desired sheet of the desired xlsx file
def ReadSamplesAndFAMEsFrom(xlsxName):
    print("Reading in",xlsxName)
    wb = open_workbook(xlsxName)

    lines = LineManager()

    for sheet in wb.sheets():
        if(sheet.name != "Output" and sheet.name != "Config"):
            # print("Reading in",sheet.name,sheet.nrows,"x",sheet.ncols)
            s = Sample()
            # s.m_name = str(sheet.name)
            s.m_name = str(sheet.cell(4,9).value)
            lineName = s.m_name.split('-')[0]
            l = lines.GetLine(lineName)
            if(not l.IsValid()):
                l = Line()
                l.m_name = lineName
                # print("Created new line:", l.m_name)
                lines.AddLine(l)
            # print("Adding sample", s.m_name)
            for r in range(24,sheet.nrows):
                if(sheet.cell_type(r,1) != xlrd.XL_CELL_EMPTY and sheet.cell_type(r,6) != xlrd.XL_CELL_EMPTY):
                    f = FAME()
                    f.m_retentionTime = float(sheet.cell(r,1).value)
                    f.m_peakArea = float(sheet.cell(r,3).value)
                    s.AddFAME(f);
                    # print("Adding FAME",f.m_retentionTime,":",f.m_peakArea)
            l.AddSample(s)
    return lines

#Reads in configXlsx and sets m_mgFADW and m_mgStd for each sample in the given lineManager according to the corresponding config entry.
def ApplyConfigTo(configXlsx,lineManager):
    sampleConfigured = False
    wb = open_workbook(configXlsx)
    for sheet in wb.sheets():
        if(sheet.name == "Config"):
            print("Reading samples from", configXlsx)
            for l in lineManager.m_lines:
                for s in l.m_samples:
                    sampleConfigured = False
                    for r in range(1,sheet.nrows):
                        if(sheet.cell_type(r,0) == xlrd.XL_CELL_EMPTY):
                            break
                        if(s.m_name == sheet.cell(r,0).value):
                            s.m_mgFADW = float(sheet.cell(r,1).value)
                            s.m_mgStd = float(sheet.cell(r,3).value)
                            # print("Added mg:",s.m_mgFADW,"and std:",s.m_mgStd,"to",s.m_name)
                            sampleConfigured = True
                            break
                    if(not sampleConfigured):
                        print("No config entry found for",s.m_name)
            break

#RETURNS: a FAMEsContainer with the standard FAMEs as provided in the specified excel sheet
def ReadLabeledFAMEsFrom(xlsxName):
    ret = FAMEsContainer()

    wb = open_workbook(xlsxName)
    for sheet in wb.sheets():
        if (sheet.name == "Config"):
            print("Reading labeled FAMEs from", xlsxName)
            for r in range(1,sheet.nrows): #skip the first row of header info
                if (sheet.cell_type(r,5) == xlrd.XL_CELL_EMPTY):
                    break;
                f = FAME()
                f.m_name = sheet.cell(r,5).value
                if (sheet.cell_type(r,6) != xlrd.XL_CELL_NUMBER):
                    print("Cell",r,"3 (",sheet.cell(r,6).value,") is not a number")
                else:
                    f.m_retentionTime = sheet.cell(r,6).value #read dist from std RT
                print("Read FAME:",f.m_name,":",f.m_retentionTime)
                ret.AddFAME(f)
            break
    return ret

def Get18_2_to_3(sample):
    p18_2 = sample.GetFAME("C18:2")
    if not p18_2.IsValid():
        print(sample.m_name, "has no valid C18:2")
        return 0
    p18_3 = sample.GetFAME("C18:3")
    if not p18_3.IsValid():
        print(sample.m_name, "has no valid C18:3")
        return 0
    return p18_2.m_peakArea / p18_3.m_peakArea

def Get18_2_and_20_0_and_22_1(sample):
    ret = 0
    p18_2 = sample.GetFAME("C18:2")
    if(p18_2.IsValid()):
        ret += p18_2.m_percentFA
    p20_0 = sample.GetFAME("C20:0")
    if(p20_0.IsValid()):
        ret += p20_0.m_percentFA
    p22_1 = sample.GetFAME("C22:1")
    if(p22_1.IsValid()):
        ret += p22_1.m_percentFA
    return ret

def WriteAnalysis(lineManager, xlsxName):
    xlFile = xlsxwriter.Workbook(xlsxName)

    #row/column pos
    r = 0
    startC = 0
    c = startC

    xlOut = xlFile.add_worksheet('Summary')
    xlOut.write(r, c, "Sorted by Name")
    r += 1
    xlOut.write(r, c, "Sample")
    xlOut.write(r, c+1, "% ΣFA")
    xlOut.write(r, c+2, "DW")
    xlOut.write(r, c+3, "18:2/3")
    r += 1
    for l in lineManager.m_lines:
        for s in l.m_samples:
            xlOut.write(r, c, s.m_name)
            xlOut.write(r, c+1, s.m_totalWeightFraction * 100)
            xlOut.write(r, c+2, s.m_mgFADW)
            xlOut.write(r, c+3, Get18_2_to_3(s))
            r += 1
    r = 0
    c = 5
    xlOut.write(r, c, "Sorted by %FA")
    r += 1
    xlOut.write(r, c, "Sample")
    xlOut.write(r, c+1, "% ΣFA")
    xlOut.write(r, c+2, "18:2/3")
    r += 1
    samples = []
    for l in lineManager.m_lines:
        samples += l.m_samples
    samples.sort(key = lambda s: s.m_totalWeightFraction)
    for s in samples:
        xlOut.write(r, c, s.m_name)
        xlOut.write(r, c+1, s.m_totalWeightFraction * 100)
        xlOut.write(r, c+2, Get18_2_to_3(s))
        r+= 1
    r = 0
    c = 9
    xlOut.write(r, c, "Sorted by Ratio")
    r += 1
    xlOut.write(r, c, "Sample")
    xlOut.write(r, c+1, "18:2/3")
    xlOut.write(r, c+2, "% ΣFA")
    r += 1
    samples = []
    for l in lineManager.m_lines:
        samples += l.m_samples
    samples.sort(key = lambda s: Get18_2_to_3(s))
    for s in samples:
        xlOut.write(r, c, s.m_name)
        xlOut.write(r, c+1, Get18_2_to_3(s))
        xlOut.write(r, c+2, s.m_totalWeightFraction * 100)
        r+= 1

    xlOut = xlFile.add_worksheet('Profiles')

    #row/column pos
    r = 0
    c = startC

    lineFields = 1
    sampleFields = 2
    FAMEFields = 3

    xlOut.write(r,c,"Line")
    xlOut.write(r,c+lineFields,"Sample")
    xlOut.write(r,c+lineFields+1,"% ΣFA")
    xlOut.write(r,c+lineFields+sampleFields,"FAME")
    xlOut.write(r,c+lineFields+sampleFields+1,"% of tissue")
    xlOut.write(r,c+lineFields+sampleFields+2,"% of total FA")
    r += 1

    for l in lineManager.m_lines:
        xlOut.write(r,c,l.m_name)
        for s in l.m_samples:
            xlOut.write(r,c+lineFields,s.m_name)
            xlOut.write(r,c+lineFields+1,s.m_totalWeightFraction)
            for f in s.m_fames:
                if(f.m_bestMatch):
                    xlOut.write(r,c+lineFields+sampleFields, f.m_name)
                    xlOut.write(r,c+lineFields+sampleFields+1, f.m_percentFA * 100)
                    xlOut.write(r,c+lineFields+sampleFields+2, f.m_percentOfTotalFA * 100)
                    r += 1

    # xlOut = xlFile.add_worksheet('Side by Side Ratios')
    # r = 0
    # c = startC

    # xlOut.write(r,c,"18-2 to 18-3")
    # r+= 1
    # xlOut.write(r,c,"Line Name")
    # xlOut.write(r,c+1,"Sample Name")
    # xlOut.write(r,c+2,"18:2/18:3 (by percentFA)")
    # xlOut.write(r,c+3,"Σ(18:2, 20:0, 22:1)")
    # r+= 1
    # for l in lineManager.m_lines:
    #     xlOut.write(r,c,l.m_name)
    #     for s in l.m_samples:
    #         xlOut.write(r,c+lineFields,s.m_name)
    #         xlOut.write(r,c+lineFields+1,Get18_2_to_3(s))
    #         xlOut.write(r,c+lineFields+2,Get18_2_and_20_0_and_22_1(s))
    #         r+= 1

    # xlOut = xlFile.add_worksheet('Sorted Ratios')
    # r = 0
    # c = startC

    # xlOut.write(r,c,"18-2 to 18-3")
    # r+= 1
    # xlOut.write(r,c,"Line Name")
    # xlOut.write(r,c+1,"Sample Name")
    # xlOut.write(r,c+2,"18:2/18:3 (by percentFA)")
    # r+= 1
    # for l in lineManager.m_lines:
    #     xlOut.write(r,c,l.m_name)
    #     samples = l.m_samples
    #     samples.sort(key = lambda s: Get18_2_to_3(s))
    #     for s in samples:
    #         xlOut.write(r,c+lineFields,s.m_name)
    #         xlOut.write(r,c+lineFields+1,Get18_2_to_3(s))
    #         r+= 1

    # xlOut.write(r,c,"Σ(18:2, 20:0, 22:1)")
    # r+= 1
    # xlOut.write(r,c,"Line Name")
    # xlOut.write(r,c+1,"Sample Name")
    # xlOut.write(r,c+2,"Σ by percentFA")
    # r+= 1
    # for l in lineManager.m_lines:
    #     xlOut.write(r,c,l.m_name)
    #     samples = l.m_samples
    #     samples.sort(key = lambda s: Get18_2_and_20_0_and_22_1(s))
    #     for s in samples:
    #         xlOut.write(r,c+lineFields,s.m_name)
    #         xlOut.write(r,c+lineFields+1,Get18_2_and_20_0_and_22_1(s))
    #         r+= 1

    # xlOut = xlFile.add_worksheet('All Samples by Ratios')
    # r = 0
    # c = startC
    # samples = lineManager.GetAllSamples()

    # xlOut.write(r,startC,"18-2 to 18-3")
    # r+= 1
    # xlOut.write(r,c,"Sample Name")
    # xlOut.write(r,c+1,"Line Name")
    # xlOut.write(r,c+2,"18:2/18:3 (by percentFA)")
    # samples.sort(key = lambda s: Get18_2_to_3(s.m_sample))
    # r+= 1
    # for s in samples:
    #     xlOut.write(r,c,s.m_sample.m_name)
    #     xlOut.write(r,c+1,s.m_line.m_name)
    #     xlOut.write(r,c+2,Get18_2_to_3(s.m_sample))
    #     r+= 1

    # xlOut.write(r,startC,"Σ(18:2, 20:0, 22:1)")
    # r+= 1
    # xlOut.write(r,c,"Sample Name")
    # xlOut.write(r,c+1,"Line Name")
    # xlOut.write(r,c+2,"Σ by percentFA")
    # r+= 1
    # samples.sort(key = lambda s: Get18_2_and_20_0_and_22_1(s.m_sample))
    # for s in samples:
    #     xlOut.write(r,c,s.m_sample.m_name)
    #     xlOut.write(r,c+1,s.m_line.m_name)
    #     xlOut.write(r,c+2,Get18_2_and_20_0_and_22_1(s.m_sample))
    #     r+= 1


    xlFile.close()

    print("Saved data to",xlsxName)

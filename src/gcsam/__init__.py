#!/usr/bin/env python3
import os
import sys
import xlsxwriter
import argparse
from .sampleManagement import *
from .samplesInExcel import *
from .excelHelpers import *
from .graphs import *

def gcsam():
    descriptionStr = "Welcome to GC SAM, your friendly, neighborhood gas chromatography sample analyzer and manager! Before using GC SAM, please make sure you have a valid Config sheet for the file to be analyzed. See the documentation for more details on what this means."

    parser = argparse.ArgumentParser(description = descriptionStr)
    parser.add_argument('-c','--config-file', type = str, metavar = 'config.xlsx',help = 'File containing sample weights, standard weights, and labeled FAMEs', dest = 'configFile')
    parser.add_argument('-i', '--input-files', metavar = ('input1.xlsx','input2.xlsx'), type = str, help = 'GC output to be analyzed.', dest = 'iFiles', nargs = '*')
    # parser.add_argument('-if', '--input-format', type = str, choices = ['Harper'], default = 'Harper', help = 'Format of the GC output to be analyzed.', dest = 'iFileFormat')
    parser.add_argument('--standard', metavar = 'standardName', type = str, default = 'C17:0', help = 'String identifier of the standard', dest = 'std')
    parser.add_argument('-o', '--output-file', metavar = 'output.xlsx', type = str, help = 'Where to output analyzed data.', dest = 'oFile')
    # parser.add_argument('-of', '--output-format', type = str, choices = ['ZeaKal'], default = 'ZeaKal', help = 'Format of the desired output.', dest = 'oFileFormat')
    parser.add_argument('-s', '--save-file', metavar = 'saveTo.xlsx', type = str, help = 'Save all data for easy access later.', dest = 'saveLoc')
    parser.add_argument('-g', '--graph-folder', metavar = 'graphs', type = str, help = 'Directory to put generated graphs in.', dest = 'gDir')
    parser.add_argument('-e', '--export', metavar = 'exportTo.xlsx', type = str, help = 'Exports limited selection of data in a more consumable format', dest = 'exportLoc')
    parser.add_argument('-l', '--load-file', metavar = ('saved1.xlsx', 'saved2.xlsx'), type = str, help = 'Recall a previously analyzed data set.', dest = 'loadLocs', nargs = '*')
    parser.add_argument('--only-lines', metavar = ('line1','line2'), type = str, help = 'Only read in lines that match the given list', dest = 'onlyLines', nargs = '*')
    parser.add_argument('--not-lines', metavar = ('line1','line2'), type = str, help = 'Read in all lines except those listed', dest = 'notLines', nargs = '*')

    args = parser.parse_args()

    def ExitDueToErr(errorStr):
        print("#################################################################\n")
        print(errorStr)
        print("\n#################################################################")
        parser.print_help()
        sys.exit()

    if(args.iFiles is None and args.loadLocs is None):
        ExitDueToErr("You must specify at least one input file via -i or -l")

    if(args.iFiles is not None and args.iFileFormat is None):
        errorStr = "Please specify the format of" + args.iFiles
        ExitDueToErr(errorStr)

    if(args.iFiles is not None and args.configFile is None):
        ExitDueToErr("Please specify a config file for the given inputs")

    if(args.oFile is None and args.saveLoc is None and args.exportLoc is None and args.gDir is None):
        ExitDueToErr("You must specify at least one output file via -e, -g, -o, or -s")

    if(args.onlyLines is not None and args.notLines is not None):
        ExitDueToErr("Please specify either --only-lines or --not-lines, not both")

    if(args.std is None):
        print("No standard specified, percentFA will not be calculated")

    lineManager = LineManager()
    if(args.iFiles is not None):
        for f in args.iFiles:
            lineManager.ImportDataFrom(ReadSamplesAndFAMEsFrom(f))
            ApplyConfigTo(args.configFile,lineManager)
            lineManager.LabelFAMEsWith(ReadLabeledFAMEsFrom(args.configFile))
            if(args.std is not None):
                lineManager.CalculatePercentFAForAllSamples(args.std)
                lineManager.CalculatePercentOfTotalFAForAllSamples(args.std)
            lineManager.Analyze()
        print("Removing unlabeled FAMEs")
        lineManager.RemoveUnlabeledFAMEs()
        lineManager.SetUniqueIDs()
        lineManager.Analyze()

    if(args.loadLocs is not None):
        for f in args.loadLocs:
            lineManager.ImportDataFrom(LoadFromFile(f))
        lineManager.SetUniqueIDs()
        lineManager.Analyze()

    if(args.onlyLines is not None):
        lineManager.RemoveAllLinesExcept(args.onlyLines)
    if(args.notLines is not None):
        lineManager.RemoveAllTheseLines(args.notLines)

    if(args.saveLoc is not None):
        SaveToFile(lineManager,args.saveLoc)
    if(args.exportLoc is not None):
        Export(lineManager,args.exportLoc)
    if(args.oFile is not None):
        WriteAnalysis(lineManager,args.oFile)
    if(args.gDir is not None):
        ggen = GraphGenerator(lineManager,args.gDir)
        ggen.GenerateGraphForProfiles()
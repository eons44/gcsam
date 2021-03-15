# GCSAM
Gas chromatography sample analyzer and manager

## Use Case

Tired of labeling GC output peaks by hand?
`gcsam` will automatically digest the output from your gas chromatograph and give you usable data so that you can spend more time analyzing what matters.
This program is written entirely in python for easy hacking and designed to specifically work with Microsoft Excel.

## Prerequisites
python >= 3.6.3

python packages:
* xlrd
* xlsxwriter

Microsoft Excel

## Installation
`pip install gcsam` (coming soon, I hope!)

## Usage
Go take a look at the `example/` folder, which should have all you need to get started.

### Configuration

A “Config” sheet (xlsx file) must be created that contains the following information:
* Sample IDs matching the sheet titles with:
    * Dry weight of the sample (in mg)
    * Amount of standard added (in mg)
* Known FAMEs
    * One column for the string identifier, this can be anything that makes sense to you (e.g. “c17.0” or “17:0”
    * Another column for the expected retention time of the FAME.

**NOTE: all retention times everywhere must be in descending order!**

Notice that the Volume column is not used, as all of the entries (in mL) are 1. If this is ever not the case, the code will need to be adjusted accordingly.

### Required Arguments

Please note that the order of the arguments supplied does not matter.

The config, from above, is required (`-c`).

At least one input and one output are required. It is possible to have multiple input and output types.

Inputs can be:
* One or more GC outputs (`-i`), read in with a particular format (`-if`)
* Previously saved files can be loaded (`-l`)

Outputs can be:
* Saved data (`-s`) for use with loading or further analysis.
* Output data, (`-o`) with a custom format (`-of`)
* Graph visualization folder (`-g`).

Specifying a standard is not required but you will likely want to include it with `--standard`. The default standard, "C17:0", may change in a future release.

### Examples

Converting raw GC output to a GCSAM file
```
gcsam -i "gcOutput.xlsx" -f "Harper" -c "configFile.xlsx" -s "savedGCdata.xlsx" --standard "C17:0"
```

Removing lines from previously saved files
```
gcsam -l "savedGCdata.xlsx" –-not-lines "1" "abc" "foxes45" -s "newSavedGCdata.xlsx"
```
Or
```
gcsam -l "savedGCdata.xlsx" --only-lines "1" "abc" "foxes45" -s "newSavedGCdata.xlsx"
```

Analyzing select lines from multiple, previously-saved files
```
gcsam -l "savedGCdata1.xlsx" "savedGCdata2.xlsx" –-only-lines "that" "I" "want" -o "analysisThatIWant.xlsx"
```

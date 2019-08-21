# XLS_file_processor

Created a XLSX/XLS file processor for alleviating ETL tasks. 

It merges input XLSX/XLS files into a single CSV file that conforms to OpenSpecimen APIs.

## How to setup

1. `git clone https://github.com/Swapnil-ingle/XLS_file_processor`

2. `gradle clean`

3. `gradle installApp`

## How to run

The shell script and batch script for running the program are generated in the dir **$XLS_REPO/build/install/XLS_file_processor/bin** after the plugin is setup.

1. `cd $XLS_REPO/build/install/XLS_file_processor/bin`

**Linux:**
> 2. `./tkids-merge-xls-script <ABSOLUTE-PATH-TO-INPUT-FILES>`

**Windows:**
> 2. `tkids-merge-xls-script <ABSOLUTE-PATH-TO-INPUT-FILES>`

**Note:**

1. This was tested on Gradle v2.0

2. The shell script and batch script for the program are in $XLS_REPO/build/install/XLS_file_processor/bin.

## Processing

All XLSX files directly under the input directory provided will be processed by the program.
> **Note:** Program won't read the directory recursively.

Upon fully processing the input files, the program will give a message as *"Done processing!!"*

Under the input directory a new folder *"merge-outputs"* will be created, if does not exists previously. 

Under **$INPUT_DIR/merge-outputs/<yyyy_mm_dd-hh-MM-ss>** two output CSVs will be created with the name 
1. **output.csv** (Merged output CSV)
2. **output-headers-origin.csv** (A helping CSV that back-tracks where each column header came from)

## Error and special case handling

1. Program expects a runtime argument - the absolute path for input files - If not provided the program will give an error **"Please provide the path for source files as runtime arguments"**

2. Any stray cells, cells that does not have any header above them will be **ignored**.

3. All rows, until EOF, will be merged to the output file. (Any empty rows, will also be merged as is.)

4. All the headers will be converted to lowercase to avoid case-sensitivity issues while matching the various headers of multiple files.

## Using the scripts on a remote machine

Once the program is setup the scripts can be moved to any machine where it needs to be executed.
Copy the **$XLS_REPO/build/install** folder to the destination server/machine where you need to execute this.

Once you have the *scripts(/bin)* and the */lib* folder on the target server you can follow the **"How to run"** section

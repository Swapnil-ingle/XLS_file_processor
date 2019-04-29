# XLS_file_processor

Created a XLSX/XLS file processor for alleviating ETL tasks. 

It merges input XLSX/XLS files into a single CSV file that conforms to OpenSpecimen APIs.

## How to setup

1. `git clone https://github.com/Swapnil-ingle/XLS_file_processor`

2. `gradle clean`

3. `gradle installApp`

4. `cd $XLS_REPO/build/install/XLS_file_processor/bin`

5. `./tkids-merge-xls-script <PATH-TO-INPUT-FILES>`

**Note:**

1. This was tested on Gradle v2.0

2. Here resides the shell script and batch script for the program.

## Processing

All XLSX files directly under the input directory provided will be processed by the program.
> **Note:** Program won't read the directory recursively.

Upon fully processing the input files, the program will give a message as *"Done processing!!"*

Under the input directory a new folder *"merge-outputs"* will be created, if does not exists previously. Under **$INPUT_DIR/merge-outputs/** the output CSV will be created with the name **output_<yyyy_mm_dd-hh-MM-ss>.csv**

## Error and special case handling

1. Program expects a runtime argument - the absolute path for input files - If not provided the program will give an error **"Please provide the path for source files as runtime arguments"**

2. Any stray cells, cells that does not have any header above them will be **ignored**.

3. All rows, until EOF, will be merged to the output file. (Any empty rows, will also be merged as is.)

4. All the headers will be converted to lowercase to avoid case-sensitivity issues while matching the various headers of multiple files.


# Copy "Base" Files and Folders in Windows

This program reads off an Excel document, copying folders and files from user-specified source paths to user-specified destination paths.

The copy process does NOT retain the full folder structure of the source folders and files. It only copies the file or folder at the bottom (or "base") of the source path and copies it to the specified destination folder.

If the destination path does NOT already exist, and the destination path is valid, the program will attempt to create the folder. Because of this feature, you can edit the destination folder path to include the desired folder structure.

Both local and long network paths should work.

The program harnesses Windows' built-in Robocopy tool to perform the copying. The following Robocoopy settings were applied to retain all data, attributes, and time stamps: 
/COPY:DAT

## Authors

- [@wlao-cyber](https://github.com/wlao-cyber)


## FEATURES, TIPS, and WARNINGS:

- Before using the program again, completely delete all Excel rows below headers in case there were previous cells left with spaces or random characters.
    - Cells with just spaces (whitespace) will be reported as empty rows.
- Save Excel document with your changes prior to running EXE.
- Cells not detected as valid paths should be reported in results at the end of execution.
- Just like with Windows Command Prompt, if you left-click the console Window, it pauses Robocopy.
    - To resume, you can right-click or press ENTER.
- The program tries to identify invalid paths and will note them in the console window.
- A size mismatch errors CSV log will be generated if any size mismatches from copying are found.

## Installation and How to Use

- Download the EXE and XLSX pair of files locally and keep them in the same folder.
- Don't change the Excel XLSX Name. Otherwise, the program will not work.
- In the XLSX document, fill out the full source path and destination paths under the labelled columns
- Double click the EXE, and the program will start.

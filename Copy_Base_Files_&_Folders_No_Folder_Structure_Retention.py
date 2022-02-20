""" This for copying files AND folders using Windows Commands built into Python to forgo using BAT scripts"""
# Copies only the BASE folders and files. Does not restore the full folder structure.
# The source folder/file path can be either mapped or long UNC.
    # Same goes for dest path
# Auto creates destination path if it doesn't already exist.
# Assumes that the dest path is a valid path. Otherwise, will not work when generating nonexisting folders.

# Limitations:
    # Does not check for illegal characters in folder paths to be created.
    # Does not automatically restore the full folder structure for files
    # destination paths set without drive letter but beginning with backslash
    # would be set as the root OS drive
        # ex.  \Nonexistent\ will be interpreted as C:\Nonexistent\
    # destinations not beginning with drive letter or server name (ex. just a string),
    # would be created in the same directory as this program.
    
# FYI:
    # Just like with Command Prompt, if you left-click the copy Window, it pauses Robocopy.
    # To resume, you can either right-click or press the ENTER key.

# Features:
# Read rows of first and second columns in Excel.
    # Ignore blank names and do not add to script
    # xlrd automatically skips unfilled blank cell entries.
# Copies file/folders to the destination path WITHOUT wiping everything in the destination path.
    # Did NOT use Robocopy's /mir flag because it will delete files which do not exist in the source.
    # Using /e to include all subdirectories including empty folders
# Works whether or not folder paths have trailing backslashes (at the end of path)
    # Use os.path.abspath() gives absolute path, while negating the trailing backslash
    # No need to remove the trailing backslash
# Tells the user which rows are empty, incomplete, or have invalid paths.
# Checks for when there are spaces that would then cause xlrd to read the cell.
    # Use .strip() method to remove the trailing and leading whitespace.
    # Also skips and reports cells that have JUST spaces.
# Uses robocopy to copy folders and files to retain metadata. Can be configured depending on desired robocopy settings.
# Copies to destination path whether or not the folder already exists.
    # Creates destination path if doesn't exist.
    # Creates folders even if destination path has folder or file names with illegal leading and trailing whitespace.                                                                                                            
# Copies file/folders to the destination path WITHOUT wiping everything in the destination path.
    # Did NOT use Robocopy's /mir flag because it will delete files which do not exist in the source.
    # Using /e to include all subdirectories including empty folders
# Works whether or not folder paths have trailing backslashes (at the end of path)
    # Use os.path.abspath() gives absolute path, while negating the trailing backslash
    # No need to remove the trailing backslash
# Prints the source path and destination path in console (for user review)
# Uses byte comparison to check if folders/files were copied successfully.
    # use Windows API to calculate folder and file size [same speed as WinExplorer], rather than use Python.
    # Python has no straightforward folder size calculation and would be inefficient making
        # several system calls, taking up much longer time.
# Tells users the number of size mismatch errors resulting from the copy program
# If there are size mismatches, generates an error CSV file with: src path, dest path, src size, and dest size
    # Create errors.csv with time and date, so it doesn't replace old errors.csv
# Calculates time spent on script execution
# Ask for specific user input, so the EXE window doesn't auto close if user wants to check results.
# If program crashes, will display the Traceback error in the console.
# Added additional pattern check for org_dest_path to prevent cases where:
    # \Nonexistent\ will be mistinterpreted as valid path after normalization by
    # os.path.abs() to C:\Nonexistent
    # AND then robocopy creates the destination path

# NOTES:
# Not normalizing the src path because we want to validate (later) src path first.
        # normalization with os.path.abspath() can cause issues with values that
        # are just alphanumeric or start with just one backslash \

# Add the following:
# Checks for invalid paths
    # source paths that do NOT exist.
    # destination paths NOT beginning with drive letter or server name (ex. just a string)
        # Only Destination Paths starting with drive letter or \\ as in for long UNC file paths
        # are accepted.
            # Alphanumeric strings no longer create folder in the same directory as the program.
                # hello for the destination path would place hello folder in the same directory as the program.
            # String following backslash \ will no longer cause folders to be generated in C Drive.
                # \hello for destination path would place this folder in the C Drive
    # Doesn't create staging folders unless path starts with drive letter or \\ (for long UNC)
    # Only accepts source paths with drive letters. Tells user if no drive letter detected.
# Checks for when there are spaces that would then cause xlrd to read the cell.
    # Skips and reports cells that have JUST spaces.
# Check for trailing and leading whitespace in the field entries
    # Use .strip() method to remove the trailing and leading whitespace.
    # Also, removes trailing and leading whitespace of folder names WITHIN folder paths b/c Windows does NOT
    # accept folder names with leading or trailing whitespace.
        # Example:
            # <D:\  whitespace folder \trailing \ leading> becomes <D:\whitespace folder\trailing\leading>
# Not normalizing the src path because we want to validate (later) src path first.
        # normalization with os.path.abspath() can cause issues with values that
        # are just alphanumeric or start with just one backslash \

# IMPORTING PACKAGES/MODULES
# os to read, direct, and create file/folder paths
# subprocess to run Windows commands
# Use re (regular expressions) to check patterns for
    # source path and dest path
# csv to generate csv file
# time to print out a unique errors.csv named by a
    # unique current date & time that doesn't overwrite older ones
# datetime to calculate the execution time
# traceback to show any program crash errors.
import os
import subprocess
import re
import csv
import time
import datetime
import traceback

# start time of script
start_time = datetime.datetime.now()

# win32com.client (part of pywin32 package) to use Windows' File Explorer
# xlrd to read from Excel file
import win32com.client
import xlrd

def variables():
    # set global variables
    global fso, file_location, workbook, sheet, invalid_paths_rows, \
    empty_rows, missing_entries_rows, \
    size_mismatch_rows, copy_errors
    
    # use Win32 API to get the folder size later.
    # set file system object using the Dispatch command
    # Provides access to a computer's file system
    fso = win32com.client.Dispatch("Scripting.FileSystemObject")

    # Specify the Excel file to read from. The long file path is unnecessary if the script is in the same directory,
    # but specifying the long file path of the excel file gives you the option to run the script from anywhere. 
    file_location = r"Copy_Base_Files_and_Folders.xlsx"

    workbook = xlrd.open_workbook(file_location)

    # This opens the first sheet (0th index) in the Excel File
    # If you know the name of the sheet, you can also open by
    # the sheet name "Sheet1" in this case.
        #sheet = workbook.sheet_by_name("Sheet1")
    sheet = workbook.sheet_by_index(0)

    # set variable out of loop to keep track of
    # empty rows, rows w/ missing entries,
    # invalid paths as well as size mismatches.
        # Use a list to keep adding rows to.
        # Then shoot out the result at the end, detailing
        # which rows are empty or are missing entries.
        # mainly using lists because lists retain order.
    empty_rows = []
    missing_entries_rows = []
    invalid_paths_rows = []
    size_mismatch_rows = 0

    # set global list of dictionaries objects that will keep growing as size mismatches are detected.
    copy_errors = []

def read_Excel():
    # set global for size_mismatch_rows again so
    # other functions like this can better identify
    # Otherwise, python may assume local variable is
    # referenced before being assigned
    global size_mismatch_rows 

    # set up pattern to check long UNC paths if path begins with
    # drive letter or \\ for long UNC paths.
        # Remember: 1 backslash is denoted as 4 backslashes in regex.
            # To check for 2 repititions use {2}
        # Check for 2 backslashes instead of 1 because
            # a path starting with just 1 backslash will be created in C:\ directory
    check_path_start = '^[a-zA-Z]:\\\\|^\\\\{2}'
    
    # Read through the spreadsheet.
    # Ignore the first row because it has the header.
    # Set the for loop to start with the second Excel row
    # Use the cell values indicated in the Excel document to
    # copy the files/folders to their destination folders
    for row in range(1, sheet.nrows):
        
        # the argument next to row is the column index.
            # Ex. 0th and 1st columns mean the first and second Excel columns.
        # set cell value as a string so pure number values would not be interpreted as floats.
            # .strip() method fails on non-string types
        # Use .strip() method so that leading and trailing whitespace will be ignored in original and destination path
            # problem is that long numbers inputted into the Excel file will be interpreted as a float.
            # Thus, the .strip() method would fail on a type that is not a string
            # safer to convert the src_path and dest_path type to str preemptively, so strip will work
        src_path = str(sheet.cell_value(row, 0)).strip()
        dest_path = str(sheet.cell_value(row, 1)).strip()

        # get the original dest_path prior to normalization
            # reason: normalization with os.path.abspath() later converts pathless
            # alphanumeric strings to paths in the current directory of this program.
            # Also, os.path.abspath() converts values beginning with
            # just one backslash to path in the C drive.
            # Additionally, also converts '' to python directory.
        # save original dest_path for using in a condition later.
        org_dest_path = dest_path

        # Remove leading and trailing whitespace in names of all folders and
        # base folder or file name
        # Problem: os.path.isabs() does not check to see if path will be valid
            # leading and trailing spaces are illegal in Windows folder/file names but
            # return True to os.path.isabs() as long as the path begins with backslash
            # after a drive letter or long path
        # Solution: Remove leading and trailing whitespace in dest path set by user input (if any)
            # Ex. remove spaces in folder like C:\   Users\ Ocelot  \Demo \
            # split dest path by its backslashes to separate all folders in a list
            # Use list comprehension to quickly strip leading and trailing whitespace in the folder names
                # is like in place redefining of the list items
            # join all list items by backslash
                # This still returns long UNC file paths to original because
                # list nulls '' from stripping are joined by '\\' which in turn
                # restores the original double backslashes '\\\\' that
                # start the long UNC paths
        split_path = dest_path.split('\\')
        strip_whitespace = [i.strip() for i in split_path]
        dest_path = '\\'.join(strip_whitespace)

        # normalize the destination path by stripping backslash at the end (if any) using os.path.abspath()
            # However, this also sets a simple alphanumeric cell value as a path corresponding to current directory of this program
                # Ex. The dest_path value "Nonexistent" becomes <C:\Users\Maker's Will\Desktop\Python Projects\Python Projects Edited 2020-08-21\Copy Files and Folders\Copy_Retain_Folder_Structure_and_Stage_Files\Nonexistent\>
        # Then add backslash along with custodian and folder_type
        # don't want to normalize the src_path because if it is '', os.path.abspath() will
        # convert it to the python directory
            # If you try os.path.abspath() on a null string, it will return the
                # path where python is installed.
                    # >>> test
                    # ''
                    # >>> os.path.abspath(test)
                    # "C:\\Users\\Maker's Will\\AppData\\Local\\Programs\\Python\\Python37" 
        dest_path = os.path.abspath(dest_path)

        # Not normalizing the src path because we want to validate (later) src path first.
        # normalization with os.path.abspath() can cause issues with values that
        # are just alphanumeric or start with just one backslash \

        # Check where both src path and dest path are not null after being stripped
            # This is to exclude cells that have only spaces, and no actual values
        if src_path != '' and org_dest_path != '':

            # Only create the stage path if:
                # the src path exists, if there's actually something valid to copy over
                # stage path does NOT already exist.
                    # Note: Sometimes, you can have different user input that
                    # reevaluates to the same path after being stripped.
                        # Ex. of two folders that will reevaluate to same path after stripping
                            # D:\  whitespace folder \trailing \ leading
                            # D:\  whitespace folder  \            trailing \          leading
                    # Otherwise, will get a traceback error.
                    # You only want to create the same path once.
                # stage path specified is a legitimate path to create
                    # os.path.isabs() can be used to see if dest_path is actually a path
                    # beginning with backslash after drive letter in Windows or double backslash
                        # Problem with dest_path = os.path.abspath(dest_path) assigned earlier.
                        # False positives for:
                            # earlier dest_path was just an alphanumeric string
                                # converts an alphanumeric string to a path in this program's directory.
                            # dest path begins with backslash \
                                # converts backslash
                        # Solution: Use the org_dest_path and pattern check_path_start to
                        # check if original user input begins with double backslash \\ or drive letter
            if os.path.exists(src_path) and os.path.exists(dest_path)==False \
            and os.path.isabs(dest_path) and \
            bool(re.match(check_path_start, org_dest_path)):
                    
                # os.makedirs() will create a path whether or not the string ends with a backslash \
                # also creates nonexisting intermediate paths.
                print("Creating nonexisting destination path: " + dest_path)
                os.makedirs(dest_path)

            # Robocopy follows different syntax for folders and for files
            # First check for when source path is a dir and
            # destination path is a dir to use robocopy format for
            # copying folder to folder
            # The following if statement only runs IF
                # both src and stage paths exist
                    # if not already existing, the dest path should already
                    # be created from earlier
                # if src_path begins with drive letter
                    # re.match('pattern', 'string') checks only if a string's
                    # pattern matches at the start.
                        # bool to return True or False
            # os.path.isdir() works whether path ends with backslash or not.
            if os.path.isdir(src_path) and os.path.isdir(dest_path) \
            and bool(re.match(check_path_start, org_dest_path)):
                # normalize the src path by stripping backslash at
                # the end (if any) using os.path.abspath()
                # works whether or not path ends with backslash \
                src_path = os.path.abspath(src_path)

                # Robocopy doesn't copy the base folder over to destination path.
                # Therefore, add the base folder to the end of the destination path.
                # Use os.path.basename() to get the base folder or file name in the source path
                # Then add the basename to the destination folder to include the src's base folder, making the final destination
                src_base_folder = os.path.basename(src_path)
                final_dest = dest_path + '\\' + src_base_folder
                
                print("Source path: " + src_path)
                print("Destination path: " + dest_path)

                # Use subprocess module to run Robocopy from the source path to the final destination
                    # Robocopy syntax for folders: Robocopy "src folder" "dest folder" [flags]
                    # /COPY:DAT copies all file properties for data, attribute, and timestamps
                    # /E flag copies all subdirectories in source path, including empty folders
                subprocess.run(["Robocopy", src_path, final_dest, "/COPY:DAT", "/E"])
                
                # point Windows API's FileSystemObject to the folder path to get the sizes in bytes
                src_fldr = fso.GetFolder(src_path)
                dest_fldr = fso.GetFolder(final_dest)

                scrc_fldr_size = src_fldr.Size
                dest_fldr_size = dest_fldr.Size
                
                print("Source Folder Size: " + str(scrc_fldr_size))
                print("Destination Folder Size: " + str(dest_fldr_size))
                print()

                # If size mismatch detected, compile dictionary terms into the copy_errors list defined globally
                    # Headers will be 'Source Path', 'Destination Path', 'Source Size', 'Destination Size'
                    # Set the values tied to these dictionary items
                        # dictionary object can have combination of variables, numbers, and strings. (i.e., mix of strings and numbers)
                if scrc_fldr_size != dest_fldr_size:
                    copy_errors.append({'Source Path' : src_path, 'Destination Path' : final_dest, \
                                        'Source Size' : scrc_fldr_size, 'Destination Size' : dest_fldr_size})
                    # for each size mismatch, add to counter
                    size_mismatch_rows += 1

            # Use different robocopy format to copy files
            # Check for when source path is a file and destination path is a dir
                # Use os.path.basename() to grab just the base file name just for the robocopy command.
                # Does NOT work if there is a trailing backslash at the end of file name.
                    # This wouldn't make sense anyway and would be placed under invalid paths
            elif os.path.isfile(src_path) and os.path.isdir(dest_path) \
            and bool(re.match(check_path_start, org_dest_path)):
                src_path = os.path.abspath(src_path)
                
                src_base_file = os.path.basename(src_path)

                print("Source path: " + src_path)
                print("Destination path: " + dest_path)

                # command syntax for files: robocopy "src folder" "dest folder" file.txt [flags]
                subprocess.run(["Robocopy", os.path.dirname(src_path), dest_path, src_base_file, "/COPY:DAT"])

                # point FileSystemObject to the file path to get the sizes in bytes
                src_file = fso.GetFile(src_path)
                dest_file = fso.GetFile(dest_path + '\\' + src_base_file)

                src_file_size = src_file.Size
                dest_file_size = dest_file.Size
                
                print("Source File Size: " + str(src_file_size))
                print("Destination File Size: " + str(dest_file_size))
                print()

                # Just like before, if size mismatch detected, append dictionary terms into the copy_errors list
                if src_file_size != dest_file_size:
                    copy_errors.append({'Source Path' : src_path, 'Destination Path' : dest_path, \
                                        'Source Size' : src_file_size, 'Destination Size' : dest_file_size})
                    size_mismatch_rows += 1

            # For all cases where, the row values do NOT match the above conditions of:
                # Copy valid folder path to valid folder path
                # Copy valid file path to valid folder path
            # Append the Excel row into the master list for invalid paths
                # Excel row is "row+1" because xlrd interprets first row as 0th row.
            else:
                invalid_paths_rows.append(row+1)
        
        # The elif statement will check when all original fields are empty and
        # append the Excel row number to the master empty_rows list
        # Use org_dest_path NOT dest_path
        # REMEMBER: dest_path was previously set to be normalized by os.path.abspath()
        # Therefore, a null string would be converted to a python path. 
            # If you try os.path.abspath() on a null string, it will return the
            # path where python is installed.
                # >>> test
                # ''
                # >>> os.path.abspath(test)
                # "C:\\Users\\Maker's Will\\AppData\\Local\\Programs\\Python\\Python37" 
        elif src_path == '' and org_dest_path == '':
            empty_rows.append(row+1)

        # The else statement will check if either the srcs or dest are missing values
        # append the Excel row number to the missing_entries_rows list variable.
        else:
            missing_entries_rows.append(row+1)

# Calculate the difference from start to end time and then convert to str for readable format.
# Reminder: Start time already declared at beginning of script.
# Changing format of the start_time and end_time to reader-friendly format.
    # %A for day of the week
    # %Y for Year
    # %m for month
    # %d for day of the month
    # %H for the hour (24-hr format) at the time
    # %M for the minutes (60 minute format) at the time
    # %S for the seconds (60 second format) at the time
# Now declaring and defining end time. And then subtracting to get timedelta.
def execution_time():
    print()
    print("Start Local Time: " + start_time.strftime("%A, %Y-%m-%d %H:%M:%S"))
    end_time = datetime.datetime.now()
    print("End Local Time: " + end_time.strftime("%A, %Y-%m-%d %H:%M:%S"))
    time_duration = end_time - start_time
    print("Total Time Elapsed in (Days) Hours:Mins:Secs: " + str(time_duration))

def end_results():
    print()
    # convert the list variables to string by type conversion to str
    # Strip the brackets from the lists by using .strip() method on the left and right bracket.

    # Only tell user to check empty Excel rows if any detected.
    if len(empty_rows) > 0:
        print("EMPTY ROWS")
        print("Check empty Excel rows: " + str(empty_rows).strip('[]'))
        print()
        
    # Only tell user to check missing entries rows if any detected.
    if len(missing_entries_rows) > 0:
        print("MISSING ENTRIES")
        print("For copy to work, each row must be filled out from source path to destination path.")
        print("Double check Excel rows: " + str(missing_entries_rows).strip('[]'))
        print()
        
    # Only tell user to Invalid paths rows if any detected.
    if len(invalid_paths_rows) > 0:
        print("INVALID PATHS")
        print("Double check Excel rows with invalid path(s): " + str(invalid_paths_rows).strip('[]'))
        print("Make sure the source paths exist and destination paths are valid.")
        print()

    # Always tell if there are any size mismatch errors resulting from copying
    print("Number of Size Mismatch Errors: " + str(size_mismatch_rows))

    # Only refer user to check errors.csv IF there are size mismatch errors.
    if size_mismatch_rows > 0:
        
        # sets the date and time in string format as a variable to later add to CSV name
        # This sets the time format in YearMonthDay_HrMinSec in Military Time
        now = time.strftime("%Y%m%d_%H%M%S")
        CSV_name = 'copy_errors_' + now + '.csv'
        
        # Generate the errors.csv
        # if you DO NOT set the newline parameter as '', then after the header, a row will be skipped.
            # This makes sure that you begin writing directly below the header.
        # It is good practice to use the 'with' keyword when dealing with file objects.
            # The advantage is that the file is properly closed after its
            # suite finishes, even if an exception is raised at some point.
        with open(CSV_name, 'w', newline='') as output_csv:
            fields = ['Source Path', 'Destination Path', 'Source Size', 'Destination Size']

            # The fieldnames parameter is a sequence of keys that identify the order in which
            # values in the dictionary passed to the writerow() method are written to the CSV file.
            # csv.DictWriter to set the header as a dictionary object that accepts values for its items
            Headers = csv.DictWriter(output_csv, fieldnames=fields)
            Headers.writeheader()

            # Make variable for writing to the CSV, so writing rows can begin
            output_writer = csv.writer(output_csv)

            # by now copy_errors should be a list of dictionary objects
                # each dictionary object with row header to values:
                # Ex. [{'Source Path': 'C:\\_Test\\new 2.txt', 'Destination Path': 'D:\\_TEST_TEST\\1\\', 'Source Size': 11, 'Destination Size': 23},
                # {'Source Path': 'C:\\_Test\\a\\', 'Destination Path': 'D:\\_TEST_TEST\\2', 'Source Size': 156, 'Destination Size': 149}]
            # .writerow() method normally accepts lists.
                # within each list element (the dictionary object), call on the value tied to the dictionary item.
            for row in copy_errors:
                output_writer.writerow([row['Source Path'], row['Destination Path'], \
                                        row['Source Size'], row['Destination Size']])

            print("See " + CSV_name + " if there were any size mismatch errors")

    print()

# Ask for user input so they can review results before closing.
    # Robocopy pauses if you right-click it. Left-click AND also the ENTER key will resume the copy.
    # However, pressing any button will cause the window to close if just asking for any user input.
    # Therefore, instead of asking simply ANY user input, ask for a specific input.
    # Only break while loop when specific input matched.
def user_close():
    while(True):
        user_input = input("To close this window, type 'exit' followed by ENTER" \
        " or click the close button: ")
        if user_input == 'exit':
            break

def main():
    variables()
    read_Excel()
    execution_time()
    end_results()
    user_close()

# Print the traceback error if the script fails.
if __name__ == "__main__":
    try:
        main()
    except Exception:
        print()
        print("ERROR:")
        print(traceback.format_exc())
        input("Let Dev know of error. Screenshot error and keep Excel records.")

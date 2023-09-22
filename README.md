# Grade-Manager-in-VBA

## Problem Statement

The Grade Manager is a tool designed to consolidate and combine student grades from multiple class section roster grade books. The class is split up into multiple sections, with each teaching assistant (TA) responsible for the grades of his or her section. The Grade Manager will take section files and consolidate/sync the information from those files into the main file (main grade sheet). 

![image](https://github.com/God-ass/Grade-Manager-in-VBA/assets/92200827/6dcf2537-353b-4cea-a11d-39fe0461138c)


## Functionalities

1. **Initialization**: If the course is new and has not yet been initialized, a directory for the course needs to be set up and the class roster file separated into section roster files (“section files”).
2. **Synchronization**: Your program/user form should have the ability to synchronize (“sync”) the information found in those section files to the main grade file.
3. **Backup**: Your project should have a way to make a backup file of the current file, saving it with the current date in the file name.
4. **Add/Delete Columns**: There should be a way for the user to add or delete entire grade columns, and these changes would automatically be made to ALL individual section files.
5. **Search Tool**: Your user form should have a way to search through by student name (selected via a drop-down list) for a certain grade on an assignment (selected from a drop-down list), and this would be displayed in a message box.

![image](https://github.com/God-ass/Grade-Manager-in-VBA/assets/92200827/aabfbd43-977a-4ca1-9145-c06f4ca5c478)


## Assumptions/Premises

- User will always place names and SID in first 2 columns, and we’ll assume that those will never be misspelled.
- New assignment headings/labels will always be placed in the top row.
- TAs will be able to modify their individual sheets and these will overwrite existing data.
- Lab sections will always be in increments of 1.
- Grades will always be on Sheet1 of the roster files
- No new students will need to be added

## More Details

1. **Initialization**: From the START PAGE, clicking on the button --- **'Set Up New Class'** will create a new course, which involves several steps such as creating a new directory for the course, saving the file with a specific name, deleting the START PAGE, placing the file path on the “Path” sheet of the main file, importing the roster file into the “Roster” sheet of the Grade Manager, creating individual section files in a subdirectory, and formatting each of the section files.

![image](https://github.com/God-ass/Grade-Manager-in-VBA/assets/92200827/1d7605a3-dbc1-4901-bd45-05865ffcdb34)

**Click the 'Modify' button to activate the Grade Manager Box**

2. **Syncing**: TAs should then be able to add grades to their individual section grade files. When needed, the teacher can import grades from the section files into the main Grade Manager file (this is known as syncing). 
3. **Backup**: You should have the ability to back up your roster file and your program should save the file as a new name that includes the current date.
4. **Add Grade Item**: Your Grade Manger must have the capability to add new grade items (columns) but it doesn’t need to be able to add students.
5. **Delete Grade Item**: Your Grade Manager project needs to be able to delete columns of data from all sheets.
6. **Search and Replace**: Finally, your Grade Manager must have a search/replace tool that allows the teacher to select the student from a drop-down list (combo box) as well as an assignment of interest (also selected from a drop-down list), and this would be displayed in a message box.

## Hints

- When syncing, also import the first row (grade items/labels) in the import.
- Whenever a combobox is used on a user form, it should be unloaded and not hidden at the end of the subroutine.  Otherwise, duplicate items are added to the comboboxes.
- When troubleshooting, make sure to delete all preexisting files and folders that you are trying to create; otherwise, you might have an error if you try to write over a preexisting folder or file.
- All of the code for everything (including user forms) need to be in the “Grade Manager – SETUP.xlsm” file to begin with


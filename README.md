# Degreee classification calculator Maths at Queen's University Belfast

This script can be used to calculate the final (averaged) marks that students get at the end of level 3 of their maths degree at Queen's University Belfast.  It takes as input:

1. An excel spreadsheet called `QSR_EXAM_RESULTS_1379.xlsx` that contains information on every module the student has ever enrolled on.  Staff in the education office can download this spreadsheet from QSIS.
2. Grade rosters for all the modules for which marks are available.  These grade rosters should be named `{module code}.xlsx`.

It outputs a directory called `Results` that contains 6 files:

1. `student_marks.json` = a JSON file that contains all the information about each student from the various spreadsheets that were input.
2. `BSC_MARKS.xlsx` = Marks and grade classifications for BSc students 		
3. `MISC_MARKS.xlsx` = Marks the script is not able to interpret because students do not have 120 CAT points at levels one and 2.		
4. `MSCI_MARKS.xlsx` = Marks and progression decisions for MSci students		
5. `OTHER_MARKS.xlsx` = Data for students who are not listed in the `QSR_EXAM_RESULTS_1379.xlsx` file.  These should all be level four students who are taking an alternating module that is taught at level 3 as well.	
6.  `PLACEMENT_MARKS.xlsx` = The list of students who are completing some sort of placement. 

The excel files are generated from `student_marks.json` so in theory there could be two separate scripts with the second taking `student_marks.json` as its input.

Headings in the excel files are self explanatory.  If the code outputs a message about a student to standard output then the reason that message is being output should be manually investigated.

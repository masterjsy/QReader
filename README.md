# QReader

This is a simple program that is designed to get data from Yahoo Finance and Market watch.

How to:
1. Run the program
2. Choose csv or excel file. 
2.1. This file should have a column "Symbol" and each row under the column should have ticker
2.2 If the file is too big, Yahoo Finance will deny getting the query result. So it is recommended to make each excel file less than 1000 rows.
2.3 You can choose multiple files. QReader will run first file and wait 5400 seconds. You can make it shorter or longer by changing a variable.
3. Wait
4. Result will be stored in a new file starting with Record_Output_.

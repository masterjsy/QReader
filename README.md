# QReader

This is a simple program that is designed to get data from Yahoo Finance and Market watch.

How to:
1. Run the program
2. Choose csv or excel file. 
2.1. This file should have a column "Symbol" and each row under the column should have ticker
2.2 If the file is too big, Yahoo Finance will deny getting the query result. Therefore, the program will split the queries. The default side per query is 700. There is waiting time between queries. The default waiting time is 4000 seconds
2.3 You can choose multiple files. QReader will run first file and wait 4000 seconds. You can make it shorter or longer by changing a variable.
3. Wait
4. Result will be stored in a new file starting with Record_Output_.

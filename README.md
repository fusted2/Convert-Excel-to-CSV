# Convert Excel to CSV

## Problem Description
Hi everyone!

When working in the financial and accounting, I noted that some Excel reports having layout that the 1<sup>st</sup> sheet is usually the reporting, signing, guidance,... and the 2<sup>nd</sup> sheet is the tabular data we need. I wanted to batch convert these reports to `csv`/`txt` files and ASAP Utilities worked for me.

However, this application only supports with the active sheet, i.e. the last place where you saved your files. Therefore, I had to open all those files to ensure the correct active sheets I wanted to convert and this is tedious!

I tried to find on the internet if there were any answer for my problem, but I haven't found any yet. Hence, I created this application to solve that niche issue.

## Usage
![excelToCSV_example](https://github.com/user-attachments/assets/a47788f1-2acd-448f-bed1-b8176e78441a)

It is quite simple and easy to use.

You copy and paste the input/output folder location to the bars, and type in the sheet index, header row, output extension, and separator.

Refer to the screenshot for example.

## Python Library Requirements
Nothing new or special, I'm using `pandas` to read Excel files and export to `csv`/`txt`, and `openpyxl`, `pyxlsb`, `xlrd` to handle `.xlsx`, `.xlsb`, `.xls` extensions.

You can find the version details in the file `requirements.txt`

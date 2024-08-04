# csv2XLsheet
Append delimited data to Microsoft Excel templates or xlsx files<br>
Works with templates that contain existing tables, pivot tables, slicers<br>
Line input errors are ignored and logged.<br>
Quotation marks are removed during processing.<br>

```
Usage: csv2XLsheet [-i,-t,-s,-o,-d,-r,-h]
```

#### Options:<br>
  -i  Input Path to the source CSV/TSV file (required)<br>
  -t  Path to the Excel XLSX/XLTX file (required)<br>
  -s  Existing sheet name to append lines (required)<br>
  -o  Output file name (required)<br>
  -d  Delimiter of input file (options: 'csv', 'tab', or character(s)) (default: 'csv')<br>
  -r  Start appending sheet from this line number (default: 1)<br>
  -h  Show this help message<br>

 #### Example:

```
csv2XLsheet -i prc.csv -t PfSlicer.xltx -s Pf-Table -r 2 -o pfoutput.xlsx
```
 Example appends the CSV file prc.csv to a sheet named Pf-Table.<br>
 The source excel template is named PfSlicer.xltx.<br>
 The import starts at line 2 (omitting the csv header)<br>
 and outputs a file named pfoutput.xlsx<br>

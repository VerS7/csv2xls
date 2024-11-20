# csv2xls
### Simple python script for writing CSV data to xls/xlsx excel file
### Default example
```commandline
    python csv2xls.py -f input.csv -o output.xlsx 
```
### Commands
> `--help`
> > Help information
> 
> `-f | -from | -file | -csv` 
> > path to CSV file  
> 
> `-o | -output | -out`
> > `path to output xls/xlsx file`
> 
> `-header`
> > `include` (default) include first line of CSV data to excel file  
> > `no` exclude first line  
> 
> `-a`
> > append CSV data to excel file
> 
> `-debug | -d`
> > `all` all debug to stdout  
> > `necessary` only necessary debug info  
> > `no` exclude debug info
> 
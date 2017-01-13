#!/usr/local/bin/python
#graig orie 2017
import glob
import csv
import xlwt
import os

## loc = "J:\- Rapportages\Maandag\Uitvoer\*.csv"
loc = "Uitvoer/*.*"

## Eerst worden alle totalen uit de csv bestanden gelezen 

row_count1 = 0
start = "Tota"
wb1 = xlwt.Workbook()
wt = wb1.add_sheet(str("Totalen"),cell_overwrite_ok=True)
for filename1 in glob.glob(loc):
    (f_path1, f_name1) = os.path.split(filename1)
    (f_short_name1, f_extension1) = os.path.splitext(f_name1)    
    spamReader1 = csv.reader(open(filename1, 'rb'), delimiter=',',quotechar='"')
## alle totalen worden naar de eerste worksheet geschreven    
    ll = [0,1,2,3,4,5,6,7,8,9]
    for row1 in spamReader1:
        if row1[0].startswith(str(start)) == True:
            row_count1 = row_count1 + 1   
            for col1 in range(len(row1)):
                wt.write(row_count1,col1,row1[col1])
        else:
            continue

## nu worden alle csv bestanden als apparte worksheets toegevoegd aan het workbook

print(loc)
for filename in glob.glob(loc):
    (f_path, f_name) = os.path.split(filename)
    (f_short_name, f_extension) = os.path.splitext(f_name)
    ws = wb1.add_sheet(str(f_short_name))
    spamReader = csv.reader(open(filename, 'rb'), delimiter=',',quotechar='"')
    row_count = 0
    row_count1 = 0
    for row in spamReader:
        for col in range(len(row)):
            ws.write(row_count,col,row[col])                                
        row_count +=1        
wb1.save("compiled.xlsx")
print("Done")
import re, os, openpyxl as xl
from datetime import datetime
from dateutil.parser import parse

newdate = "%Y-%m-%d"
name= input('Please enter name of the excel file')

assignment_regex = re.compile(r'ACTIVE \w+')
interior_regex = re.compile(r'^INTERIOR')
wb = xl.load_workbook(name +'.xlsx')
ws = wb.active

#uppercasing 
print('starting cleanup')
for row in range(2, ws.max_row + 1):
    if ws.cell(row=row,column=2).value==None:
       ws.cell(row=row,column=2).value='NONE'
    else:   
        ws.cell(row=row,column=2).value=ws.cell(row=row,column=2).value.strip().upper()
        
    if ws.cell(row=row,column=4).value==None:
       ws.cell(row=row,column=4).value='NONE'
    else:   
        ws.cell(row=row,column=4).value=ws.cell(row=row,column=4).value.strip().upper()

    if ws.cell(row=row,column=6).value==None:
       ws.cell(row=row,column=6).value='NONE'
    else:   
        ws.cell(row=row,column=6).value=ws.cell(row=row,column=6).value.strip().upper()
		
    if ws.cell(row=row,column=9).value==None:
       ws.cell(row=row,column=9).value='NONE'
    else:   
        ws.cell(row=row,column=9).value=ws.cell(row=row,column=10).value.strip().upper()		
        
    if ws.cell(row=row,column=10).value==None:
       ws.cell(row=row,column=10).value='NONE'
    else:   
        ws.cell(row=row,column=10).value=ws.cell(row=row,column=10).value.strip().upper()
    
           
    print(str(row) + 'upper casing')

#making exemptions
print('starting active status cleanup')
for row in range(2,ws.max_row+1):
    
    
    if ws.cell(row=row, column=4).value.startswith('ACTIVE') or ws.cell(row=row,column=4).value == assignment_regex  :
        ws.cell(row=row, column=4).value = 'ACTIVE'

    if 'RECRUIT' in ws.cell(row=row,column=6).value:
        ws.cell(row=row,column=4).value = 'RECRUIT'
    
    if ws.cell(row=row, column=10).value == interior_regex or 'INTERIOR' in ws.cell(row=row,column=10).value or 'GAEC' in ws.cell(row=row, column=7).value:
        ws.cell(row=row,column=17).value = 'Y'
    else:
        ws.cell(row=row,column=17).value='N'
        
    
    if 'PUPIL' in ws.cell(row=row,column=6).value and 'TEACHER' in ws.cell(row=row,column=6).value:
        ws.cell(row=row,column=18).value = 'Y'
    else:
        ws.cell(row=row, column=18).value = 'N'
        
    
    if 'TEACHER' in ws.cell(row=row,column=6).value and 'TRAINEE' in ws.cell(row=row,column=6).value:
        ws.cell(row=row,column=19).value = 'Y'
    else:
        ws.cell(row=row,column=19).value ='N'
        print(str(row) + 'exemptions')
		
		
#setting ssn nulls
for row in range(2, ws.max_row + 1):
    if ws.cell(row=row,column=11).value==None:
       ws.cell(row=row,column=11).value='NONE'
              
    print(str(row) + 'ssn')

print('saving sql upload')
wb.save(name + '_sql_upload.xlsx')



print('starting date formatting')
for row in range(2,ws.max_row+1):
    
    if ws.cell(row=row,column=3).value == 'NONE':
       ws.cell(row=row,column=3).value='1111-11-01'
    else:    
        current_date = parse(str(ws.cell(row=row,column=3).value))
        ws.cell(row=row,column=3).value=current_date.strftime(newdate)
        print(str(row))
        
    if ws.cell(row=row,column=8).value == 'NONE':   
       ws.cell(row=row,column=8).value ='1111-11-01'
    else:
        current_date = parse(str(ws.cell(row=row, column=8).value))
        ws.cell(row=row, column=8).value = current_date.strftime(newdate)
        print(str(row))
print('saving data formatting')
wb.save(name+'_CUT_UPPERED_FORMATTED_DATED_DONE.xlsx');
print('saved')

from openpyxl import Workbook
from openpyxl import load_workbook

#load mark.xlsx
wb = load_workbook('mark.xlsx')
#select mark.xlsx
sheet = wb.active

# d2 = sheet['D2'] to access D2 cell

#getting max rows
max_row = sheet.max_row
#getting max column
max_column = sheet.max_column
print(max_row)
print(max_column)

students = []

#inputing value into students array
#iterate over all cells
for i in range(2,max_row+1):
    # for j in range (0,max_column)
        student = [0,0,0,0,0,0,0,0,0,0]
        student[0]=sheet.cell(row=i,column=1).value
        student[1]=(sheet.cell(row=1,column=2).value,sheet.cell(row =i,column=2).value)
        student[2]=sheet.cell(row=1,column=3).value,sheet.cell(row =i,column=3).value
        student[3]=sheet.cell(row=1,column=4).value,sheet.cell(row =i,column=4).value
        student[4]=sheet.cell(row=1,column=5).value,sheet.cell(row =i,column=5).value
        student[5]=sheet.cell(row=1,column=6).value,sheet.cell(row =i,column=6).value
        student[6]=sheet.cell(row=1,column=7).value,sheet.cell(row =i,column=7).value
        student[7]=sheet.cell(row=1,column=8).value,sheet.cell(row =i,column=8).value
        student[8]=sheet.cell(row=1,column=9).value,sheet.cell(row =i,column=9).value
        student[9]=sheet.cell(row=1,column=10).value,sheet.cell(row =i,column=10).value
        students.append(student)

for stud in students:
    print(stud)

fail_count = 0
marks=[]
#traversing through each student
for student_item in students:
    total_mark = 0
    #traversing through each subject column of student_item
    for j in range (1,len(student_item)):
        if(student_item[j][1]=='F' or student_item[j][1] == 'FE'):
            fail_count+=1
        if(student_item[j][1]=='O'):total_mark+=10
        elif(student_item[j][1]=='A+'):total_mark+=9
        elif(student_item[j][1]=='A'):total_mark+=8.5
        elif(student_item[j][1]=='B+'):total_mark+=8
        elif(student_item[j][1]=='B'):total_mark+=7
        elif(student_item[j][1]=='C'):total_mark+=6
        elif(student_item[j][1]=='P'):total_mark+=5
        else: pass
    marks.append((student_item[0],total_mark))

print(fail_count)
print("total marks : ",marks)

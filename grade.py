import csv
from xlrd import open_workbook
import xlsxwriter

grades={}
name_map={}

def id_of_student(name):
    for key in name_map:
        if key.lower().strip() in name.lower().strip():
            return name_map[key]
    return ""

def filter_grades():
    for key in grades:

        if 'id' not in grades[key]:
            grades[key]['Id']=key
        if 'Assignment' not in grades[key]:
            grades[key]['Assignment']=0.0
        if 'Attendance' not in grades[key]:
            grades[key]['Attendance']=0.0
        if 'Quiz1' not in grades[key]:
            grades[key]['Quiz1']=0.0
        if 'Quiz2' not in grades[key]:
            grades[key]['Quiz2']=0.0
        if 'Quiz3' not in grades[key]:
            grades[key]['Quiz3']=0.0
        besttwo=[grades[key]['Quiz3'],grades[key]['Quiz2'],grades[key]['Quiz1']]
        besttwo=sorted(besttwo, key = lambda x:float(x))
        grades[key]['Quiz']=besttwo[0]+besttwo[1]


def add_grades():
    for key in grades:

        total=0
        grade='F'
        for sub_key in grades[key]:
            if sub_key!='Id':
               total+=float(grades[key][sub_key])
        if total>= 50 and total<60 :
            grade='D'
        elif total>= 60 and total<65 :
            grade='D+'
        elif total>= 65 and total<70 :
            grade='C'
        elif total>= 70 and total<75 :
            grade='C+'
        elif total>= 75 and total<80 :
            grade='B'
        elif total>= 80 and total<85 :
            grade='B+'
        elif total>= 85 and total<90 :
            grade='A'
        elif total>= 90 :
            grade='A+'

        grades[key]['Grade']=grade



# open csv file and assign assignment
with open('Assignment.csv') as csv_file:
  csv_reader = csv.reader(csv_file, delimiter=',')
  line_count = 0
  for row in csv_reader:
    if line_count > 0:
        if row[1] not in grades:
          grades[row[1]]={}
        grades[row[1]]['Assignment']=row[3]
        name_split=row[2].split(',')
        if len(name_split)>1:
            name_map[name_split[1]]=row[1]
        else:
            name_map[name_split[0]] = row[1]
    line_count+=1
# print(name_map)
# # open csv file and assign week 1 lab
with open('Week 1 Lab .csv',encoding='utf-16') as csv_file:
  csv_reader = csv.reader(csv_file, delimiter='\t')
  line_count = 0
  for row in csv_reader:
    if line_count > 0:
        id=id_of_student(row[0])
        if id !="" and id not in grades:
            grades[id] = {}
            if "join" in row[1].lower() :
                if "Attendance" not in grades[id]:
                    grades[id]['Attendance']=0
                if grades[id]['Attendance']<2:
                    grades[id]['Attendance']+=2
                # grades[row[1]]['Assignment']=row[3]
    line_count+=1

# # open csv file and assign week 1 theory
with open('Week 1 Theory.csv',encoding='utf-16') as csv_file:
  csv_reader = csv.reader(csv_file, delimiter='\t')
  line_count = 0
  for row in csv_reader:
    if line_count > 0:
        id=id_of_student(row[0])
        if id !="":
            if id not in grades:
                grades[id] = {}
            if "join" in row[1].lower().strip() :
                if "Attendance" not in grades[id]:
                    grades[id]['Attendance']=0
                if grades[id]['Attendance']<4:
                    grades[id]['Attendance']+=2
                # grades[row[1]]['Assignment']=row[3]
    line_count+=1

# # open csv file and assign Week 2 Theory.csv
with open('Week 2 Theory.csv',encoding='utf-16') as csv_file:
  csv_reader = csv.reader(csv_file, delimiter='\t')
  line_count = 0
  for row in csv_reader:
    if line_count > 0:
        id=id_of_student(row[0])
        if id !="" :
            if id not in grades:
                grades[id] = {}
            if "join" in row[1].lower().strip():
                if "Attendance" not in grades[id]:
                    grades[id]['Attendance']=0
                if grades[id]['Attendance']<6:
                    grades[id]['Attendance']+=2
                # grades[row[1]]['Assignment']=row[3]
    line_count+=1

# # open csv file and assign Week 4 Lab (Makeup).csv
with open('Week 4 Lab (Makeup).csv',encoding='utf-16') as csv_file:
  csv_reader = csv.reader(csv_file, delimiter='\t')
  line_count = 0
  for row in csv_reader:
    if line_count > 0:
        id=id_of_student(row[0])
        if id !="" :
            if id not in grades:
                grades[id] = {}
            if "join" in row[1].lower().strip():
                if "Attendance" not in grades[id]:
                    grades[id]['Attendance']=0
                if grades[id]['Attendance']<8:
                    grades[id]['Attendance']+=2
                # grades[row[1]]['Assignment']=row[3]
    line_count+=1

# # open csv file and assign Week 5 Lab.csv
with open('Week 5 Lab.csv',encoding='utf-16') as csv_file:
  csv_reader = csv.reader(csv_file, delimiter='\t')
  line_count = 0
  for row in csv_reader:
    if line_count > 0:
        id=id_of_student(row[0])
        if id !="" :
            if id not in grades:
                grades[id] = {}
            if "join" in row[1].lower().strip():
                if "Attendance" not in grades[id]:
                    grades[id]['Attendance']=0
                if grades[id]['Attendance']<10:
                    grades[id]['Attendance']+=2
                # grades[row[1]]['Assignment']=row[3]
    line_count+=1

# open xlsx file and assign quize1 number
workbook = open_workbook(filename="Quiz 1.xlsx")
worksheet = workbook.sheet_by_name("Sheet1")
line_count=0
for curr_row in range(0, worksheet.nrows):
    if line_count>0:
        student_id = worksheet.cell_value(curr_row, 4)
        student_id=student_id.partition('@')[0]
        mark=worksheet.cell_value(curr_row, 5)
        if student_id not in grades:
              grades[student_id]={}
        grades[student_id]['Quiz1']=mark
    line_count+=1

# open xlsx file and assign quize2 number
workbook = open_workbook(filename="Quiz 2.xlsx")
worksheet = workbook.sheet_by_name("Sheet1")
line_count=0
for curr_row in range(0, worksheet.nrows):
    if line_count>0:
        student_id = worksheet.cell_value(curr_row, 4)
        student_id=student_id.partition('@')[0]
        mark=worksheet.cell_value(curr_row, 5)
        if student_id not in grades:
              grades[student_id]={}
        grades[student_id]['Quiz2']=mark
    line_count+=1

# open xlsx file and assign quize3 number
workbook = open_workbook(filename="Quiz 3.xlsx")
worksheet = workbook.sheet_by_name("Sheet1")
line_count=0
for curr_row in range(0, worksheet.nrows):
    if line_count>0:
        student_id = worksheet.cell_value(curr_row, 4)
        student_id=student_id.partition('@')[0]
        mark=worksheet.cell_value(curr_row, 5)
        if student_id not in grades:
              grades[student_id]={}
        grades[student_id]['Quiz3']=mark
    line_count+=1

# open xlsx file and assign Lab number
workbook = open_workbook(filename="Lab Exam.xlsx")
worksheet = workbook.sheet_by_name("Sheet1")
line_count=0
for curr_row in range(0, worksheet.nrows):
    if line_count>0:
        student_id = worksheet.cell_value(curr_row, 4)
        student_id=student_id.partition('@')[0]
        mark=worksheet.cell_value(curr_row, 5)
        if student_id not in grades:
              grades[student_id]={}
        grades[student_id]['Lab']=mark
    line_count+=1

filter_grades()
add_grades()
print(grades)

workbook = xlsxwriter.Workbook('Grades.xlsx')
worksheet = workbook.add_worksheet()
row=0
column=0
items=['Id','Assignment','Attendance','Lab','Quiz1','Quiz2','Quiz3','Quiz','Grade']

for item in items:
    worksheet.write(row, column, item)
    column+=1
column=0
row=0
for key in grades:
    row+=1
    column=0
    for item in items:
        worksheet.write(row, column,grades[key][item])
        column+=1
workbook.close()
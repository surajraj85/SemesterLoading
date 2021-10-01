import json
import math
import xlsxwriter
import pandas as pd
# https://xlsxwriter.readthedocs.io/index.html
# with open('raw.json') as f:
#   data = json.load(f)

excel_data_df = pd.read_excel('input/raw.xlsx', sheet_name='Sheet1')
print(excel_data_df.head())

my_dict = {}
len_course = excel_data_df['courseList'].count()
print(len_course)
course_list = excel_data_df['courseList'].tolist()
print(course_list)
pattern = excel_data_df['pattern'].tolist()
print(pattern)

dictionary = dict(zip(course_list, pattern))
print(dictionary)
program_name = excel_data_df['programName'].values[0]
print(program_name)
year = excel_data_df['year'].values[0]
print(year)
term = excel_data_df['term'].values[0]
print(term)
totalEnrollmentplanned = excel_data_df['totalEnrollmentplanned'].values[0]
print(totalEnrollmentplanned)
plannedStudents = excel_data_df['plannedStudents'].values[0]
print(plannedStudents)

course_lst=[]
lec_and_lab=[]
hrs_week = []
planned_students = []
section = []
pattern = []
room_type = []
final_exam = []
instructor = []
comments = []
len_sections = math.ceil(int(totalEnrollmentplanned)/int(plannedStudents))
print(len_sections)

for key,val in dictionary.items():
    for i in range(len_sections):
        if "-" in str(val):
            course_lst.append(key)
            lec_and_lab.append("Lecture - Faculty")
            hrs_week.append(str(val.split("-")[0]))
            planned_students.append(plannedStudents)
            section.append("00"+str(i+1))
            pattern.append(str(val))
            room_type.append("Class room with Wi-Fi")
            final_exam.append(" ")
            instructor.append("Instructor Name(WIP)")
            comments.append(" ")

            course_lst.append(key)
            lec_and_lab.append("Lab - Faculty")
            hrs_week.append(str(val.split("-")[1]))
            planned_students.append(plannedStudents)
            section.append("00" + str(i + 1))
            pattern.append("")
            room_type.append("Class room with Wi-Fi")
            final_exam.append(" ")
            instructor.append("Instructor Name(WIP)")
            comments.append(" ")
        else:
            course_lst.append(key)
            lec_and_lab.append("Lecture - Faculty")
            hrs_week.append(str(val))
            planned_students.append(plannedStudents)
            section.append("00" + str(i + 1))
            pattern.append(str(val))
            room_type.append("Class room with Wi-Fi")
            final_exam.append(" ")
            instructor.append("Instructor Name(WIP)")
            comments.append(" ")

print(course_lst)
print(section)
print(planned_students)
print(lec_and_lab)
print(hrs_week)
print(pattern)
print(room_type)
print(final_exam)
print(instructor)
print(comments)

my_dict['Course ID']=course_lst
my_dict['Section']=section
my_dict['Planned Students']=planned_students
my_dict['Component']=lec_and_lab
my_dict['hrs/Wk']=hrs_week
my_dict['Pattern']=pattern
my_dict['Room Type']=room_type
my_dict['Final Exam?']=final_exam
my_dict['Recommended Instructor']=instructor
my_dict['Comments']=comments

print(my_dict)

workbook = xlsxwriter.Workbook('output/SemesterLoadingExcelOutput.xlsx')
worksheet = workbook.add_worksheet()
header_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'valign': 'top',
    'fg_color': '#D7E4BC',
    'border': 1})

col_num = 0
for key, value in my_dict.items():
    worksheet.write(9, col_num, key,header_format)
    worksheet.write_column(10, col_num, value)
    col_num += 1

# new additions-1
bold_course = workbook.add_format({'bold': True})
bold_course.set_font_size(25)
center = workbook.add_format({'align': 'center'})

worksheet.write_rich_string('A1',bold_course,
                            'Data Analytics for Business',
                            bold_course,'(Post Graduate)')

bold_schoolname = workbook.add_format({'bold': True})
bold_schoolname.set_font_size(15)

worksheet.write_rich_string('A3',bold_schoolname,
                            'Zekelman School of Business ',
                            bold_schoolname,'& Information Technology')

worksheet.write_rich_string('A5',bold_schoolname,
                            'Program',
                            bold_schoolname,': ')
worksheet.write_rich_string('A6',bold_schoolname,
                            'Campus',
                            bold_schoolname,': ')

worksheet.write_rich_string('B5',bold_schoolname,
                            'B018',
                            bold_schoolname,' ')
worksheet.write_rich_string('B6',bold_schoolname,
                            'Downtown Campus',
                            bold_schoolname,' ')

worksheet.write_rich_string('A8',bold_schoolname,
                            'AAL:',
                            bold_schoolname,' 01')

worksheet.write_rich_string('C8',bold_schoolname,
                            'Total Enrollment Planned: ',
                            bold_schoolname,str(totalEnrollmentplanned))

worksheet.write_rich_string('I3',bold_schoolname,
                            'YEAR: ',
                            bold_schoolname,'2022')


bold_normaltext = workbook.add_format({'bold': True})
bold_normaltext.set_font_size(11)

worksheet.write_rich_string('F7',bold_normaltext,
                            'Our program is BYOD(Laptop) and our students ',
                            bold_normaltext,'require both Wi-Fi and power in the classroom.')

format1 = workbook.add_format({'bg_color': '#87CEFA'})

worksheet.conditional_format('D2:D1000', {'type':     'text',
                                    'criteria': 'containing',
                                    'value':    "Lab - Faculty",
                                    'format':   format1})
workbook.close()
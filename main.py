import pandas as pd
import openpyxl
import xlrd
from pathlib import Path

df = pd.read_csv('D:/Work/Enrollment Management/Retention Data/Copy of Retention Data short_data.csv')
pd.set_option('display.max_columns', None)
df.sort_values(by=['EMPLID', 'STRM'], ascending=[True, False])
df['Course'] = df['SUBJECT'] + ' ' + df['CATALOG_NBR']
print(df)


student_id = []
for i in range(len(df)-1):
    if df.loc[i, 'EMPLID'] not in student_id:
        student_id.append(df.loc[i,'EMPLID'])

one_semester_student_list = []
for id in student_id:
    for i in range(len(df)-1):
        if id == df.loc[i, 'EMPLID']:
            if df.loc[i, 'STRM'] > 1209:
                break
            else:
                if id not in one_semester_student_list:
                    one_semester_student_list.append(id)


boolean_df = df.EMPLID.isin(one_semester_student_list)
one_semester_student_df = df[boolean_df]
home = Path.home()
save_file = Path(home, 'Desktop', 'One Semester Students.xlsx')
one_semester_student_df.to_excel(save_file)
print(one_semester_student_df)
# one_semester_student_df['Course'] = one_semester_student_df['SUBJECT'] + ' ' + one_semester_student_df['CATALOG_NBR']
print(one_semester_student_df)
Fall2020_Only = one_semester_student_df[one_semester_student_df['STRM']==1209]
Fall2020_Only_Reset = Fall2020_Only.reset_index(drop=True)
print(Fall2020_Only_Reset)
Grade_list = ['A','B', 'C', 'D', 'F', 'FW', 'NG', 'P', 'NP', 'W']
grade_boolean_df = Fall2020_Only_Reset.CRSE_GRADE_OFF.isin(Grade_list)
Grade_df = Fall2020_Only_Reset[grade_boolean_df]
Reset_Grade_df = Grade_df.reset_index(drop=True)

fail_grade_list = ['F', 'FW', 'NG', 'NP', 'W']
for i in range(len(Reset_Grade_df)):
    # if Reset_Grade_df.loc[i, 'CRSE_GRADE_OFF'] in fail_grade_list:
    #     Reset_Grade_df['Grade Point'] = 0
    if Reset_Grade_df.loc[i, 'CRSE_GRADE_OFF'] == 'A':
        Reset_Grade_df.loc[i,'Grade Point'] = 4
    elif Reset_Grade_df.loc[i, 'CRSE_GRADE_OFF'] == 'B':
        Reset_Grade_df.loc[i, 'Grade Point'] = 3
    elif Reset_Grade_df.loc[i, 'CRSE_GRADE_OFF'] == 'C' or Reset_Grade_df.loc[i, 'CRSE_GRADE_OFF'] == 'P':
        Reset_Grade_df.loc[i, 'Grade Point'] = 2
    elif Reset_Grade_df.loc[i, 'CRSE_GRADE_OFF'] == 'D':
        Reset_Grade_df.loc[i, 'Grade Point'] = 1
    else:
        Reset_Grade_df.loc[i, 'Grade Point'] = 0
print(Reset_Grade_df)

fall_student_id = []

for i in range(len(Reset_Grade_df)-1):
    if Reset_Grade_df.loc[i, 'EMPLID'] not in fall_student_id:
        fall_student_id.append(Reset_Grade_df.loc[i, 'EMPLID'])

class CourseCount:

    def __init__(self, id):
        self.id = id

    def number_courses_enrolled_in(self):
        course_count = 0
        passed_count = 0
        for i in range(len(Reset_Grade_df)):
            print(i, id, Reset_Grade_df.loc[i, 'EMPLID'])
            if id == Reset_Grade_df.loc[i, 'EMPLID']:
                course_count += 1
                if Reset_Grade_df.loc[i, 'Grade Point'] >= 3:
                    passed_count += 1
                print(Reset_Grade_df.loc[i, 'Grade Point'])
                print(course_count)
                print(passed_count)
                if id != Reset_Grade_df.loc[i+1,'EMPLID']:
                    break
        if course_count == 1:
             course_taken = Reset_Grade_df.loc[i,'Course']
        else:
            course_taken = ''
        return course_count, passed_count, course_taken




class CourseCountReport:

    columns = ['Student_ID', 'Course_Count', 'Passed_Count', 'Course_Taken']
    count_df = pd.DataFrame(columns=columns)

    def __init__(self, student_id, course_count, passed_count, course_taken):
        self.student_id = student_id
        self.course_count = course_count
        self.passed_count = passed_count
        self.course_taken = course_taken

    def count_report(self):
        length = len(CourseCountReport.count_df)
        CourseCountReport.count_df.loc[length, 'Student_ID'] = self.student_id
        CourseCountReport.count_df.loc[length, 'Course_Count'] = self.course_count
        CourseCountReport.count_df.loc[length, 'Passed_Count'] = self.passed_count
        if course_count == 1:
            CourseCountReport.count_df.loc[length, 'Course_Taken'] = self.course_taken
        print(CourseCountReport.count_df)
        return CourseCountReport.count_df


for id in fall_student_id:
    counts = CourseCount(id=id)
    course_count, passed_count, course_taken = counts.number_courses_enrolled_in()
    report = CourseCountReport(student_id=id, course_count=course_count, passed_count=passed_count, course_taken=course_taken)
    count_df = report.count_report()

home = Path.home()
save_file = Path(home, 'Desktop', 'Courses and Grades.xlsx')
count_df.to_excel(save_file)

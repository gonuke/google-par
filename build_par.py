import gspread
from oauth2client.service_account import ServiceAccountCredentials


# use creds to create a client to interact with the Google Drive API
scope = ['https://spreadsheets.google.com/feeds']
creds = ServiceAccountCredentials.from_json_keyfile_name('ep-par-processing.json', scope)
client = gspread.authorize(creds)

# Find a workbook by name and open the first sheet
# Make sure you use the right name here.
book = client.open("PAR Data")


def print_course_list(book,year):

    semester_list = ['Spring','Summer','Fall']

    single_rule = " \\\\ \hline"
    part_rule = " \\\\ \cline{2-4}"
    double_rule = single_rule + "\hline"
    header = "\section{List of Courses Taught}\n" + \
             "\\begin{centering}\n" + \
             "\\begin{tabular}{|l|l|c|l|}\\hline\n" + \
             "\\textbf{Semester} & \\textbf{Course} & \\textbf{\# of Students} & \\textbf{Role}" + double_rule
    empty_semester = "\multicolumn{3}{|c|}{\emph{none}} " + double_rule
    footer = "\\end{tabular}\n \end{centering}"

    full_history = book.worksheet('CourseHistory').get_all_records()

    print(header)

    for semester in semester_list:
        course_list = [e for e in full_history if (e['YEAR'],e['SEMESTER']) == (year,semester)]
        if (len(course_list) > 0):
            first_col = "\multirow{" + str(len(course_list)) + "}{*}{" + semester + "} & "
            rows = [" & ".join([entry['COURSEID'],str(entry['STUDENTS']),entry['ROLE']]) for entry in course_list]
            print(first_col + (part_rule + "\n & ").join(rows) + double_rule)
        else:
            print(semester + " & " + empty_semester)
    print(footer)

def print_future_courses(book):

    all_courses = book.worksheet('CourseInfo').get_all_records()

    header = "\section{List of Courses for Future}\n"

    status_list = ['PREP','REPEAT','INTEREST']
    status_strings = {'PREP':'Courses I am prepared to teach',
                      'REPEAT':'Courses I have taught and could teach again',
                      'INTEREST':'Courses I am interested in but not prepared to teach'}
    print(header)
    for status in status_list:
        print(status_strings[status] + ": " + \
              ", ".join([c['COURSEID'] for c in all_courses if c['PREPSTATUS'] == status]) + \
              "\\\\")

section_sep = "\n\n%%\n%% "



print("""\
\documentclass{article}

\usepackage{ep_par}

\\newcommand{\paryear}{2017}
\\newcommand{\parperson}{Paul P.\ H.\ Wilson}
\\begin{document}

\partitle
""")

print(section_sep + "Course List")    
print_course_list(book,2017)

print(section_sep + "Course Prep")    
print_future_courses(book)



print("\end{document}")

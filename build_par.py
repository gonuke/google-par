import gspread
from oauth2client.service_account import ServiceAccountCredentials

year = 2017

# use creds to create a client to interact with the Google Drive API
scope = ['https://spreadsheets.google.com/feeds']
creds = ServiceAccountCredentials.from_json_keyfile_name('ep-par-processing.json', scope)
client = gspread.authorize(creds)

# Find a workbook by name and open the first sheet
# Make sure you use the right name here.
book = client.open("PAR Data")

timeunit = [("WEEK", "wk"),
            ("MONTH", "month"),
            ("SEMESTER", "sem"),
            ("YEAR", "yr")]

def get_course_list(book,year):

    semester_list = ['Spring','Summer','Fall']

    single_rule = " \\\\ \hline"
    part_rule = " \\\\ \cline{2-4}"
    double_rule = single_rule + "\hline"
    empty_semester = "\multicolumn{3}{c|}{\emph{none}} " + double_rule

    course_list_str = "\section{Courses Taught}\n" + \
             "\\begin{centering}\n" + \
             "\\begin{tabular}{|l|l|c|l|}\\hline\n" + \
             "\\textbf{Semester} & \\textbf{Course} & \\textbf{\# of Students} & \\textbf{Role}" + double_rule
    footer = "\\end{tabular}\n \end{centering}"

    full_history = book.worksheet('CourseHistory').get_all_records()

    for semester in semester_list:
        course_list = [e for e in full_history if (e['YEAR'],e['SEMESTER']) == (year,semester)]
        if (len(course_list) > 0):
            first_col = "\multirow{" + str(len(course_list)) + "}{*}{" + semester + "} & "
            rows = [" & ".join([entry['COURSEID'],str(entry['STUDENTS']),entry['ROLE']]) for entry in course_list]
            course_list_str += first_col + (part_rule + "\n & ").join(rows) + double_rule + "\n"
        else:
            course_list_str += semester + " & " + empty_semester + "\n"

    return course_list_str + footer

def get_future_courses(book):

    all_courses = book.worksheet('CourseInfo').get_all_records()

    future_course_str = "\section{Course Interest for Future}\n"

    status_list = ['PREP','REPEAT','INTEREST']
    status_strings = {'PREP':'Courses I am prepared to teach',
                      'REPEAT':'Courses I have taught and could teach again',
                      'INTEREST':'Courses I am interested in but not prepared to teach'}
    for status in status_list:
        future_course_str += status_strings[status] + ": " + \
              ", ".join([c['COURSEID'] for c in all_courses if c['PREPSTATUS'] == status]) + \
              "\\\\ \n"

    return future_course_str

def is_current(record,year):

    if 'YEAR' in record.keys():
        return record['YEAR'] == year
    elif 'STARTYEAR' in record.keys():
        return (record['STARTYEAR'] <= year and record['ENDYEAR'] >= year)
    else:
        return False

def extend_advising_list(advising_str, num_students, student_type_str, add_comma):

    if (num_students>0):
        if (add_comma):
            advising_str += ", "
        advising_str += str(num_students) + student_type_str
        add_comma = True

    return (advising_str, add_comma)


def get_student_advising_info(book,year):

    # Student advisees
    all_advising = book.worksheet('AdviseeList').get_all_records()
    
    current_advisees = [s for s in all_advising if is_current(s,year)]

    ne_ugrads = [s for s in current_advisees if (s['TYPE'] == "U" and s['PROGRAM'] == "NE")]
    neep_grads = [s for s in current_advisees if (s['TYPE'] == "G" and s['PROGRAM'] == "NEEP")]

    other_grads = [s for s in current_advisees if (s['TYPE'] == "G" and s['PROGRAM'] != "NEEP")]
    other_grad_majors = {s['PROGRAM'] for s in other_grads}
    other_grad_string = " other graduate students (" + ", ".join(other_grad_majors) + ")"
    
    add_comma = False
    if (len(current_advisees)>0):
        advising_str = "Academic advisor for "
    (advising_str, add_comma) = extend_advising_list(advising_str, len(ne_ugrads), " NE undergraduates", add_comma)
    (advising_str, add_comma) = extend_advising_list(advising_str, len(neep_grads), " NE graduate students", add_comma)
    (advising_str, add_comma) = extend_advising_list(advising_str, len(other_grads), other_grad_string, add_comma)

    return advising_str + "."

def get_org_advising_info(book,year):

    # Org advisees

    org_history = book.worksheet('StudentOrgAdvising').get_all_records()
    org_info_list = book.worksheet('StudentOrgList').get_all_records()
    
    current_orgs = [o for o in org_history if is_current(o,year)]

    if (len(current_orgs) > 0):
        org_strs = []
        for org in current_orgs:
            code = org['ORGCODE']
            name = [org_info['ORGNAME'] for org_info in org_info_list if org_info['ORGCODE'] == code]
            org_strs.append(name[0] + " (" + str(org['WEEKLYHOURS']) + " hours/wk)")
        advising_str = "Faculty advisor for the " + ", ".join(org_strs)
        
    return advising_str + "."


def get_advising_info(book,year):


    advising_str = "\section{Advising Responsibilities}\n\n"

    advising_str += get_student_advising_info(book,year) + "\\\\ \n"

    advising_str += get_org_advising_info(book,year)
    
    return advising_str

def expand_table(table,cat_table,join_key):

    catalog = {e[join_key] : e for e in cat_table}
    
    for e in table:
        for (k,v) in catalog[e[join_key]].items():
            e[k] = v

    return table

def commit_list(services):

    commit_str = ""
    # list in order of weekly, monthly, semesterly, annually
    for time_key, time_str in timeunit:
        for svc in [s for s in services if s['COMMITMENTUNIT'] == time_key]:
            commit_str += svc['NAME'] + " (" + str(svc['COMMITMENTQUANTITY']) + " hrs/" + \
                          time_str + ")\\\\\n"

    return commit_str

def soc_svc(services,society):

    name = [svc['NAME'] for svc in services if svc['SERVICECODE'] == society]
    
    cat_svc_str = "\subsubsection{" + name[0] + "}\n\n"

    cat_svc_str += commit_list([s for s in services if s['SOCIETY'] == society])

    return cat_svc_str

def get_service_cat(services,cat,cat_str):

    cat_svc = [s for s in services if s['CATEGORY'] == cat]
    
    cat_svc_str = "\subsection{To " + cat_str + "}\n\n"

    if (len(cat_svc) < 1):
        cat_svc_str += "\emph{none} \\\\ \n"
    else:
        if (cat != "SOCIETY"):
            cat_svc_str += commit_list(cat_svc)
        else:
            for soc in {s['SOCIETY'] for s in cat_svc if s['SOCIETY'] != ""}:
                cat_svc_str += soc_svc(cat_svc,soc)
            
    return cat_svc_str + "\n\n"

def get_service(book,year):

    current_services = [s for s in book.worksheet('Service').get_all_records() if is_current(s,year) ]
    service_catalog = book.worksheet('ServiceList').get_all_records()

    current_services = expand_table(current_services, service_catalog, "SERVICECODE")
    current_services.extend([s for s in service_catalog if (s['CATEGORY'] == 'SOCIETY' and s['SOCIETY'] == "")])
    
    service_str = "\section{Service}\n\n"

    service_cat_list = [("EP", "the Engineering Physics Department"),
                        ("COE", "the College of Engineering"),
                        ("UW", "the UW Campus and State of Wisconsin"),
                        ("NATIONAL", "National Groups and Other Universities/Institutions"),
                        ("SOCIETY", "Professional Societies")]
    
    for (cat,cat_str) in service_cat_list:
        service_str += get_service_cat(current_services, cat, cat_str)
        
    return service_str

section_sep = "\n\n%%\n%% "

print("""\
\documentclass[12pt]{article}

\usepackage{ep_par}

\\newcommand{\paryear}{
""" + str(year) + """\
}
\\newcommand{\parperson}{Paul P.\ H.\ Wilson}
\\begin{document}

\partitle
""")

print(section_sep + "Course List")    
print(get_course_list(book,year))

print(section_sep + "Course Prep")    
print(get_future_courses(book))

print(section_sep + "Course Dev")
print("""\
\section{Course Development Activities}
\\todo{Add course development}
""")

print(section_sep + "Student Advising")
print(get_advising_info(book,year))

print(section_sep + "Service")
print(get_service(book,year))

print("\end{document}")

import gspread
from oauth2client.service_account import ServiceAccountCredentials

from datetime import datetime

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

section_sep = "\n\n%%\n%% "

newline = "\\\\ \n"

# utility function can check whether a record is current, either because
# YEAR == year, or
# STARTYEAR <= year <= ENDYEAR
def is_current(record,year):

    if 'YEAR' in record.keys():
        return record['YEAR'] == year
    elif 'STARTYEAR' in record.keys():
        return (record['STARTYEAR'] <= year and record['ENDYEAR'] >= year)
    elif 'DATE' in record.keys():
        date_obj = datetime.strptime(record['DATE'],"%m/%d/%y")
        return date_obj.year == year
    else:
        return False

# utility function to get all current record from worksheetx
def filter_for_current(worksheet,year):

    return [s for s in worksheet.get_all_records() if is_current(s,year)]

# utility function to start a LaTeX table with some set of headings
def build_table_header(col_headings):

    single_rule = " \\\\ \hline"
    double_rule = single_rule + "\hline"
    col_widths = [("p{" + str(frac) + "\\textwidth}|") for (head,frac) in col_headings]
    col_heads = [("\\textbf{" + head + "}") for (head,frac) in col_headings]
    header =  "\\begin{centering}\n" + \
              "\\begin{tabular}" + \
              "{|" + "".join(col_widths)  + "}" + \
              "\\hline\n" +  " & ".join(col_heads) + double_rule + " \n"

    return header


# utility function to add columns to one table based on a shared key with another table
# for every row in `table` add all the columns from `cat_table`
#   from the row with matching `join_key`
def expand_table(table,cat_table,join_key):

    catalog = {e[join_key] : e for e in cat_table}
    
    for e in table:
        for (k,v) in catalog[e[join_key]].items():
            e[k] = v

    return table

# list of courses
def get_course_list(book,year):

    semester_list = ['Spring','Summer','Fall']

    single_rule = " \\\\ \hline"
    part_rule = " \\\\ \cline{2-4}"
    double_rule = single_rule + "\hline\n"
    empty_semester = "\multicolumn{3}{c|}{\emph{none}} " + double_rule

    column_headings = [("Semester",0.25),
                       ("Course",0.25),
                       ("\# of Students",0.25),
                       ("Role",0.25)]
    
    course_list_str = "\section{Courses Taught}\n" + \
                      build_table_header(column_headings)
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

# info about future course interests
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
              newline

    return future_course_str


# information about individual advisees
def get_student_advising_info(book,year):

    adv_str_list = []
    # Student advisees
    current_advisees = filter_for_current(book.worksheet('AdviseeList'), year)
    
    num_ne_ugrads = len([s for s in current_advisees if (s['TYPE'] == "U" and s['PROGRAM'] == "NE")])
    if num_ne_ugrads > 0:
        adv_str_list.append(str(num_ne_ugrads) + " NE undergraduates")

    num_neep_grads = len([s for s in current_advisees if (s['TYPE'] == "G" and s['PROGRAM'] == "NEEP")])
    if num_neep_grads > 0:
        adv_str_list.append(str(num_neep_grads) + " NE graduate students")
        
    other_grads = [s for s in current_advisees if (s['TYPE'] == "G" and s['PROGRAM'] != "NEEP")]
    num_other_grads = len(other_grads)
    if num_other_grads > 0:
        other_grad_majors = {s['PROGRAM'] for s in other_grads}
        other_grad_string = " other graduate students (" + ", ".join(other_grad_majors) + ")"
        adv_str_list.append(str(num_other_grads) + other_grad_string)

    return ", ".join(adv_str_list) + "."

# information about advising organizations
def get_org_advising_info(book,year):

    
    current_orgs = filter_for_current(book.worksheet('StudentOrgAdvising'),year)
    org_info_list = book.worksheet('StudentOrgList').get_all_records()

    advising_str = ""
    
    if (len(current_orgs) > 0):
        org_strs = []
        for org in current_orgs:
            code = org['ORGCODE']
            name = [org_info['ORGNAME'] for org_info in org_info_list if org_info['ORGCODE'] == code]
            org_strs.append(name[0] + " (" + str(org['WEEKLYHOURS']) + " hours/wk)")
        advising_str = "Faculty advisor for the " + ", ".join(org_strs)
        
    return advising_str + "."

# all advising info
def get_advising_info(book,year):


    advising_str = "\section{Advising Responsibilities}\n\n"

    advising_str += get_student_advising_info(book,year) + newline

    advising_str += get_org_advising_info(book,year)
    
    return advising_str

# string to expand the list of committees for a professional society
def committee_list(services):

    commit_str = ""
    # list in order of weekly, monthly, semesterly, annually
    for time_key, time_str in timeunit:
        for svc in [s for s in services if s['COMMITMENTUNIT'] == time_key]:
            commit_str += svc['NAME'] + " (" + str(svc['COMMITMENTQUANTITY']) + " hrs/" + \
                          time_str + ")" + newline

    return commit_str

# special formatting for service to professional societies
def soc_svc(services,society):

    name = [svc['NAME'] for svc in services if svc['SERVICECODE'] == society]
    
    cat_svc_str = "\subsubsection{" + name[0] + "}\n\n"

    cat_svc_str += committee_list([s for s in services if s['SOCIETY'] == society])

    return cat_svc_str

# get the formatted version of each different category of service
def get_service_cat(services,cat,cat_str):

    cat_svc = [s for s in services if s['CATEGORY'] == cat]
    
    cat_svc_str = "\subsection{To " + cat_str + "}\n\n"

    if (len(cat_svc) < 1):
        cat_svc_str += "\emph{none}" + newline
    else:
        if (cat != "SOCIETY"):
            cat_svc_str += committee_list(cat_svc)
        else:
            for soc in {s['SOCIETY'] for s in cat_svc if s['SOCIETY'] != ""}:
                cat_svc_str += soc_svc(cat_svc,soc)
            
    return cat_svc_str + "\n\n"

# get formatted list of reviews
def get_reviews(book,year):

    current_reviews = filter_for_current(book.worksheet('Reviews'),year)
    review_catalog = book.worksheet('ReviewSource').get_all_records()

    current_reviews = expand_table(current_reviews,review_catalog,"SOURCE")

    review_types = [('PAPER','Paper Reviews for Technical Journals and Conferences'),
                    ('PROPOSAL','Proposal Reviews')]
    
    review_strs = []
    for r_type,r_type_str in review_types:
        reviews = [s for s in current_reviews if s['REVIEWTYPE'] == r_type]
        if len(reviews) > 0:
            review_strs.append("\subsection{" + r_type_str + "}\n\n")
            for review_data in reviews:
                review_strs[-1] += review_data['NAME'] + " (" + str(review_data['NUMBER']) + ") " + newline

    return "\n\n".join(review_strs)

# get all service obligations
def get_service(book,year):

    current_services = filter_for_current(book.worksheet('Service'),year)
    service_catalog = book.worksheet('ServiceList').get_all_records()

    current_services = expand_table(current_services, service_catalog, "SERVICECODE")

    # find all the parent societies
    parent_societies = {s['SOCIETY'] for s in current_services if (s['SOCIETY'] != "")}
    current_services.extend([s for s in service_catalog if (s['SERVICECODE'] in parent_societies)])
    
    service_str = "\section{Service}\n\n"

    service_cat_list = [("EP", "the Engineering Physics Department"),
                        ("COE", "the College of Engineering"),
                        ("UW", "the UW Campus and State of Wisconsin"),
                        ("NATIONAL", "National Groups and Other Universities/Institutions"),
                        ("SOCIETY", "Professional Societies")]
    
    for (cat,cat_str) in service_cat_list:
        service_str += get_service_cat(current_services, cat, cat_str)


    service_str += get_reviews(book,year)
    
    return service_str

# get outreach
def get_outreach(book,year):

    current_outreach = filter_for_current(book.worksheet('Outreach'),year)

    outreach_strs = []
    for outreach in current_outreach:
        date_obj = datetime.strptime(outreach['DATE'],"%m/%d/%y")
        outreach_strs.append(datetime.strftime(date_obj,"%B") + " " + str(date_obj.day) +
                             ": " + outreach['AUDIENCE'] + ", \"" + outreach['TITLE'] + "\"")

    return "\section{Educational Outreach Activities}\n\n" + newline.join(outreach_strs)

def get_course_dev(book,year):

    course_dev_str = "\section{Course Development Activities}\n\n" + \
                     "\\todo{Add course development}\n\n"

    return course_dev_str

def get_honors(book,year):

    honor_str = "\section{Honors and Awards received in " + str(year) + "}\n\n" + \
                "\\todo{Add honors and awards here}\n\n"

    return honor_str

def get_patents(book,year):

    patent_str = "\section{Patents applied for or granted in " + str(year) + "}\n\n" + \
                "\\todo{Add patents here}\n\n"

    return patent_str

def get_proposal_submissions(book,year):

    submission_list = filter_for_current(book.worksheet('ProposalsAndGrants'), year)

    single_rule = " \\\\ \hline"
    double_rule = single_rule + "\hline"
    part_rule = " \\\\ \cline{2-4}"
    empty_semester = "\multicolumn{3}{c|}{\emph{none}} " + double_rule

    footer = "\\end{tabular}\n \end{centering}"

    column_headings = [("Begin Date",0.1),
                       ("End Date",0.1),
                       ("Amount [\$k]",0.1),
                       ("Topic",0.3),
                       ("Agency",0.1),
                       ("Authors",0.1),
                       ("Role",0.05)]
    column_data = ['BEGINDATE','ENDDATE','AMOUNT','TOPIC','AGENCY','CO-AUTHORS','ROLE']
    
    if len(submission_list) > 0:
        submission_list_str = "\section{Research proposals submitted during " + str(year) + "}\n\n" + \
                              build_table_header(column_headings)
        for s in submission_list:
            submission_list_str += " & ".join([str(s[d]) for d in column_data]) + single_rule
        
    return submission_list_str + footer


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
print(get_course_dev(book,year))

print(section_sep + "Student Advising")
print(get_advising_info(book,year))

print(section_sep + "Service")
print(get_service(book,year))

print(section_sep + "Educational Outreach Activities")
print(get_outreach(book,year))

print(section_sep + "Awards/Honors")
print(get_honors(book,year))

print(section_sep + "Patents")
print(get_patents(book,year))

print(section_sep + "Submitted Proposals")
print(get_proposal_submissions(book,year))


print("\end{document}")

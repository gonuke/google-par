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

newline = "\\\\ \n"

section_sep = "\n\n%%\n%% "
emph_none = "\emph{none}\n\n"

table_single_rule = " \\\\ \hline"
table_double_rule = table_single_rule + "\hline"
table_footer = "\\end{tabular}\n \end{centering}"

# add a section heading with some optional guidance
def heading(section, guidance=""):

    head_str = "\section{" + section + "}\n\n"
    if len(guidance) > 0:
        head_str += "\emph{(" + guidance + ")}\n\n "

    return head_str

# unmissible placeholder to be completed later by author
def todo(todo_str):

    return "\\todo{" + todo_str + "}\n\n"

# utility function can check whether a record is current, either because
# YEAR == year, or
# STARTYEAR <= year <= ENDYEAR
def is_current(record,year):

    if 'YEAR' in record.keys():
        return record['YEAR'] == year
    elif 'STARTYEAR' in record.keys():
        return (record['STARTYEAR'] <= year and record['ENDYEAR'] >= year)
    elif 'STARTDATE' in record.keys():
        start_date_obj = datetime.strptime(record['STARTDATE'],"%m/%d/%y")
        end_date_obj = datetime.strptime(record['ENDDATE'],"%m/%d/%y")
        return (start_date_obj.year == year or end_date_obj.year == year)
    elif 'DATE' in record.keys():
        date_obj = datetime.strptime(record['DATE'],"%m/%d/%y")
        return date_obj.year == year
    else:
        return False

# utility function to get all current record from worksheetx
def filter_for_current(worksheet,year):

    return [s for s in worksheet.get_all_records() if is_current(s,year)]

# utility function to start a LaTeX table with some set of headings
def build_table_header(col_info):

    col_widths = [("p{" + str(frac) + "\\textwidth}|") for (head,frac,key) in col_info]
    col_heads = [("\\textbf{" + head + "}") for (head,frac,key) in col_info]
    header =  "\\begin{centering}\n" + \
              "\\begin{tabular}" + \
              "{|" + "".join(col_widths)  + "}" + \
              "\\hline\n" +  " & ".join(col_heads) + table_double_rule + " \n"

    return header

# utility function to add rows to table
def build_table_rows(record_list,col_info):

    rows_str = ""
    for r in record_list:
        rows_str += " & ".join([str(r[key]) for (head,frac,key) in col_info]) + table_single_rule + "\n"
    return rows_str + table_footer


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

    part_rule = " \\\\ \cline{2-4}"
    empty_semester = "\multicolumn{3}{c|}{" + emph_none + "} " + table_double_rule

    column_info = [("Semester",0.25,''),
                   ("Course",0.25,'COURSEID'),
                   ("\# of Students",0.25,'STUDENTS'),
                   ("Role",0.25,'ROLE')]
    
    course_list_str = heading("Courses Taught","") + build_table_header(column_info)

    full_history = book.worksheet('CourseHistory').get_all_records()

    for semester in semester_list:
        course_list = [e for e in full_history if (e['YEAR'],e['SEMESTER']) == (year,semester)]
        if (len(course_list) > 0):
            first_col = "\multirow{" + str(len(course_list)) + "}{*}{" + semester + "} & "
            rows = [" & ".join([entry['COURSEID'],str(entry['STUDENTS']),entry['ROLE']]) for entry in course_list]
            course_list_str += first_col + (part_rule + "\n & ").join(rows) + table_double_rule + "\n"
        else:
            course_list_str += semester + " & " + empty_semester + "\n"

    return course_list_str + table_footer

# info about future course interests
def get_future_courses(book):

    all_courses = book.worksheet('CourseInfo').get_all_records()

    future_course_str = heading("Course Interest for Future","")

    status_strings = {'PREP':'Courses I am prepared to teach',
                      'REPEAT':'Courses I have taught and could teach again',
                      'INTEREST':'Courses I am interested in but not prepared to teach'}

    for status in ['PREP','REPEAT','INTEREST']:
        future_course_str += status_strings[status] + ": " + \
                             ", ".join([c['COURSEID'] for c in all_courses if c['PREPSTATUS'] == status]) + \
                             newline
        
    return future_course_str


# information about individual advisees
def get_student_advising_info(book,year):

    adv_str_list = []
    # Student advisees
    current_advisees = filter_for_current(book.worksheet('AdviseeList'), year)
    
    num_ne_ugrads = len([s for s in current_advisees if (s['TYPE'],s['PROGRAM']) == ("U","NE")])
    if num_ne_ugrads > 0:
        adv_str_list.append(str(num_ne_ugrads) + " NE undergraduates")

    num_neep_grads = len([s for s in current_advisees if (s['TYPE'],s['PROGRAM']) == ("G","NEEP")])
    if num_neep_grads > 0:
        adv_str_list.append(str(num_neep_grads) + " NEEP graduate students")
        
    other_grads = [s for s in current_advisees if (s['TYPE'] == "G" and s['PROGRAM'] != "NEEP")]
    num_other_grads = len(other_grads)
    if num_other_grads > 0:
        other_grad_majors = {s['PROGRAM'] for s in other_grads}
        other_grad_string = " other graduate student(s) (" + ", ".join(other_grad_majors) + ")"
        adv_str_list.append(str(num_other_grads) + other_grad_string)

    if len(adv_str_list) > 0:
        return newline.join(adv_str_list) + "."
    else:
        return emph_none

# information about advising organizations
def get_org_advising_info(book,year):

    current_orgs = expand_table(filter_for_current(book.worksheet('StudentOrgAdvising'),year),
                                book.worksheet('StudentOrgList').get_all_records(),
                                'ORGCODE')
    if (len(current_orgs) > 0):
        advising_str = "Faculty advisor for the " + \
                       ", ".join([(o['ORGNAME'] + " (" + str(o['WEEKLYHOURS']) + " hours/wk)") for o in current_orgs])
    else:
        advising_str = emph_none
        
    return advising_str

# all advising info
def get_advising_info(book,year):


    advising_str = heading("Advising Responsibilities")

    advising_str += get_student_advising_info(book,year) + newline + \
                    get_org_advising_info(book,year)
    
    return advising_str

# string to expand the list of committees for a professional society
def committee_list(services):

    timeunit = [("WEEK", " hrs/wk"),
                ("MONTH", " hrs/month"),
                ("SEMESTER", " hrs/semester"),
                ("YEAR", " hrs/yr"),
                ("COUNT", "")]

    commit_str_list = []
    # list in order of weekly, monthly, semesterly, annually
    for time_key, time_str in timeunit:
        for svc in [s for s in services if s['COMMITMENTUNIT'] == time_key]:
            commit_str_list.append(svc['NAME'] + " (" + str(svc['COMMITMENTQUANTITY']) + \
                                   time_str + ")")

    return newline.join(commit_str_list)

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

    if (len(cat_svc) > 0):
        if (cat != "SOCIETY"):
            cat_svc_str += committee_list(cat_svc)
        else:
            for soc in {s['SOCIETY'] for s in cat_svc if s['SOCIETY'] != ""}:
                cat_svc_str += soc_svc(cat_svc,soc)
    else:
        cat_svc_str += emph_none

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
            review_strs.append("\subsection{" + r_type_str + "}")
            review_str_list = [(r['NAME'] + " (" + str(r['NUMBER']) + ") ") for r in reviews]
            review_strs.append(newline.join(review_str_list))
                               
    return "\n\n".join(review_strs)

def other_student_committees(book,year):

    other_comms = filter_for_current(book.worksheet('OtherStudentCommittees'),year)

    comm_type_list = ["BS Defense", "MS Oral", "MS Defense", "PhD Prelim", "PhD Defense"]

    comm_summary_list = []
    for comm_type in comm_type_list:
        for cat in ["EP","UW","NATIONAL"]:
            name_list = [s['NAME'] for s in other_comms if (s['TYPE'],s['DEPT']) == (comm_type,cat)]
            if len(name_list) > 0:
                comm_summary_list.append({'CATEGORY':cat,
                                          'NAME': comm_type + " committees (" + ",".join(name_list) + ")",
                                          'COMMITMENTUNIT': 'COUNT',
                                          'COMMITMENTQUANTITY': len(name_list)})

    return comm_summary_list
    

# get all service obligations
def get_service(book,year):

    current_services = filter_for_current(book.worksheet('Service'),year)
    service_catalog = book.worksheet('ServiceList').get_all_records()

    current_services = expand_table(current_services, service_catalog, "SERVICECODE")

    # find all the parent societies and insert them in the service list
    parent_societies = {s['SOCIETY'] for s in current_services if (s['SOCIETY'] != "")}
    current_services.extend([s for s in service_catalog if (s['SERVICECODE'] in parent_societies)])

    current_services.extend(other_student_committees(book,year))
    
    service_str = heading("Service")

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

    return heading("Educational Outreach Activities") + newline.join(outreach_strs)

def get_course_dev(book,year):
    
    return heading("Course Development Activities") + todo("Add course development")

def get_honors(book,year):

    return heading("Honors and Awards received in " + str(year)) + todo("Add honors and awards here")


def get_patents(book,year):

    patent_list = filter_for_current(book.worksheet("Patents"),year)

    patent_str = heading("Patents applied for or granted in " + str(year))

    if len(patent_list) > 0:
        patent_str += newline.join([("Patent No. " + p['PATENTNUMBER'] + ": " + p['TITLE'] + "(" + p['STATUS'] + ")") for p in patent_list]) + newline
    else:
        patent_str += emph_none

    return patent_str

def make_grant_list(grants):

    column_info = [("Begin Date",0.1,'BEGINDATE'),
                   ("End Date",0.1,'ENDDATE'),
                   ("Amount [\$k]",0.1,'AMOUNT'),
                   ("Topic",0.3,'TOPIC'),
                   ("Agency",0.1,'AGENCY'),
                   ("Co-Authors",0.1,'CO-AUTHORS'),
                   ("Role",0.05,'ROLE')]
    
    grants_str = build_table_header(column_info) + \
                 build_table_rows(grants,column_info)

    return grants_str

def get_proposal_submissions(book,year):

    submission_list = filter_for_current(book.worksheet('ProposalsAndGrants'), year)

    submission_list_str = heading("Research proposals submitted during " + str(year))

    if len(submission_list) > 0:
        submission_list_str +=  make_grant_list(submission_list)
    else:
        submission_list_str += emph_none

    return submission_list_str

def active_grant(grant_record,year):

    begin_date_obj = datetime.strptime(grant_record['BEGINDATE'],"%m/%d/%y")
    end_date_obj = datetime.strptime(grant_record['BEGINDATE'],"%m/%d/%y")
    return begin_date_obj.year <= year and end_date_obj.year >= 2017
    

def get_active_grants(book,year):

    grant_list = book.worksheet('ProposalsAndGrants').get_all_records()
    active_grant_list = [g for g in grant_list if g['STATUS'] == 'FUNDED' and active_grant(g,year)]

    active_grant_list_str = heading("Research grants and contracts active during " + str(year))
    
    if len(active_grant_list) > 0:
        active_grant_list_str += make_grant_list(active_grant_list)
    else:
        active_grant_list_str += emph_none

    return active_grant_list_str
    
def get_consulting(book,year):

    consult_list = filter_for_current(book.worksheet("Consulting"),year)

    consulting_str = heading("Consulting agreements held in " + str(year))

    if len(consult_list) > 0:
        consulting_str += newline.join([("\\textbf{" + c['ORGANIZATION'] + ":}" + c['TOPIC']) for c in consult_list]) + newline
    else:
        consulting_str += emph_none

    return consulting_str

def make_date_range(record):

    start_date_obj = datetime.strptime(record['STARTDATE'],"%m/%d/%y")
    end_date_obj = datetime.strptime(record['ENDDATE'],"%m/%d/%y")

    return start_date_obj.strftime("%m/%d") + "-" + end_date_obj.strftime("%m/%d")
    
    
def get_meetings(book,year):

    meeting_list = filter_for_current(book.worksheet("Meetings"),year)

    for meeting in meeting_list:
        meeting['DATERANGE'] = make_date_range(meeting)
    
    column_info = [("Dates",0.15,'DATERANGE'),
                   ("Location",0.2,'LOCATION'),
                   ("Meeting",0.55,'MEETINGNAME')]

    meeting_list_str = heading("Professional Meetings and Conferences Attended in " + str(year),"")

    if len(meeting_list) > 0:
        meeting_list_str += build_table_header(column_info) + \
                            build_table_rows(meeting_list,column_info)
    else:
        meeting_list_str += emph_none

    return meeting_list_str

def get_current_grad_students(book,year):

    grad_types = {"G", "V"}
    
    research_student_list = expand_table(filter_for_current(book.worksheet("ActiveResearchStudents"),year),
                                         [s for s in filter_for_current(book.worksheet("AdviseeList"),year) if s['TYPE'] in grad_types],
                                         'LASTNAME')

    for student in research_student_list:
        student['FULLNAME'] = student['FIRSTNAME'] + " " + student['LASTNAME']

    column_info = [("Student",0.25,'FULLNAME'),
                   ("Program",0.11,'DEGREE'),
                   ("Research Topic",0.35,'TOPIC'),
                   ("Source of Support",0.2,'SOURCE')]

    research_student_str = heading("Graduate Students",
                           "Current grad students doing thesis research, " + \
                           "and non-thesis students who are active in research and " + \
                           "taking as much time as though doing there theses.")

    if len(research_student_list) > 0:
        research_student_str += build_table_header(column_info) + \
                                build_table_rows(research_student_list,column_info)
    else:
        research_student_str += emph_none
        
    return research_student_str

def get_graduated_students(book,year):

    research_student_list = expand_table(filter_for_current(book.worksheet("ActiveResearchStudents"),year),
                                         filter_for_current(book.worksheet("AdviseeList"),year),
                                         'LASTNAME')
    graduated_student_list = [s for s in research_student_list if (s['ENDYEAR'], s['TYPE']) == (year,"G")]

    for student in graduated_student_list:
        student['FULLNAME'] = student['FIRSTNAME'] + " " + student['LASTNAME']

    column_info = [("Student",0.3,'FULLNAME'),
                   ("Program",0.15,'DEGREE'),
                   ("Date",0.15,'DEFENSEDATE'),
                   ("Employer",0.3,'CURRENTEMPLOYER')]

    graduated_student_str = heading("Graduate students who graduated in " + str(year),"")
    
    if len(graduated_student_list) > 0:
        graduated_student_str += build_table_header(column_info) + \
                                 build_table_rows(graduated_student_list, column_info)
    else:
        graduated_student_str += emph_none

    return graduated_student_str

def get_staff_list(book,year):

    staff_positions = {'Scientist','Researcher'}

    staff_list = [s for s in filter_for_current(book.worksheet("Staff"),year) if s['TITLE'] in staff_positions]

    
    staff_str = heading("Research Staff",
                "Post-PhD and academic staff supervised in " + str(year))

    column_info = [("Name", 0.25,'NAME'),
                   ("Title",0.11,'TITLE'),
                   ("Research Topic",0.35,'TOPIC'),
                   ("Source of Support",0.2,'SUPPORT')]
    
    if len(staff_list) > 0:
        staff_str += build_table_header(column_info) + \
                     build_table_rows(staff_list, column_info)
    else:
        staff_str += emph_none

    return staff_str

def get_ug_list(book,year):

    ug_positions = {'U/G Hourly'}

    ug_list = [s for s in filter_for_current(book.worksheet("Staff"),year) if s['TITLE'] in ug_positions]
    
    ug_str = heading("Undergraduate researchers")

    column_info = [("Name", 0.25,'NAME'),
                   ("Position",0.11,'TITLE'),
                   ("Research Topic",0.35,'TOPIC'),
                   ("Source of Support",0.2,'SUPPORT')]
    
    if len(ug_list) > 0:
        ug_str += build_table_header(column_info) + \
                  build_table_rows(ug_list, column_info)
    else:
        ug_str += emph_none

    return ug_str

def get_personal_research(book,year):

    pr_list = filter_for_current(book.worksheet("PersonalResearch"),year)

    pr_str = heading("Personal Research",
                     "Brief description of the extent and nature of any personal research " + \
                     "(defined here as research performed independent of graduate students rather " + \
                     "than through them) and indicate the project on which this research is done.")

    if len(pr_list) > 0:
        pr_str += "\\\\ \n\n".join([pr['DESCRIPTION'] for pr in pr_list])
    else:
        pr_str += emph_none

    return pr_str

def build_publications():

    pub_str = heading("Publications list","")

    for section in ['books','journalssubmitted','journalsaccepted','journalspublished','conference','reports','invited']:
        pub_str += "\\nocite{0}{{*}}\n".format(section)
        pub_str += "\\bibliographystyle{0}{{ep_par.bst}}\n".format(section)
        pub_str += "\\bibliography{0}{{articles.bib}}\n\n".format(section)

    return pub_str

        
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

print(section_sep + "Active Grants")
print(get_active_grants(book,year))

print(section_sep + "Consulting")
print(get_consulting(book,year))

print(section_sep + "Meetings")
print(get_meetings(book,year))

print(section_sep + "Current Grad Students")
print(get_current_grad_students(book,year))

print(section_sep + "Graduated Students")
print(get_graduated_students(book,year))

print(section_sep + "Staff")
print(get_staff_list(book,year))

print(section_sep + "Undergrads")
print(get_ug_list(book,year))

print(section_sep + "Personal Research")
print(get_personal_research(book,year))

print(section_sep + "Publications")
print(build_publications())

print(section_sep + "Other Activites")
print(heading("Other Important Activities","Comment on any important acitivities not covered above.") + 
      todo("add any important activities not covered here"))

print(section_sep + "Significant Accomplishments")
print(heading("Significant Accomplishments",
              "Your own view of your most significant accomplishments during the past year") + 
      todo("add any significant accomplishments"))
      
print("\end{document}")

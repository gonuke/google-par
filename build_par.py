import gspread
import argparse
from oauth2client.service_account import ServiceAccountCredentials

from datetime import datetime

# Some standard formatting to use throughout
newline = "\\\\ \n"

emph_none = "\emph{none}\n\n"

table_single_rule = " \\\\ \hline"
table_double_rule = table_single_rule + "\hline"
table_footer = "\\end{tabular}\n \end{centering}"


def section_sep(title):
    """
    Add a comment that serves as a section separator and a title for human
    readability of the ultimate TeX file.
    
    inputs
    ------
    title : (str) title for this section

    outputs
    -------
    section_sep_str : (str) containing the separator comment
    """
    
    return  "\n\n%%\n%% " + title + "\n%% " + "-"*len(title) + "\n"

def heading(section, guidance=""):
    """
    Add a section heading and optionally some additional guidance for 
    how to interpret the section.

    inputs
    ------
    section : (str) the title of the section
    guidance : (str) guidance to be added in italics after the section heading
                default : ""

    outputs
    -------
    head_str : (str) formatted section heading and guidance
    """

    head_str = "\section{" + section + "}\n\n"
    if len(guidance) > 0:
        head_str += "\emph{(" + guidance + ")}\n\n "

    return head_str

def is_current(record,year):
    """
    Utility function to check whether a record is valid for the current year.
    Since records contain different information about dates, as appropriate,
    there are multiple conditions that must be tested.

    a) if a record contains only a YEAR, it's value must match the given `year`
    b) if a record contains a STARTYEAR (and therefore an ENDYEAR), the given `year`
       must lie in the interval [STARTYEAR, ENDYEAR]
    c) if a record contains a STARTDATE (and therefore an ENDDATE), the given `year`
       must lie in the interval [STARTDATE, ENDDATE]
    d) if a record contains a DATE, the DATE must be in the given `year`

    inputs
    ------
    year : (int) year of interest to be matched

    output
    ------
    Boolean per the above conditions
    """

    if 'YEAR' in record.keys():
        return record['YEAR'] == year
    elif 'STARTYEAR' in record.keys():
        return (record['STARTYEAR'] <= year and record['ENDYEAR'] >= year)
    elif 'STARTDATE' in record.keys():
        start_date_obj = datetime.strptime(record['STARTDATE'],"%m/%d/%y")
        end_date_obj = datetime.strptime(record['ENDDATE'],"%m/%d/%y")
        return (start_date_obj.year <= year and end_date_obj.year >= year)
    elif 'DATE' in record.keys():
        date_obj = datetime.strptime(record['DATE'],"%m/%d/%y")
        return date_obj.year == year
    else:
        return False

def filter_for_current(worksheet,year):
    """
    Utility function to extract all the records from a given worksheet that 
    are valid for the current year, as defined by the `is_current()` method.

    inputs
    ------
    worksheet : a Google sheets worksheet object
    year      : the year to test for currency


    output
    ------
    a list of records that are current
    """
    
    return [s for s in worksheet.get_all_records() if is_current(s,year)]

def build_table_header(col_info):
    """
    Utility function to start a LaTeX table with some set of headings.

    The format information is passed in a list of tuples, with one tuple for
    each column, in order.  Each tuple contains 
    * a string for the header, 
    * a float for the relative width of the column
    * a key for that column based on the database table being used

    inputs
    ------
    col_info : a list of tuples for each 

    output
    ------
    a string that creates a correctly formatted LaTeX table

    """
    
    col_widths = [("p{" + str(frac) + "\\textwidth}|") for (head,frac,key) in col_info]
    col_heads = [("\\textbf{" + head + "}") for (head,frac,key) in col_info]
    header =  "\\begin{centering}\n" + \
              "\\begin{tabular}" + \
              "{|" + "".join(col_widths)  + "}" + \
              "\\hline\n" +  " & ".join(col_heads) + table_double_rule + " \n"

    return header

def build_table_rows(record_list,col_info):
    """
    Utility function to add rows to a table

    inputs
    ------
    record_list : a list of dictionaries.  

                  Each dictionary represents one record.  The keys of the
                  dictionary match the keys in the col_info to ensure that
                  each item goes in the correct column

    col_info : a list of tuples.

               Each tuple represents one column.  The last entry in each tuple
               is a key that matches one of the columns in the original 
               database table.

    output
    ------
    a string that creates multiple correctly formatted LaTeX table rows, 
    with the data in the correct columns

    """

    rows_str = ""
    for r in record_list:
        rows_str += " & ".join([str(r[key]) for (head,frac,key) in col_info]) + table_single_rule + "\n"
    return rows_str + table_footer


def expand_table(table,cat_table,join_key):
    """
    Utility function to add columns to one table based on the shared `join_key`
    with another table for every row in `table` add all the columns 
    from `cat_table` from the row with matching `join_key`

    inputs
    ------

    table : a list of dictionaries, each with one key equal to `join_key`
    cat_table : a list of dictionaries, each with one key equal to `join_key`
    join_key : a key that appears in all dictionaries of each table

    outputs
    -------
    a list of dictionaries, each with all the entries of both original dictionaries

    """

    # create a dictionary that maps the value of `join_key` for each record
    # onto that full record
    catalog = {e[join_key] : e for e in cat_table}

    # update each dictionary in `table` with the data from `cat_table`
    # with the matching value of `join_key`
    for e in table:
        e.update(catalog[e[join_key]])

    return table

def get_course_list(book,year):
    """
    Build a table with the list of courses.
    
    inputs
    ------
    book : a google sheet workbook object
    year : the current year

    output
    ------
    A correctly formatted LaTeX string with the section header and table
    listing the courses that have been taught.
    """

    # the semester list is defined in the order in which they occur in a
    # calendar year
    semester_list = ['Spring','Summer','Fall']

    # a LaTeX table separator within each semester
    part_rule  = " \\\\ \cline{2-4}"
    # an entry to use in every semester with no teaching
    empty_semester = "\multicolumn{3}{c|}{" + emph_none + "} " + table_double_rule

    # map the database column names to the LaTeX table headers and widths
    column_info = [("Semester",0.25,''),
                   ("Course",0.25,'COURSEID'),
                   ("\# of Students",0.25,'STUDENTS'),
                   ("Role",0.25,'ROLE')]

    # Start with the heading & table heading
    course_list_str = heading("Courses Taught","") + build_table_header(column_info)

    # get all course history
    full_history = book.worksheet('CourseHistory').get_all_records()

    for semester in semester_list:
        # get this semester's list of courses
        course_list = [e for e in full_history if (e['YEAR'],e['SEMESTER']) == (year,semester)]

        # special formatting for multirow with semester heading
        if (len(course_list) > 0):
            first_col = "\multirow{" + str(len(course_list)) + "}{*}{" + semester + "} & "
            rows = [" & ".join([entry['COURSEID'],str(entry['STUDENTS']),entry['ROLE']]) for entry in course_list]
            course_list_str += first_col + (part_rule + "\n & ").join(rows) + table_double_rule + "\n"
        else:
            course_list_str += semester + " & " + empty_semester + "\n"

    return course_list_str + table_footer

# info about future course interests
def get_future_courses(book):
    """
    Provide list of future coures interest in three groupings.

    inputs
    ------
    book : a google workbook object

    output
    ------
    LaTeX formatted list of courses in three groupings.
    """
    
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


def get_student_advising_info(book,year):
    """
    Read through lists of advisees to determine number of advisees at each stage.

    inputs
    ------
    book : a google workbook object
    year : the current year to extract

    output
    ------
    A string with a count of current year advisees grouped as:
    * NE undergraduates
    * NEEP graduates
    * other graduate students
    """

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

def get_org_advising_info(book,year):
    """
    Information about advising student organizations

    inputs
    ------
    book : a google workbook object
    year : the current year to extract

    output
    ------
    A string that indicates which student orgs are advised and the time commitment.
    """
    
    current_orgs = expand_table(filter_for_current(book.worksheet('StudentOrgAdvising'),year),
                                book.worksheet('StudentOrgList').get_all_records(),
                                'ORGCODE')
    if (len(current_orgs) > 0):
        advising_str = "Faculty advisor for the " + \
                       ", ".join([(o['ORGNAME'] + " (" + str(o['WEEKLYHOURS']) + " hours/wk)") for o in current_orgs])
    else:
        advising_str = emph_none
        
    return advising_str

def get_advising_info(book,year):
    """
    Create a report section with both individual advisee counts and student org advising.

    inputs
    ------
    book : a google workbook object
    year : the current year to extract

    output
    ------
    A string that combines the output of two advising functions
    """
    
    advising_str = heading("Advising Responsibilities")

    advising_str += get_student_advising_info(book,year) + newline + \
                    get_org_advising_info(book,year)
    
    return advising_str

def committee_list(services):
    """
    Create a string by expanding the service obligations and sorting them
    from shortest frequency unit to longest.

    inputs
    ------
    services : a list of record for service obligations

    output
    ------
    A LaTeX string with a new line for each service obligation.
    """
        
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

def soc_svc(services,society):
    """
    Create a string by expanding the service obligations to a specific society 
    and sorting them from shortest frequency unit to longest.

    inputs
    ------
    services : a list of record for service obligations
    society : a string indicating which professional society

    output
    ------
    A LaTeX string with a subheading for the society and a list of entries 
    for each service obligation within that society.
    """

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
            review_strs.append(", ".join(review_str_list))
                               
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
                                          'NAME': comm_type + " committees (" + ", ".join(name_list) + ")",
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
                             ": " + outreach['AUDIENCE'] + ", ``" + outreach['TITLE'] + "''")

    return heading("Educational Outreach Activities") + newline.join(outreach_strs)


def get_patents(book,year):

    patent_list = filter_for_current(book.worksheet("Patents"),year)

    patent_str = heading("Patents applied for or granted in " + str(year))

    if len(patent_list) > 0:
        patent_str += newline.join([("Patent No. " + p['PATENTNUMBER'] + ": " + p['TITLE'] + "(" + p['STATUS'] + ")") for p in patent_list]) + newline
    else:
        patent_str += emph_none

    return patent_str

def make_grant_list(grants):

    column_info = [("Begin Date",0.1,'STARTDATE'),
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

    begin_date_obj = datetime.strptime(grant_record['STARTDATE'],"%m/%d/%y")
    end_date_obj = datetime.strptime(grant_record['ENDDATE'],"%m/%d/%y")
    return begin_date_obj.year <= year and end_date_obj.year >= year
    

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

    staff_positions = {'Scientist','Researcher','Academic Staff'}

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

def get_narrative(book, year, tab, title, desc):

    item_list = filter_for_current(book.worksheet(tab),year)

    item_str = heading(title, desc)

    if len(item_list) > 0:
        item_str += " \n\n".join([item['DESCRIPTION'] for item in item_list])
    else:
        item_str += emph_none

    return item_str

def build_publications():

    pub_str = heading("Publications list","")

    for section in ['books','journalssubmitted','journalsaccepted','journalspublished','conference','reports','invited']:
        pub_str += "\\nocite{0}{{*}}\n".format(section)
        pub_str += "\\bibliographystyle{0}{{ep_par.bst}}\n".format(section)
        pub_str += "\\bibliography{0}{{{0}.bib}}\n\n".format(section)

    return pub_str


def build_par(book,year):

    par_tex = """\
\documentclass[12pt]{article}

\usepackage{ep_par}
"""
    
    par_tex += "\\newcommand{\paryear}{" +  str(year) + "}\n"
    par_tex += """\
\\newcommand{\parperson}{Paul P.\ H.\ Wilson}
\\begin{document}

\partitle
"""
    
    par_tex += section_sep("Course List")
    par_tex += get_course_list(book,year) + "\n"

    par_tex += section_sep("Course Prep")
    par_tex += get_future_courses(book) + "\n"

    par_tex += section_sep("Course Dev")
    par_tex += get_narrative(book,year,"CourseDevelopment", "Course Development Activities","") + "\n"

    par_tex += section_sep("Student Advising")
    par_tex += get_advising_info(book,year) + "\n"
    
    par_tex += section_sep("Service")
    par_tex += get_service(book,year) + "\n"
    
    par_tex += section_sep("Educational Outreach Activities")
    par_tex += get_outreach(book,year) + "\n"
    
    par_tex += section_sep("Awards/Honors")
    par_tex += get_narrative(book,year,"HonorsAwards", "Honors and Awards received in " + str(year),"") + "\n"

    par_tex += section_sep("Patents")
    par_tex += get_patents(book,year) + "\n"

    par_tex += section_sep("Submitted Proposals")
    par_tex += get_proposal_submissions(book,year) + "\n"

    par_tex += section_sep("Active Grants")
    par_tex += get_active_grants(book,year) + "\n"
    
    par_tex += section_sep("Consulting")
    par_tex += get_consulting(book,year) + "\n"
    
    par_tex += section_sep("Meetings")
    par_tex += get_meetings(book,year) + "\n"
    
    par_tex += section_sep("Current Grad Students")
    par_tex += get_current_grad_students(book,year) + "\n"
    
    par_tex += section_sep("Graduated Students")
    par_tex += get_graduated_students(book,year) + "\n"
    
    par_tex += section_sep("Staff")
    par_tex += get_staff_list(book,year) + "\n"
    
    par_tex += section_sep("Undergrads")
    par_tex += get_ug_list(book,year) + "\n"
    
    par_tex += section_sep("Personal Research")
    par_tex += get_narrative(book,year,"PersonalResearch", "Personal Research",
                             "Brief description of the extent and nature of any personal research " + \
                             "(defined here as research performed independent of graduate students rather " + \
                             "than through them) and indicate the project on which this research is done.")+ "\n"

    par_tex += section_sep("Publications")
    par_tex += build_publications() + "\n"
    

    par_tex += section_sep("Other Activites")
    par_tex += get_narrative(book,year,"ImportantActivities", "Other Important Activities",
                             "Comment on any important acitivities not covered above.") + "\n"

    par_tex += section_sep("Significant Accomplishments")
    par_tex += get_narrative(book,year,"SignificantAccomplishments","Significant Accomplishments",
                    "Your own view of your most significant accomplishments during the past year") + "\n"
    
      
    par_tex += "\end{document}\n"

    return par_tex


def open_book(credentials,filename):

    # use creds to create a client to interact with the Google Drive API
    scope = ['https://spreadsheets.google.com/feeds']
    creds = ServiceAccountCredentials.from_json_keyfile_name(credentials, scope)
    client = gspread.authorize(creds)

    # Find a workbook by name and open the first sheet
    # Make sure you use the right name here.
    book = client.open(filename)

    return book


if __name__ == '__main__':

    parser = argparse.ArgumentParser(formatter_class=argparse.ArgumentDefaultsHelpFormatter)

    parser.add_argument("-y", "--year", type=int, default=datetime.now().year -1 ,
                        help="Year of report")

    parser.add_argument("-c", "--credentials", type=str, default='ep-par-processing.json',
                        help="JSON file with secret credentials for accessing a users Google files")

    parser.add_argument("-f", "--filename", type=str, default='PAR Data',
                        help="Name of Google Sheets file")
                        
    args = parser.parse_args()

    book = open_book(args.credentials, args.filename)
    
    print(build_par(book,args.year))

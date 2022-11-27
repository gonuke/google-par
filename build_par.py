import gspread
import argparse
from oauth2client.service_account import ServiceAccountCredentials

from datetime import datetime

# Some standard formatting to use throughout
newline = "\\\\ \n"
doubleblank = "\n\n"

emph_none = "\emph{none}"

date_fmt = "%m/%d/%y"

table_single_rule = " \\\\ \hline"
table_double_rule = table_single_rule + "\hline"
table_footer = "\\end{tabular}\n\\end{centering}"
bind_page_start = "\\noindent\\begin{minipage}{\\textwidth}\n"
bind_page_end = "\n\\end{minipage}\n"

grantlist_column_info = [("Begin Date",0.1,'STARTDATE'),
                   ("End Date",0.1,'ENDDATE'),
                   ("Amount [\$k]",0.1,'AMOUNT'),
                   ("Topic",0.3,'TOPIC'),
                   ("Agency",0.1,'AGENCY'),
                   ("Co-Authors",0.1,'CO-AUTHORS'),
                   ("Role",0.05,'ROLE')]
 
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
    
    return  doubleblank + "%%\n%% " + title + "\n%% " + '-' * len(title) + doubleblank

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

    head_str = "\section{" + section "}" + doubleblank
    if len(guidance) > 0:
        head_str += "\emph{(" + guidance + ")}" + doubleblank

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
        startyear = int(record['YEAR'])
        endyear = startyear
    elif 'STARTYEAR' in record.keys():
        startyear = int(record['STARTYEAR'])
        endyear = int(record['ENDYEAR'])
    elif 'STARTDATE' in record.keys():
        startyear = datetime.strptime(record['STARTDATE'],date_fmt).year
        endyear = datetime.strptime(record['ENDDATE'],date_fmt).year
    # check for DATE last because some records (e.g. proposals) contain DATE and STARTDATE
    elif 'DATE' in record.keys():
        startyear = datetime.strptime(record['DATE'],date_fmt).year
        endyear = startyear
    else:
        startyear = 9999
        endyear = 0

    return (startyear <= year and endyear >= year)

def make_date_range(record):
    """
    Return a formatted string of the date range indicated in the record

    input:
    ------

    record : a single record with a STARTDATE and ENDDATE entry

    outputs:
    ---------
    A LaTeX formatted string representing the date range of the record

    """
    start_date_obj = datetime.strptime(record['STARTDATE'],date_fmt)
    end_date_obj = datetime.strptime(record['ENDDATE'],date_fmt)

    return start_date_obj.strftime("%m/%d") + "-" + end_date_obj.strftime("%m/%d")

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
    
    col_widths = [("p{" + frac + "\\textwidth}|") for (head,frac,key) in col_info]
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

def build_table_or_none(row_data, column_info):
    """
    Create a complete table with a header and row data from a list of data.
    If the list is empty, return the italicized "none" in its place.


    inputs
    ------
    record_list : a list of dictionaries.  

                  Each dictionary represents one record.  The keys of the
                  dictionary match the keys in the col_info to ensure that
                  each item goes in the correct column

    column_info : a list of tuples.

               Each tuple represents one column.  The last entry in each tuple
               is a key that matches one of the columns in the original 
               database table.

    output
    ------
    A LaTeX formatted string with a complete table, or an entry specifying none.

    """

    row_data_str = ""

    if len(row_data) > 0:
        row_data_str += build_table_header(column_info) + \
                        build_table_rows(row_data,column_info)
    else:
        row_data_str += emph_none + doubleblank

    return row_data_str

def build_textlist_or_none(list_data):
    """
    Create a list of entries if there are any, otherwise none.


    inputs
    ------
    list_data : a list of text entries

    output
    ------
    A LaTeX formatted string with a simple list of one line per entry.

    """

    list_str = ""

    if len(list_data) > 0:
        list_str += newline.join(list_data)
    else:
        list_str += emph_none + doubleblank

    return list_str



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

    # get all course history
    full_history = book.worksheet('CourseHistory').get_all_records()

    # map the database column names to the LaTeX table headers and widths
    column_info = [("Semester",0.25,''),
                   ("Course",0.25,'COURSEID'),
                   ("\# of Students",0.25,'STUDENTS'),
                   ("Role",0.25,'ROLE')]

    # Start with the heading & table heading
    course_list_header = heading("Courses Taught") + build_table_header(column_info)

    # a LaTeX table separator within each semester
    part_rule  = " \\\\ \cline{2-4}"

    # an entry to use in every semester with no teaching
    empty_semester = "\multicolumn{3}{c|}{" + emph_none + "} " + table_double_rule

    course_list_str = ""

    # loop through semesters in the order in which they occur in a
    # calendar year    
    for semester in ['Spring','Summer','Fall']:
        # extract necessary data from this semester's list of courses
        course_list = [ (e['COURSEID'],str(e['STUDENTS']),e['ROLE']) for e in full_history if (e['YEAR'],e['SEMESTER']) == (year,semester)]

        # special formatting for multirow with semester heading
        if (len(course_list) > 0):
            first_col = "\multirow{" + str(len(course_list)) + "}{*}{" + semester + "} \n    & "
            rows = [" & ".join(entry) for entry in course_list]
            course_list_str += first_col + (part_rule + "\n    & ").join(rows) + table_double_rule + "\n"
        else:
            course_list_str += semester + " & " + empty_semester + "\n"

    return course_list_header + course_list_str + table_footer

# info about future course interests
def get_future_courses(book):
    """
    Provide list of future courses interest in three groupings.

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

    # Student advisees
    current_advisees = filter_for_current(book.worksheet('AdviseeList'), year)
    
    adv_str_list = []

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

    return build_textlist_or_none(adv_str_list)

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
        advising_str = "Student Organizations: " + emph_none + doubleblank
        
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
    
    cat_svc_str = "\subsubsection{" + name[0] + "}" + doubleblank

    cat_svc_str += committee_list([s for s in services if s['SOCIETY'] == society])

    return cat_svc_str

# get the formatted version of each different category of service
def get_service_cat(services,cat,cat_str):
    """
    Create a summary of the service obligations for a single category of service.

    inputs
    ------

    services : a list of records for service obligations
    cat      : the category key to extract for this block
    cat_str  : the long text description of this category

    output
    ------
    A LaTeX string with a subheading for the category and a list of entries 
    for each service obligation within that category.

    """
    cat_svc = [s for s in services if s['CATEGORY'] == cat]
    
    cat_svc_str = "\subsection{To " + cat_str + "}\n\n"

    if (len(cat_svc) > 0):
        if (cat != "SOCIETY"):
            cat_svc_str += committee_list(cat_svc)
        else:
            for soc in {s['SOCIETY'] for s in cat_svc if s['SOCIETY'] != ""}:
                cat_svc_str += soc_svc(cat_svc,soc)
    else:
        cat_svc_str += emph_none + doubleblank

    return cat_svc_str + doubleblank

# get formatted list of reviews
def get_reviews(book,year):
    """
    Create a string to summarize proposal and paper reviews

    inputs
    ------
    book : a google workbook object
    year : the current year to extract

    output
    ------
    A LaTeX string with a subheading for each of paper reviews and proposals and an inline list
    of the number of proposals for each source.
    """

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
                               
    return doubleblank.join(review_strs)

def other_student_committees(book,year):
    """
    Create table entries that contain correct information for a service commitment
    based on the participation in student BS/MS/PhD committees

    inputs
    ------
    book : a google workbook object
    year : the current year to extract

    output
    ------
    comm_summary_list : a dictionary with the same records as the entries in the list of committees,
                        with one entry for each type of committee (BS, MS, PhD) 
                        and each organization (EP, UW, NATIONAL)

    """

    # get list of other student committees for this year
    other_comms = filter_for_current(book.worksheet('OtherStudentCommittees'),year)

    # set of possible committees
    comm_type_list = ["BS Defense", "MS Oral", "MS Defense", "PhD Prelim", "PhD Defense"]

    comm_summary_list = []
    for comm_type in comm_type_list:
        for cat in ["EP","UW","NATIONAL"]:
            name_list = [s['NAME'] for s in other_comms if (s['TYPE'],s['DEPT']) == (comm_type,cat)]
            if len(name_list) > 0:
                comm_summary_list.append({'CATEGORY': cat,
                                          'NAME': comm_type + " committees (" + ", ".join(name_list) + ")",
                                          'COMMITMENTUNIT': 'COUNT',
                                          'COMMITMENTQUANTITY': len(name_list)})

    return comm_summary_list
    

# get all service obligations
def get_service(book,year):
    """
    Create a section with all service obligations

    inputs
    ------
    book : a google workbook object
    year : the current year to extract

    output
    ------
    A LaTeX formatted string describing the service obligations in various categories
    and subcategories.

    """

    current_services = filter_for_current(book.worksheet('Service'),year)
    service_catalog = book.worksheet('ServiceList').get_all_records()

    current_services = expand_table(current_services, service_catalog, "SERVICECODE")

    # find all the parent societies and insert them in the service list
    parent_societies = {s['SOCIETY'] for s in current_services if (s['SOCIETY'] != "")}
    current_services.extend([s for s in service_catalog if (s['SERVICECODE'] in parent_societies)])

    # append service list with summaries of student committees
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
    """
    Create a section with outreach experiences

    inputs
    ------
    book : a google workbook object
    year : the current year to extract

    output
    ------
    A LaTeX formatted string describing the outreach activities

    """

    current_outreach = filter_for_current(book.worksheet('Outreach'),year)

    outreach_strs = []
    for outreach in current_outreach:
        date_obj = datetime.strptime(outreach['DATE'],date_fmt)
        outreach_strs.append(datetime.strftime(date_obj,"%B") + " " + str(date_obj.day) +
                             ": " + outreach['AUDIENCE'] + ", ``" + outreach['TITLE'] + "''")

    return heading("Educational Outreach Activities") + newline.join(outreach_strs)


def get_patents(book,year):
    """
    Create a section with all patents

    inputs
    ------
    book : a google workbook object
    year : the current year to extract

    output
    ------
    A LaTeX formatted string describing the patents

    """

    patent_list = filter_for_current(book.worksheet("Patents"),year)

    patent_str = heading("Patents applied for or granted in " + str(year))

    patent_str_list = ["Patent No. %(PATENTNUMBER) : %(TITLE) (%(STATUS))" % p for p in patent_list]

    patent_str += build_textlist_or_none(patent_str_list)

    return patent_str

def make_grant_list(grants):
    """
    Create a table out of a list of grants with information about dates, amount, topic, agency, etc

    inputs
    ------
    grants : a list of records that each has the appropriate information for describing either an
             active grant or a proposal

    output
    ------
    A LaTeX formatted string with a table containing all the grant info

    """

    # layout the columns
   
    grants_str = build_table_or_none(grants,column_info)

    return grants_str

def get_proposal_submissions(book,year):
    """
    Create a table for grant submissions during the year of interest

    inputs
    ------
    book : a google workbook object
    year : the current year to extract

    output
    ------
    A LaTeX formatted string with a table of grant proposal information.
    """

    submission_list = filter_for_current(book.worksheet('ProposalsAndGrants'), year)

    submission_list_str = heading("Research proposals submitted during " + str(year))

    submission_list_str += build_table_or_none(submission_list, grantlist_column_info)

    return submission_list_str

def active_grant(grant_record,year):

    begin_date_obj = datetime.strptime(grant_record['STARTDATE'],date_fmt)
    end_date_obj = datetime.strptime(grant_record['ENDDATE'],date_fmt)
    return begin_date_obj.year <= year and end_date_obj.year >= year
    

def get_active_grants(book,year):
    """
    Create a table for grant submissions during the year of interest

    inputs
    ------
    book : a google workbook object
    year : the current year to extract

    output
    ------
    A LaTeX formatted string with a table of active grant information.
    """

    grant_list = book.worksheet('ProposalsAndGrants').get_all_records()

    active_grant_list = [g for g in grant_list if g['STATUS'] == 'FUNDED' and active_grant(g,year)]

    active_grant_list_str = heading("Research grants and contracts active during " + str(year))
    
    active_grant_list_str += build_table_or_none(active_grant_list, grantlist_column_info)

    return active_grant_list_str
    
def get_consulting(book,year):
    """
    Create a list of consulting engagements

    inputs
    ------
    book : a google workbook object
    year : the current year to extract

    output
    ------
    A LaTeX formatted string with one line for each consulting arrangement.
    """

    consult_list = filter_for_current(book.worksheet("Consulting"),year)

    consulting_str = heading("Consulting agreements held in " + str(year))

    consulting_str_list = [("\\textbf{" + c['ORGANIZATION'] + ":}" + c['TOPIC']) for c in consult_list]

    consulting_str += build_textlist_or_none(consulting_str_list)

    return consulting_str

def get_meetings(book,year):
    """
    Create a list of meetings that were attended

    inputs
    ------
    book : a google workbook object
    year : the current year to extract

    output
    ------
    A LaTeX formatted table with one row per meeting
    """
 
    meeting_list = filter_for_current(book.worksheet("Meetings"),year)

    # extend the data with a formatted date range
    for meeting in meeting_list:
        meeting['DATERANGE'] = make_date_range(meeting)
    
    column_info = [("Dates",0.15,'DATERANGE'),
                   ("Location",0.2,'LOCATION'),
                   ("Meeting",0.55,'MEETINGNAME')]

    meeting_list_str = heading("Professional Meetings and Conferences Attended in " + str(year),"")

    meeting_list_str += build_table_or_none(meeting_list,column_info)

    return meeting_list_str

def get_current_grad_students(book,year):
    """
    Create a table with current grad students, their topic and funding

    inputs
    ------
    book : a google workbook object
    year : the current year to extract

    output
    ------
    A LaTeX formatted table with one row per student
    """
 
    all_current_students = filter_for_current(book.worksheet("ActiveResearchStudents"),year)
    all_current_advisees = filter_for_current(book.worksheet("AdviseeList"),year)

    grad_types = {"G", "V"}
    research_student_list = expand_table( all_current_students,
                                         [s for s in all_current_advisees if s['TYPE'] in grad_types],
                                         'LASTNAME')

    # extend record with full name to use standard table building capability
    for student in research_student_list:
        student['FULLNAME'] = student['FIRSTNAME'] + " " + student['LASTNAME']

    column_info = [("Student",0.25,'FULLNAME'),
                   ("Program",0.11,'DEGREE'),
                   ("Research Topic",0.35,'TOPIC'),
                   ("Source of Support",0.2,'SOURCE')]

    # force this to start on a new page
    research_student_str = bind_page_start
    research_student_str += heading("Graduate Students",
                           "Current grad students doing thesis research, " + \
                           "and non-thesis students who are active in research and " + \
                           "taking as much time as though doing there theses.")

    research_student_str += build_table_or_none(research_student_list,column_info)

    research_student_str += bind_page_end
    
    return research_student_str

def get_graduated_students(book,year):
    """
    Create a table with students who graduated in this time period.

    inputs
    ------
    book : a google workbook object
    year : the current year to extract

    output
    ------
    A LaTeX formatted table with one row per student who graduated
    """
 
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
    
    graduated_student_str += build_table_or_none(graduated_student_list, column_info)
    
    return graduated_student_str


def get_staff_list(book,year, staff_positions):
    """
    Create a table with staff with a given set of titles

    inputs
    ------
    book : a google workbook object
    year : the current year to extract
    staff_positions : a set of strings with valid titles to include in this table

    output
    ------
    A LaTeX formatted table with one row per staff member
    """
 
    staff_list = [s for s in filter_for_current(book.worksheet("Staff"),year) if s['TITLE'] in staff_positions]

    for staff in staff_list:
        staff['FULLNAME'] = staff['FIRSTNAME'] + " " + staff['LASTNAME']

    column_info = [("Name", 0.25,'FULLNAME'),
                   ("Title",0.11,'TITLE'),
                   ("Research Topic",0.35,'TOPIC'),
                   ("Source of Support",0.2,'SUPPORT')]
    
    return build_table_or_none(staff_list, column_info)

def get_prof_list(book, year):
    """
    Create a table of professional staff: Scientists, Researches, Academic Staff

    inputs
    ------
    book : a google workbook object
    year : the current year to extract

    output
    ------
    A LaTeX formatted table with one row per staff member
    """
 
    staff_positions = {'Scientist','Researcher','Academic Staff'}

    prof_list_str = bind_page_start + heading("Research Staff",
                    "Post-PhD and academic staff supervised in " + str(year))
    
    prof_list_str += get_staff_list(book, year, staff_positions)
    
    return prof_list_str + bind_page_end

def get_ug_list(book,year):
    """
    Create a table of undergraduate students

    inputs
    ------
    book : a google workbook object
    year : the current year to extract

    output
    ------
    A LaTeX formatted table with one row per staff member
    """
 
    ug_positions = {'U/G Hourly'}
  
    ug_list_str = heading("Undergraduate researchers")

    ug_list_str += get_staff_list(book, year, ug_positions)

    return ug_list_str

def get_narrative(book, year, tab, title, desc):
    """
    Create a text list from a given spreadsheet tab composed of titles and descriptions

    inputs
    ------
    book : a google workbook object
    year : the current year to extract

    output
    ------
    A LaTeX formatted list wth one item per list
    """
 
    item_list = filter_for_current(book.worksheet(tab),year)

    item_str = heading(title, desc)

    item_str_list = [item['DESCRIPTION'] for item in item_list]

    item_str += build_textlist_or_none(item_str_list)

    return item_str

def build_publications():
    """
    Create a bibliography entry for each of the separate bibtex files.
    """


    pub_str = heading("Publications list","")

    for section in ['journalssubmitted','journalsaccepted','journalspublished','conference','reports','books','invited']:
        pub_str += "\\nocite{0}{{*}}\n".format(section)
        pub_str += "\\bibliographystyle{0}{{ep_par.bst}}\n".format(section)
        pub_str += "\\bibliography{0}{{{0}.bib}}\n\n".format(section)

    return pub_str


def build_par(book,year):

    par_tex = """\
\documentclass[12pt]{article}

\\usepackage{ep_par}
"""
    
    par_tex += "\\newcommand{\paryear}{" +  str(year) + "}\n"
    par_tex += """\
\\newcommand{\parperson}{Paul P.\ H.\ Wilson}
\\begin{document}

\partitle
"""
    par_tex += section_sep("Professional Summary")
    par_tex += "\input{professional_summary}\n\n"

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
    par_tex += get_prof_list(book,year) + "\n"
    
    par_tex += section_sep("Undergrads")
    par_tex += get_ug_list(book,year) + "\n"
    
    par_tex += section_sep("Personal Research")
    par_tex += get_narrative(book,year,"PersonalResearch", "Personal Research",
                             "Brief description of the extent and nature of any personal research " + \
                             "(defined here as research performed independent of graduate students rather " + \
                             "than through them) and indicate the project on which this research is done.")+ "\n"

    par_tex += section_sep("Publications")
    par_tex += build_publications() + "\n"
    
    par_tex += section_sep("Climate Improvement")
    par_tex += get_narrative(book,year,"ClimateImprovement", "Climate Improvement Activities",
                             "Please comment on any ways that you have worked to enhance the climate " + \
                             "and culture of the department and/or college.  This may include " + \
                             "strategies to increase inclusivity in your courses, approaches to " + \
                             "maintain a healthy climate in your research group, activities that " + \
                             "contribute to a collaborative working environment with other faculty " + \
                             "and staff, and workshops/trainings on leadership, mentoring, recruiting, " + \
                             "or diversity.  If you have attended workshops/trainings on leadership, " + \
                             "mentoring, recruiting, or diversity, please list those and identify actions " +\
                             "you have taken as a result of what you learned.") + "\n"

    par_tex += section_sep("Other Activities")
    par_tex += get_narrative(book,year,"ImportantActivities", "Other Important Activities",
                             "Comment on any important activities not covered above.") + "\n"

    par_tex += section_sep("Significant Accomplishments")
    par_tex += get_narrative(book,year,"SignificantAccomplishments","Significant Accomplishments",
                    "Your own view of your most significant accomplishments during the past year") + "\n"
    
      
    par_tex += "\end{document}\n"

    return par_tex


def open_book(credentials,filename):

    # use creds to create a client to interact with the Google Drive API
    client = gspread.service_account(filename=credentials)

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

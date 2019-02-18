# google-par
Generate a LaTeX based Professional Activity Report from a Google spreadsheet-based DB.

The format of this professional activity report is based on that of the
Engineering Physics Department at the University of Wisconsin-Madison.  The
formatting is hard coded in the python, but should be straightforward to
change for most python programmers.

Base in part from information learned [here](https://www.twilio.com/blog/2017/02/an-easy-way-to-read-and-write-to-a-google-spreadsheet-in-python.html).

A reference on splitting a bibliography:
https://tex.stackexchange.com/questions/20246/sectioning-bibliography-by-type-of-referred-item

## Usage

Note: Currently relies on a docker image to simplify python dependencies.

To build a docker image with the required dependencies, build docker image
from `google-par.docker` and name the image `gonuke/google-par`

To build a new PAR in a batch mode, run `docker.par.build`.

To build a new PAR in an interactive mode, run `docker.par.interactive` and
consult `build_par.sh` for build steps.

## Google Credentials
Instructions for getting the necessary google credentials are available in [this reference](https://www.twilio.com/blog/2017/02/an-easy-way-to-read-and-write-to-a-google-spreadsheet-in-python.html).

These instructions suggest saving the file as `client_secret.json`, but the
default name expected by this script is `ep-par-processing.json`.

## Google Spreadsheet Structure

The structure of the Google Sheet is as follows:

* CourseInfo:  One entry per possible course
   * COURSEID - unique string for reference in table CourseHistory
   * TITLE - the name of the course
   * CREDITS - the number of credit hours for the course
   * PREPSTATUS - an indiciation of how prepared the individual is for a given course with options:
       * PREP - courses fully prepared to teach at any time
       * REPEAT - courses that have been taught before and could be taught again
       * INTEREST - courses that have never been taught, but there is some iterest

* CourseHistory: One entry for each time a course is taught
   * YEAR - the calendar year that the course was taught
   * SEMESTER - the semester that the course was taught, one of: Spring, Summer, Fall
   * COURSEID - the string for cross-referencing from CourseInfo
   * STUDENTS - the number of students enrolled in that session
   * ROLE - the role of the individual during that session

* CourseDevelopment: One entry for each notable instance of course development
   * YEAR - the calendar year of course development
   * DESCRIPTION - an arbitrarily long description of the course development activities

* AdviseeList: One entry for each student-degree combination.  That is, if a
  student pursues two different degrees (e.g. BS and MS) as an advisee, then they appear twice.
   * STARTYEAR: first year that they were an advisee under the given program
   * ENDYEAR: last year that they were an advisee under the given program
   * LASTNAME: last name of student
   * FIRSTNAME: first name of student
   * TYPE: G for graduate student, U for undergraduate, V for visiting student
   * DEGREE: BS, MS or PhD
   * PROGRAM: A list of possible values can/should be given to help with data validation
   * TOPIC: For students pursuing research, a phrase indicating the topic of that reasearch
   * DEFENSEDATE: For students who pursued research, the date that they defended that research
   * CURRENTEMPLOYER: The current employer if know, using a unique ID that is cross-references against the EmployerList table
   * FIRSTEMPLOYER: The first employer following graduation, using a unique ID that is cross-references against the EmployerList table

* EmployerList: One entry for each employer of graduates
   * EMPLOYERCODE: A unique identifier for each employer that is used in AdviseeList
   * NAME: The full name of the employer

(More documentation coming)


import subprocess
import sys
# This uses system terminal to install python modules
subprocess.call([sys.executable, "-m", "pip", "install", 'xlrd'])
subprocess.call([sys.executable, "-m", "pip", "install", 'XlsxWriter'])
subprocess.call([sys.executable, "-m", "pip", "install", 'py3-validate-email'])


import os
import xlrd as xread
import xlsxwriter as xwrite
from Classes import Person, Department


def init():

    # opening database file
    database_loc = os.path.join(sys.path[0], 'database.xlsx')
    try:
        database_wb = xread.open_workbook(database_loc)
        sheets = [database_wb.sheet_by_index(i) for i in range(database_wb.nsheets)]
    except:
        print("\n" + "Error: No Excel Sheet Named 'database.xlsx' Found :( ")
        input("Please paste the database sheet into this folder, and run the script again!")

    # creating diagnostic sheets
    diagnostic_loc = os.path.join(sys.path[0], 'Actions Needed.xlsx')
    diagnostic_wb = xwrite.Workbook(diagnostic_loc)
    website_sheet = diagnostic_wb.add_worksheet('Update Invalid Websites')
    email_sheet = diagnostic_wb.add_worksheet('Fix Broken Emails')
    status_sheet = diagnostic_wb.add_worksheet('Remove Inactive Faculty')
    date_sheet = diagnostic_wb.add_worksheet('Renew Outdated Entries')
    duplicates_sheet = diagnostic_wb.add_worksheet('Revise Duplicates')

    # sheet formats
    bold = diagnostic_wb.add_format({'bold': True, 'font_color': 'red'})

    # turning database into data
    people = make_person(sheets)

    # PERFORMING DIAGNOSTICS
    # ///////////////////////////////////////////////////////////////
    # checking websites
    website_tasks = check_websites(people, website_sheet, bold)
    # checking email
    email_tasks = check_email(people, email_sheet, bold)
    # checking status
    status_tasks = check_status(people, status_sheet, bold)
    # checking date modified
    date_tasks = check_date(people, date_sheet, bold)
    # checking duplicates
    duplicates_tasks = check_duplicates(people, duplicates_sheet, bold)
    # ///////////////////////////////////////////////////////////////
    diagnostic_wb.close()

    # printing diagnostic result
    print("\n" + "ACTIONS NEEDED:" + "\n" + "" + "\n"
          + "Website Tasks = " + str(website_tasks) + "\n"
          + "Email Tasks = " + str(email_tasks) + "\n"
          + "Status Tasks = " + str(status_tasks) + "\n"
          + "Date Tasks = " + str(date_tasks) + "\n"
          + "Duplicates Tasks = " + str(duplicates_tasks) + "\n" + "" + "\n"
          + "(Please see the 'Actions Needed.xlsx' sheet for details)" + "\n")

    # exit
    input("\n" + "Completed! Press any key to exit... ")
    exit()


def make_person(sheets):

    # status
    print("\n" + "Making People...")

    # list that stores people
    departments = []

    for sheet in sheets:

        people = []
        for i in range(1, sheet.nrows):
            if sheet.cell_value(i, 0) != "":

                last_name = sheet.cell_value(i, 0)
                first_name = sheet.cell_value(i, 1)
                institution = sheet.cell_value(i, 2)
                program = sheet.cell_value(i, 3)
                position = sheet.cell_value(i, 4)
                knowledge = sheet.cell_value(i, 5)
                email = sheet.cell_value(i, 6)
                website = sheet.cell_value(i, 7)
                gender = sheet.cell_value(i, 8)
                urm = sheet.cell_value(i, 9)
                date_modified = sheet.cell_value(i, 10)
                status = sheet.cell_value(i, 11)

                new_person = Person(last_name, first_name, institution, program,
                                    position, knowledge, email, website, gender, urm, date_modified, status)
                people.append(new_person)

                # status
                print(new_person.name + "|| completed")

        new_department = Department(sheet.name, people)
        departments.append(new_department)

        # status
        print(new_department.name + "|| completed")

    return departments


from urllib.request import urlopen
from urllib.error import HTTPError, URLError


def check_websites(departments, website_sheet, style, any_false=[], i=0):

    # status
    print("\n" + "Checking Websites...")

    # check if a website is valid
    for department in departments:
        website_sheet.write(i, 0, department.name, style)
        i += 1
        for person in department.people:
            try:
                response = urlopen(person.website)
                if response.code != 200:
                    print("Webpage Error: " + person.name)
                    any_false.append(False)
            except HTTPError as e:
                codes = [403, 999]
                if e.code not in codes:
                    print("Webpage Error: " + person.name + " || " + str(e.code) + "http")
                    website_sheet.write(i, 0, person.name)
                    i += 1
                    any_false.append(False)
            except URLError as e:
                print("Webpage Error: " + person.name + " || " + str(e.args) + "url")
                website_sheet.write(i, 0, person.name)
                i += 1
                any_false.append(False)
            except:
                pass

    if all(any_false):
        print("\n" + "Congrats, All websites are working! ")
        return 0

    return i - len(departments)


from validate_email import validate_email


def check_email(departments, email_sheet, style, any_false=[], i=0):

    # status
    print("\n" + "Checking emails...")

    # check if a website is valid
    for department in departments:
        email_sheet.write(i, 0, department.name, style)
        i += 1
        for person in department.people:
            try:
                is_valid = validate_email(person.email, check_regex=True, check_mx=False)
                if is_valid is False:
                    print("Email Error: " + person.name)
                    email_sheet.write(i, 0, person.name)
                    i += 1
                    any_false.append(is_valid)
            except:
                pass

    if all(any_false):
        print("\n" + "Congrats, All emails are valid! ")
        return

    return i - len(departments)


def check_status(departments, status_sheet, style, any_false=[], i=0):

    # status
    print("\n" + "Checking status...")

    for department in departments:
        status_sheet.write(i, 0, department.name, style)
        i += 1
        for person in department.people:
            non_active = ['Inactive', 'Unsure', 'Transferred']
            if person.status in non_active:
                print("Inactive Status: " + person.name)
                status_sheet.write(i, 0, person.name)
                i += 1
                any_false.append(False)

    if all(any_false):
        print("\n" + "Congrats, All faculty members are active! ")
        return 0

    return i - len(departments)


from datetime import datetime


def check_date(departments, date_sheet, style, any_false=[], i=0, semester=183):

    # status
    print("\n" + "Checking dates...")

    for department in departments:
        date_sheet.write(i, 0, department.name, style)
        i += 1
        for person in department.people:

            if person.date_modified == "":
                print("Missing Date: " + person.name)
                date_sheet.write(i, 0, person.name)
                i += 1
                any_false.append(False)

            else:
                try:
                    calculate = lambda d1, d2: abs((d1 - d2).days)
                    today = datetime.today()
                    last_date = xread.xldate_as_datetime(person.date_modified, 0)
                    period = calculate(today, last_date)

                    if period > semester:
                        print("Outdated Entry: " + person.name)
                        date_sheet.write(i, 0, person.name)
                        i += 1
                        any_false.append(False)
                except:
                    print("Invalid Entry: " + person.name)
                    date_sheet.write(i, 0, person.name)
                    i += 1
                    any_false.append(False)

    if all(any_false):
        print("\n" + "Congrats, All people are up-to-date! ")
        return 0

    return i - len(departments)


def check_duplicates(departments, duplicates_sheet, style, any_false=[], i=0):

    # status
    print("\n" + "Checking duplicates...")

    for department in departments:
        duplicates_sheet.write(i, 0, department.name, style)
        i += 1

        unique_persons = {}
        for person in department.people:
            if person.name not in unique_persons:
                unique_persons[person.name] = 1
            else:
                print("Duplicated Entry: " + person.name)
                duplicates_sheet.write(i, 0, person.name)
                i += 1
                any_false.append(False)

    if all(any_false):
        print("\n" + "Congrats, All people are up-to-date! ")
        return 0

    return i - len(departments)


init()

# Miscellaneous Functions


def keyword_search(keywords, people):

    # keyword search
    for person in people:
        if keywords in person.knowledge:
            print(person.name)


def university_search(unviersity, people):

    # university search
    for person in people:
        if unviersity in person.institution:
            print(person.name)


def urm_search(people):

    # urm search
    for person in people:
        if person.urm == 'URM' and person.gender == 'Female':
            print(person.name)




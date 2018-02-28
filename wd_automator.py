import openpyxl
from selenium import webdriver
import datetime
import traceback
# traceback.print_exc()
import time
import os

from selenium.common.exceptions import NoSuchElementException

wb = openpyxl.load_workbook("test2.xlsx")
sheet = wb["Sheet1"]
driver = webdriver.Firefox()


class ExcelParser(object):
    def parse_last_name(self, row):
        lastname_cell = sheet["A" + str(row)]
        return lastname_cell.value

    def parse_sin(self, row):
        sin_cell = sheet["D" + str(row)]
        return sin_cell.value

    def parse_end_month(self, row):
        course_end_cell = sheet["Y" + str(row)].value
        course_end_cell = str(course_end_cell)[:10]
        return datetime.datetime.strptime(course_end_cell, "%Y-%m-%d").month

    def parse_end_year(self, row):
        course_end_cell = sheet["Y" + str(row)].value
        course_end_cell = str(course_end_cell)[:10]
        return datetime.datetime.strptime(course_end_cell, "%Y-%m-%d").year

    def parse_type(self, row):
        type_cell = sheet["AB" + str(row)]
        if type_cell.value[0:2].upper() == "WD" or type_cell.value[0:2].upper() == "UC" or type_cell.value[
                                                                                           0:2].upper() == "EC":
            return type_cell.value[0:2]
        else:
            print("Error, couldn't parse type")
    # TODO: The following don't really work if the input isn't perfect, take out the slicing and make it so you put one option per cell
    # 1: Student did not successfully complete all funded courses. If student withdrew, provide end date.
    # 2: Student withdrew at end of first term (provide end date).
    # 3: Student successfully completed studies early (provide end date).
    # 4: Student increased course load to full-time (provide end date).
    def parse_reason(self, row):
        reason_cell = sheet["AB" + str(row)]
        if reason_cell.value[4] == "1" or reason_cell.value[4] == "2" or reason_cell.value[4] == "3" or \
                reason_cell.value[4] == "4":
            return reason_cell.value[4]
        else:
            print("Error, couldn't parse reason")

    def parse_nonpunitive(self, row):
        nonpunitive_cell = sheet["AB" + str(row)]
        if nonpunitive_cell.value[7] == "Y".upper():
            return True
        else:
            return False

    def parse_date(self, row):
        date_cell = sheet["AB" + str(row)]

        if sheet["AB" + str(row)].value[7] == "Y".upper():
            if date_cell.value[10]:
                return date_cell.value[10:]
            else:
                print("No date specified")
                return
        else:
            if date_cell.value[7]:
                return date_cell.value[7:]
            else:
                return


class StudentInfo(object):

    def __init__(self, last_name, sin, end_month, end_year, type, reason, nonpunitive=False,
                 last_attended_date=None):
        self.last_name = last_name
        self.sin = sin
        self.end_month = end_month
        self.end_year = end_year
        self.type = type
        self.reason = reason

        if nonpunitive is False:
            self.nonpunitive = False
        else:
            self.nonpunitive = True

        if last_attended_date is None:
            self.last_attended_date = ""
        else:
            self.last_attended_date = last_attended_date

    def get_last_name(self):
        return self.last_name

    def get_sin(self):
        return self.sin

    def get_end_month(self):
        return self.end_month

    def get_end_year(self):
        return self.end_year

    def get_type(self):
        return self.type

    def get_reason(self):
        return self.reason

    def get_nonpunitive(self):
        return self.nonpunitive

    def get_last_attended_date(self):
        return self.last_attended_date

    def print_student_info(self):
        print(self.last_name, self.sin, self.end_month, self.end_year, self.type, self.reason, self.nonpunitive,
              self.last_attended_date)

    def write_error(self, item, excel_row):
        with open("errors.txt", 'a+') as f:
            f.write("\n" + item + " could not be verified on excel row " + str(excel_row))


class WebsiteNavigator(object):

    def open_sail(self, link, username, password):
        driver.get(link)

        driver.find_element_by_id("user").send_keys(username)
        driver.find_element_by_id("password").send_keys(password)
        time.sleep(1)
        driver.find_element_by_name("btnSubmit").click()

    def enter_sin(self, sin):
        driver.find_element_by_id("sin").send_keys(sin)
        driver.find_element_by_id("searchButton").click()

    def check_identity(self, name, end_year, end_month, current_row):

        months = {"Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6, "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10,
                  "Nov": 11, "Dec": 12}

        month_string = ""

        for string, num in months.items():
            if end_month == str(num):
                month_string = string

        container = driver.find_element_by_id("main_container")
        base_table = container.find_element_by_class_name("content")

        try:
            if name.upper() not in str(base_table.find_element_by_xpath('//table[@class="border"]/tbody/tr[2]/td[2]').text):
                s.write_error(name, current_row)
                return

            if end_year not in str(base_table.find_element_by_xpath('//table[@class="border"]/tbody/tr[2]/td[6]').text):
                s.write_error("End date (year)", current_row)
                return

            if month_string not in str(base_table.find_element_by_xpath('//table[@class="border"]/tbody/tr[2]/td[6]').text):
                s.write_error("End date (month)", current_row)
                return

            if str(base_table.find_element_by_xpath('//table[@class="border"]/tbody/tr[2]/td[7]').text) != \
                    "Award Notice Sent":
                s.write_error("App status", current_row)
                return

            # If it gets this far then all tests passed, move on to withdrawal reporting page
            driver.find_element_by_xpath('//*[@title="Application Number"]').click()
            driver.find_element_by_xpath('//ul[@class="sailTabs"]/li[6]//div[@class="greyRt"]').click()

        except:
            driver.find_element_by_id("sin").clear()
            s.write_error("SIN", current_row)

    def get_name(self):
        return input("Enter username: ")

    def get_pass(self):
        return input("Enter password: ")


p = ExcelParser()
wn = WebsiteNavigator()

wn.open_sail("xxx", wn.get_name(), wn.get_pass())
print("Logging in....")

while True:
    try:
        if driver.find_element_by_id("user"):
            print("Incorrect credentials, try again...")
            driver.find_element_by_id("user").send_keys(wn.get_name())
            driver.find_element_by_id("password").send_keys(wn.get_pass())
            time.sleep(1)
            driver.find_element_by_name("btnSubmit").click()
    except NoSuchElementException:
        while True:
            begin = input("Login successful. Begin final entry into SAIL? [y] or [n]")
            if begin.upper() == "Y":
                break
            elif begin.upper() == "N":
                print("Exiting...")
                quit()
            else:
                print("Enter [y] or [n]")

    break

for r in range(2, sheet.max_row + 1):
    s = StudentInfo(p.parse_last_name(r), p.parse_sin(r), p.parse_end_month(r),
                    p.parse_end_year(r), p.parse_type(r), p.parse_reason(r), p.parse_nonpunitive(r), p.parse_date(r))

    wn.enter_sin(s.get_sin())
    wn.check_identity(s.get_last_name(), str(s.get_end_year()), str(s.get_end_month()), r)

# Delete unnecessary error file if none exist
if os.path.exists("/errors.txt") and os.path.getsize("/errors.txt") > 0:
    os.remove("errors.txt")

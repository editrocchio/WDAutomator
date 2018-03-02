import openpyxl
from selenium import webdriver
from selenium.webdriver.support.ui import Select
import datetime
import time
import os
import getpass
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementNotInteractableException

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

    def parse_wd_type(self, row):
        wd_type_cell = sheet["AB" + str(row)]
        if wd_type_cell.value.upper() == "WD" or wd_type_cell.value.upper() == "UC" or wd_type_cell.value.upper() == \
                "EC":
            return wd_type_cell.value

    # 1: Student did not successfully complete all funded courses. If student withdrew, provide end date.
    # 2: Student withdrew at end of first term (provide end date).
    # 3: Student successfully completed studies early (provide end date).
    # 4: Student increased course load to full-time (provide end date).
    def parse_reason(self, row):
        reason_cell = sheet["AC" + str(row)]
        if reason_cell.value == 1 or reason_cell.value == 2 or reason_cell.value == 3 or reason_cell.value == 4:
            return reason_cell.value

    def parse_nonpunitive(self, row):
        nonpunitive_cell = sheet["AD" + str(row)]
        if nonpunitive_cell.value is not None and nonpunitive_cell.value.upper() == "Y":
            return True
        else:
            return False

    def parse_date(self, row):
        date_cell = sheet["AE" + str(row)]

        if date_cell.value is not None:
            return date_cell.value


class StudentInfo(object):

    def __init__(self, last_name, sin, end_month, end_year, wd_type, reason, nonpunitive=False,
                 last_attended_date=None):
        self.last_name = last_name
        self.sin = sin
        self.end_month = end_month
        self.end_year = end_year
        self.wd_type = wd_type
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

    def get_wd_type(self):
        return self.wd_type

    def get_reason(self):
        return self.reason

    def get_nonpunitive(self):
        return self.nonpunitive

    def get_last_attended_date(self):
        return self.last_attended_date

    def print_student_info(self):
        print(self.last_name, self.sin, self.end_month, self.end_year, self.wd_type, self.reason, self.nonpunitive,
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

            # If it gets this far then all tests passed, move on to reporting the withdrawal
            wn.enter_withdrawal_info(current_row)

        # Throws exception if SIN didn't work, clear the box and write error
        except:
            driver.find_element_by_id("sin").clear()
            s.write_error("SIN", current_row)

    def enter_withdrawal_info(self, row):
        driver.find_element_by_xpath('//*[@title="Application Number"]').click()
        driver.find_element_by_xpath('//ul[@class="sailTabs"]/li[6]//div[@class="greyRt"]').click()

        if s.get_wd_type() is not None and s.get_wd_type().upper() == "WD":
            driver.find_element_by_id("withdrawalForm_withdrawalBean_withdrawalTypeId790").click()
        elif s.get_wd_type() is not None and s.get_wd_type().upper() == "EC":
            driver.find_element_by_id("withdrawalForm_withdrawalBean_withdrawalTypeId791").click()
        elif s.get_wd_type() is not None and s.get_wd_type().upper() == "UC":
            driver.find_element_by_id("withdrawalForm_withdrawalBean_withdrawalTypeId792").click()
        else:
            s.write_error("Type of withdrawal", row)
            try:
                driver.find_element_by_link_text("Search Applications").click()
            except ElementNotInteractableException:
                driver.find_element_by_link_text("Search").click()
                driver.find_element_by_link_text("Search Applications").click()
            return

        if s.get_reason() == 1:
            driver.find_element_by_id("withdrawalForm_withdrawalBean_withdrawalReasonId793").click()
        elif s.get_reason() == 2:
            driver.find_element_by_id("withdrawalForm_withdrawalBean_withdrawalReasonId794").click()
        elif s.get_reason() == 3:
            driver.find_element_by_id("withdrawalForm_withdrawalBean_withdrawalReasonId795").click()
        elif s.get_reason() == 4:
            driver.find_element_by_id("withdrawalForm_withdrawalBean_withdrawalReasonId796").click()
        else:
            s.write_error("Reason for withdrawal", row)
            try:
                driver.find_element_by_link_text("Search Applications").click()
            except ElementNotInteractableException:
                driver.find_element_by_link_text("Search").click()
                driver.find_element_by_link_text("Search Applications").click()
            return

        # Remaining entries are optional, so no error reporting if they don't exist
        if s.get_nonpunitive():
            Select(driver.find_element_by_id("withdrawal_nonPunitive")).select_by_visible_text("Yes")
        elif not s.get_nonpunitive():
            Select(driver.find_element_by_id("withdrawal_nonPunitive")).select_by_visible_text("No")
        else:
            pass

        if s.get_last_attended_date() is not None:
            driver.find_element_by_id("withdrawal_dateLastAttended").send_keys(s.get_last_attended_date())
        else:
            pass

        #TODO: Uncomment this when testing on real records. Also will need to go back to SIN entry page for next record
        # driver.find_element_by_id("saveButton").click()

    def get_name(self):
        return input("Enter username: ")

    def get_pass(self):
        return getpass.getpass()


p = ExcelParser()
wn = WebsiteNavigator()

wn.open_sail("https://www.sail.aved.gov.bc.ca/sail/search/applicationSearch.action", wn.get_name(), wn.get_pass())
print("Logging in....")

# Look for the username box and if it throws an exception then the credentials were correct
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
                    p.parse_end_year(r), p.parse_wd_type(r), p.parse_reason(r), p.parse_nonpunitive(r), p.parse_date(r))

    wn.enter_sin(s.get_sin())
    wn.check_identity(s.get_last_name(), str(s.get_end_year()), str(s.get_end_month()), r)

# Delete unnecessary error file if none exist
if os.path.exists("/errors.txt") and os.path.getsize("/errors.txt") > 0:
    os.remove("errors.txt")

import os
import sys
import time
from RPA.FileSystem import FileSystem
from RPA.PDF import PDF
import xlwt
from datetime import timedelta
from RPA.Browser.Selenium import Selenium
from SeleniumLibrary.errors import ElementNotFound

browser = Selenium()
lib = FileSystem()
pdf = PDF()

try:
    WEB_URL = os.environ['WEB_URL']
    AGENCY = os.environ['AGENCY']
except KeyError:
    WEB_URL = "https://itdashboard.gov/"
    AGENCY = "National Archives and Records Administration"

OUTPUT_DIR = f"{os.getcwd()}/output/"


def open_the_website(url):
    browser.open_available_browser(url, headless=True)


def get_all_spendings():
    """
    Function to get spendings of all agencies
    """
    browser.click_element("link:DIVE IN")
    element = 'id:agency-tiles-widget'
    browser.wait_until_element_is_visible(element)
    count_agencies = browser.get_element_count("class:seals")
    spendings_info = browser.get_text(element).split("\n")
    excel_workbook = xlwt.Workbook()
    excel_worksheet = excel_workbook.add_sheet('Agencies')
    excel_worksheet.write(0, 0, "Department Title")
    excel_worksheet.write(0, 1, "Total Spendings")
    step = 4
    dept_name = 0
    spending_amount = 2
    for i in range(count_agencies):
        excel_worksheet.write(i + 1, 0, spendings_info[dept_name])
        excel_worksheet.write(i + 1, 1, spendings_info[spending_amount])
        dept_name += step
        spending_amount += step
    excel_workbook.save(f'{OUTPUT_DIR}/Agencies.xls')


def individual_spendings(agency):
    try:
        browser.click_element(f"partial link:{agency}")
    except ElementNotFound:
        print("Incorrect value entered. Please enter a valid US Agency title")
        sys.exit()
    selector_element = "css:select"
    browser.wait_until_element_is_visible(selector_element, timeout=timedelta(seconds=20))
    browser.select_from_list_by_index(selector_element, "3")
    browser.wait_until_page_does_not_contain_element("//*[@id='investments-table-object_paginate']/span/a[2]",
                                                     timeout=timedelta(seconds=15))
    table = 'css:table#investments-table-object'

    row_count = browser.get_element_count("xpath://*[@id='investments-table-object']/tbody/tr")

    excel_workbook = xlwt.Workbook()
    excel_worksheet = excel_workbook.add_sheet('Investments')
    excel_worksheet.write(0, 0, "UII")
    excel_worksheet.write(0, 1, "Bureau")
    excel_worksheet.write(0, 2, "Investment Title")
    excel_worksheet.write(0, 3, "Total FY2021 Spending ($M)")
    excel_worksheet.write(0, 4, "Type")
    excel_worksheet.write(0, 5, "CIO Rating")
    excel_worksheet.write(0, 6, "# of Projects")
    excel_worksheet.write(0, 7, "Compare with PDF")

    for row in range(row_count):
        try:
            browser.wait_until_page_contains_element(table, timeout=timedelta(seconds=30))
        except AssertionError:
            browser.reload_page()
            selector_element = "css:select"
            browser.wait_until_element_is_visible(selector_element, timeout=timedelta(seconds=30))
            browser.select_from_list_by_index(selector_element, "3")
            browser.wait_until_page_does_not_contain_element("//*[@id='investments-table-object_paginate']/span/a[2]",
                                                             timeout=timedelta(seconds=25))
        finally:
            unique_investment_identifier = browser.get_table_cell(table, 3 + row, 1)
        for column in range(7):
            excel_worksheet.write(1 + row, column, browser.get_table_cell(table, 3 + row, 1 + column))

        try:
            link = browser.get_element_attribute(f"link:{unique_investment_identifier}", attribute="href")
            investment_title = browser.get_table_cell(table, 3 + row, 3)
            browser.execute_javascript(f"window.open('{link}')")
            browser.switch_window("NEW")
            time.sleep(5)
            browser.click_element("link:Download Business Case PDF")
            time.sleep(10)
            browser.close_window()
            text_from_pdf = pdf.get_text_from_pdf(f"{unique_investment_identifier}.pdf")
            uii = text_from_pdf[1].split('Section')[1].split('2. ')[-1].split(': ')[1]
            investment_name = text_from_pdf[1].split('Section')[1].split('2. ')[0].split('. ')[-1].split(': ')[1]
            if unique_investment_identifier == uii and investment_title == investment_name:
                excel_worksheet.write(1 + row, 7, "Equal")
            else:
                excel_worksheet.write(1 + row, 7, "Not Equal")
            # assert unique_investment_identifier == uii, 'Error'
            # assert investment_title == investment_name, 'Error'
            pdf.close_pdf()

            browser.switch_window("MAIN")
        except ElementNotFound:
            excel_worksheet.write(1 + row, 7, "--")

    excel_workbook.save(f'{OUTPUT_DIR}/Individual Investments of {agency}.xls')

    pdf_files = lib.find_files("**/*.pdf")
    if pdf_files:
        try:
            lib.move_files(pdf_files, f"{OUTPUT_DIR}")
        except FileExistsError:
            pass


def main():
    try:
        open_the_website(WEB_URL)
        get_all_spendings()
        individual_spendings(AGENCY)

    finally:
        browser.close_all_browsers()


if __name__ == "__main__":
    main()

import time
import os
from robocorp import browser
from robocorp.tasks import task
from RPA.Excel.Files import Files as Excel

# Set your username and password. Robocorp Vault is recommended 
# Set the HEADLESS argument to True for faster execution without the visible browser. 
USERNAME = "eki@example.com"
PASSWORD = "MyPassword"
HEADLESS = False 

@task
def solve_challenge():
    """Solve the automation challenge"""
    browser.configure(
        browser_engine="chrome",
        screenshot="only-on-failure",
        headless=HEADLESS,
    )

    page = browser.goto("https://www.theautomationchallenge.com/")
    page.get_by_role("button", name="SIGN UP OR LOGIN").click()
    page.get_by_role("button", name="OR LOGIN", exact=True).click()
    page.get_by_role("textbox", name="Email").fill(USERNAME)
    page.get_by_role("textbox", name="Password").fill(PASSWORD)
    page.get_by_role("button", name="LOG IN").click()

    with page.expect_download() as download_info:
        page.get_by_text("Download Excel Spreadsheet").click()
    download = download_info.value
    filepath = os.path.join(os.getenv("ROBOT_ROOT"), download.suggested_filename)
    download.save_as(filepath)

    excel = Excel()
    excel.open_workbook(filepath)
    rows = excel.read_worksheet("data", header=True)
    excel.close_workbook()

    field_names = [["Company Name", "company_name"], ["Address", "company_address"], ["EIN", "employer_identification_number"], ["Sector", "sector"], ["Automation Tool", "automation_tool"], ["Annual Saving", "annual_automation_saving"], ["Date", "date_of_first_project"]]
    page.get_by_role("button", name="Start").click()

    for row in rows:
        # remove possible capthca elements from the DOM :)
        page.add_script_tag(content="document.querySelector('.CustomElement').remove();")
        page.add_script_tag(content="document.querySelector('.greyout').remove();")
        page.add_script_tag(content="document.querySelector('.Popup').remove();")
        for field in field_names:
            text_content, spreadsheet_header = field
            input_selector = f'//div[@class="bubble-element Group"][.="{text_content}"][not(ancestor-or-self::*[contains(@style,"display: none;")])]//input'
            element = page.query_selector(input_selector)
            page.fill(input_selector,row[spreadsheet_header])
       
        page.get_by_role("button", name="Submit").click()

    os.remove(filepath)
    time.sleep(5)
    browser.screenshot()
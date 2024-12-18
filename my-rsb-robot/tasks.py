from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF
from RPA.Robocorp.WorkItems import WorkItems
import logging
import json

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

@task
def robot_spare_bin_python():
    """Insert the sales data for the week and export it as a PDF"""
    try:
        browser.configure(slowmo=100)
        open_the_intranet_website()
        log_in()
        download_excel_file()
        fill_form_with_excel_data()
        collect_results()
        export_as_pdf()
        log_out()
        logging.info("Process completed successfully!")
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")

def open_the_intranet_website():
    """Navigates to the given URL"""
    try:
        browser.goto("https://robotsparebinindustries.com/")
        logging.info("Website opened successfully.")
    except Exception as e:
        logging.error(f"Failed to open the website: {e}")
        raise

def read_credentials(file_path="credentials.json"):
    """open and load configuration file"""
    try:
        with open(file_path, 'r') as file:
            return json.load(file)
    except Exception as e:
        # Handle the exception
        print(f"Could not read credentials: {e}")

def log_in():
    """Fills in the login form and clicks the 'Log in' button"""
    credentials = read_credentials()
    username = credentials.get("username")
    password = credentials.get("password")
    try:
        page = browser.page()
        page.fill("#username", username)
        page.fill("#password", password)
        page.click("button:text('Log in')")
        logging.info("Logged in successfully.")
    except Exception as e:
        logging.error(f"Failed to log in: {e}")
        raise

def fill_and_submit_sales_form(sales_rep):
    """Fills in the sales data and clicks the 'Submit' button"""
    try:
        page = browser.page()
        page.fill("#firstname", sales_rep["First Name"])
        page.fill("#lastname", sales_rep["Last Name"])
        page.select_option("#salestarget", str(sales_rep["Sales Target"]))
        page.fill("#salesresult", str(sales_rep["Sales"]))
        page.click("text=Submit")
        logging.info(f"Form submitted for {sales_rep['First Name']} {sales_rep['Last Name']}")
    except Exception as e:
        logging.error(f"Failed to submit the form for {sales_rep['First Name']} {sales_rep['Last Name']}: {e}")
        raise

def download_excel_file():
    """Downloads Excel file from the given URL"""
    try:
        http = HTTP()
        http.download(url="https://robotsparebinindustries.com/SalesData.xlsx", overwrite=True)
        logging.info("Excel file downloaded successfully.")
    except Exception as e:
        logging.error(f"Failed to download Excel file: {e}")
        raise

def fill_form_with_excel_data():
    """Read data from Excel and fill in the sales form"""
    try:
        excel = Files()
        excel.open_workbook("SalesData.xlsx")
        worksheet = excel.read_worksheet_as_table("data", header=True)
        excel.close_workbook()
        logging.info("Excel file read successfully.")

        for row in worksheet:
            try:
                fill_and_submit_sales_form(row)
            except Exception as e:
                logging.warning(f"Skipping row due to error: {e}")
                continue  # Continue with the next row if one fails
    except Exception as e:
        logging.error(f"Failed to process Excel data: {e}")
        raise

def collect_results():
    """Take a screenshot of the page"""
    try:
        page = browser.page()
        page.screenshot(path="output/sales_summary.png")
        logging.info("Screenshot saved successfully.")
    except Exception as e:
        logging.error(f"Failed to take a screenshot: {e}")
        raise

def export_as_pdf():
    """Export the data to a PDF file"""
    try:
        page = browser.page()
        sales_results_html = page.locator("#sales-results").inner_html()

        pdf = PDF()
        pdf.html_to_pdf(sales_results_html, "output/sales_results.pdf")
        logging.info("PDF generated successfully.")
    except Exception as e:
        logging.error(f"Failed to export PDF: {e}")
        raise

def log_out():
    """Presses the 'Log out' button"""
    try:
        page = browser.page()
        page.click("text=Log out")
        logging.info("Logged out successfully.")
    except Exception as e:
        logging.error(f"Failed to log out: {e}")
        raise

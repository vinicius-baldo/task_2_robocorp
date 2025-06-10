from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables

@task
def robot_spare_bin_python():
    """Insert the sales data for the week and export it as a PDF"""
    browser.configure(
        slowmo=500,
    )
    download_excel_file()
    open_the_intranet_website()
    fill_form_with_csv_data()
    archive_receipts()

def download_excel_file():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def fill_form_with_csv_data():
    """Read data from csv and fill in the sales form"""
    tables = Tables()
    orders = tables.read_table_from_csv("orders.csv", header=True)

    for row in orders:
        fill_robot_form(row)

def retry_order_submission(page, max_attempts=3):
    """Retry the order submission if it fails"""
    for attempt in range(max_attempts):
        try:
            page.click("#order")
            page.wait_for_selector("#receipt", state="visible", timeout=5000)
            return True
        except Exception as e:
            if attempt == max_attempts - 1:  # Last attempt
                raise Exception(f"Failed to submit order after {max_attempts} attempts: {str(e)}")
            page.wait_for_timeout(100)  # Wait a bit before retrying

def fill_robot_form(row):
    """Fills the robot form with data from a row"""
    page = browser.page()
    
    # Wait for the form to be visible
    page.wait_for_selector("#head")
    
    # Fill in the form fields with proper selectors
    page.select_option("#head", row["Head"])
    page.click(f"#id-body-{row['Body']}")
    page.fill("input[placeholder='Enter the part number for the legs']", row["Legs"])
    page.fill("#address", str(row["Address"]))
    
    # Click the preview button
    page.click("#preview")
    
    # Wait for preview to be generated
    page.wait_for_selector("#robot-preview-image")
    
    # Submit the order with retry mechanism
    retry_order_submission(page)
    
    # Store the receipt as a PDF
    order_number = row["Order number"]
    store_receipt_as_pdf(order_number)

    clear_screen()

def clear_screen():
    """Clears the screen for the next order"""
    page = browser.page()
    page.click("#order-another")
    page.wait_for_timeout(100)  # Wait a bit to reload page
    page.click("button:text('OK')")

def store_receipt_as_pdf(order_number):
    """Stores the receipt as a PDF file"""
    page = browser.page()
    page.click("#receipt")
    
    # Wait for the receipt to be visible
    page.wait_for_selector("#receipt", state="visible")
    
    # Save the receipt as PDF
    page.pdf(path=f"receipt-{order_number}.pdf")


def take_screenshot(order_number):
    """Takes a screenshot of the current page"""
    page = browser.page()
    page.screenshot(path=f"screenshot-{order_number}.png")


def archive_receipts():
    """Archive all PDF receipts in a zip file"""
    import zipfile
    import os

    with zipfile.ZipFile("receipts.zip", "w") as zipf:
        for filename in os.listdir("."):
            if filename.startswith("receipt-") and filename.endswith(".pdf"):
                zipf.write(filename)

def open_the_intranet_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    page = browser.page()
    page.click("button:text('OK')")
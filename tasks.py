from typing import Any
from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.PDF import PDF
from RPA.Tables import Tables
from RPA.Archive import Archive


@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(slowmo=100)
    orders = get_orders()
    open_robot_order_website()
    for order in orders:
        close_constitutional_rights_modal()
        fill_form(order)
        pdf_path = store_receipt_as_pdf(order["Order number"])
        screenshot_path = screenshot_robot(order["Order number"])
        embed_screenshot_to_receipt(
            screenshot_path=screenshot_path, receipt_path=pdf_path
        )
        order_another_robot()
    archive_receipts()


def open_robot_order_website():
    """Opens the RobotSpareBin website"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")


def download_orders_csv():
    """Downloads the orders .csv file"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)


def get_orders() -> list[dict[str, Any]]:
    """Get orders from .csv"""
    download_orders_csv()
    library = Tables()
    orders = library.read_table_from_csv("orders.csv")
    return orders.to_list()


def close_constitutional_rights_modal():
    """Closes annoying modal"""
    page = browser.page()
    page.click("text=Ok")


def fill_form(row: dict[str, Any]):
    """Fill order form"""
    page = browser.page()
    page.select_option("#head", row["Head"])
    page.click(f"#id-body-{row["Body"]}")
    # page.click("text=Show model info")
    page.fill('input[placeholder="Enter the part number for the legs"]', row["Legs"])
    page.fill("#address", row["Address"])
    # page.click("text=Preview")
    page.click("button:text('ORDER')")
    while page.is_visible(
        "#root > div > div.container > div > div.col-sm-7 > div.alert.alert-danger"
    ):
        page.click("button:text('ORDER')")


def store_receipt_as_pdf(order_number: str):
    """Export the data to a pdf file"""
    page = browser.page()
    receipt_html = page.locator("#receipt").inner_html()

    pdf = PDF()
    path = f"output/receipts/order_{order_number}_receipt.pdf"
    pdf.html_to_pdf(receipt_html, path)
    return path


def screenshot_robot(order_number: str):
    """Take a screenshot of the robot page"""
    page = browser.page()
    path = f"output/screenshots/order_{order_number}_robot.png"
    page.query_selector("#robot-preview-image").screenshot(path=path, scale="css")
    return path


def embed_screenshot_to_receipt(screenshot_path: str, receipt_path: str):
    """Add screenshot to receipt pdf"""
    pdf = PDF()
    pdf.add_files_to_pdf(
        files=[screenshot_path], target_document=receipt_path, append=True
    )


def order_another_robot():
    """Go to the order page again"""
    page = browser.page()
    page.click("#order-another")


def archive_receipts():
    """Creates a zip of receipts pdfs"""
    library = Archive()
    library.archive_folder_with_zip(
        folder="output/receipts", archive_name="output/receipts.zip"
    )

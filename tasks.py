from robocorp.tasks import task
from robocorp import browser
#from RPA.Browser.Selenium import Selenium
import time
from RPA.Excel.Files import Files
import os
import re
import subprocess
from datetime import datetime, timedelta
import logging
from tenacity import retry, wait_exponential, stop_after_attempt, retry_if_exception_type
from datetime import datetime
import argparse



# Setup logging
logging.basicConfig(filename="automation.log", level=logging.INFO, 
                    format="%(asctime)s - %(levelname)s - %(message)s")

# Retry settings
@retry(wait=wait_exponential(multiplier=1, min=4, max=10), stop=stop_after_attempt(3),
       retry=retry_if_exception_type(Exception))
def retryable_function(func, *args, **kwargs):
    return func(*args, **kwargs)



@task
def ROBOT_RAUL_ANOTONIO_HERNANDEZ_MOJICA_Thoughtful_Challenge():
    """Extracting data from a news website (LOS ANGELES TIMES)"""

    args = parse_arguments()

    search_text = args.search_text
    news_category = args.news_category
    months = args.months

    # Hardcoded values for testing
    #search_text = "Bitcoin"
    #news_category = "newest"
    #months = "1"

    try:
        kill_excel_process()
        browser.configure(slowmo=500)
        open_news_website()
        search_phrase(search_text)
        choose_latest(news_category)
        extract_news(search_text, months)
    except Exception as e:
        logging.error(f"Error during task execution: {e}")
    finally:
        logging.info("Task completed. Closing all resources.")
        close_browser()
        kill_excel_process()



def parse_arguments():
    parser = argparse.ArgumentParser(description="Extract news data from a news website.")
    parser.add_argument("--search_text", type=str, default="Bitcoin", help="Text to search for")
    parser.add_argument("--news_category", type=str, default="all", help="Category of news to search in. eg.:newest")
    parser.add_argument("--months", type=int, default=1, help="Number of months to search within")
    return parser.parse_args()



def open_news_website():
    """Navigates to the given URL"""
    retryable_function(browser.goto, "https://www.latimes.com/")

def search_phrase(search_text):
    """Clicks the 'Search' icon, fills in the search input textbox, and clicks the 'Search' button"""
    page = browser.page()
    retryable_function(page.wait_for_selector, "[data-element='search-button']", state="visible", timeout=20000)
    page.click("[data-element='search-button']")
    page.fill("[data-element='search-form-input']", search_text)
    page.click("[type='submit']")

def choose_latest(news_category):
    """Opens the Sort By selector, and selects NEWEST"""

    # Check if the news_category is "newest"
    if news_category.lower() == "newest":
        page = browser.page()
        retryable_function(page.wait_for_selector, "select[name='s']", state="visible", timeout=20000)
        time.sleep(5)
        retryable_function(page.select_option, "select[name='s']", label="Newest")
        time.sleep(5)

def extract_news(search_text, months):
    """Extract Title, Description, and Date for each record on the result list"""
    page = browser.page()

    output_dir = os.path.join(os.getcwd(), "output")
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    images_dir = os.path.join(output_dir, "images")
    if not os.path.exists(images_dir):
        os.makedirs(images_dir)

    excel_file_path = os.path.join(output_dir, "news_data.xlsx")

    excel = Files()
    excel.create_workbook(excel_file_path)
    excel.create_worksheet("Sheet1")
    excel.append_rows_to_worksheet([["Title", "Description", "Date", "Image Filename", "Search Count", "Contains Money"]], "Sheet1")

    months = int(months) + 1
    threshold_date = datetime.now() - timedelta(days=months * 30)
    threshold_date = threshold_date.replace(day=1)

    stop_collection = False

    while True:
        try:
            retryable_function(page.wait_for_selector, "li ps-promo", state="visible", timeout=10000)
            items = browser.page().locator("li ps-promo").all()

            data = []
            for item in items:
                title = item.locator(".promo-title a").inner_text()
                description = item.locator(".promo-description").inner_text()
                date = item.locator(".promo-timestamp").inner_text()
                news_date = parse_news_date(date)

                if news_date < threshold_date:
                    logging.info(f"Stopping collection as news date {news_date} is older than threshold {threshold_date}.")
                    stop_collection = True
                    break

                img_src = item.locator(".image").get_attribute("src")

                # Generate the image filename using the current datetime
                current_time = datetime.now().strftime("IMG_%Y%m%d_%H%M%S%f")
                image_extension = ".jpg"

                # Create the final image filename
                img_filename = current_time + image_extension

                img_filepath = os.path.join(images_dir, img_filename)
                download_image(img_src, img_filepath)

                search_count = count_occurrences(search_text, title, description)
                contains_money = contains_money_amount(title, description)
                data.append([title, description, date, img_filename, search_count, contains_money])

            excel.append_rows_to_worksheet(data, "Sheet1")

            if stop_collection:
                break

            next_button = browser.page().locator("a[rel='nofollow']:has-text('Next')")
            if next_button.is_visible():
                next_button.click()
            else:
                break
        except Exception as e:
            logging.error(f"Error during news extraction: {e}")
            break

    excel.save_workbook()
    excel.close_workbook()

def download_image(image_url, save_path):
    page = browser.page()
    new_page = page.context.new_page()
    retryable_function(new_page.goto, image_url)
    new_page.locator("img").screenshot(path=save_path)
    logging.info(f"Image saved to {save_path}")
    new_page.close()




def count_occurrences(search_text, title, description):
    """Count occurrences of search_text in title and description"""
    return title.lower().count(search_text.lower()) + description.lower().count(search_text.lower())

def contains_money_amount(title, description):
    """Check if the title or description contains any amount of money"""
    money_patterns = [
        r"\$\d+(?:,\d{3})*(?:\.\d{1,2})?",   # Matches $11.1, $111,111.11
        r"\d+\s*dollars",                    # Matches 11 dollars
        r"\d+\s*USD"                         # Matches 11 USD
    ]
    combined_text = title + " " + description
    for pattern in money_patterns:
        if re.search(pattern, combined_text, re.IGNORECASE):
            return True
    return False


def parse_news_date(date_str):
    """Attempt to parse the news date using multiple formats"""
    formats = [
        "%B %d, %Y",   # Full month name, e.g., "July 18, 2024"
        "%b. %d, %Y"   # Abbreviated month with a period, e.g., "Aug. 14, 2024"
    ]
    
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    
    # If no format matches, raise an error
    raise ValueError(f"Date format for '{date_str}' not recognized.")



def kill_excel_process():
    """Kill all running Excel processes"""
    try:
        if os.name == 'nt':  # For Windows
            subprocess.call(["taskkill", "/F", "/IM", "EXCEL.EXE"])
        else:  # For Unix-based systems (Linux, macOS)
            subprocess.call(["pkill", "-f", "Excel"])
        print("Excel processes killed successfully.")
    except Exception as e:
        print(f"Error killing Excel processes: {e}")


def close_browser():
    """Close the browser"""
    browser.page().close()  # Closes the current page
    #browser.close()  # Close the entire browser context if needed
import io

import os
import cv2
import re
import numpy as np
import openpyxl
import pandas as pd
import psycopg2
import selenium
from PIL import Image
import pytesseract
from openpyxl.styles import PatternFill, Border, Side
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import time
from openpyxl.styles import PatternFill, Border, Side
from xlsxwriter import Workbook

# Define chart type, month, and year - modify these as needed
CHART_TYPE = "Current"  # Options: "Voltage", "Current", "Energy", "Demand"
MONTH = 6  # Month number (1-12)
YEAR = 2024  # Year (e.g., 2023)

# Parameter configurations by chart type
PARAMETER_CONFIGS = {
    "Voltage": {
        "chart_params": ["Date", "Phase 1", "Phase 2", "Phase 3", "Avg"],
        "db_columns": ["dateandtime", "avgvoltage_ph1", "avgvoltage_ph2", "avgvoltage_ph3", "avgvoltage_avg"],
        "tab_selector": "//span[@class='dx-tab-text-span' and text()='Voltage']"
    },
    "Current": {
        "chart_params": ["Date", "Line 1", "Line 2", "Line 3", "Avg", "Neutral"],
        "db_columns": ["dateandtime", "avgcurrent_ph1", "avgcurrent_ph2", "avgcurrent_ph3", "avgcurrent_avg",
                      "avgneutralcurrent"],
        "tab_selector": "//span[@class='dx-tab-text-span' and text()='Current']"
    },
    "Energy": {
        "chart_params": ["Date", "Active", "Apparent", "Reactive"],
        "db_columns": ["dateandtime", "activepowersum", "apparentpowersum", "reactivepowersum"],
        "tab_selector": "//span[@class='dx-tab-text-span' and text()='Energy']"
    },
    "Demand": {
        "chart_params": ["Date", "Active", "Apparent", "Reactive"],
        "db_columns": ["dateandtime", "avgdemand_kw", "avgdemand_kva", "avgdemand_kvar"],
        "tab_selector": "//span[@class='dx-tab-text-span' and text()='Demand']"
    }
}
# Get the configuration for the selected chart type
# Get the configuration for the selected chart type
if CHART_TYPE in PARAMETER_CONFIGS:
    config = PARAMETER_CONFIGS[CHART_TYPE]
    CHART_PARAMETERS = config["chart_params"]
    DB_COLUMNS = config["db_columns"]
    TAB_SELECTOR = config["tab_selector"]
else:
    # Default if chart type not recognized
    print(f"⚠️ Warning: Chart type '{CHART_TYPE}' not recognized, using Voltage parameters")
    CHART_PARAMETERS = PARAMETER_CONFIGS["Demand"]["chart_params"]
    DB_COLUMNS = PARAMETER_CONFIGS["Demand"]["db_columns"]
    TAB_SELECTOR = PARAMETER_CONFIGS["Demand"]["tab_selector"]
# **Step 1: Set Up WebDriver**
options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ["enable-logging"])

driver = webdriver.Chrome(options=options)
driver.get("https://networkmonitoring.secure.online:10122/Account/Login")
time.sleep(3)  # Allow page to load
driver.maximize_window()

# **Step 2: Enter Login Credentials**
username_input = WebDriverWait(driver, 5).until(
    EC.presence_of_element_located((By.XPATH, "//input[@name='UserName']"))
)
username_input.clear()
username_input.send_keys("SanyamU")

password_input = driver.find_element(By.XPATH, "//input[@name='Password']")
password_input.clear()
password_input.send_keys("Secure@12345")

login_button = driver.find_element(By.ID, 'btnlogin')
login_button.click()


# it will select tenanant name
# driver.find_element(By.XPATH,"//button[@id='btnGo']").click()
# **Step 3: Function to Process CAPTCHA Image**
def preprocess_captcha(image_bytes):
    """Preprocess CAPTCHA image for better OCR accuracy."""
    img = Image.open(io.BytesIO(image_bytes))
    img = img.convert("L")  # Convert to grayscale
    img = np.array(img)

    # Apply adaptive thresholding for better clarity
    img = cv2.adaptiveThreshold(img, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)

    # Morphological operations to clean up noise
    kernel = np.ones((2, 2), np.uint8)
    img = cv2.erode(img, kernel, iterations=1)
    img = cv2.dilate(img, kernel, iterations=1)

    return Image.fromarray(img)


# **Step 4: Function to Extract Text from CAPTCHA**
def extract_captcha_text(image):
    """Extract text from processed CAPTCHA using OCR with regex filtering."""
    pytesseract.pytesseract.tesseract_cmd = r"C:\\Users\\110573\\AppData\\Local\\Programs\\Tesseract-OCR\\tesseract.exe"
    custom_config = r"--psm 8 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"

    for _ in range(3):  # Try multiple times to get a valid CAPTCHA
        raw_text = pytesseract.image_to_string(image, config=custom_config).strip()
        match = re.search(r"(\d{3}[A-Z]{3})", raw_text)  # Match 'NNNLLL' pattern
        if match:
            return match.group(0)  # Return only the valid part

    return ""  # Return empty if no valid match found


# **Step 5: Function to Refresh CAPTCHA**
def refresh_captcha():
    """Refresh CAPTCHA if login fails."""
    try:
        refresh_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//i[@class='float-right fas fa-sync-alt']"))
        )
        refresh_button.click()
        time.sleep(2)
        print("🔄 CAPTCHA Refreshed")
    except (NoSuchElementException, TimeoutException):
        print("⚠️ CAPTCHA refresh button not found!")


# **Step 6: Function to Attempt Login**
def attempt_login():
    """Attempts login and ensures CAPTCHA is always filled."""
    try:
        # Locate CAPTCHA Image and Capture Screenshot
        captcha_element = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, "//canvas[@id='canvas']"))
        )
        captcha_png = captcha_element.screenshot_as_png

        # Process and Extract CAPTCHA Text
        processed_captcha = preprocess_captcha(captcha_png)
        processed_captcha.save("debug_captcha.png")  # Save for debugging
        captcha_text = extract_captcha_text(processed_captcha)

        if not captcha_text:
            print("❌ Failed to extract valid CAPTCHA. Refreshing...")
            return False

        print(f"🔹 Attempting login with CAPTCHA: {captcha_text}")

        # Locate and fill CAPTCHA input field
        captcha_input = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Enter Captcha here..']"))
        )
        captcha_input.clear()
        captcha_input.send_keys(captcha_text)
        time.sleep(1)  # Allow field to register input

        # Ensure CAPTCHA is filled
        filled_captcha = captcha_input.get_attribute("value")
        print(f"🔍 CAPTCHA field value: '{filled_captcha}'")

        if not filled_captcha:
            print("⚠️ CAPTCHA field is empty! Retrying input...")
            captcha_input.send_keys(captcha_text)
            time.sleep(1)

        # Click Login Button using JavaScript (More Reliable)

        driver.execute_script("arguments[0].click();", login_button)

        # Wait for response
        time.sleep(5)

        # Check if login was successful
        current_url = driver.current_url
        print(f"🌍 Current URL after login: {current_url}")

        if "dashboard" in current_url.lower():
            print("✅ Login Successful!")
            return True
        else:
            print("❌ Login Failed. Trying again...")
            return False

    except NoSuchElementException as e:
        print(f"⚠️ Element not found: {e}")
    except TimeoutException:
        print("⏳ Page took too long to load!")

    return False  # Return False if login is not successful


# **Step 7: Execute Login Attempts**
MAX_ATTEMPTS = 4
for attempt in range(MAX_ATTEMPTS):
    login_successful = attempt_login()

    if login_successful:
        break  # Stop immediately if login succeeds
    else:
        refresh_captcha()  # Refresh only if login fails

# **Final Handling**
if not login_successful:
    print("❌ All attempts failed. Please try manually.")

time.sleep(1)


def select_dropdown_option(dropdown_id, option_name):
    """Selects an option from a dropdown dynamically."""
    try:
        # 🔹 Click dropdown to open it
        dropdown = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.ID, dropdown_id)))
        dropdown.click()

        # 🔹 Wait for dropdown options to be visible
        options_list_css = ".dx-list-item"
        WebDriverWait(driver, 3).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, options_list_css)))

        # 🔹 Find all dropdown options
        options = driver.find_elements(By.CSS_SELECTOR, options_list_css)

        # Debugging: Print all available options
        print(f"🔍 Available options for {dropdown_id}: {[opt.text.strip() for opt in options]}")

        for option in options:
            if option.text.strip().lower() == option_name.lower():
                option.click()
                print(f"✅ Selected: {option_name}")
                return

        print(f"❌ Option '{option_name}' not found in {dropdown_id}!")

    except Exception as e:
        print(f"⚠️ Error selecting '{option_name}' in {dropdown_id}: {e}")


# def select_month():
#     """Selects the 'Month' button from the button group."""
#     try:
#         # 🔹 Find and click the 'Month' button
#         month_button = WebDriverWait(driver, 3).until(EC.element_to_be_clickable(
#             (By.XPATH, "//div[contains(@class, 'dx-buttongroup-item') and .//span[text()='Month']]")))
#         month_button.click()
#         print("✅ Selected: Month")
#
#     except Exception as e:
#         print(f"⚠️ Error selecting Month: {e}")


def select_day():
    """Selects the 'Month' button from the button group."""
    try:
        # 🔹 Find and click the 'Month' button
        day_button = WebDriverWait(driver, 3).until(EC.element_to_be_clickable(
            (By.XPATH, "//div[contains(@class, 'dx-buttongroup-item') and .//span[text()='Day']]")))
        day_button.click()
        print("✅ Selected: Day")

    except Exception as e:
        print(f"⚠️ Error selecting Month: {e}")


def fill_date_input(date_value):
    """Fills the date input field with the given value (e.g., 'November 2024')."""
    try:
        # 🔹 Locate the date input field
        date_input = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@aria-label='Date']")))

        # 🔹 Clear existing value and enter new date
        date_input.clear()
        date_input.send_keys(date_value)
        ##date_input.send_keys(Keys.TAB)  # 🔹 Ensures the input is registered

        print(f"✅ Entered Date: {date_value}")

    except Exception as e:
        print(f"⚠️ Error entering date '{date_value}': ")


# 🔹 Select Area
select_dropdown_option("ddl-area", "PIA   ")  # Replace with the actual Area option

# 🔹 Select Substation
select_dropdown_option("ddl-substation", "SS_001")  # Replace with the actual Substation option

# 🔹 Select MV Feeder
select_dropdown_option("ddl-feeder", "SS_001")  # Replace with the actual MV Feeder option

# 🔹 Select Month/day
# select_month()
select_day()
# 🔹 Fill Date Input
fill_date_input("01/06/2024")

# 🔹 Click Right Arrow Button
try:
    arrow_button = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "i.dx-icon.dx-icon-arrowright"))
    )
    arrow_button.click()
    print("✅ Clicked Right Arrow Button")
except Exception as e:
    print(f"⚠️ Error clicking right arrow button: {e}")

time.sleep(8)

# Find the row containing both the specific meter number and MV feeder
meter_number = "Q0792230_1"
mv_feeder = "MV_001"

row_xpath = f"//tr[td[contains(text(), '{meter_number}')] and td[contains(text(), '{mv_feeder}')]]"

try:
    # Locate the row
    row_element = driver.find_element(By.XPATH, row_xpath)

    # Find the "View" link within the same row
    view_link = row_element.find_element(By.XPATH, ".//a[contains(text(), 'View')]")

    # Click the "View" link
    view_link.click()
    print("Successfully clicked on View.")
    time.sleep(4)

except Exception as e:
    print(f"Error: {e}")
time.sleep(4)
# click the "Detailed view for Voltage profile "
driver.find_element(By.XPATH, "//a[@id='VPDetailedLink']").click()

time.sleep(4)

# Click the appropriate chart tab based on chart type
try:
    chart_tab = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, TAB_SELECTOR))
    )
    chart_tab.click()
    print(f"✅ Selected {CHART_TYPE} tab")
except Exception as e:
    print(f"⚠️ Error selecting {CHART_TYPE} tab: {e}")

time.sleep(5)
import os
import re
import time
import numpy as np
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

# Global parameters - update these as needed
CHART_TYPE = "Current"  # Options: "Voltage", "Current", "Energy", "Demand"
# Define parameters for the chart type
CHART_PARAMETERS = ["Time", "Line 1", "Line 2", "Line 3", "Avg", "Neutral"]


class TimeChartDataExtractor:
    def __init__(self, driver):
        self.driver = driver
        self.chart_container = None
        self.chart_x = None
        self.chart_y = None
        self.chart_width = None
        self.chart_height = None
        self.collected_data = []
        self.target_parameters = None

        # Create directories for screenshots and data
        os.makedirs("chart_screenshots", exist_ok=True)
        os.makedirs("extracted_data", exist_ok=True)

    def find_and_setup_chart(self):
        """Find the chart container and set up initial chart properties"""
        try:
            self.chart_container = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "(//*[local-name()='svg' and @class='dxc dxc-chart'])[1]"))
            )
            print("Chart container found successfully")

            chart_rect = self.driver.find_element(By.XPATH,
                                                  "(//*[local-name()='svg' and @class='dxc dxc-chart'])[1]//*[name()='rect']")
            self.chart_width = chart_rect.size['width']
            self.chart_height = chart_rect.size['height']
            self.chart_x = chart_rect.location['x']
            self.chart_y = chart_rect.location['y']

            # Scroll chart into view
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", self.chart_container)
            time.sleep(2)  # Allow time for scrolling and rendering

            # Take a screenshot of the chart area
            self.driver.save_screenshot("chart_screenshots/full_chart.png")
            print(f"Chart dimensions: Width={self.chart_width}, Height={self.chart_height}")
            return True
        except Exception as e:
            print(f"Chart container not found: {str(e)}")
            return False

    def extract_tooltip_data(self, target_params=None):
        """Extract data from tooltip with improved detection for time-based parameters"""
        if target_params is None:
            target_params = CHART_PARAMETERS

        # First, look for tooltip elements with specific class names
        tooltip_selectors = [
            ".dxc-tooltip", ".chart-tooltip", "[role='tooltip']",
            ".highcharts-tooltip", ".tooltip", ".c3-tooltip",
            ".nvtooltip", ".chartjs-tooltip"
        ]

        tooltip_text = None
        for selector in tooltip_selectors:
            try:
                tooltip_elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                if tooltip_elements and tooltip_elements[0].is_displayed():
                    tooltip_text = tooltip_elements[0].text
                    break
            except:
                pass

        if not tooltip_text:
            try:
                # If no tooltip found by selector, try to find any element that might be a tooltip
                tooltip_elements = self.driver.execute_script("""
                    return Array.from(document.querySelectorAll('*')).filter(el => {
                        const style = window.getComputedStyle(el);
                        return style.display !== 'none' &&
                               style.visibility !== 'hidden' &&
                               style.opacity !== '0' &&
                               el.innerText &&
                               (el.innerText.includes(':') || 
                                el.innerText.match(/\\d+\\:\\d+/) || 
                                el.innerText.match(/\\d+\\s*\\w+/));
                    }).map(el => el.innerText);
                """)

                if tooltip_elements and len(tooltip_elements) > 0:
                    tooltip_text = " | ".join(tooltip_elements)
            except:
                pass

        extracted_data = {}

        if tooltip_text:
            # Debug the tooltip text to understand its structure
            print(f"Tooltip text detected: {tooltip_text}")

            # First try to extract time from the tooltip
            time_pattern = r"(\d{1,2}:\d{2})"
            time_match = re.search(time_pattern, tooltip_text)
            if time_match:
                extracted_data["Time"] = time_match.group(1)

            # Extract the explicitly specified parameters
            for param in target_params:
                if param == "Time" and "Time" in extracted_data:
                    continue  # Already extracted time

                match = re.search(rf"{param}:\s*([^\n|]+)", tooltip_text, re.IGNORECASE)
                if match:
                    extracted_data[param] = match.group(1).strip()
                else:
                    # Alternative formats like "Line 1 - 45.6"
                    alt_match = re.search(rf"{param}\s*[\-]\s*([^\n|]+)", tooltip_text, re.IGNORECASE)
                    if alt_match:
                        extracted_data[param] = alt_match.group(1).strip()
                    else:
                        extracted_data[param] = None

            # Look for patterns like "Parameter: Value" or "Parameter - Value"
            additional_params = re.findall(r"([A-Za-z0-9\s]+)(?:\:|\-)\s*([^\n|]+)", tooltip_text)
            for param, value in additional_params:
                param_key = param.strip()
                if param_key not in extracted_data and param_key.lower() not in [p.lower() for p in
                                                                                 extracted_data.keys()]:
                    extracted_data[param_key] = value.strip()

        # Debug what we extracted
        if extracted_data:
            print(f"Extracted data: {extracted_data}")
        else:
            print("No data could be extracted from tooltip")

        return extracted_data, tooltip_text

    def detect_available_parameters(self):
        """Try to detect what parameters are available in the chart by scanning tooltips"""
        print("Detecting available chart parameters...")

        # Move to the middle of the chart to get a sample tooltip
        middle_x = self.chart_x + (self.chart_width / 2)
        middle_y = self.chart_y + (self.chart_height / 2)

        try:
            js_script = f"""
            var evt = new MouseEvent('mousemove', {{
                'view': window,
                'bubbles': true,
                'cancelable': true,
                'clientX': {int(middle_x)},
                'clientY': {int(middle_y)}
            }});
            document.elementFromPoint({int(middle_x)}, {int(middle_y)}).dispatchEvent(evt);
            """
            self.driver.execute_script(js_script)
            time.sleep(1)

            # Try to get the tooltip text
            _, tooltip_text = self.extract_tooltip_data()

            # Try to extract parameter names from the tooltip text
            detected_params = ["Time"]  # Always include Time for time-based charts

            if tooltip_text:
                # Look for patterns like "Parameter: Value" or "Parameter - Value"
                pattern_matches = re.findall(r"([A-Za-z0-9\s]+)(?:\:|\-)\s*([^\n|]+)", tooltip_text)
                for param, _ in pattern_matches:
                    param_key = param.strip()
                    if param_key.lower() != "time" and param_key not in detected_params:
                        detected_params.append(param_key)

            if len(detected_params) > 1:
                print(f"Detected {len(detected_params)} parameters: {', '.join(detected_params)}")
                return detected_params
        except Exception as e:
            print(f"Error detecting parameters: {e}")

        # If we couldn't detect parameters, return the default ones
        print(f"Using default parameters: {', '.join(CHART_PARAMETERS)}")
        return CHART_PARAMETERS

    def scan_time_points(self, scan_density=200):
        """Scan across the chart to discover all time points with data"""
        print("Scanning chart to discover available time points...")

        # Set parameters for scanning
        left_buffer = self.chart_x + (self.chart_width * 0.01)  # 1% buffer from left edge
        right_edge = self.chart_x + self.chart_width - (self.chart_width * 0.01)
        middle_y = self.chart_y + (self.chart_height / 2)

        # Create a high-resolution scan for precise results
        positions = np.linspace(left_buffer, right_edge, scan_density)

        found_time_points = {}  # Dictionary to store {time_str: (x_position, extracted_info, tooltip_text)}

        # Reset mouse position before starting
        actions = ActionChains(self.driver)
        actions.move_to_element(self.chart_container).move_by_offset(-100, -100).perform()
        time.sleep(1)

        # Track the last successful time to avoid duplicates with minimal differences
        last_time = None
        last_position = None
        min_position_diff = self.chart_width / (scan_density * 0.5)  # Minimum distance between captured points

        for idx, x_pos in enumerate(positions):
            try:
                # Move to the position
                js_script = f"""
                var evt = new MouseEvent('mousemove', {{
                    'view': window,
                    'bubbles': true,
                    'cancelable': true,
                    'clientX': {int(x_pos)},
                    'clientY': {int(middle_y)}
                }});
                document.elementFromPoint({int(x_pos)}, {int(middle_y)}).dispatchEvent(evt);
                """
                self.driver.execute_script(js_script)
                time.sleep(0.1)  # Small delay to let tooltip appear

                # Extract tooltip data
                extracted_info, full_tooltip_text = self.extract_tooltip_data(self.target_parameters)

                if extracted_info and "Time" in extracted_info and extracted_info["Time"]:
                    time_value = extracted_info["Time"]

                    # Check if this time point is significantly different from the last one
                    is_new_point = True
                    if last_time == time_value and last_position is not None:
                        if abs(x_pos - last_position) < min_position_diff:
                            is_new_point = False

                    if is_new_point:
                        found_time_points[time_value] = (x_pos, extracted_info, full_tooltip_text)
                        last_time = time_value
                        last_position = x_pos

                        # Take a screenshot periodically (every 10th point)
                        if idx % 10 == 0:
                            screenshot_path = f"chart_screenshots/time_point_{time_value.replace(':', '_')}.png"
                            self.driver.save_screenshot(screenshot_path)

                        print(f"Found time point {time_value} at position {x_pos:.1f}")
            except Exception as e:
                print(f"Error at position {x_pos:.1f}: {str(e)}")

        print(f"Scan complete. Found {len(found_time_points)} unique time points: {sorted(found_time_points.keys())}")
        return found_time_points

    def extract_time_series_data(self):
        """Main function to extract time series data from the chart"""
        print(f"Extracting time series data for {CHART_TYPE} chart...")

        # Find and set up the chart
        if not self.find_and_setup_chart():
            return None

        # Detect available parameters or use default ones
        self.target_parameters = self.detect_available_parameters()
        print(f"Using parameters: {self.target_parameters}")

        # Reset mouse position before starting
        actions = ActionChains(self.driver)
        actions.move_to_element(self.chart_container).move_by_offset(-100, -100).perform()
        time.sleep(1)

        # Scan the chart to discover available time points
        # Higher density (300+) gives more precise results but takes longer
        time_points = self.scan_time_points(scan_density=300)

        if not time_points:
            print("No time points were found in the chart.")
            return None

        # Clear previously collected data
        self.collected_data = []

        # Extract data for each available time point
        count = 0
        for time_str, (x_pos, extracted_info, full_tooltip_text) in time_points.items():
            count += 1
            # Only take a screenshot for every 5th point to avoid too many images
            if count % 5 == 0:
                # Move to the position to ensure tooltip is visible
                js_script = f"""
                var evt = new MouseEvent('mousemove', {{
                    'view': window,
                    'bubbles': true,
                    'cancelable': true,
                    'clientX': {int(x_pos)},
                    'clientY': {int(self.chart_y + (self.chart_height / 2))}
                }});
                document.elementFromPoint({int(x_pos)}, {int(self.chart_y + (self.chart_height / 2))}).dispatchEvent(evt);
                """
                self.driver.execute_script(js_script)
                time.sleep(0.2)

                # Take a screenshot for verification
                safe_time = time_str.replace(':', '_')
                screenshot_path = f"chart_screenshots/data_time_{safe_time}.png"
                self.driver.save_screenshot(screenshot_path)

            # Initialize time point data
            point_data = {
                'Time': time_str
            }

            # Add data for each parameter (excluding Time which we formatted above)
            for param in self.target_parameters:
                if param != "Time":
                    point_data[param] = self.clean_numeric_value(extracted_info.get(param, ''))

            self.collected_data.append(point_data)

        # Save to Excel
        return self.save_to_excel()

    def clean_numeric_value(self, value):
        """Extract numeric value from string like '29.88 A'"""
        if value and isinstance(value, str):
            numeric_match = re.search(r'([\d.]+)', value)
            if numeric_match:
                return numeric_match.group(1)
        return value

    def save_to_excel(self):
        """Save collected data to Excel file"""
        print("Saving data to Excel...")

        if not self.collected_data:
            print("No data to save.")
            return None

        # Create a pandas DataFrame
        df = pd.DataFrame(self.collected_data)

        # Ensure we have the 'Time' column
        if 'Time' not in df.columns:
            print("Warning: 'Time' column not found in collected data.")
            return None

        # Convert numeric columns
        for col in df.columns:
            if col != 'Time':
                try:
                    df[col] = pd.to_numeric(df[col])
                except (ValueError, TypeError):
                    pass  # Keep as is if conversion fails

        # Sort by time
        def time_to_minutes(time_str):
            try:
                if ':' in time_str:
                    hours, minutes = time_str.split(':')
                    return int(hours) * 60 + int(minutes)
                return 0
            except:
                return 0

        df['_time_sort'] = df['Time'].apply(time_to_minutes)
        df = df.sort_values('_time_sort')
        df = df.drop('_time_sort', axis=1)

        # Save to Excel
        date_str = time.strftime("%Y%m%d")
        filename = f"extracted_data/time_series_{CHART_TYPE}_data_{date_str}.xlsx"
        df.to_excel(filename, index=False)
        print(f"Data saved to {filename}")

        return filename


# Example usage:
# extractor = TimeChartDataExtractor(driver)
# excel_file = extractor.extract_time_series_data()


    # Main execution
# Main execution
if __name__ == "__main__":
    try:
        # Initialize the time chart data extractor
        extractor = TimeChartDataExtractor(driver)

        # Extract the time series data
        excel_file = extractor.extract_time_series_data()

        if excel_file:
            print(f"✅ Chart data extraction completed successfully!")
            print(f"📊 Data saved to: {excel_file}")
        else:
            print("❌ Failed to extract chart data.")

    except Exception as e:
        print(f"❌ Error in main execution: {str(e)}")
    finally:
        # Cleanup
        print("🔄 Closing browser...")
        driver.quit()
        # driver.quit()


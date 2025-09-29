"""
NY Courts WebCivil Scraper - Manual Mode
This version works by opening a browser and letting you manually navigate,
then automatically extracts the data once you're on the results page.
"""
import time
import logging
import pandas as pd
import random
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
# selenium 3
from selenium import webdriver
from webdriver_manager.microsoft import EdgeChromiumDriverManager

driver = webdriver.Edge(EdgeChromiumDriverManager().install())
import threading
import tkinter as tk
from tkinter import messagebox

# CONFIG
INPUT_XLSX = r"F:\download data\py\webcivil_mockup.xlsx"
INPUT_SHEET = "I. Input Sheet"
OUTPUT_SHEET = "II. Output Sheet"
BASE_URL = "https://iapps.courts.state.ny.us/webcivilLocal/LCMain"
HEADLESS = False  # Must be False for manual mode

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.FileHandler("scraper.log"), logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

OUTPUT_COLS = [
    "IndexNumber", "AppearanceDate", "FirstPlaintiffFirm", "CaseName", "Time", "Purpose",
    "OutcomeType", "Judge", "Part", "Classification", "FilingDate", "DispositionDate",
    "AppearanceDate_1", "Time_1", "Purpose_1", "OutcomeType_1", "Judge_1", "Part_1",
    "MotSeq_1", "AppearanceDate_2", "Time_2", "Purpose_2", "OutcomeType_2", "Judge_2", "Part_2"
]

class ManualScraper:
    def __init__(self):
        self.driver = None
        self.current_index = None
        self.results = []
        self.root = None

    def init_driver(self):
        """Initialize Chrome driver with stealth settings"""
        chrome_options = webdriver.ChromeOptions()

        # Add realistic user agent and settings
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)

        # Optional: Add these if still getting detected
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--no-sandbox")

        self.driver = webdriver.Chrome(
            service=ChromeService(ChromeDriverManager().install()),
            options=chrome_options
        )

        # Execute anti-detection scripts
        self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        self.driver.execute_cdp_cmd('Network.setUserAgentOverride', {
            "userAgent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        })

        self.driver.set_page_load_timeout(30)
        self.driver.implicitly_wait(3)

    def create_gui(self):
        """Create a simple GUI for manual control"""
        self.root = tk.Tk()
        self.root.title("NY Courts Scraper - Manual Mode")
        self.root.geometry("500x400")

        # Instructions
        instructions = tk.Text(self.root, height=15, wrap=tk.WORD)
        instructions.insert(tk.END, """
INSTRUCTIONS:

1. The browser window will open to the NY Courts website
2. For each index number, you need to:
   - Solve any CAPTCHAs manually
   - Navigate to Index Search if needed
   - Enter the index number: {INDEX_NUMBER}
   - Click Find Cases/Search
   - Wait for results to load
   - Click 'Data Ready' button below

3. The script will automatically extract data from the current page
4. Repeat for all index numbers

Current Status: Ready to start
        """)
        instructions.config(state=tk.DISABLED)
        instructions.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        # Status label
        self.status_label = tk.Label(self.root, text="Ready to start", fg="blue")
        self.status_label.pack(pady=5)

        # Buttons
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=10)

        self.start_btn = tk.Button(button_frame, text="Start Scraping", command=self.start_scraping, bg="green", fg="white")
        self.start_btn.pack(side=tk.LEFT, padx=5)

        self.ready_btn = tk.Button(button_frame, text="Data Ready", command=self.extract_current_data, bg="blue", fg="white", state=tk.DISABLED)
        self.ready_btn.pack(side=tk.LEFT, padx=5)

        self.skip_btn = tk.Button(button_frame, text="Skip Current", command=self.skip_current, bg="orange", fg="white", state=tk.DISABLED)
        self.skip_btn.pack(side=tk.LEFT, padx=5)

        self.stop_btn = tk.Button(button_frame, text="Stop & Save", command=self.stop_scraping, bg="red", fg="white", state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)

    def start_scraping(self):
        """Start the scraping process"""
        try:
            # Read index numbers
            df = pd.read_excel(INPUT_XLSX, sheet_name=INPUT_SHEET)
            possible_cols = ["IndexNumber", "Index #:", "Index Number", "Index", "Case Number"]
            col = None
            for possible_col in possible_cols:
                if possible_col in df.columns:
                    col = possible_col
                    break

            if col is None:
                col = df.columns[0]

            self.index_numbers = df[col].dropna().astype(str).str.strip().tolist()
            self.current_idx = 0

            if not self.index_numbers:
                messagebox.showerror("Error", "No index numbers found in Excel file!")
                return

            # Initialize driver
            self.init_driver()

            # Open the website
            self.driver.get(BASE_URL)

            # Update GUI
            self.start_btn.config(state=tk.DISABLED)
            self.ready_btn.config(state=tk.NORMAL)
            self.skip_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.NORMAL)

            self.show_current_index()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to start: {e}")
            logger.error(f"Start error: {e}")

    def show_current_index(self):
        """Show current index number to search"""
        if self.current_idx < len(self.index_numbers):
            current = self.index_numbers[self.current_idx]
            self.current_index = current
            self.status_label.config(
                text=f"Search for: {current} ({self.current_idx + 1}/{len(self.index_numbers)})",
                fg="blue"
            )

            # Also set clipboard for easy copying
            try:
                self.root.clipboard_clear()
                self.root.clipboard_append(current)
                logger.info(f"Index number {current} copied to clipboard")
            except:
                pass
        else:
            self.finish_scraping()

    def extract_current_data(self):
        """Extract data from current page"""
        if not self.driver or not self.current_index:
            return

        try:
            self.status_label.config(text="Extracting data...", fg="orange")
            self.root.update()

            # Extract data using multiple strategies
            result = self.scrape_page_data(self.current_index)
            self.results.append(result)

            logger.info(f"Extracted data for {self.current_index}")

            # Move to next
            self.current_idx += 1

            if self.current_idx < len(self.index_numbers):
                self.show_current_index()
            else:
                self.finish_scraping()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to extract data: {e}")
            logger.error(f"Extract error: {e}")

    def skip_current(self):
        """Skip current index number"""
        if self.current_index:
            # Add empty record
            empty_result = dict.fromkeys(OUTPUT_COLS, None)
            empty_result["IndexNumber"] = self.current_index
            self.results.append(empty_result)

            logger.info(f"Skipped {self.current_index}")

            # Move to next
            self.current_idx += 1

            if self.current_idx < len(self.index_numbers):
                self.show_current_index()
            else:
                self.finish_scraping()

    def scrape_page_data(self, index_number):
        """Extract data from current page"""
        result = dict.fromkeys(OUTPUT_COLS, None)
        result["IndexNumber"] = index_number

        try:
            # Wait a moment for page to fully load
            time.sleep(1)

            # Strategy 1: Look for tables
            tables = self.driver.find_elements(By.TAG_NAME, "table")
            for table in tables:
                try:
                    self.extract_table_data(table, result)
                except:
                    continue

            # Strategy 2: Look for specific text patterns
            try:
                page_text = self.driver.find_element(By.TAG_NAME, "body").text
                self.extract_text_patterns(page_text, result)
            except:
                pass

            # Strategy 3: Look for form fields and divs
            try:
                self.extract_form_data(result)
            except:
                pass

            return result

        except Exception as e:
            logger.error(f"Error extracting data for {index_number}: {e}")
            return result

    def extract_table_data(self, table, result):
        """Extract data from HTML tables"""
        rows = table.find_elements(By.TAG_NAME, "tr")

        for row in rows:
            cells = row.find_elements(By.TAG_NAME, "td")
            if len(cells) < 2:
                continue

            for i, cell in enumerate(cells):
                try:
                    text = cell.text.strip()
                    if not text:
                        continue

                    # Check for links (often contain important data)
                    links = cell.find_elements(By.TAG_NAME, "a")
                    if links:
                        link_text = links[0].text.strip()
                        if link_text:
                            text = link_text

                    # Pattern matching
                    if self.is_date_pattern(text):
                        if not result["AppearanceDate"]:
                            result["AppearanceDate"] = text
                        elif not result["FilingDate"]:
                            result["FilingDate"] = text
                    elif self.is_judge_pattern(text):
                        if not result["Judge"]:
                            result["Judge"] = text
                    elif self.is_firm_pattern(text):
                        if not result["FirstPlaintiffFirm"]:
                            result["FirstPlaintiffFirm"] = text
                    elif self.is_time_pattern(text):
                        if not result["Time"]:
                            result["Time"] = text

                except:
                    continue

    def extract_text_patterns(self, page_text, result):
        """Extract data using text pattern matching"""
        import re

        # Common patterns
        patterns = {
            "Classification": [r"Classification[:\s]+([^\n]+)", r"Case Type[:\s]+([^\n]+)"],
            "FilingDate": [r"Filing Date[:\s]+([^\n]+)", r"Filed[:\s]+([^\n]+)"],
            "DispositionDate": [r"Disposition[:\s]+([^\n]+)", r"Disposed[:\s]+([^\n]+)"],
            "CaseName": [r"Case Name[:\s]+([^\n]+)", r"Caption[:\s]+([^\n]+)"],
        }

        for field, field_patterns in patterns.items():
            if result[field]:
                continue

            for pattern in field_patterns:
                match = re.search(pattern, page_text, re.IGNORECASE)
                if match:
                    result[field] = match.group(1).strip()
                    break

    def extract_form_data(self, result):
        """Extract data from form elements and structured divs"""
        # Look for common field IDs and classes
        field_mappings = {
            "CaseName": ["caseName", "case-name", "caption"],
            "Judge": ["judge", "hon", "honorable"],
            "Part": ["part", "room", "courtroom"],
            "Classification": ["classification", "case-type", "type"],
        }

        for field, selectors in field_mappings.items():
            if result[field]:
                continue

            for selector in selectors:
                try:
                    # Try ID
                    element = self.driver.find_element(By.ID, selector)
                    text = element.text.strip() or element.get_attribute('value')
                    if text:
                        result[field] = text
                        break
                except:
                    try:
                        # Try class name
                        element = self.driver.find_element(By.CLASS_NAME, selector)
                        text = element.text.strip()
                        if text:
                            result[field] = text
                            break
                    except:
                        continue

    def is_date_pattern(self, text):
        """Check if text looks like a date"""
        import re
        date_patterns = [
            r'\d{1,2}/\d{1,2}/\d{2,4}',
            r'\d{1,2}-\d{1,2}-\d{2,4}',
            r'\d{4}-\d{1,2}-\d{1,2}'
        ]
        return any(re.match(pattern, text) for pattern in date_patterns)

    def is_judge_pattern(self, text):
        """Check if text looks like a judge name"""
        return "hon." in text.lower() or "judge" in text.lower()

    def is_firm_pattern(self, text):
        """Check if text looks like a law firm"""
        firm_indicators = ["law", "attorney", "esq", "llc", "pllc", "pc", "&"]
        return any(indicator in text.lower() for indicator in firm_indicators)

    def is_time_pattern(self, text):
        """Check if text looks like a time"""
        import re
        return bool(re.match(r'\d{1,2}:\d{2}', text))

    def finish_scraping(self):
        """Finish scraping and save results"""
        try:
            self.status_label.config(text="Saving results...", fg="green")
            self.root.update()

            # Save results
            if self.results:
                df = pd.DataFrame(self.results)
                if Path(INPUT_XLSX).exists():
                    with pd.ExcelWriter(INPUT_XLSX, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                        df.to_excel(writer, sheet_name=OUTPUT_SHEET, index=False)
                else:
                    with pd.ExcelWriter(INPUT_XLSX, engine="openpyxl") as writer:
                        df.to_excel(writer, sheet_name=OUTPUT_SHEET, index=False)

                logger.info(f"Saved {len(self.results)} records to {INPUT_XLSX}")
                messagebox.showinfo("Success", f"Scraping completed! Saved {len(self.results)} records.")
            else:
                messagebox.showinfo("Info", "No data was collected.")

            self.status_label.config(text="Completed!", fg="green")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to save results: {e}")
            logger.error(f"Save error: {e}")
        finally:
            self.stop_scraping()

    def stop_scraping(self):
        """Stop scraping and cleanup"""
        try:
            if self.driver:
                self.driver.quit()
        except:
            pass

        # Reset GUI
        self.start_btn.config(state=tk.NORMAL)
        self.ready_btn.config(state=tk.DISABLED)
        self.skip_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.DISABLED)
        self.status_label.config(text="Stopped", fg="red")

    def run(self):
        """Run the GUI"""
        self.create_gui()
        self.root.protocol("WM_DELETE_WINDOW", self.stop_scraping)
        self.root.mainloop()

def main():
    """Main function"""
    try:
        scraper = ManualScraper()
        scraper.run()
    except Exception as e:
        logger.error(f"Main error: {e}")
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
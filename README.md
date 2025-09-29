README: Web Civil Scraper Automation
This document provides instructions on how to set up and run the web scraping script to extract case data from the New York Courts website.

1. Initial Setup
Create the Folder: On your E: drive, create a folder with the following path:
E:\GrapeTask\Web_Civil\Final

Place the Files: Ensure all four project files (cases_input_output.xlsm, chrome_debug.bat, scraper_final.py, 
and cases_input_output.xlsx) are placed inside this newly created folder.

Launch Chrome for Debugging: Right-click on the chrome_debug.bat file (Run as administrator) to open a special instance of the Chrome browser. 
This browser is required for the script to function correctly.

One-Time Captcha/Turnstile Solve: When the browser opens, go to the following 
URL: https://iapps.courts.state.ny.us/webcivilLocal/LCSearch?param=I. 
If you see a "Turnstile" or a "Captcha" security check, please solve it manually. This step is only required once. 
After solving it, don't close browser

2. Running the Scraper
Open the Excel File: Open the cases_input_output.xlsm file. This file contains the list of Index Numbers that the script will process.


Run the Script: On the Excel sheet, locate and click the "Run Scraper" button.

Wait for Confirmation: Once you click the button, a message box will appear with the text:
"Python script is Running Successfully. Please Check cases_input_output.xlsx file for Data.

3. How the Script Works
Input Data: The script reads the list of Index Numbers from the IndexNumbers sheet in the cases_input_output.xlsm file.

Web Scraping: It connects to the web browser you launched earlier, navigates to the website, and automatically 
fills in the search forms for each Index Number.

Output: All the scraped data is saved to the cases_input_output.xlsx file.

"Not Found" Handling: If the script is unable to find data for a specific Index Number, 
it will write "Not Found" in the Case Status column and immediately move on to the next one, ensuring the process is efficient and doesn't get stuck.

# Web-scraping

# CMMI Institute PARS Data Scraper

A Python script that automates data extraction from the CMMI Institute's Published Appraisal Results System (PARS) using Selenium and BeautifulSoup.

## Features

- Extracts comprehensive appraisal data including:
  - Organization names and IDs
  - Appraisal team leaders
  - Sponsors and partners
  - Organizational unit (OU) names
  - Appraisal validity dates
  - Model views/domains and maturity levels
- Handles pagination to scrape all available results
- Saves data to a structured Excel file
- Filters results by country (default: United States)

## Prerequisites

- Python 3.8+
- Chrome browser installed
- ChromeDriver matching your Chrome version

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/cmmi-pars-scraper.git
   cd cmmi-pars-scraper
Install required packages:

bash
pip install selenium beautifulsoup4 pandas openpyxl
Download ChromeDriver from https://chromedriver.chromium.org/ and place it in your project directory or system PATH.

## Usage
Update the ChromeDriver path in the script if needed:

s = Service(r"C:\chromedriver.exe")  # Update this path
Run the script:

bash
python webscraping.py

The script will:

Launch a Chrome browser window
Navigate to the CMMI PARS site
Apply the country filter
Scrape all available pages
Save results to cmmi_data.xlsx

## ustomization Options
Change the target country by modifying:


if(dropdownbox[i].text == "United States"):
Adjust wait times if experiencing timeouts:

time.sleep(2)  # Increase if needed
Modify extracted fields by editing the BeautifulSoup selectors

Output
The script generates an Excel file (cmmi_data.xlsx) with columns for all extracted fields:

Organization Name	ID	Appraisal Team Leader	Sponsors	...
ABC Corporation	12345	John Smith	XYZ Inc.	...

## Troubleshooting
ChromeDriver issues: Ensure the ChromeDriver version matches your installed Chrome version
Element not found: The website structure may have changed - update the selectors
Slow performance: Increase sleep intervals if running on slow connections

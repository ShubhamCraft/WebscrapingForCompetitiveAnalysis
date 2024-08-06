# WebscrapingForCompetitiveAnalysis


## Overview

This script automates the process of scraping host information from a list of URLs and saving the data into an Excel file. The script uses Selenium for web scraping and Pandas for data manipulation and storage.

## Features

1. **Web Scraping:** The script uses Selenium to automate the process of visiting each host URL and extracting the necessary information.
2. **Data Storage:** The scraped data is saved incrementally into an Excel file using Pandas and Openpyxl.
3. **Error Handling:** The script includes error handling for each data extraction step to ensure robustness and continuity in case of failures.

## Detailed Workflow

1. **Setup WebDriver:**
    - Initializes a Chrome WebDriver instance for web automation.

2. **Read Host Links:**
    - Reads a CSV file (`D:/host_info.csv`) containing host URLs.

3. **File Handling:**
    - Checks if the output Excel file (`D:/scraped_host_info.xlsx`) exists.
    - Opens the existing workbook or creates a new one.

4. **Scraping Process:**
    - Iterates through each host URL.
    - Extracts host details such as name, description, total reviews, rating, years hosting, and listing count.
    - Handles potential errors during the extraction of each piece of information.

5. **Data Writing:**
    - Uses `ExcelWriter` to append the scraped data to the Excel file without overwriting existing data.
    - Writes the data incrementally to ensure efficient memory usage.

6. **Browser Management:**
    - Closes the browser after the scraping is completed.

## Columns Scraped

- `hostlink`: The URL of the host's page.
- `name`: The name of the host.
- `description`: The description of the host.
- `total_reviews`: The total number of reviews the host has received.
- `rating`: The rating of the host.
- `years_hosting`: The number of years the host has been active.
- `listing_count`: The number of listings the host has.

## Usage

- Ensure the Chrome WebDriver is installed and accessible.
- Place the host URLs in a CSV file (`D:/host_info.csv`) with a column named `hostlink`.
- Run the script to start the scraping process.
- The scraped data will be saved in `D:/scraped_host_info.xlsx`.

# Web-Product-Data-Scraping-Tool

This Python script provides a flexible framework for extracting product data from various manufacturer websites based on a list of identifiers (e.g., MPNs, product codes) in an Excel file. It supports configuration for different website structures using either **XPath** or **CSS Selectors** for locating data elements.

**⚠️ Work in Progress / Active Development ⚠️**

Please note that while the current version of this script is functional, it is still under active development and modification. The goal is to continuously improve its capabilities, robustness, and adaptability to a wider range of websites. Expect new features, enhanced error handling, and optimizations in future updates. Your feedback and suggestions are welcome!

## Overview

The script reads product identifiers from an Excel file and then uses user-defined configurations to scrape product information from the corresponding websites. The key to its generality lies in the user's ability to specify how to find different data points (product links, family information, images) on each target website using either XPath or CSS selectors.

The process involves:

1.  **Reading Identifiers:** Taking a column of product identifiers from an Excel file.
2.  **Constructing Search URLs:** Using a user-provided format to create search URLs for each identifier on the target website.
3.  **Navigating and Extracting:**
    * Finding the link to the specific product page from the search results using a chosen selector (XPath or CSS).
    * Navigating to the product page.
    * Extracting relevant data (family hierarchy, image links) using specified selectors.
4.  **Saving Results:** Writing the extracted data back to the Excel file in new columns.

The GUI allows users to configure the script for different websites by specifying the search URL format, the type of selector (XPath or CSS), and the selector strings for each data point they want to extract.

## Features

* **Functional Core:** The script currently works for extracting product links, family information, and image links based on user-provided configurations.
* **Generalized Scraping:** Designed to work with diverse website structures through user-defined configurations.
* **XPath and CSS Selector Support:** Offers flexibility in targeting HTML elements.
* **Configurable Extraction:** Users define how to find product links, family data, and images using selectors.
* **Batch Processing:** Handles multiple product identifiers from an Excel file.
* **Clear GUI:** Provides an interface for easy configuration and operation.
* **Progress Tracking:** Shows the status and progress of the scraping process.
* **Basic Error Handling:** Includes mechanisms to catch common errors during web requests and data extraction.

**Ongoing Development / Planned Improvements:**

* **Enhanced Dynamic Content Handling:** Exploring and implementing methods to better scrape websites that heavily rely on JavaScript (e.g., integration with Selenium or Puppeteer).
* **Improved Pagination Handling:** Developing logic to automatically navigate and extract data from multi-page search results and product listings.
* **More Robust Error Management:** Implementing more comprehensive error detection, logging, and potential recovery mechanisms.
* **Polite Scraping Features:** Adding built-in delays and options for respecting `robots.txt` more effectively.
* **Advanced Configuration Options:** Potentially adding features for handling cookies, headers, and more complex website interactions.
* **Configuration Management:** Exploring ways to save and load configurations for different target websites.
* **Data Validation and Cleaning:** Implementing optional steps to validate and clean the extracted data.
* **More Detailed Documentation and Examples:** Expanding the documentation with more use cases and examples for various website structures.

## Requirements

* Python 3.x
* The following Python libraries:
    * `requests` (`pip install requests`)
    * `beautifulsoup4` (`pip install beautifulsoup4`)
    * `lxml` (`pip install lxml`)
    * `pandas` (`pip install pandas`)
    * `tkinter` (usually included with Python)

## Usage

1.  **Save the code:** Save the Python code as a `.py` file (e.g., `generalized_scraper.py`).
2.  **Run the script:** Execute from your terminal: `python generalized_scraper.py`
3.  **GUI Interaction:**
    * **Select Excel File:** Browse and select the Excel file containing product identifiers.
    * **Manufacturer Configuration:**
        * **Identifier Column Name:** Enter the name of the column with product identifiers.
        * **Search URL Format:** Provide the URL template for searching. Use `{mpn}` (or a similar placeholder) for the identifier.
        * **Product Link Selector:** Enter either an XPath expression or a CSS selector to find the product link. Choose the corresponding radio button ("XPath" or "CSS Selector").
        * **Product Link Base URL (optional):** If product links are relative, provide the base URL.
        * **Family Hierarchy Selector (optional):** Enter the XPath or CSS selector for the family/category path. Choose the selector type.
        * **Image Links Selector (optional):** Enter the XPath or CSS selector to find image elements (e.g., `//img/@src` for XPath or `img` for CSS). Choose the selector type.
        * **Output Column Prefix (optional):** Set a prefix for the output columns (default: "Product").
    * **Start Data Extraction:** Click to begin the scraping process.
    * **Progress:** Monitor the status and progress bar.
    * **Save Output:** Save the modified Excel file when prompted.

## Input Excel File Format

The Excel file should have a column containing unique product identifiers. The name of this column should be specified in the GUI.

## Output Excel File Format

The output Excel file will contain the original data along with new columns:

* **{Output Prefix} Link:** The extracted product page URL.
* **Family:** The extracted family hierarchy (if configured).
* **Image Link 1**, **Image Link 2**, ...: URLs of the extracted product images (if configured).

## Alternatives and Considerations for Website Dynamics

To adapt to different website dynamics, consider the following alternatives and techniques:

1.  **Selector Flexibility (XPath vs. CSS):**
    * **XPath:** More powerful for navigating complex HTML structures and traversing up and down the DOM tree. Useful when elements are identified by attributes or relationships.
    * **CSS Selectors:** Often simpler and more readable for selecting elements based on tags, classes, and IDs. Can be more straightforward for basic selections. The script currently supports choosing between these.

2.  **Handling Dynamic Content (JavaScript Rendering):**
    * **Selenium:** A powerful library that can automate web browsers. It can render JavaScript-heavy websites, allowing you to scrape content that is loaded dynamically. This is a significant area for future development. (`pip install selenium`)
    * **Puppeteer (via `pyppeteer`):** Another browser automation library, primarily for Chromium. Similar to Selenium, it can execute JavaScript and scrape dynamically loaded content. **(Potential Future Enhancement)**

3.  **API Usage (If Available):**
    * If the manufacturer provides a public API, it is often a more reliable and efficient way to retrieve product data compared to scraping the HTML. Check the website's developer section or look for API documentation.

4.  **Regular Expressions (for specific patterns):**
    * The `re` library in Python can be used to extract data based on specific text patterns within the HTML. This can be useful for extracting information that is not easily targeted with XPath or CSS selectors.

5.  **Handling Pagination:**
    * Many websites display search results or product listings across multiple pages. Logic to handle pagination will be a key area of future development.

6.  **Dealing with AJAX Requests:**
    * Websites might load data asynchronously using AJAX. Techniques to identify and potentially simulate these requests may be explored in future versions.

7.  **Robots.txt and Terms of Service:**
    * Always check the `robots.txt` file of the website (e.g., `https://www.example-manufacturer.com/robots.txt`) to understand which parts of the site are disallowed for scraping. Respect the website's terms of service.

8.  **Rate Limiting and Polite Scraping:**
    * Implementing more sophisticated rate limiting and polite scraping strategies is planned for future updates.

9.  **Error Handling and Retries:**
    * The current error handling is basic. Future development will focus on making it more robust and providing informative feedback.

10. **User-Agent Rotation and Proxy Usage:**
    * Exploring the integration of user-agent rotation and proxy support to enhance anonymity and avoid potential blocking is a possible future direction.

## Disclaimer

This script is a generalized tool for web scraping and should be used responsibly and ethically. The author is not responsible for any misuse or violation of website terms of service. Always ensure you have the legal right to scrape data from a website before doing so. As this project is under active development, functionality and features may evolve over time.

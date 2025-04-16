import requests
from bs4 import BeautifulSoup
from lxml import etree  # For XPath evaluation
import pandas as pd  # For reading Excel files
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import time
from urllib.parse import urlparse, urljoin
import re  # For regular expressions (CSS selector fallback)

# Global flag to track if a file has been selected
file_selected = False
excel_file_path_global = ""

def extract_link(html_content, method='xpath', selector=None, base_url=None, url=None, progress_callback=None):
    """
    Extracts a link from HTML content using either XPath or CSS selector.

    Args:
        html_content (bytes): The HTML content of the webpage.
        method (str, optional): The extraction method ('xpath' or 'css'). Defaults to 'xpath'.
        selector (str, optional): The XPath or CSS selector to use. Defaults to None.
        base_url (str, optional): The base URL to prepend to relative links. Defaults to None.
        url (str, optional): The URL of the current page (for resolving relative URLs). Defaults to None.
        progress_callback (callable, optional): Function to call with progress updates.

    Returns:
        str: The extracted link, or None if not found.
    """
    if not selector:
        return None

    try:
        if method == 'xpath':
            tree = etree.HTML(html_content)
            link_elements = tree.xpath(selector)
            if link_elements:
                href = link_elements[0].get('href')
                if href:
                    return urljoin(base_url if base_url else urlparse(url).scheme + "://" + urlparse(url).netloc, href)
            elif progress_callback:
                progress_callback(f"Link not found with XPath: {selector}", 90)
        elif method == 'css':
            soup = BeautifulSoup(html_content, 'html.parser')
            link_element = soup.select_one(selector)
            if link_element and link_element.has_attr('href'):
                href = link_element['href']
                return urljoin(base_url if base_url else urlparse(url).scheme + "://" + urlparse(url).netloc, href)
            elif progress_callback:
                progress_callback(f"Link not found with CSS selector: {selector}", 90)
        else:
            if progress_callback:
                progress_callback(f"Invalid extraction method: {method}", 100)
            return None
    except (etree.XPathEvalError, Exception) as e:
        if progress_callback:
            progress_callback(f"Error during link extraction ({method}): {e}", 100)
        return None
    return None

def extract_data(html_content, method='xpath', selector=None):
    """
    Extracts text data from HTML content using either XPath or CSS selector.

    Args:
        html_content (bytes): The HTML content.
        method (str, optional): 'xpath' or 'css'. Defaults to 'xpath'.
        selector (str, optional): The XPath or CSS selector. Defaults to None.

    Returns:
        list: A list of extracted text strings.
    """
    if not selector:
        return []
    try:
        if method == 'xpath':
            tree = etree.HTML(html_content)
            elements = tree.xpath(selector)
            return [element.text.strip() for element in elements if element.text is not None]
        elif method == 'css':
            soup = BeautifulSoup(html_content, 'html.parser')
            elements = soup.select(selector)
            return [element.get_text(strip=True) for element in elements]
        else:
            return []
    except (etree.XPathEvalError, Exception) as e:
        print(f"Error during data extraction ({method}): {e}")
        return []

def extract_image_links(html_content, method='xpath', selector=None, base_url=None, url=None):
    """
    Extracts image source URLs using either XPath or CSS selector.

    Args:
        html_content (bytes): The HTML content.
        method (str, optional): 'xpath' or 'css'. Defaults to 'xpath'.
        selector (str, optional): The XPath or CSS selector. Defaults to None.
        base_url (str, optional): The base URL for resolving relative paths. Defaults to None.
        url (str, optional): The URL of the current page. Defaults to None.

    Returns:
        list: A list of image URLs.
    """
    if not selector:
        return []
    try:
        if method == 'xpath':
            tree = etree.HTML(html_content)
            img_elements = tree.xpath(selector)
            return [urljoin(base_url if base_url else urlparse(url).scheme + "://" + urlparse(url).netloc, img) for img in img_elements if img]
        elif method == 'css':
            soup = BeautifulSoup(html_content, 'html.parser')
            img_elements = soup.select(selector)
            return [urljoin(base_url if base_url else urlparse(url).scheme + "://" + urlparse(url).netloc, img['src']) for img in img_elements if 'src' in img.attrs]
        else:
            return []
    except (etree.XPathEvalError, Exception) as e:
        print(f"Error during image link extraction ({method}): {e}")
        return []

def get_html_content(url, progress_callback=None):
    """Fetches HTML content with potential retries."""
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'}
    max_retries = 3
    retry_delay = 5  # seconds

    for attempt in range(max_retries):
        try:
            if progress_callback:
                progress_callback(f"Fetching HTML from: {url} (Attempt {attempt + 1}/{max_retries})", 20)
            response = requests.get(url, headers=headers, verify=False, timeout=15)
            response.raise_for_status()
            if progress_callback:
                progress_callback("HTML fetched successfully.", 40)
            return response.content
        except requests.exceptions.RequestException as e:
            error_message = f"Error fetching HTML from {url} (Attempt {attempt + 1}/{max_retries}): {e}"
            if attempt < max_retries - 1:
                print(f"{error_message}. Retrying in {retry_delay} seconds...")
                time.sleep(retry_delay)
            else:
                if progress_callback:
                    progress_callback(error_message, 100)
                print(error_message)
                return None
        except Exception as e:
            error_message = f"An error occurred while getting HTML from {url}: {e}"
            if progress_callback:
                progress_callback(error_message, 100)
            print(error_message)
            return None
    return None

def process_manufacturer(excel_file, config, progress_var, progress_percent_label, status_label, root):
    try:
        df = pd.read_excel(excel_file)
        mpn_column = config.get('mpn_column')
        search_url_format = config.get('search_url_format')
        product_link_selector = config.get('product_link_selector')
        product_link_selector_type = config.get('product_link_selector_type', 'xpath')
        product_link_base_url = config.get('product_link_base_url')
        family_selector = config.get('family_selector')
        family_selector_type = config.get('family_selector_type', 'xpath')
        image_selector = config.get('image_selector')
        image_selector_type = config.get('image_selector_type', 'xpath')
        output_prefix = config.get('output_prefix', 'Product')

        if not all([mpn_column, search_url_format, product_link_selector]):
            error_message = "Error: Configuration is missing required fields (mpn_column, search_url_format, product_link_selector)."
            messagebox.showerror("Error", error_message)
            status_label.config(text=error_message, foreground="red")
            return

        if mpn_column not in df.columns:
            error_message = f"Error: MPN Column '{mpn_column}' not found in the Excel file."
            messagebox.showerror("Error", error_message)
            status_label.config(text=error_message, foreground="red")
            return

        total_items = len(df[mpn_column])
        extracted_data = {
            f'{output_prefix} Link': [None] * total_items,
            'Family': [None] * total_items,
            'Image Links': [[]] * total_items
        }

        for index, identifier in enumerate(df[mpn_column]):
            try:
                search_url = search_url_format.format(mpn=identifier) # Assuming 'mpn' is the generic identifier key
                progress_text = f"Processing Identifier: {identifier} ({index + 1}/{total_items})"
                status_label.config(text=progress_text, foreground="blue")
                progress_percent = int(((index + 1) / total_items) * 20)
                progress_var.set(progress_percent)
                progress_percent_label.config(text=f"{progress_percent}%")
                root.update()

                html_search_content = get_html_content(search_url,
                                                       progress_callback=lambda msg, p: status_label.config(text=f"Processing {identifier}: Searching - {msg}", foreground="blue"))

                if html_search_content:
                    product_link = extract_link(html_search_content, product_link_selector_type, product_link_selector, product_link_base_url, search_url,
                                                progress_callback=lambda msg, p: status_label.config(text=f"Processing {identifier}: Finding Product Link - {msg}", foreground="blue"))
                    extracted_data[f'{output_prefix} Link'][index] = product_link

                    if pd.notna(product_link):
                        html_product_content = get_html_content(product_link,
                                                               progress_callback=lambda msg, p: status_label.config(text=f"Processing {identifier}: Fetching Product Page - {msg}", foreground="blue"))
                        if html_product_content:
                            if family_selector:
                                family_data = extract_data(html_product_content, family_selector_type, family_selector)
                                extracted_data['Family'][index] = " > ".join(family_data) if family_data else None

                            if image_selector:
                                image_links = extract_image_links(html_product_content, image_selector_type, image_selector, urlparse(product_link).scheme + "://" + urlparse(product_link).netloc, product_link)
                                extracted_data['Image Links'][index] = image_links

                            status_label.config(text=f"Processing {identifier}: Product page fetched and data extracted.", foreground="green")
                            progress_percent = int(((index + 1) / total_items) * 80)
                            progress_var.set(progress_percent)
                            progress_percent_label.config(text=f"{progress_percent}%")
                            root.update()
                        else:
                            status_label.config(text=f"Processing {identifier}: Failed to fetch product page.", foreground="orange")
                    else:
                        status_label.config(text=f"Processing {identifier}: Product link not found.", foreground="orange")
                else:
                    status_label.config(text=f"Processing Identifier: {identifier} - Could not retrieve search results.", foreground="orange")

                time.sleep(0.1)

            except Exception as e:
                print(f"Error processing Identifier {identifier}: {e}")
                status_label.config(text=f"Error processing Identifier {identifier}: {e}", foreground="red")

        for col, data in extracted_data.items():
            df[col] = data

        max_images = max(len(links) for links in extracted_data['Image Links']) if extracted_data['Image Links'] else 0
        for i in range(max_images):
            df[f"Image Link {i+1}"] = [links[i] if len(links) > i else None for links in extracted_data['Image Links']]

        output_excel_file = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                        initialfile="Output_extracted_data.xlsx",
                                                        title="Save Output Excel File")
        if output_excel_file:
            df.to_excel(output_excel_file, index=False)
            success_message = f"Data extracted and saved to '{output_excel_file}'"
            status_label.config(text=success_message, foreground="green")
            messagebox.showinfo("Success", success_message)
        else:
            status_label.config(text="Saving cancelled.", foreground="orange")

    except FileNotFoundError:
        error_message = f"Error: Excel file '{excel_file}' not found."
        messagebox.showerror("Error", error_message)
        status_label.config(text=error_message, foreground="red")
    except Exception as e:
        error_message = f"An unexpected error occurred: {e}"
        messagebox.showerror("Error", error_message)
        status_label.config(text=error_message, foreground="red")
    finally:
        progress_var.set(0)
        progress_percent_label.config(text="0%")
        browse_button.config(state=tk.NORMAL)
        start_process_button.config(state=tk.NORMAL)

def load_config():
    config = {}
    config['mpn_column'] = mpn_column_entry.get()
    config['search_url_format'] = search_url_entry.get()
    config['product_link_selector'] = product_link_selector_entry.get()
    config['product_link_selector_type'] = product_link_selector_type_var.get()
    config['product_link_base_url'] = product_link_base_url_entry.get()
    config['family_selector'] = family_selector_entry.get()
    config['family_selector_type'] = family_selector_type_var.get()
    config['image_selector'] = image_selector_entry.get()
    config['image_selector_type'] = image_selector_type_var.get()
    config['output_prefix'] = output_prefix_entry.get()
    return config

def start_processing_generalized():
    global file_selected
    global excel_file_path_global

    if not file_selected:
        messagebox.showerror("Error", "Please select an Excel file.")
        return
    if not excel_file_path_global:
        messagebox.showerror("Error", "Please select an Excel file.")
        return

    config = load_config()

    if not all([config.get('mpn_column'), config.get('search_url_format'), config.get('product_link_selector')]):
        messagebox.showerror("Error", "Please fill in the required configuration fields.")
        return

    browse_button.config(state=tk.DISABLED)
    start_process_button.config(state=tk.DISABLED)
    status_label.config(text="Processing...", foreground="blue")
    threading.Thread(target=process_manufacturer, args=(excel_file_path_global, config, progress_var, progress_percent_label, status_label, root)).start()

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Generalized Product Data Extractor")
    root.geometry("750x700")
    root.resizable(False, False)

    style = ttk.Style()
    style.theme_use('clam')
    style.configure('TLabel', padding=8)
    style.configure('TEntry', padding=8)
    style.configure('TButton', padding=10)
    style.configure('TLabelframe.Label', font=('TkDefaultFont', 12, 'bold'))

    input_frame = ttk.LabelFrame(root, text="1. Input Settings", padding=15)
    input_frame.pack(padx=20, pady=15, fill=tk.X)

    excel_label = ttk.Label(input_frame, text="Select Excel File:")
    excel_label.grid(row=0, column=0, padx=10, pady=8, sticky=tk.W)
    excel_entry = ttk.Entry(input_frame, width=50, state='readonly')
    excel_entry.grid(row=0, column=1, padx=10, pady=8, sticky=tk.EW)
    browse_button = ttk.Button(input_frame, text="Browse", command=browse_file)
    browse_button.grid(row=0, column=2, padx=10, pady=8, sticky=tk.E)
    file_selected_label = ttk.Label(input_frame, text="No file selected.", foreground="red")
    file_selected_label.grid(row=1, column=0, columnspan=3, padx=10, pady=5, sticky=tk.W)

    config_frame = ttk.LabelFrame(root, text="2. Manufacturer Configuration", padding=15)
    config_frame.pack(padx=20, pady=10, fill=tk.X)

    mpn_column_label = ttk.Label(config_frame, text="Identifier Column Name:")
    mpn_column_label.grid(row=0, column=0, padx=10, pady=8, sticky=tk.W)
    mpn_column_entry = ttk.Entry(config_frame, width=40)
    mpn_column_entry.insert(0, "MPN")
    mpn_column_entry.grid(row=0, column=1, padx=10, pady=8, sticky=tk.EW)

    search_url_label = ttk.Label(config_frame, text="Search URL Format (e.g., https://example.com/search?q={mpn}):")
    search_url_label.grid(row=1, column=0, padx=10, pady=8, sticky=tk.W)
    search_url_entry = ttk.Entry(config_frame, width=40)
    search_url_entry.grid(row=1, column=1, padx=10, pady=8, sticky=tk.EW)

    product_link_label = ttk.Label(config_frame, text="Product Link Selector:")
    product_link_label.grid(row=2, column=0, padx=10, pady=8, sticky=tk.W)
    product_link_selector_entry = ttk.Entry(config_frame, width=40)
    product_link_selector_entry.grid(row=2, column=1, padx=10, pady=8, sticky=tk.EW)
    product_link_selector_type_var = tk.StringVar(value='xpath')
    product_link_xpath_radio = ttk.Radiobutton(config_frame, text="XPath", variable=product_link_selector_type_var, value='xpath')
    product_link_xpath_radio.grid(row=3, column=0, padx=10, pady=2, sticky=tk.W)
    product_link_css_radio = ttk.Radiobutton(config_frame, text="CSS Selector", variable=product_link_selector_type_var, value='css')
    product_link_css_radio.grid(row=3, column=1, padx=10, pady=2, sticky=tk.W)

    product_link_base_url_label = ttk.Label(config_frame, text="Product Link Base URL (optional):")
    product_link_base_url_label.grid(row=4, column=0, padx=10, pady=8, sticky=tk.W)
    product_link_base_url_entry = ttk.Entry(config_frame, width=40)
    product_link_base_url_entry.grid(row=4, column=1, padx=10, pady=8, sticky=tk.EW)

    family_label = ttk.Label(config_frame, text="Family Hierarchy Selector (optional):")
    family_label.grid(row=5, column=0, padx=10, pady=8, sticky=tk.W)
    family_selector_entry = ttk.Entry(config_frame, width=40)
    family_selector_entry.grid(row=5, column=1, padx=10, pady=8, sticky=tk.EW)
    family_selector_type_var = tk.StringVar(value='xpath')
    family_xpath_radio = ttk.Radiobutton(config_frame, text="XPath", variable=family_selector_type_var, value='xpath')
    family_xpath_radio.grid(row=6, column=0, padx=10, pady=2, sticky=tk.W)
    family_css_radio = ttk.Radiobutton(config_frame, text="CSS Selector", variable=family_selector_type_var, value='css')
    family_css_radio.grid(row=6, column=1, padx=10, pady=2, sticky=tk.W)

    image_label = ttk.Label(config_frame, text="Image Links Selector (optional):")
    image_label.grid(row=7, column=0, padx=10, pady=8, sticky=tk.W)
    image_selector_entry = ttk.Entry(config_frame, width=40)
    image_selector_entry.grid(row=7, column=1, padx=10, pady=8, sticky=tk.EW)
    image_selector_type_var = tk.StringVar(value='xpath')
    image_xpath_radio = ttk.Radiobutton(config_frame, text="XPath", variable=image_selector_type_var, value='xpath')
    image_xpath_radio.grid(row=8, column=0, padx=10, pady=2, sticky=tk.W)
    image_css_radio = ttk.Radiobutton(config_frame, text="CSS Selector", variable=image_selector_type_var, value='css')
    image_css_radio.grid(row=8, column=1, padx=10, pady=2, sticky=tk.W)

    output_prefix_label = ttk.Label(config_frame, text="Output Column Prefix (optional):")
    output_prefix_label.grid(row=9, column=0, padx=10, pady=8, sticky=tk.W)
    output_prefix_entry = ttk.Entry(config_frame, width=40)
    output_prefix_entry.insert(0, "Product")
    output_prefix_entry.grid(row=9, column=1, padx=10, pady=8, sticky=tk.EW)

    start_process_button_frame = ttk.Frame(root, padding=10)
    start_process_button_frame.pack(fill=tk.X, padx=20, pady=5)
    start_process_button = ttk.Button(start_process_button_frame, text="3. Start Data Extraction", state=tk.DISABLED,
                                        command=start_processing_generalized)
    start_process_button.pack(pady=8, fill=tk.X)

    progress_frame = ttk.LabelFrame(root, text="4. Extraction Progress", padding=15)
    progress_frame.pack(padx=20, pady=10, fill=tk.X)

    progress_var = tk.IntVar()
    progress_bar = ttk.Progressbar(progress_frame, variable=progress_var, mode='determinate')
    progress_bar.pack(fill=tk.X, pady=10)

    progress_percent_label = ttk.Label(progress_frame, text="0%", anchor=tk.E)
    progress_percent_label.pack(fill=tk.X, padx=10, pady=5)

    status_label = ttk.Label(progress_frame, text="Ready to select Excel file and configure settings.", anchor=tk.W)
    status_label.pack(fill=tk.X, pady=8)

    root.mainloop()
  

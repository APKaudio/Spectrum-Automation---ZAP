import csv
import subprocess
import sys
import xml.etree.ElementTree as ET
import os
# messagebox is no longer needed here as the GUI will handle messages
# from tkinter import messagebox 

# --- BeautifulSoup Installation Check ---
# This block checks if BeautifulSoup4 is installed. If not, it attempts to install it.
# This is crucial for parsing HTML files.
try:
    from bs4 import BeautifulSoup
except ImportError:
    print("BeautifulSoup4 not found. Attempting to install...")
    try:
        # Use subprocess to run pip install for BeautifulSoup4
        subprocess.check_call([sys.executable, "-m", "pip", "install", "beautifulsoup4"])
        from bs4 import BeautifulSoup
        print("BeautifulSoup4 installed successfully.")
    except subprocess.CalledProcessError as e:
        # If installation fails, we need to exit or ensure the main app handles it.
        # For a utility file, it's better to raise an error or print a critical message.
        # Keeping messagebox for critical installation error, but generally avoid in utils.
        from tkinter import messagebox
        messagebox.showerror("Installation Error", f"Error installing BeautifulSoup4: {e}\\nPlease install it manually by running: pip install beautifulsoup4")
        sys.exit(1)
    except Exception as e:
        from tkinter import messagebox
        messagebox.showerror("Installation Error", f"An unexpected error occurred during BeautifulSoup4 installation: {e}")
        sys.exit(1)


def convert_html_report_to_csv(html_content):
    """
    Converts the HTML frequency coordination report into a list of dictionaries
    suitable for CSV output, handling multiple zones.

    Args:
        html_content (str): The full HTML content of the report.

    Returns:
        tuple: A tuple containing:
               - list: A list of strings representing the CSV headers.
               - list: A list of dictionaries, where each dictionary represents a row
                       in the CSV and keys are column headers.
    """
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Define the headers for the CSV file
    csv_headers = [
        "ZONE",
        "GROUP",
        "DEVICE",
        "NAME",
        "FREQ"
    ]
    
    data_rows = [] # List to store all extracted data rows

    main_content_container = None
    
    # Attempt to find the main content container based on specific HTML structures
    first_zone_p = soup.find('p', style=lambda value: value and 'font-size: large' in value and 'text-decoration: underline' in value)

    if first_zone_p:
        main_content_container = first_zone_p.find_parent('span')
    
    # Fallback search for the main content container if the first attempt fails
    if not main_content_container:
        main_table = soup.find('table', class_='MainTable')
        if main_table:
            main_table_trs = main_table.find_all('tr')
            if len(main_table_trs) > 1:
                second_tr_td = main_table_trs[1].find('td')
                if second_tr_td:
                    potential_span_wrapper = second_tr_td.find('span')
                    if potential_span_wrapper:
                        main_content_container = potential_span_wrapper
                    else:
                        main_content_container = second_tr_td
    
    # If the main content container is still not found, print a warning and return empty data
    if not main_content_container:
        print("Warning: Could not find the main content container. No data will be extracted.")
        return csv_headers, data_rows

    current_zone_type = "" # Variable to hold the current zone being processed
    # Iterate through the children of the identified main content container
    for element in main_content_container.children:
        # Identify new zones based on specific paragraph styles
        if element.name == 'p' and element.get('style') and \
           'font-size: large' in element.get('style') and \
           'text-decoration: underline' in element.get('style'):
            zone_text = element.get_text(strip=True)
            if zone_text.startswith("Zone:"):
                current_zone_type = zone_text.replace("Zone:", "").strip()
        
        # Process "Assignment" tables to extract device and frequency information
        elif element.name == 'table' and 'Assignment' in element.get('class', []):
            table = element # This is an assignment table
            
            # Extract the group name from the table's header
            device_name_tag = table.find('th')
            current_group_name = device_name_tag.get_text(strip=True) if device_name_tag else ""

            # Find all data rows within the table (skipping the header row)
            rows = table.find_all('tr')[1:]

            for row in rows:
                # Handle rows with inner <span> wrappers
                data_spans = row.find_all('span')
                
                if data_spans:
                    for data_span in data_spans:
                        cells = data_span.find_all('td')
                        if len(cells) >= 4: # Ensure enough cells for data extraction
                            band_type = cells[0].get_text(strip=True)
                            
                            channel_frequency_tag = cells[3].find('b')
                            channel_frequency = channel_frequency_tag.get_text(strip=True) if channel_frequency_tag else ""

                            channel_name = cells[1].get_text(strip=True)
                            if not channel_name:
                                channel_name = cells[2].get_text(strip=True)
                            
                            # Create a dictionary for the current row
                            row_data = {
                                "ZONE": current_zone_type,
                                "GROUP": current_group_name,
                                "DEVICE": band_type,
                                "NAME": channel_name,
                                "FREQ": channel_frequency
                            }
                            # Add row only if it contains meaningful data
                            if band_type or channel_frequency or channel_name:
                                data_rows.append(row_data)
                # Handle rows with direct <td> elements
                else:
                    cells = row.find_all('td')
                    if len(cells) >= 4:
                        band_type = cells[0].get_text(strip=True)
                        channel_frequency_tag = cells[3].find('b')
                        channel_frequency = channel_frequency_tag.get_text(strip=True) if channel_frequency_tag else ""
                        
                        channel_name = cells[1].get_text(strip=True)
                        if not channel_name:
                            channel_name = cells[2].get_text(strip=True)

                        row_data = {
                            "ZONE": current_zone_type,
                            "GROUP": current_group_name,
                            "DEVICE": band_type,
                            "NAME": channel_name,
                            "FREQ": channel_frequency
                        }
                        if band_type or channel_frequency or channel_name:
                            data_rows.append(row_data)
    
    return csv_headers, data_rows

def generate_csv_from_shw(xml_file_path):
    """
    Generates CSV data from a .shw XML file, extracting frequency entry information.

    Args:
        xml_file_path (str): The path to the input .shw XML file.

    Returns:
        tuple: A tuple containing:
               - list: A list of strings representing the CSV headers.
               - list: A list of dictionaries, where each dictionary represents a row
                       in the CSV and keys are column headers.
    Raises:
        FileNotFoundError: If the specified XML file does not exist.
        ET.ParseError: If there is an error parsing the XML file.
        Exception: For any other unexpected errors during processing.
    """
    try:
        tree = ET.parse(xml_file_path) # Parse the XML file
        root = tree.getroot()

        csv_data = [] # List to store extracted data
        headers = ["ZONE", "GROUP", "DEVICE", "NAME", "FREQ"] # Define CSV headers

        # Iterate through each 'freq_entry' element in the XML
        for freq_entry in root.findall('.//freq_entry'):
            # Extract ZONE
            zone_element = freq_entry.find('compat_key/zone')
            zone = zone_element.text if zone_element is not None else "N/A"

            # Extract GROUP from the 'tag' attribute
            group = freq_entry.get('tag', "N/A")

            # Extract DEVICE by combining manufacturer, model, and band
            manufacturer = freq_entry.find('manufacturer').text if freq_entry.find('manufacturer') is not None else "N/A"
            model = freq_entry.find('model').text if freq_entry.find('model') is not None else "N/A"
            band_element = freq_entry.find('compat_key/band')
            band = band_element.text if band_element is not None else "N/A"
            device = f"{manufacturer} - {model} - {band}"

            # Extract NAME
            name_element = freq_entry.find('source_name')
            name = name_element.text if name_element is not None else "N/A"

            # Extract FREQ and convert to kHz (divide by 1000)
            freq_element = freq_entry.find('value')
            freq = "N/A"
            if freq_element is not None and freq_element.text is not None:
                try:
                    freq = float(freq_element.text) / 1000.0
                except ValueError:
                    freq = "Invalid Frequency"

            # Append the extracted data to the list
            csv_data.append({
                "ZONE": zone,
                "GROUP": group,
                "DEVICE": device,
                "NAME": name,
                "FREQ": freq
            })
        return headers, csv_data

    except FileNotFoundError:
        raise FileNotFoundError(f"The file '{xml_file_path}' was not found.")
    except ET.ParseError as e:
        raise ET.ParseError(f"Error parsing XML file: {e}")
    except Exception as e:
        raise Exception(f"An unexpected error occurred during SHW conversion: {e}")


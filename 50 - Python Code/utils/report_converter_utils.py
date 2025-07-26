import csv
import subprocess
import sys
import xml.etree.ElementTree as ET
import os
import re 
from bs4 import BeautifulSoup # BeautifulSoup is imported here, no need for try-except block in this file

# --- BeautifulSoup Installation Check (moved to main_app.py or handled by environment) ---
# This block is typically handled at the application's entry point (e.g., main_app.py)
# to ensure all dependencies are met before utility functions are called.
# Keeping it here for standalone testing purposes, but it's redundant if main_app.py handles it.
try:
    from bs4 import BeautifulSoup
except ImportError:
    print("BeautifulSoup4 not found. Attempting to install...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "beautifulsoup4"])
        from bs4 import BeautifulSoup
        print("BeautifulSoup4 installed successfully.")
    except subprocess.CalledProcessError as e:
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
    suitable for CSV output, handling multiple zones. This version is based on
    the IAS HTML to CSV.py prototype for accurate extraction.

    Inputs:
        html_content (str): The full HTML content of the report.

    Returns:
        tuple: A tuple containing:
               - list: A list of strings representing the CSV headers.
               - list: A list of dictionaries, where each dictionary represents a row
                       in the CSV and keys are column headers.
    """
    soup = BeautifulSoup(html_content, 'html.parser')
    
    csv_headers = [
        "ZONE",
        "GROUP",
        "DEVICE",
        "NAME",
        "FREQ"
    ]
    
    data_rows = []

    # Find the main content area within the HTML, based on the IAS prototype.
    main_content_container = None
    
    first_zone_p = soup.find('p', style=lambda value: value and 'font-size: large' in value and 'text-decoration: underline' in value)

    if first_zone_p:
        main_content_container = first_zone_p.find_parent('span')
    
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
    
    if not main_content_container:
        print("Warning: Could not find the main content container. No data will be extracted.")
        return csv_headers, data_rows

    current_zone_type = ""
    # Iterate through the children of the identified main content container
    for element in main_content_container.children:
        if element.name == 'p' and element.get('style') and \
           'font-size: large' in element.get('style') and \
           'text-decoration: underline' in element.get('style'):
            zone_text = element.get_text(strip=True)
            if zone_text.startswith("Zone:"):
                current_zone_type = zone_text.replace("Zone:", "").strip()
        
        elif element.name == 'table' and 'Assignment' in element.get('class', []):
            table = element
            
            device_name_tag = table.find('th')
            current_group_name = device_name_tag.get_text(strip=True) if device_name_tag else ""

            rows_in_table = table.find_all('tr')[1:] # Skip the first row as it contains the <th> (device_name)

            for row in rows_in_table:
                data_spans = row.find_all('span')
                
                if data_spans:
                    for data_span in data_spans:
                        cells = data_span.find_all('td')
                        if len(cells) >= 4:
                            band_type = cells[0].get_text(strip=True)
                            
                            channel_frequency_tag = cells[3].find('b')
                            channel_frequency_str = channel_frequency_tag.get_text(strip=True) if channel_frequency_tag else ""

                            channel_name = cells[1].get_text(strip=True)
                            if not channel_name:
                                channel_name = cells[2].get_text(strip=True)
                            
                            # Convert frequency string to kHz (prototype does not use regex here)
                            freq_khz = "N/A"
                            try:
                                # Assuming the frequency string is like "490.900 MHz" or "1000 kHz"
                                # We need to extract the number and convert to kHz
                                freq_match = re.search(r'(\d+(\.\d+)?)\s*(kHz|MHz|GHz)', channel_frequency_str, re.IGNORECASE)
                                if freq_match:
                                    value = float(freq_match.group(1))
                                    unit = freq_match.group(3).lower()
                                    if unit == 'mhz':
                                        freq_khz = value * 1000
                                    elif unit == 'ghz':
                                        freq_khz = value * 1_000_000
                                    else: # Assume kHz if no unit or 'khz'
                                        freq_khz = value
                                else:
                                    # Fallback if regex doesn't match, try direct float conversion (assuming kHz)
                                    freq_khz = float(channel_frequency_str) # Assume it's already in kHz if no unit
                                    print(f"WARNING (HTML): No unit found for '{channel_frequency_str}'. Assuming kHz.")
                            except ValueError:
                                print(f"WARNING (HTML): Could not convert frequency '{channel_frequency_str}' to float.")
                                freq_khz = "Invalid Frequency"

                            row_data = {
                                "ZONE": current_zone_type,
                                "GROUP": current_group_name,
                                "DEVICE": band_type,
                                "NAME": channel_name,
                                "FREQ": freq_khz
                            }
                            if band_type or channel_frequency_str or channel_name:
                                data_rows.append(row_data)
                else:
                    # Process rows that have <td>s directly (e.g., blank rows or specific structures without inner spans)
                    cells = row.find_all('td')
                    if len(cells) >= 4: 
                        band_type = cells[0].get_text(strip=True)
                        channel_frequency_tag = cells[3].find('b')
                        channel_frequency_str = channel_frequency_tag.get_text(strip=True) if channel_frequency_tag else ""
                        
                        channel_name = cells[1].get_text(strip=True)
                        if not channel_name:
                            channel_name = cells[2].get_text(strip=True)

                        # Convert frequency string to kHz (prototype does not use regex here)
                        freq_khz = "N/A"
                        try:
                            freq_match = re.search(r'(\d+(?:\.\d+)?)\s*(?:(k|m|g)?hz)?', channel_frequency_str, re.IGNORECASE)
                            if freq_match:
                                value = float(freq_match.group(1))
                                unit = freq_match.group(3).lower()
                                if unit == 'mhz':
                                    freq_khz = value * 1000
                                elif unit == 'ghz':
                                    freq_khz = value * 1_000_000
                                else: 
                                    freq_khz = value
                            else:
                                freq_khz = float(channel_frequency_str) # Assume it's already in kHz if no unit
                                print(f"WARNING (HTML): No unit found for '{channel_frequency_str}'. Assuming kHz.")
                        except ValueError:
                            print(f"WARNING (HTML): Could not convert frequency '{channel_frequency_str}' to float.")
                            freq_khz = "Invalid Frequency"

                        row_data = {
                            "ZONE": current_zone_type,
                            "GROUP": current_group_name,
                            "DEVICE": band_type,
                            "NAME": channel_name,
                            "FREQ": freq_khz
                        }
                        if band_type or channel_frequency_str or channel_name:
                            data_rows.append(row_data)
    
    return csv_headers, data_rows


def generate_csv_from_shw(xml_file_path):
    """
    Parses an SHW (XML) file and extracts frequency data, converting it
    into a standardized CSV format. This version is based on the SHOW to CSV.py
    prototype for accurate extraction of ZONE and GROUP.

    Inputs:
        xml_file_path (str): The full path to the SHW (XML) file.
    Outputs:
        tuple: A tuple containing:
               - headers (list): A list of strings representing the CSV header row.
               - csv_data (list): A list of dictionaries, where each dictionary
                                  represents a row of data with keys matching the headers.
    Raises:
        FileNotFoundError: If the specified XML file does not exist.
        xml.etree.ElementTree.ParseError: If the XML file is malformed.
        Exception: For other parsing or data extraction errors.
    """
    headers = ["ZONE", "GROUP", "DEVICE", "NAME", "FREQ"]
    csv_data = []

    try:
        with open(xml_file_path, 'r', encoding='utf-8') as f:
            tree = ET.parse(f)
        root = tree.getroot()

        # Iterate through 'freq_entry' elements
        for freq_entry in root.findall('.//freq_entry'):
            # Reverting ZONE and GROUP extraction to match SHOW to CSV.py prototype
            zone_element = freq_entry.find('compat_key/zone')
            zone = zone_element.text if zone_element is not None else "N/A"

            group = freq_entry.get('tag', "N/A") # Extract GROUP from the 'tag' attribute of freq_entry
            
            # Extract DEVICE (manufacturer, model, band)
            manufacturer = freq_entry.find('manufacturer').text if freq_entry.find('manufacturer') is not None else "N/A"
            model = freq_entry.find('model').text if freq_entry.find('model') is not None else "N/A"
            band_element = freq_entry.find('compat_key/band') 
            band = band_element.text if band_element is not None else "N/A"
            device = f"{manufacturer} - {model} - {band}"

            # Extract NAME
            name_element = freq_entry.find('source_name')
            name = name_element.text if name_element is not None else "N/A"

            # Extract FREQ from value and convert to kHz (assuming value is in Hz)
            freq_element = freq_entry.find('value')
            freq = "N/A"
            if freq_element is not None and freq_element.text is not None:
                freq_str = freq_element.text 
                
                print(f"DEBUG (SHW): Processing freq_str: '{freq_str}' for device '{name}'")

                # The prototype directly converts to float and divides by 1000.0
                # Assuming SHW 'value' is consistently in Hz and numeric.
                try:
                    freq = float(freq_str) / 1000.0 # Convert Hz to kHz
                except ValueError:
                    print(f"WARNING (SHW): Could not convert SHW frequency value '{freq_str}' to float.")
                    freq = "Invalid Frequency"

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
        raise ET.ParseError(f"Error parsing XML (SHW) file '{xml_file_path}': {e}")
    except Exception as e:
        print(f"Error during SHW conversion data extraction: {e}")
        raise 
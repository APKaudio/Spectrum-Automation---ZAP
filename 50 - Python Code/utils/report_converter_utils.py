import csv
import subprocess
import sys
import xml.etree.ElementTree as ET
import os
import re
from bs4 import BeautifulSoup # BeautifulSoup is imported here, no need for try-except block in this file
import pdfplumber # Import pdfplumber for PDF conversion

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
        messagebox.showerror("Installation Error", f"Error installing BeautifulSoup4: {e}\nPlease install it manually by running: pip install beautifulsoup4")
        sys.exit(1)
    except Exception as e:
        from tkinter import messagebox
        messagebox.showerror("Installation Error", f"An unexpected error occurred during BeautifulSoup4 installation: {e}")
        sys.exit(1)

# --- PDFPlumber Installation Check (New) ---
try:
    import pdfplumber
except ImportError:
    print("pdfplumber not found. Attempting to install...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pdfplumber"])
        import pdfplumber
        print("pdfplumber installed successfully.")
    except subprocess.CalledProcessError as e:
        from tkinter import messagebox
        messagebox.showerror("Installation Error", f"Error installing pdfplumber: {e}\nPlease install it manually by running: pip install pdfplumber")
        sys.exit(1)
    except Exception as e:
        from tkinter import messagebox
        messagebox.showerror("Installation Error", f"An unexpected error occurred during pdfplumber installation: {e}")
        sys.exit(1)


def convert_html_report_to_csv(html_content):
    """
    Converts the HTML frequency coordination report into a list of dictionaries
    suitable for CSV output, handling multiple zones. This version is based on
    the IAS HTML to CSV.py prototype for accurate extraction.
    All frequencies are converted to MHz for consistency.

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
                            
                            # Convert frequency string to MHz
                            freq_mhz = "N/A"
                            try:
                                freq_match = re.search(r'(\d+(\.\d+)?)\s*(kHz|MHz|GHz)', channel_frequency_str, re.IGNORECASE)
                                if freq_match:
                                    value = float(freq_match.group(1))
                                    unit = freq_match.group(3).lower()
                                    if unit == 'mhz':
                                        freq_mhz = value
                                    elif unit == 'ghz':
                                        freq_mhz = value * 1000 # GHz to MHz
                                    elif unit == 'khz':
                                        freq_mhz = value / 1000 # kHz to MHz
                                else:
                                    # Fallback if regex doesn't match, assume MHz
                                    freq_mhz = float(channel_frequency_str) # Assume it's already in MHz
                                    print(f"WARNING (HTML): No unit found for '{channel_frequency_str}'. Assuming MHz.")
                            except ValueError:
                                print(f"WARNING (HTML): Could not convert frequency '{channel_frequency_str}' to float.")
                                freq_mhz = "Invalid Frequency"

                            row_data = {
                                "ZONE": current_zone_type,
                                "GROUP": current_group_name,
                                "DEVICE": band_type,
                                "NAME": channel_name,
                                "FREQ": freq_mhz # Store in MHz
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

                        # Convert frequency string to MHz
                        freq_mhz = "N/A"
                        try:
                            freq_match = re.search(r'(\d+(?:\.\d+)?)\s*(?:(k|m|g)?hz)?', channel_frequency_str, re.IGNORECASE)
                            if freq_match:
                                value = float(freq_match.group(1))
                                unit_group = freq_match.group(3)
                                if unit_group:
                                    unit = unit_group.lower()
                                    if unit == 'm': # MHz
                                        freq_mhz = value
                                    elif unit == 'g': # GHz
                                        freq_mhz = value * 1000
                                    elif unit == 'k': # kHz
                                        freq_mhz = value / 1000
                                else: # No unit specified, assume MHz
                                    freq_mhz = value
                            else:
                                # Fallback if regex doesn't match, assume MHz
                                freq_mhz = float(channel_frequency_str) # Assume it's already in MHz
                                print(f"WARNING (HTML): No unit found for '{channel_frequency_str}'. Assuming MHz.")
                        except ValueError:
                            print(f"WARNING (HTML): Could not convert frequency '{channel_frequency_str}' to float.")
                            freq_mhz = "Invalid Frequency"

                        row_data = {
                            "ZONE": current_zone_type,
                            "GROUP": current_group_name,
                            "DEVICE": band_type,
                            "NAME": channel_name,
                            "FREQ": freq_mhz # Store in MHz
                        }
                        if band_type or channel_frequency_str or channel_name:
                            data_rows.append(row_data)
    
    return csv_headers, data_rows


def generate_csv_from_shw(xml_file_path):
    """
    Parses an SHW (XML) file and extracts frequency data, converting it
    into a standardized CSV format. This version is based on the SHOW to CSV.py
    prototype for accurate extraction of ZONE and GROUP.
    All frequencies are converted to MHz for consistency.

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

            # Extract FREQ from value and convert to MHz (assuming value is in Hz)
            freq_element = freq_entry.find('value')
            freq_mhz = "N/A"
            if freq_element is not None and freq_element.text is not None:
                freq_str = freq_element.text 
                
                print(f"DEBUG (SHW): Processing freq_str: '{freq_str}' for device '{name}'")

                try:
                    freq_mhz = float(freq_str) / 1_000_000.0 # Convert Hz to MHz
                except ValueError:
                    print(f"WARNING (SHW): Could not convert SHW frequency value '{freq_str}' to float.")
                    freq_mhz = "Invalid Frequency"

            csv_data.append({
                "ZONE": zone,
                "GROUP": group,
                "DEVICE": device,
                "NAME": name,
                "FREQ": freq_mhz # Store in MHz
            })
        return headers, csv_data

    except FileNotFoundError:
        raise FileNotFoundError(f"The file '{xml_file_path}' was not found.")
    except ET.ParseError as e:
        raise ET.ParseError(f"Error parsing XML (SHW) file '{xml_file_path}': {e}")
    except Exception as e:
        print(f"Error during SHW conversion data extraction: {e}")
        raise

def convert_pdf_report_to_csv(pdf_file_path):
    """
    Parses a PDF file (Sound Base format) and extracts frequency data, converting it
    into a standardized CSV format. This function maps PDF fields to the MARKERS.CSV
    structure as follows:
    - PDF 'Group' -> CSV 'ZONE'
    - PDF 'Model' -> CSV 'GROUP'
    - PDF 'Name' -> CSV 'NAME'
    - PDF 'Frequency' -> CSV 'FREQ' (in MHz)
    - CSV 'DEVICE' is constructed from PDF 'Model', 'Band', and 'Preset'.

    Inputs:
        pdf_file_path (str): The full path to the PDF file.
    Outputs:
        tuple: A tuple containing:
               - headers (list): A list of strings representing the CSV header row.
               - csv_data (list): A list of dictionaries, where each dictionary
                                  represents a row of data with keys matching the headers.
    Raises:
        FileNotFoundError: If the specified PDF file does not exist.
        Exception: For other parsing or data extraction errors.
    """
    headers = ["ZONE", "GROUP", "DEVICE", "NAME", "FREQ"]
    csv_data = []

    try:
        with pdfplumber.open(pdf_file_path) as pdf:
            last_known_group = "Uncategorized" # Default group if not found

            for page in pdf.pages:
                # Extract text for group headers
                lines = page.extract_text().splitlines()
                lines = [line.strip() for line in lines if line.strip()]

                group_headers = [(i, line) for i, line in enumerate(lines)
                                 if re.match(r".+\(\d+ frequencies\)", line)]

                tables = page.extract_tables()

                group_index = 0
                for table in tables:
                    if group_index < len(group_headers):
                        last_known_group = group_headers[group_index][1]
                        group_index += 1

                    current_zone = last_known_group # PDF Group -> CSV ZONE

                    for row in table:
                        if not row or all(cell is None or cell.strip() == "" for cell in row):
                            continue

                        if "Model" in row[0] and "Frequency" in row[-1]: # Skip header rows
                            continue

                        clean_row = [cell.replace("\n", " ").strip() if cell else "" for cell in row]
                        # Ensure row has at least 6 elements to unpack safely
                        while len(clean_row) < 6:
                            clean_row.append("")

                        model_pdf, band_pdf, name_pdf, preset_pdf, spacing_pdf, frequency_pdf_str = clean_row

                        if model_pdf.strip() == current_zone.strip(): # Skip rows that mistakenly repeat the group name
                            continue

                        # Map PDF fields to CSV fields
                        zone_csv = current_zone
                        group_csv = model_pdf # PDF Model -> CSV GROUP

                        # Construct DEVICE from PDF Model, Band, Preset
                        device_csv = f"{model_pdf}"
                        if band_pdf:
                            device_csv += f" - {band_pdf}"
                        if preset_pdf:
                            device_csv += f" - {preset_pdf}"
                        
                        name_csv = name_pdf # PDF Name -> CSV NAME

                        freq_mhz_csv = "N/A"
                        try:
                            # Assume frequency_pdf_str is already in MHz or can be directly converted
                            freq_mhz_csv = float(frequency_pdf_str)
                        except ValueError:
                            print(f"WARNING (PDF): Could not convert PDF frequency value '{frequency_pdf_str}' to float (MHz).")
                            freq_mhz_csv = "Invalid Frequency"

                        csv_data.append({
                            "ZONE": zone_csv,
                            "GROUP": group_csv,
                            "DEVICE": device_csv,
                            "NAME": name_csv,
                            "FREQ": freq_mhz_csv
                        })
        return headers, csv_data

    except FileNotFoundError:
        raise FileNotFoundError(f"The file '{pdf_file_path}' was not found.")
    except Exception as e:
        print(f"Error during PDF conversion data extraction: {e}")
        raise


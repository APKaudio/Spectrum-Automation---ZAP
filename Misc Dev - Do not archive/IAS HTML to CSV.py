import csv
import subprocess
import sys

# Attempt to import BeautifulSoup. If it fails, try to install it.
try:
    from bs4 import BeautifulSoup
except ImportError:
    print("BeautifulSoup4 not found. Attempting to install...")
    try:
        # Use subprocess to run pip install
        subprocess.check_call([sys.executable, "-m", "pip", "install", "beautifulsoup4"])
        from bs4 import BeautifulSoup
        print("BeautifulSoup4 installed successfully.")
    except subprocess.CalledProcessError as e:
        print(f"Error installing BeautifulSoup4: {e}")
        print("Please install it manually by running: pip install beautifulsoup4")
        sys.exit(1) # Exit if installation fails
    except Exception as e:
        print(f"An unexpected error occurred during BeautifulSoup4 installation: {e}")
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
    
    # Define the headers for the CSV file with updated names and order
    csv_headers = [
        "ZONE",
        "GROUP",
        "DEVICE",
        "NAME",
        "FREQ"
    ]
    
    # List to store all the extracted data rows
    data_rows = []

    # Find the main content area within the HTML.
    # The report content is typically inside a <span> that contains the first Zone paragraph.
    main_content_container = None
    
    # Try to find the first paragraph that looks like a zone header
    first_zone_p = soup.find('p', style=lambda value: value and 'font-size: large' in value and 'text-decoration: underline' in value)

    if first_zone_p:
        # The actual data tables and zone headers are within a <span> parent of this paragraph
        main_content_container = first_zone_p.find_parent('span')
    
    # Fallback if the above doesn't work (e.g., if the first zone isn't directly in a span under the main td)
    if not main_content_container:
        main_table = soup.find('table', class_='MainTable')
        if main_table:
            # Look for the second <tr> in the MainTable, as the first <tr> often contains header info
            main_table_trs = main_table.find_all('tr')
            if len(main_table_trs) > 1:
                # Get the first <td> of the second <tr>
                second_tr_td = main_table_trs[1].find('td')
                if second_tr_td:
                    # Check if this td contains the span with "Report Body Elements"
                    potential_span_wrapper = second_tr_td.find('span')
                    if potential_span_wrapper:
                        main_content_container = potential_span_wrapper
                    else:
                        # If no span wrapper, the elements might be direct children of this td
                        main_content_container = second_tr_td
    
    if not main_content_container:
        print("Warning: Could not find the main content container. No data will be extracted.")
        return csv_headers, data_rows # Return empty if main content area not found

    current_zone_type = ""
    # Iterate through the children of the identified main content container
    for element in main_content_container.children:
        # Check if the element is a <p> tag that signifies a new zone
        # using the specific style attributes for identification
        if element.name == 'p' and element.get('style') and \
           'font-size: large' in element.get('style') and \
           'text-decoration: underline' in element.get('style'):
            zone_text = element.get_text(strip=True)
            if zone_text.startswith("Zone:"):
                current_zone_type = zone_text.replace("Zone:", "").strip()
        
        # Check if the element is an "Assignment" table
        elif element.name == 'table' and 'Assignment' in element.get('class', []):
            table = element # This is an assignment table
            
            # The device_name (or category) is in the <th> tag within each Assignment table
            device_name_tag = table.find('th')
            # The value for 'group' column is extracted from here
            current_group_name = device_name_tag.get_text(strip=True) if device_name_tag else ""

            # Find all rows (tr) within the current assignment table
            # Skip the first row as it contains the <th> (device_name)
            rows = table.find_all('tr')[1:] # Start from the second row to get data

            for row in rows:
                # Each 'span' tag within a 'tr' contains a set of <td> elements for one entry
                data_spans = row.find_all('span')
                
                # Process rows that have inner <span> wrappers
                if data_spans:
                    for data_span in data_spans:
                        cells = data_span.find_all('td')
                        if len(cells) >= 4: # Ensure there are enough cells to extract data
                            band_type = cells[0].get_text(strip=True)
                            
                            # channel/frequency is in a <b> tag within the 4th td
                            channel_frequency_tag = cells[3].find('b')
                            channel_frequency = channel_frequency_tag.get_text(strip=True) if channel_frequency_tag else ""

                            # channel/channel_name: Prioritize 2nd td, then 3rd td.
                            channel_name = cells[1].get_text(strip=True)
                            if not channel_name:
                                # get_text(strip=True) on cells[2] should correctly handle the hidden '
                                channel_name = cells[2].get_text(strip=True)
                            
                            # Create a dictionary for the current row with updated keys
                            row_data = {
                                "ZONE": current_zone_type,
                                "GROUP": current_group_name,
                                "DEVICE": band_type,
                                "NAME": channel_name,
                                "FREQ": channel_frequency
                            }
                            # Only add the row if at least one of the core data fields has content
                            if band_type or channel_frequency or channel_name:
                                data_rows.append(row_data)
                # Process rows that have <td>s directly (e.g., blank rows or specific structures without inner spans)
                else:
                    cells = row.find_all('td')
                    if len(cells) >= 4: # Ensure there are enough cells to extract data
                        band_type = cells[0].get_text(strip=True)
                        channel_frequency_tag = cells[3].find('b')
                        channel_frequency = channel_frequency_tag.get_text(strip=True) if channel_frequency_tag else ""
                        
                        channel_name = cells[1].get_text(strip=True)
                        if not channel_name:
                            channel_name = cells[2].get_text(strip=True)

                        # Create a dictionary for the current row with updated keys
                        row_data = {
                            "ZONE": current_zone_type,
                            "GROUP": current_group_name,
                            "DEVICE": band_type,
                            "NAME": channel_name,
                            "FREQ": channel_frequency
                        }
                        # Only add the row if at least one of the core data fields has content
                        if band_type or channel_frequency or channel_name:
                            data_rows.append(row_data)
    
    return csv_headers, data_rows

# Define the input HTML file name
input_html_file = "report.html"
# Define the output CSV file name
output_csv_file = "report.csv"

try:
    # Read the HTML content from the specified file
    with open(input_html_file, 'r', encoding='utf-8') as f:
        html_report_content = f.read()

    # Get the headers and data rows
    headers, rows = convert_html_report_to_csv(html_report_content)

    # Write the data to a CSV file
    with open(output_csv_file, 'w', newline='', encoding='utf-8') as csvfile:
        csv_writer = csv.DictWriter(csvfile, fieldnames=headers)
        
        csv_writer.writeheader() # Write the header row
        csv_writer.writerows(rows) # Write all the data rows

    print(f"Successfully converted HTML report from '{input_html_file}' to '{output_csv_file}'")

except FileNotFoundError:
    print(f"Error: The file '{input_html_file}' was not found. Please make sure the HTML report is in the same directory as the script and named '{input_html_file}'.")
except Exception as e:
    print(f"An error occurred: {e}")


import xml.etree.ElementTree as ET
import csv
import os

def generate_csv_from_shw(xml_file_path, csv_file_path="output.csv"):
    """
    Generates a CSV file from a .shw XML file, extracting frequency entry information.

    Args:
        xml_file_path (str): The path to the input .shw XML file.
        csv_file_path (str): The path for the output CSV file. Defaults to "output.csv".
    """
    try:
        # Parse the XML file
        tree = ET.parse(xml_file_path)
        root = tree.getroot()

        # Prepare data for CSV
        csv_data = []
        # Updated headers as per your request
        headers = ["ZONE", "GROUP", "DEVICE", "NAME", "FREQ"]

        # Iterate through each 'freq_entry' in the XML file
        # Using .// to find all freq_entry elements anywhere in the document
        for freq_entry in root.findall('.//freq_entry'):
            # Extract ZONE from compat_key/zone
            zone_element = freq_entry.find('compat_key/zone')
            zone = zone_element.text if zone_element is not None else "N/A"

            # Extract GROUP from the 'tag' attribute of freq_entry
            group = freq_entry.get('tag', "N/A")

            # Extract DEVICE from manufacturer - model - compat_key/band
            manufacturer = freq_entry.find('manufacturer').text if freq_entry.find('manufacturer') is not None else "N/A"
            model = freq_entry.find('model').text if freq_entry.find('model') is not None else "N/A"
            band_element = freq_entry.find('compat_key/band')
            band = band_element.text if band_element is not None else "N/A"
            device = f"{manufacturer} - {model} - {band}"

            # Extract NAME from source_name
            name_element = freq_entry.find('source_name')
            name = name_element.text if name_element is not None else "N/A"

            # Extract FREQ from value and divide by 1000, retaining significant figures
            freq_element = freq_entry.find('value')
            freq = "N/A"
            if freq_element is not None and freq_element.text is not None:
                try:
                    freq = float(freq_element.text) / 1000.0
                except ValueError:
                    freq = "Invalid Frequency"

            csv_data.append({
                "ZONE": zone,
                "GROUP": group,
                "DEVICE": device,
                "NAME": name,
                "FREQ": freq
            })

        # Write data to CSV file
        with open(csv_file_path, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=headers)
            writer.writeheader()
            writer.writerows(csv_data)

        print(f"CSV file '{csv_file_path}' generated successfully!")

    except FileNotFoundError:
        print(f"Error: The file '{xml_file_path}' was not found.")
    except ET.ParseError as e:
        print(f"Error parsing XML file: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

# Example usage:
# Assuming 'venue.shw' is in the same directory as the script
xml_file = 'venue.shw'
output_csv = 'venue_data.csv'

# Check if the XML file exists before attempting to process
if os.path.exists(xml_file):
    generate_csv_from_shw(xml_file, output_csv)
else:
    print(f"The file '{xml_file}' does not exist in the current directory. Please ensure it's there.")

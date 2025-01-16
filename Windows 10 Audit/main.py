import csv
from openpyxl import Workbook

# This code generates an Excel file with:
# - Dealerships listed in alphabetical order.
# - Excludes "Foundation," "Corwin," and "Colorado Comptech."

# Open the CSV file
with open("dealership_data.csv", "r") as csv_file:
    # Read the CSV file
    csv_reader = csv.reader(csv_file)
    # Skip the header rows and the irrelevant sections
    header_skipped = False
    data_lines = []
    
    for row in csv_reader:
        if header_skipped:
            # Collect actual data rows
            data_lines.append(row)
        elif "Customer" in row and "Devices" in row:
            # Identify the start of the relevant data
            header_skipped = True
    
    # Create a dictionary to store data by dealership
    dealership_data = {}
    
    # Process each row
    for row in data_lines:
        customer, site, property, description, instances, devices = row
        devices = devices.split()  # Split device IDs into a list

        # Skip dealerships that should be excluded
        if customer in ["Foundation", "Corwin", "Colorado Comptech"]:
            continue
        
        # Group data by customer
        if customer not in dealership_data:
            dealership_data[customer] = []
        
        dealership_data[customer].append({
            "property": property,
            "description": description,
            "devices": devices,
        })
    
    # Filter and process data
    output_data = {}
    
    for customer, entries in dealership_data.items():
        output_data[customer] = []
        # Find all devices on Windows 10
        windows_10_devices = [
            entry for entry in entries if entry["property"] == "Operating system by version"
            and "Windows 10" in entry["description"]
        ]
        
        # Get the list of devices running Windows 10
        devices_on_windows_10 = []
        for entry in windows_10_devices:
            devices_on_windows_10.extend(entry["devices"])
        
        # Find the CPU for each Windows 10 device
        for device_id in devices_on_windows_10:
            cpu_entry = next(
                (entry for entry in entries if device_id in entry["devices"] and "CPU by type" in entry["property"]),
                None,
            )
            if cpu_entry:
                output_data[customer].append({
                    "device_id": device_id,
                    "cpu": cpu_entry["description"],
                })

# Sort the output_data dictionary by dealership names (keys)
sorted_output_data = dict(sorted(output_data.items()))

# Write the output to an Excel file
workbook = Workbook()
sheet = workbook.active
sheet.title = "Windows 10 Audit"

# Write the headers
sheet.append(["Dealership Name", "PC Name", "CPU"])

# Write the sorted data
for customer, devices in sorted_output_data.items():
    for device in devices:
        sheet.append([customer, device["device_id"], device["cpu"]])

# Save the Excel file
workbook.save("Windows10_audit.xlsx")



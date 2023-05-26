import csv
import sys
import tkinter as tk
from tkinter import filedialog

def transform_csv(input_file, output_file):
    with open(input_file, 'r') as inp, open(output_file, 'w', newline='') as out:
        reader = csv.reader(inp)
        writer = csv.writer(out)

        headers = next(reader, None)  
        if headers:
            new_headers = ['First Name', 'Last Name', 'E-mail Address', 'Business Phone', 'Home Phone', 'Business Fax', 
                           'Mobile Phone', 'Home Street', 'Home City', 'Home State', 'Business Street', 'Business Street 2', 
                           'Business City', 'Business State', 'Business Postal Code', 'Business Country/Region', 'Job Title', 
                           'Company', 'Web Page', 'Notes']  # Outlook expected headers
            writer.writerow(new_headers)

        for row in reader:
            name_parts = row[headers.index('Display Name')].split(' ', 1)
            first_name = name_parts[0] if name_parts else ''
            last_name = name_parts[1] if len(name_parts) > 1 else ''
            
            notes = "\n".join([row[headers.index(field)] for field in ['Custom 1', 'Custom 2', 'Custom 3', 'Custom 4', 'Notes'] if field in headers])

            new_row = [first_name, last_name, row[headers.index('Primary Email')], row[headers.index('Work Phone')], 
                       row[headers.index('Home Phone')], row[headers.index('Fax Number')], row[headers.index('Mobile Number')], 
                       row[headers.index('Home Address')], row[headers.index('Home City')], row[headers.index('Home State')], 
                       row[headers.index('Work Address')], row[headers.index('Work Address 2')], row[headers.index('Work City')], 
                       row[headers.index('Work State')], row[headers.index('Work ZipCode')], row[headers.index('Work Country')], 
                       row[headers.index('Job Title')], row[headers.index('Organization')], row[headers.index('Web Page 1')], 
                       notes]
            writer.writerow(new_row)

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()

    input_file = filedialog.askopenfilename(title="Select the input CSV file",
                                            filetypes=(("CSV Files", "*.csv"), ("All Files", "*.*")))

    if not input_file: 
        print("No input file chosen. Exiting.")
        sys.exit(1)

    output_file = filedialog.asksaveasfilename(defaultextension=".csv",
                                               filetypes=(("CSV Files", "*.csv"), ("All Files", "*.*")),
                                               title="Choose the output CSV file")

    if not output_file: 
        print("No output file chosen. Exiting.")
        sys.exit(1)

    transform_csv(input_file, output_file)

import csv
import sys  
import tkinter as tk
from tkinter import filedialog


def transform_csv(input_file, output_file):
    with open(input_file, 'r') as inp, open(output_file, 'w', newline='') as out:
        reader = csv.reader(inp)
        writer = csv.writer(out)

        headers = next(reader, None)  # get the headers from input
        if headers:
            # rearrange the header for Outlook
            new_headers = ['First Name', 'Last Name', 'E-mail Address']  # Outlook expected headers
            writer.writerow(new_headers)

        for row in reader:
            # split Display Name into first and last name
            name_parts = row[headers.index('Display Name')].split(' ', 1)
            first_name = name_parts[0] if name_parts else ''
            last_name = name_parts[1] if len(name_parts) > 1 else ''

            # rearrange the columns
            new_row = [first_name, last_name, row[headers.index('Primary Email')]]
            writer.writerow(new_row)


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # this will hide the main window

    # open dialog to choose input file
    input_file = filedialog.askopenfilename(title="Select the input CSV file",
                                            filetypes=(("CSV Files", "*.csv"), ("All Files", "*.*")))

    if not input_file:  # if no file is chosen
        print("No input file chosen. Exiting.")
        sys.exit(1)

    # open dialog to choose output file
    output_file = filedialog.asksaveasfilename(defaultextension=".csv",
                                               filetypes=(("CSV Files", "*.csv"), ("All Files", "*.*")),
                                               title="Choose the output CSV file")

    if not output_file:  # if no file is chosen
        print("No output file chosen. Exiting.")
        sys.exit(1)

    # call the function with the chosen files
    transform_csv(input_file, output_file)

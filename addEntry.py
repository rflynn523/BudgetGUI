# Responsible for getting the use inputs and correctly adding it to the entry table
import tkinter as tk

# Get the open row and fill in the columns (Date, Item, Vendor, Amount, Cat)
def add_entry():
    window = tk.Tk()
    # Labels for the new entry window.
    tk.Label(window, text="Date:").grid(row=0, padx=5, pady=5)
    tk.Label(window, text="Item:").grid(row=1, padx=5, pady=5)
    tk.Label(window, text="Vendor:").grid(row=2, padx=5, pady=5)
    tk.Label(window, text="Amount:").grid(row=3, padx=5, pady=5)
    tk.Label(window, text="Category").grid(row=4, padx=5, pady=5)

    # Get the correct sheet
    # sheet = wb['Monthly']

    # Find the next open row
    open_row = 6
    # while(sheet.cell(row= open_row, column=4).value is not None):
    #     open_row += 1

    # Write the collected data into their respective columns

import tkinter as tk
from tkinter import messagebox
import pandas as pd

# Function to read data from Excel
def read_data(file_path):
    try:
        data = pd.read_excel('machines.xlsx')
        return data
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

# Function to search for location and description based on Material number
def search_material(material_number, data):
    try:
        # Perform case-insensitive search for Material number
        filtered_data = data[data['material_number'].astype(str).str.contains(material_number, case=False, na=False)]

        if len(filtered_data) == 0:
            return "Material not found."

        # Prepare result string with location and description
        result = ""
        for index, row in filtered_data.iterrows():
            result += f"Material Number: {row['material_number']}\n"
            result += f"Location: {row['location']}\n"
            result += f"Description: {row['material_description']}\n\n"

        return result.strip()
    except Exception as e:
        print(f"Error in search: {e}")
        return "Error occurred during search."

# Function to handle button click and perform search
def perform_search():
    material_number = entry.get().strip()  # Get Material number from entry widget
    if material_number == "":
        messagebox.showerror("Error", "Please enter a Material number.")
        return
    
    # Assuming 'data' is the DataFrame containing your Materialry data
    file_path = 'machines.xlsx'
    data = read_data(file_path)

    if data is not None:
        result = search_material(material_number, data)
        messagebox.showinfo("Search Results", result)
    else:
        messagebox.showerror("Error", "Failed to read data from Excel file.")

# Create main window
root = tk.Tk()
root.title("Material Locator")

# Create input field
label = tk.Label(root, text="Enter Material Number:")
label.pack()
entry = tk.Entry(root)
entry.pack()

# Create search button
search_button = tk.Button(root, text="Search", command=perform_search)
search_button.pack()

# Run the main loop
root.mainloop()

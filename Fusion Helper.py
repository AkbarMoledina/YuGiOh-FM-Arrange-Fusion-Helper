import pandas as pd
from itertools import combinations
import tkinter as tk
from tkinter import ttk

# Excel file path and sheet name
file_path = "Yu-Gi-Oh! FMA Game Data.xlsx"
sheet_name = "Fusion List by Result"

# Read data into a DataFrame
df = pd.read_excel(file_path, sheet_name=sheet_name)
columns_to_extract = ["#", "CARD", "ATK", "Type A", "Type B"]
extracted_data = df[columns_to_extract]

# Get all unique types
all_types = sorted(set(extracted_data["Type A"]).union(extracted_data["Type B"]))

# Create a Tkinter window
window = tk.Tk()
window.title("Type Selection")

selected_types = []


# Function to handle type selection
def select_type(type):
    if type in selected_types:
        selected_types.remove(type)
    else:
        selected_types.append(type)


# Create checkboxes for type selection
type_checkboxes = []
num_columns = 4
num_rows = -(-len(all_types) // num_columns)
font_size = 16
window.style = ttk.Style()
window.style.configure("Custom.TCheckbutton", font=("Arial", font_size, "bold"))

for i, cardType in enumerate(all_types):
    checkbox_var = tk.BooleanVar()
    checkbox = ttk.Checkbutton(window, text=cardType, variable=checkbox_var, command=lambda t=cardType: select_type(t),
                               style="Custom.TCheckbutton")
    checkbox.grid(row=i // num_columns, column=(i % num_columns) * 2, sticky="w", padx=10)
    type_checkboxes.append(checkbox_var)


def close_window():
    window.destroy()


# Button to close window
finish_button = ttk.Button(window, text="Finish Selection", command=close_window)
finish_button.grid(row=num_rows, column=0, columnspan=2*num_columns, pady=10)

# Center the window on the screen
window.update_idletasks()
window_width = window.winfo_width()
window_height = window.winfo_height()
screen_width = window.winfo_screenwidth()
screen_height = window.winfo_screenheight()
x = (screen_width - window_width) // 2
y = (screen_height - window_height) // 2
window.geometry(f"+{x}+{y}")

window.mainloop()

# Filter the data based on the selected types
filtered_data = extracted_data.copy()
combinations_to_check = []
for r in range(2, len(selected_types) + 1):
    combinations_to_check.extend(combinations(selected_types, r))

matched_rows = pd.Series(False, index=filtered_data.index)

for type_combination in combinations_to_check:
    filters = []
    for t in type_combination:
        filters.append((filtered_data['Type A'] == t) | (filtered_data['Type B'] == t))
    matched_rows = matched_rows | pd.concat(filters, axis=1).all(axis=1)

filtered_data = filtered_data[matched_rows]

# Filter out the types not mentioned in the user input from both "Type A" and "Type B" columns
filtered_data = filtered_data[
    filtered_data["Type A"].isin(selected_types) & filtered_data["Type B"].isin(selected_types)
]

# Sort the filtered data by ATK value
sorted_data = filtered_data.sort_values("ATK", ascending=False)

# Display the results in columns: Type A, Type B, ATK
if not sorted_data.empty:
    corresponding_values = sorted_data.loc[:, ["Type A", "Type B", "ATK"]]
    print(corresponding_values)
else:
    print("No matches")

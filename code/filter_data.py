# # *** 1. Import the required modules
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# *** 2. Read the Excel sheet using pandas:
# Load the employee Excel file into a DataFrame
df = pd.read_excel("./data/sample.xlsx")

# Initialize an empty list to store the rows that match the condition
matching_rows = []

# Iterate through the rows of the DataFrame
for index, row in df.iterrows():
    # Check if the value in the 'SubSec' column contains 'Sewing Line'
    if "Sewing Line" in row["SubSec"] and "Operator" in row["Designation"]:
        # Check the value in the 'Designation' column
        if row["Designation"] == "Asst. Operator - Sewing M/C":
            post = "helper"
        else:
            post = "operator"

        # Create a new row with the 'Desig' column set to the post value
        new_row = row.copy()
        new_row["Desig"] = post

        # Add the new row to the list of matching rows
        matching_rows.append(new_row)

# Create a new DataFrame from the matching rows
new_df = pd.DataFrame(matching_rows)

# Print the new DataFrame
print(new_df)

new_df.to_excel("./data/filtered_data.xlsx", index=False)

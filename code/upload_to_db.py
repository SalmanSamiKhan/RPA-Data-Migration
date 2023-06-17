# #*** 1. Import the required modules
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# *** 2. Read the Excel sheet using pandas:

# Specify the path to the Excel file
excel_file = "./data/filteredData.xlsx"

# Read the Excel sheet into a DataFrame
df = pd.read_excel(excel_file)

# *** 3. Set up the Selenium WebDriver:

# Prevent closing the browser immediately

options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)
options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=options)


# *** 4. Iterate over the DataFrame and input values into the website form:
for index, row in df.iterrows():
    # Go to insert page
    driver.get(
        "http://localhost/phpmyadmin/index.php?route=/table/change&db=emp&table=employee"
    )

    # Get data from df
    id = row["ID"]
    name = row["Name"]
    desig = row["Desig"]

    # Get id field
    id_field = driver.find_element(By.XPATH, "//input[@id='field_1_3']")
    id_field.clear()
    id_field.send_keys(id)

    # Get name field
    name_field = driver.find_element(By.XPATH, "//input[@id='field_2_3']")
    name_field.clear()
    name_field.send_keys(str(name))

    # Get designation field
    desig_field = driver.find_element(By.XPATH, "//input[@id='field_3_3']")
    desig_field.clear()
    desig_field.send_keys(str(desig))

    # Save to database
    submit_button = driver.find_element(
        By.XPATH, "//div[1]//table[1]//tfoot[1]//tr[1]//th[1]//input[1]"
    )
    submit_button.send_keys(Keys.ENTER)

# View table
driver.get(
    "http://localhost/phpmyadmin/index.php?route=/sql&server=1&db=emp&table=employee&pos=0"
)

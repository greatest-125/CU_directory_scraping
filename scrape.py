import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# Set the path to ChromeDriver
CHROMEDRIVER_PATH = "Your path here"
# Load the Excel file
file_path = "Your path here"
faculty_df = pd.read_excel(file_path)

# Function to extract obfuscated email
def extract_email(encoded_email):
    # Decode obfuscated email
    decoded_email = encoded_email.replace('&#64;', '@').replace('&#46;', '.').replace('mailto:', '')
    return decoded_email

# Iterate through each professor and fetch their details
for index, row in faculty_df.iterrows():
    professor_name = row['Professor']
    
    # Navigate to search results
    driver.get(f"https://directory.columbia.edu/people/search?filter.searchTerm={professor_name.replace(' ', '+')}")
    time.sleep(2)
    
    try:
        # Check if multiple results are found
        result_info = driver.find_element(By.CLASS_NAME, "search_return_info").text
        if "returned" in result_info and "Results" in result_info:
            num_results = int(result_info.split("returned ")[1].split(" Results")[0])
            if num_results > 1:
                print(f"Multiple results found for {professor_name}")
                faculty_df.loc[index, 'Email'] = 'Multiple Found'
                continue

        # Find the first search result
        first_result = driver.find_element(By.PARTIAL_LINK_TEXT, professor_name.split(' ')[-1])
        first_result.click()
        time.sleep(2)
        
        # Extract obfuscated email
        email_element = driver.find_element(By.CSS_SELECTOR, "a[href^='mailto:']")
        encoded_email = email_element.get_attribute("href")
        decoded_email = extract_email(encoded_email)
        
        # Save email to the DataFrame
        faculty_df.loc[index, 'Email'] = decoded_email
        
    except Exception as e:
        print(f"Error for {professor_name}: {e}")
        faculty_df.loc[index, 'Email'] = 'Not Found'
    
    time.sleep(2)

# Close the browser
driver.quit()

# Save the updated DataFrame back to Excel
faculty_df.to_excel(file_path, index=False)
print("Emails updated successfully!")

import time  # Used for adding sleep delays
import pandas as pd  # Used for reading and writing Excel files
from selenium import webdriver  # Provides the web driver to control the browser
from selenium.webdriver.chrome.service import Service  # Manages the Chrome driver service
from selenium.webdriver.chrome.options import Options  # Allows setting options for Chrome
from selenium.webdriver.common.by import By  # Enables using different locating strategies for elements

# Function to search for the annual report URL of the company using Google
def search_annual_report_url(company_name, driver):
    # Open Google in the web driver
    driver.get("https://www.google.com")
    
    # Locate the search box by its name attribute and enter the search query
    search_box = driver.find_element(By.NAME, "q")
    query = f"{company_name} annual report"  # Query for the company's annual report
    search_box.send_keys(query)  # Input query into the search box
    search_box.submit()  # Submit the search form
    
    # Wait for the search results page to load
    time.sleep(6)
    
    # Locate the first result and retrieve its URL
    first_result = driver.find_element(By.CSS_SELECTOR, "div.g a")
    url = first_result.get_attribute("href")
    return url  # Return the URL of the first result

# Function to extract PDF links from the annual report webpage
def extract_pdf_links(report_url, driver):
    # Open the report URL in the web driver
    driver.get(report_url)
    
    # Wait for the page to load completely
    time.sleep(5)  # Increased wait time to ensure the page is fully loaded
    
    # Scroll to the bottom of the page to ensure all elements are loaded
    driver.execute_script('window.scrollTo(0, document.body.scrollHeight);')
    
    # Wait for the scrolling to complete
    time.sleep(3)
    
    # JavaScript to find all links that end with .pdf on the page
    js_command = '''
    var link_elements = document.querySelectorAll("a[href$='.pdf']");
    var pdf_links = [];
    for (var i = 0; i < link_elements.length; i++) {
        pdf_links.push(link_elements[i].href);
    }
    return pdf_links;
    '''
    
    try:
        # Execute the JavaScript to extract all PDF links
        pdf_links = driver.execute_script(js_command)
        if not pdf_links:
            print("No PDF links found on the page.")
    except Exception as e:
        print(f"Error extracting PDF links: {e}")
        pdf_links = []  # Return an empty list in case of an error
    return pdf_links  # Return the list of PDF links

# Function to save the extracted PDF links to an Excel file
def save_pdf_links_to_excel(company_name, pdf_links):
    # Create the output file name based on the company name
    output_file = f"{company_name}_url_links.xlsx"
    
    # Convert the list of PDF links to a DataFrame
    df = pd.DataFrame(pdf_links, columns=["PDF Links"])
    
    # Save the DataFrame to an Excel file
    df.to_excel(output_file, index=False)
    print(f"PDF links saved to {output_file}")  # Inform the user that the file has been saved

# Main function that reads company names from an input Excel file and processes each company
def main(input_filename):
    # Set up Chrome options for the Selenium driver
    chrome_options = Options()
    # chrome_options.add_argument("--headless")  # Comment out to run in non-headless mode (i.e., visible mode)
    
    # Initialize the Chrome driver service with the path to the ChromeDriver executable
    service = Service('C:\\windows\\chromedriver.exe')
    
    # Create a Selenium WebDriver object
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    company_urls = []  # List to store company names and their corresponding URLs
    
    try:
        # Read the input Excel file containing the list of company names
        df = pd.read_excel(input_filename)
        
        # Loop through each company in the Excel file
        for index, row in df.iterrows():
            company_name = row['Company Name']  # Extract the company name from the row
            print(f"Searching for: {company_name}")  # Inform the user which company is being processed
            
            try:
                # Search for the annual report URL of the company
                url = search_annual_report_url(company_name, driver)
                print(f"Found URL: {url}")  # Inform the user of the found URL
                
                # Extract PDF links from the company's annual report webpage
                pdf_links = extract_pdf_links(url, driver)
                
                # Save the extracted PDF links to an Excel file
                save_pdf_links_to_excel(company_name, pdf_links)
                
                # Add the company name and URL to the list
                company_urls.append((company_name, url))
            except Exception as e:
                # Handle errors and inform the user if something goes wrong
                print(f"Failed to retrieve data for {company_name}: {e}")
                company_urls.append((company_name, "Failed to retrieve URL or PDF links"))
    finally:
        # Ensure the web driver is closed after processing is completed
        driver.quit()
    
    # Save the company names and URLs to an output Excel file
    company_url_output_file = 'company_urls_output1.xlsx'
    df_urls = pd.DataFrame(company_urls, columns=["Company Name", "Website URL"])
    df_urls.to_excel(company_url_output_file, index=False)
    print(f"Company URLs saved to {company_url_output_file}")  # Inform the user that the URLs have been saved

# Entry point for the script
if __name__ == '__main__':
    # Specify the input Excel file containing the list of company names
    input_excel_file = 'input_companies_list.xlsx'
    
    # Call the main function with the input file
    main(input_excel_file)


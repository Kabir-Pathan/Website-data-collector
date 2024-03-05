import pypyodbc as odbc
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import InvalidArgumentException
from selenium.webdriver.chrome.options import Options


DriverName = 'SQL Server'
ServerName = '192.168.0.207\CRM2017'
DatabaseName = 'DownloadedSites'
Username = ''
Password = ''

connectionString = f"""
Driver={{{DriverName}}};
Server={ServerName};
Database={DatabaseName};
UID={Username};
PWD={Password};
"""

conn = odbc.connect(connectionString)

cursor = conn.cursor()

#Get first 10 URLs
cursor.execute("SELECT TOP 10 WebsiteURL FROM Websites")

result = cursor.fetchall()

#Run in headless mode
chrome_options = Options()
chrome_options.headless = True
driver = webdriver.Chrome(options=chrome_options)


output_file = "websitedata.txt"  

try:
    with open(output_file, "w") as file:
        for website in result:
            url = website[0]
            file.write(f"Website: {url}\n")
            try:
                #Check if site is redirected
                driver.get(url)
                current_url = driver.current_url
                if url == current_url:
                    file.write("Requested URL and opened URL are the same\n")
                else:
                    file.write("Requested URL and opened URL are not the same\n")

                #Get status code
                response = requests.get(url)
                status_code = response.status_code

                #Get title
                title = driver.title
                description_element = driver.find_element(By.CSS_SELECTOR, "meta[name='description']")
                description = description_element.get_attribute("content")

                #Get body    
                body = driver.find_element(By.TAG_NAME, "body")
                body_text = body.text

                #Get links with anchor tags
                links = driver.find_elements(By.TAG_NAME, "a")
                link_texts = [link.text for link in links]
                link_urls = [link.get_attribute("href") for link in links]

                #Get all headings
                headings = driver.find_elements(By.CSS_SELECTOR, "h1, h2, h3, h4, h5, h6")
                heading_texts = [heading.text for heading in headings]

                #Write everything to .txt file
                file.write(f"URL: {url}\n")
                file.write(f"Status code: {status_code}\n")
                file.write(f"Title: {title}\n")
                file.write(f"Body: {body_text}\n")
                file.write(f"Description: {description}\n")
                file.write(f"Links: {link_texts}\n")
                file.write(f"Headings: {heading_texts}\n")
                file.write("\n")
            
            #Handle exception
            except InvalidArgumentException:
                file.write(f"Skipping URL '{url}' due to element not found error.\n")
    
            finally:
                driver.quit()
                print("Moving to next URL")         #This and the end print statement is for my convenience
finally:
    cursor.close()
    conn.close()
    print("Loop is complete")
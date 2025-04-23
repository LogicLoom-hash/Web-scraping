import json
import time
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

#set location the location of the webdriver
s = Service(r"C:\chromedriver.exe")

driver = webdriver.Chrome(service=s)
driver.get("https://cmmiinstitute.com/pars?StateId=1936e21e-c4ca-4113-91b9-c176a1013315&PageNumber=1&Handler=ApplyFilters")

# Interact with dropdown options
dropdownbox = driver.find_elements(By.TAG_NAME, "option")

i = 0
while i < len(dropdownbox):
    if(dropdownbox[i].text == "United States"):
        dropdownbox[i].click()
    i = i + 1

# Click the "APPLY" button
apply_filter = driver.find_element(By.XPATH, "//button[contains(text(), 'APPLY')]")
apply_filter.click()

all_card_data = []

# Loop through all pages until there is no next page
while True:
    time.sleep(2)
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    item_cards = soup.find_all('div', class_='item-card')

    for card in item_cards:
        card_data = {}

        # Extract Organization Name
        org_name = card.find('div', class_='item-card__title appraisal-card__org-names').get_text(strip=True)
        card_data['Organization Name'] = org_name

        # Extract ID
        card_id = card.find('div', class_='appraisal-card__id')
        card_data['ID'] = card_id.get_text(strip=True).replace("ID:", "").strip() if card_id else ""

        # Extract Appraisal Team Leader
        leader_section = card.find('h3', string='Appraisal Team Leader')
        card_data['Appraisal Team Leader'] = leader_section.find_next_sibling('small').get_text(strip=True) if leader_section else ""

        # Extract Sponsors
        sponsors_section = card.find('h3', string='Sponsors')
        if sponsors_section:
            small_tags = sponsors_section.find_all_next('small')
            for tag in small_tags:
                sponsor_name = tag.get_text(strip=True)
                if sponsor_name:
                    card_data['Sponsors'] = sponsor_name
                    break
        else:
            card_data['Sponsors'] = ""

        # Extract Partner
        partner_section = card.find('h3', string='Partner')
        card_data['Partner'] = partner_section.find_next('small').get_text(strip=True) if partner_section else ""

        # Extract OU Name
        ou_name_section = card.find('div', class_='appraisal-card__target')
        card_data['OU Name'] = ou_name_section.find_next('small').get_text(strip=True) if ou_name_section else ""

        # Extract Appraisal Validity
        validity_section = card.find('h3', string='Appraisal Validity')
        card_data['Appraisal Validity'] = validity_section.find_next('small').get_text(strip=True) if validity_section else ""


        # Extract Model View / Domain and Level from the row
        rows = card.find_all('div', class_='row')
        if len(rows) >= 2:
            # Extract Model View / Domain
            model_view_row = rows[3]
            col5 = model_view_row.find('div', class_='col5')
            model_view_div = col5.find('div', class_='appraisal-card__target') if col5 else None
            card_data['Model View / Domain'] = model_view_div.find('small').get_text(strip=True) if model_view_div and model_view_div.find('small') else ""

            # Extract Level
            level_row = rows[3] if len(rows) > 3 else None
            if level_row:
                col7 = level_row.find('div', class_='col7')
                level_div = col7.find('div', class_='appraisal-card__target') if col7 else None
                card_data['Level'] = level_div.find('small').get_text(strip=True) if level_div and level_div.find('small') else ""


        all_card_data.append(card_data)



    # Check if there's a next page and navigate to it
    next_page_button = driver.find_elements(By.XPATH, "//a[@class='button next']")
    if next_page_button:
        driver.execute_script("arguments[0].click();", next_page_button[0])
        time.sleep(2)  # Add some delay for the page to load
    else:
        break  # No more pages

driver.quit()

# Convert the list of dictionaries to a pandas DataFrame
df = pd.DataFrame(all_card_data)

# Save the DataFrame to an Excel file
df.to_excel('cmmi_data.xlsx', index=False)

print("Data extracted from all pages and saved to 'cmmi_data.xlsx'")

#loop the filter after every page


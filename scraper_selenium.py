import time

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

from bs4 import BeautifulSoup
import requests

import pandas as pd

## Initialize Selenium
web = "https://bcres.paragonrels.com/paragonls/publink/view.mvc/?GUID=43125b63-6c9a-457f-a8a5-725d467dc1f9"
path = "./chromedriver-win64/chromedriver.exe"
service = Service(executable_path=path)
driver = webdriver.Chrome(service=service)

## Link Variables
link_part_1 = 'https://bcres.paragonrels.com/ParagonLS/publink/view.mvc/Report?outputtype=HTML&GUID=43125b63-6c9a-457f-a8a5-725d467dc1f9&ListingID='
link_part_2 = "&Report=Yes&view=29&layout_id=63&screenWidth=1162"

driver.get(web)

elements = driver.find_elements(By.XPATH, "/html/frameset/frameset/frame")
driver.switch_to.frame(elements[0])

start_idx = 3

main_body = driver.find_element(By.TAG_NAME, "body")
tr_list = main_body.find_elements(By.TAG_NAME, "tr")

house_id_list = []
row_number = 3

for i in range(start_idx, len(tr_list), 3):
    td = tr_list[i].find_elements(By.TAG_NAME, "td")
    a_tag = td[0].find_element(By.TAG_NAME, "a")
    house_id = a_tag.get_attribute("onclick")
    house_id = house_id[12:].replace(f"','Row{row_number}'); return false;", "")
    row_number += 4
    house_id_list.append(house_id)


column_value_index = {
    "HouseType": 393,
    "Address": 159,
    "PostalCode": 160,
    "SalePrice": 163,
    "Bedrooms": 168,
    "Bathrooms": 169,
    "YearBuilt": 172,
    "Age": 173,
    "Zoning": 174,
    "GrossTaxes": 175,
    "LotAreaSqft": 166,
    "FloodPlain": 289,
    "View": 288,
    "FrontageFt": 167,
    "DepthSizeFt": 165,
    "Parking": 380,
    "FinishedFloorTotalSqft": 239,
    "Suite": 381,
    "Basement": 386,
    "DistToTransit": 269,
    "DistToSchoolBus": 270,
    "SoldDate": 266
}

columns = []

for key, value in column_value_index.items():
    columns.append(key)

df = pd.DataFrame(columns = columns)

for house in house_id_list:
    full_link = link_part_1 + house + link_part_2
    print("Link: ", full_link)
    page = requests.get(full_link)
    time.sleep(0.3)
    soup = BeautifulSoup(page.content, 'html.parser')

    house_info = []

    main_div = soup.find_all("div", {"id": "divHtmlReport"})
    for key, value in column_value_index.items():

        div = main_div[0].find_all("div", {"tabindex": value})
        div_text = div[0].text
        if div_text == "":
            div_text = "N/A"

        house_info.append(div_text)

    df.loc[len(df)] = house_info




print(df)
df.to_excel("house_price.xlsx", index=False)
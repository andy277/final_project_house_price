import time
from dataclasses import dataclass
from typing import Dict, List

import pandas as pd
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By


@dataclass
class ScraperConfig:
    """Configuration settings for the house data scraper."""
    base_url: str = "https://bcres.paragonrels.com/paragonls/publink/view.mvc"
    guid: str = "43125b63-6c9a-457f-a8a5-725d467dc1f9"
    chrome_driver_path: str = "./chromedriver-win64/chromedriver.exe"
    start_row_index: int = 3
    row_increment: int = 3
    request_delay: float = 0.3
    output_file: str = "house_price.csv"


class HouseDataScraper:
    """Class to handle scraping of house listing data."""

    COLUMN_VALUE_INDEX = {
        "HouseType": 393,
        "Address": 159,
        "PostalCode": 160,
        "SalePrice": 163,
        "Bedrooms": 168,
        "Bathrooms": 169,
        "YearBuilt": 172,
        "Age": 173,
        "LotAreaSqft": 166,
        "FloodPlain": 289,
        "View": 288,
        "FrontageFt": 167,
        "DepthSizeFt": 165,
        "Parking": 380,
        "FinishedFloorTotalSqft": 239,
        "Suite": 381,
        "Basement": 386,
        "SoldDate": 266,
        "NumberOfLevels": 254,
        "NumberOfRooms": 253
    }

    def __init__(self, config: ScraperConfig):
        self.config = config
        self.driver = self._setup_selenium()

    def _setup_selenium(self) -> webdriver.Chrome:
        """Initialize and configure the Selenium WebDriver."""
        service = Service(executable_path=self.config.chrome_driver_path)
        return webdriver.Chrome(service=service)

    def _get_listing_url(self, listing_id: str) -> str:
        """Generate the full URL for a specific listing."""
        return (f"{self.config.base_url}/Report?outputtype=HTML&"
                f"GUID={self.config.guid}&ListingID={listing_id}&"
                f"Report=Yes&view=29&layout_id=63&screenWidth=1162")

    def _extract_house_ids(self) -> List[str]:
        """Extract house IDs from the main listing page."""
        initial_url = f"{self.config.base_url}/?GUID={self.config.guid}"
        self.driver.get(initial_url)

        # Switch to the correct frame
        elements = self.driver.find_elements(By.XPATH, "/html/frameset/frameset/frame")
        self.driver.switch_to.frame(elements[0])

        # Find all listing rows
        main_body = self.driver.find_element(By.TAG_NAME, "body")
        tr_list = main_body.find_elements(By.TAG_NAME, "tr")

        house_ids = []
        row_number = 3

        for i in range(self.config.start_row_index, len(tr_list), self.config.row_increment):
            td = tr_list[i].find_elements(By.TAG_NAME, "td")
            a_tag = td[0].find_element(By.TAG_NAME, "a")
            house_id = a_tag.get_attribute("onclick")
            house_id = house_id[12:].replace(f"','Row{row_number}'); return false;", "")
            row_number += 4
            house_ids.append(house_id)

        return house_ids

    def _scrape_house_details(self, house_id: str) -> List[str]:
        """Scrape details for a specific house listing."""
        url = self._get_listing_url(house_id)
        print(f"Scraping: {url}")

        page = requests.get(url)
        time.sleep(self.config.request_delay)
        soup = BeautifulSoup(page.content, 'html.parser')

        house_info = []
        main_div = soup.find_all("div", {"id": "divHtmlReport"})

        for value in self.COLUMN_VALUE_INDEX.values():
            div = main_div[0].find_all("div", {"tabindex": value})
            div_text = div[0].text if div[0].text else "N/A"
            house_info.append(div_text)

        return house_info

    def scrape_and_save(self):
        """Main method to scrape all listings and save to Excel."""
        try:
            # Get all house IDs
            house_ids = self._extract_house_ids()

            # Create DataFrame
            df = pd.DataFrame(columns=self.COLUMN_VALUE_INDEX.keys())

            # Scrape details for each house
            for house_id in house_ids:
                house_info = self._scrape_house_details(house_id)
                df.loc[len(df)] = house_info

            # Save to Excel
            df.to_csv(self.config.output_file, index=False)
            print(f"Data successfully saved to {self.config.output_file}")

        finally:
            self.driver.quit()


def main():
    """Main entry point for the scraper."""
    config = ScraperConfig()
    scraper = HouseDataScraper(config)
    scraper.scrape_and_save()


if __name__ == "__main__":
    main()
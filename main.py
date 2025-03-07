import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
import re
import os

class BarbadosPropertyScraper:
    def __init__(self):
        self.base_url = 'https://www.barbadospropertysearch.com'
        self.start_url = 'https://www.barbadospropertysearch.com/for-rent'
        self.data = []
        self.scraped_urls = set()
        self.transaction_type = self.determine_transaction_type()  # Determine transaction type at initialization

    def determine_transaction_type(self):
        """Determines the transaction type based on the start URL."""
        if 'sale' in self.start_url.lower():
            return 'Sale'
        elif 'rent' in self.start_url.lower():
            return 'Rent'
        else:
            return None

    def scrape(self):
        # Scrape the first page and then follow pagination links
        self.scrape_page(self.start_url)
        
        print('Data has been collected, saving to Excel...')
        self.save_to_excel()

    def scrape_page(self, url):
        print("Scraping URL: ", url)
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')

        # Extract property URLs and avoid duplicates using a set
        listing_elements = soup.select('div.field-item.even h5 a')
        for element in listing_elements:
            property_url = self.base_url + element['href']
            if property_url not in self.scraped_urls:
                self.scraped_urls.add(property_url)
                print(f"Total unique property URLs on this page: {len(self.scraped_urls)}")
                self.scrape_listing(property_url)

        # Find the next page link
        next_page_element = soup.select_one('li.pager-next a')
        if next_page_element:
            next_page_url = self.base_url + next_page_element['href']
            self.scrape_page(next_page_url)

    def scrape_listing(self, url):
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')

        # Initialize the more_info dictionary
        more_info = {}

        # Extract rental price
        rental_price = soup.select_one('div.field-name-rental-price .field-item')
        rental_price = rental_price.get_text(strip=True) if rental_price else '-'
        more_info['Rental Price'] = rental_price

        # Extract property reference
        property_reference = soup.select_one('div.field-name-field-property-reference .field-item')
        property_reference = property_reference.get_text(strip=True) if property_reference else '-'
        more_info['Property Reference'] = property_reference

        sale_price_div = soup.select_one('div.field-name-sale-price .field-item')
        sale_price = sale_price_div.get_text(strip=True) if sale_price_div else '-'
        more_info['Sale Price'] = sale_price

        # Extract external link
        external_link_tag = soup.select_one('div.field-name-external-url-ds- .field-item a')
        external_link = external_link_tag['href'] if external_link_tag else '-'
        more_info['External Link'] = external_link

        # Extract additional fields if needed
        name = soup.select_one('meta[property="og:title"]')
        name = name['content'] if name else '-'

        latitude = soup.select_one('meta[property="og:latitude"]')
        latitude = latitude['content'] if latitude else '-'

        longitude = soup.select_one('meta[property="og:longitude"]')
        longitude = longitude['content'] if longitude else '-'

        # Extract address with fallback to new div structure if not found
        address = soup.select_one('a[href*="maps.google.com"]')
        if address:
            address = address.get_text(strip=True)
        else:
            # Fallback address extraction
            fallback_address = soup.select_one('div.field-name-location-ds- .field-item')
            address = fallback_address.get_text(strip=True) if fallback_address else '-'

        # Extract property type
        property_type = soup.select_one('div.field-name-field-property-type .field-item')
        property_type = property_type.get_text(strip=True) if property_type else '-'

        # Use the transaction type already determined at initialization
        transaction_type = self.transaction_type

        # Extract description
        description_div = soup.find('div', class_='field field-name-body field-type-text-with-summary field-label-hidden')
        description = description_div.get_text(strip=True) if description_div else '-'

        # Extract characteristics data
        characteristics = {}
        further_info_div = soup.select_one('div.group-further-information')
        if further_info_div:
            field_labels = further_info_div.select('.field-label')
            field_items = further_info_div.select('.field-item')
            for label, item in zip(field_labels, field_items):
                key = label.get_text(strip=True).replace(":", "")
                value = item.get_text(strip=True)
                characteristics[key] = value

        # Extract amenities data
        amenities_list = soup.select_one('div.field-name-field-amenities .field-items')
        amenities = [li.get_text(strip=True) for li in amenities_list.select('li')] if amenities_list else []

        # Store the extracted data
        property_data = {
            'url': url,
            'name': name,
            'address': address,
            'Sale Price': more_info.get("Sale Price", "-"),
            'Rent Price': more_info.get("Rental Price", "-"),
            'Area': characteristics.get('Land Area') or characteristics.get('Floor Area', 'N/A'),
            'description': description,
            'latitude': latitude,
            'longitude': longitude,
            'property_type': property_type,
            'transaction_type': transaction_type,
            'characteristics': characteristics,
            'amenities': amenities,
            'more_information': more_info,  # Add the more_info dictionary to property data
        }
        print(property_data)

        self.data.append(property_data)

    def sanitize_filename(self, url):
        """Sanitize the URL to make it a valid file name."""
        filename = re.sub(r'[^\w\-_.]', '_', url.replace('https://', '').replace('http://', ''))
        return filename

    def save_to_excel(self):
        """Save the scraped data to an Excel file in the output folder."""
        df = pd.DataFrame(self.data)
        sanitized_filename = self.sanitize_filename(self.start_url) + '.xlsx'
        
        # Create the output directory if it does not exist
        output_dir = os.path.join(os.getcwd(), 'output')
        os.makedirs(output_dir, exist_ok=True)
        
        file_path = os.path.join(output_dir, sanitized_filename)
        
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        print(f"File saved as: {file_path}")

# Run the scraper
if __name__ == "__main__":
    scraper = BarbadosPropertyScraper()
    scraper.scrape()

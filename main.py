from argparse import ArgumentError
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.chrome.options import Options
from openpyxl import Workbook

# Reference to stackoverflow questions:
# https://stackoverflow.com/questions/29858752/error-message-chromedriver-executable-needs-to-be-available-in-the-path
from webdriver_manager.chrome import ChromeDriverManager 

# Standard library imports
import configparser
import os
import csv
from datetime import date, datetime, timedelta
import time
import logging

logger = logging.getLogger('dragonpay-scaper')
logger.setLevel(logging.DEBUG)

# create console handler and set level to debug
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)

# create formatter
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

# add formatter to ch
ch.setFormatter(formatter)

# add ch to logger
logger.addHandler(ch)

current_dir = os.path.dirname(__file__)
dt = datetime.now()
chrome_options = Options()
chrome_options.add_argument('--headless')

def _read_config():
    """Returns and reads from the configuration object"""
    config = None
    if os.path.exists(os.path.join(current_dir, 'config.cfg')):
        config = configparser.ConfigParser()
        config.read('config.cfg')  
        return config
    else:
        pass
    
class DragonPayScraper(webdriver.Chrome):
    def __init__(self, driver_path=_read_config().get('selenium', 'driver_path'), quit=False, wait_timeout=10, headless=True):
        self.url = _read_config().get('dragonpay', 'url')
        self.quit = quit
        self.driver_path = driver_path
        self.wait_timeout = wait_timeout
        if headless:
            super().__init__(executable_path=self.driver_path, options=chrome_options)
        else:
            super().__init__(executable_path=self.driver_path)
        self.wait = WebDriverWait(self, self.wait_timeout)
        self.data = []

    def goto_url(self):
        logger.info("Accessing {} ... done".format(self.url))
        self.get(self.url)
        self.set_window_size(1500, 500)

    def login(self, id=_read_config().get('dragonpay', 'id'), 
                password=_read_config().get('dragonpay', 'password')):
        logger.info("Authenticating user {}".format(id))
        id_dom_elem = self.find_element(By.ID, 'ContentPlaceHolder1_UserId')
        id_dom_elem.clear()
        id_dom_elem.send_keys(id)
        pw_dom_elem = self.find_element(By.ID, 'ContentPlaceHolder1_UserPass')
        ActionChains(self, 10).send_keys_to_element(pw_dom_elem, password) \
            .send_keys_to_element(pw_dom_elem, Keys.ENTER).perform()

    def change_date_from_to(self, backdate=4):
        logger.info("Modifying dates ... done")
        from_dt_dom_elem = self.wait.until(
            expected_conditions.presence_of_element_located((By.ID, 'ContentPlaceHolder1_FromDateTextBox'))
        )
        from_dt_dom_elem.clear()
        if backdate < 0:
            self.close()
            raise ArgumentError("backdate: backdate cannot be negative")
        elif not backdate > 5:
            from_dt_dom_elem.send_keys(datetime.strftime(dt - timedelta(days=backdate), "%m/%d/%y"))
        else:
            print("You've reached the maximum number of days which is 4.")
            self.close()
        to_dt_dom_elem = self.wait.until(
            expected_conditions.presence_of_element_located(
                (By.ID, 'ContentPlaceHolder1_ToDateTextBox')
            )
        )
        to_dt_dom_elem.clear()
        to_dt_dom_elem.send_keys(datetime.strftime(dt, "%m/%d/%y"))

    def change_time_from_to(self):
        logger.info("Modifying time ... done")
        from_time = timedelta(hours=0, minutes=0, seconds=0)
        to_time = timedelta(hours=23, minutes=59, seconds=59)
        time_from_dom_elem = self.find_element(By.ID, 'ContentPlaceHolder1_TimeFromTextBox')
        time_from_dom_elem.clear()
        time_from_dom_elem.send_keys(str(from_time))
        time_to_dom_elem = self.find_element(By.ID, 'ContentPlaceHolder1_TimeToTextBox')
        time_to_dom_elem.clear()
        time_to_dom_elem.send_keys(str(to_time))
    
    def change_transaction_status(self):
        logger.info("Changing transaction status to 'Success' ... done")
        select_status_dom_elem = Select(self.find_element(By.ID, 'ContentPlaceHolder1_StatusList'))
        select_status_dom_elem.select_by_value("S")

    def change_date_type(self):
        logger.info("Changing date type to 'Settle' ... done")
        select_datetype_dom_elem = Select(self.find_element(By.ID, 'ContentPlaceHolder1_DateTypeList'))
        select_datetype_dom_elem.select_by_value("Settle")
        
    def scrape_data(self):
        logger.info("Scraping data from {}".format(self.url))
        headers = []
        row = []
        get_trans_dom_elem = self.find_element(By.ID, 'ContentPlaceHolder1_GetTxnBtn')
        get_trans_dom_elem.click()
        time.sleep(1)
        # Get transactions table
        logger.info("Extracting table element from the HTML dom ... done")
        table = self.find_element(By.ID, 'ContentPlaceHolder1_TxnGrid')
        # Get table headers
        logger.info("Extracting table header element from the HTML dom ... done")
        table_headers = table.find_elements(By.TAG_NAME, "th")
        for th in table_headers:
            if not th.text == " ":
                headers.append(th.text)
        # Get table rows
        logger.info("Extracting table data from the HTML dom ... done")
        for tds in table.find_elements(By.TAG_NAME, 'tr'):
            row.clear()
            for td in tds.find_elements(By.TAG_NAME, 'td'):
                if not td.text == "View":
                    row.append(td.text)
            self.data.append(list(row))
        self.data.remove([])
        self.data.insert(0, headers) # Append headers
        return self.data

    def data_to_csv(self):
        logger.info("Importing scraped data to csv 'dragonpay-data-{}.csv'.format(date.today()) ... done".format(date.today()))
        data = self.scrape_data()
        with open('dragonpay-data-{}.csv'.format(date.today()), 'w') as c:
            writer = csv.writer(c, delimiter=",", quotechar="'", lineterminator="\n")
            for row in data:
                writer.writerow(row)
            
    def data_to_excel(self):
        logger.info("Importing scraped data to csv 'dragonpay-data-{}.xlsx'.format(date.today()) ... done".format(date.today()))
        data = self.scrape_data()
        wb = Workbook()
        ws = wb.active
        for row in data:
            ws.append(row)
        wb.save("dragonpay-data-{}.xlsx".format(date.today()))
    
    def logout(self):
        pass

    def exit(self):
        logger.info("Exiting ...")
        if self.quit:
            self.close()
        else:
            pass

if __name__ == "__main__":
    d = DragonPayScraper(quit=True, headless=True)
    d.goto_url()
    d.login()
    d.change_date_from_to(backdate=0)
    d.change_time_from_to()
    d.change_transaction_status()
    d.change_date_type()
    d.data_to_csv()
    d.data_to_excel()
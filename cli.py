import argparse
from main import DragonPayScraper
import os

parser = argparse.ArgumentParser(add_help=True)
parser.add_argument("-o", "--output", help="Load scaped data to csv", metavar="", choices=["csv", "excel"], required=True)
parser.add_argument("-hl", "--headless", help="Run scraper in headless mode (default: False)", action='store_true', default=False)
parser.add_argument("-bd", "--backdate", help="", type=int, metavar="", default=0)
args = parser.parse_args()

if args.output == "excel":
    print("Extracting to excel ...")
    d = DragonPayScraper(quit=True, headless=args.headless)
    d.goto_url()
    d.login()
    d.change_date_from_to(backdate=args.backdate)
    d.change_time_from_to()
    d.change_transaction_status()
    d.change_date_type()
    d.data_to_excel()
    d.exit()
elif args.output == "csv":
    print("Extracting to csv ...")
    d = DragonPayScraper(quit=True, headless=args.headless)
    d.goto_url()
    d.login()
    d.change_date_from_to(backdate=args.backdate)
    d.change_time_from_to()
    d.change_transaction_status()
    d.change_date_type()
    d.data_to_csv()
    d.exit()
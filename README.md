# Product-Web-Scraper-Suite
Suite to pull information from SEETS, KEEPA and AMP.

All of the programs use Python along with Selenium to manipulate the browser. All of the data pulled is placed into excel files which was easiest for the user to quickly
take the information and analyze it once pulled.


SEETS_Scrape.py must be run first as it pulls the original products from their item catalog. This is where the bulk of the information for the products will be found. While
also the index (ASIN) which is used by the following two programs is also pulled from the SEETS catalog. This program can be run with multiple processes controlling 
multiple browsers to scrape concurrently. However, it was turned off as it was determined to be unnessacery for the user.


KEEPA_Scrape.py must run after SEETS_Scrape.py but will not be affected if it runs before or after AMP_Scrape.py. This program will use the ASIN information from the SEETS
scraper to pull the average sales price, the amount of stock in the market of the product, the total sales of the product, and the total sales of the product for the last 30 days.
This just gives more insight into the products if needed by the user to better their decision-making.


AMP_Scrape.py must run after SEETS_Scrape.py but will not be affected if it runs before or after KEEPA_Scrape.py. This program will use the ASIN information from the SEETS
scraper to pull the maximum cost of the product, the lowest FBA (fulfilled by Amazon) seller price for the product, and the lowest FBM (fulfilled by merchant) seller price for the product.
This just gives more insight into the products if needed by the user to better their decision-making.

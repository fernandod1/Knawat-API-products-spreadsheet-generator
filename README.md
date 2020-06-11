# Knawat API products spreadsheet generator

This python script will generate an excel spreadsheet file containing products list details from client account asociated to Knawat API (https://www.knawat.com - Turkish & European wholesale products source & Fulfillment service for Ecommerce business).

------------------------------------------------------------------
 REQUIREMENTS:
------------------------------------------------------------------

Python v3

Libraries modules json and xlwt

------------------------------------------------------------------
 CONFIGURATION:
------------------------------------------------------------------

Open script with text editor and modify only lines 168, 169, 170, 171 and 172:

APICONSUMERKEY = ""

APICONSUMERSECRET = ""

CONVERSION_VALUE_SGD_TO_USD = 0.74

SKUSFILENAME = '/path/to/skus.txt'

EXCEL_FILE = '/path/to/products-shop.xls'

------------------------------------------------------------------
 EXECUTION:
------------------------------------------------------------------

Write command:

python knawat-api.py

It will generate an excel spreadsheet with details about list of SKUs selected.


Note:
Author of script does not have any relation with Knawat enterprise.

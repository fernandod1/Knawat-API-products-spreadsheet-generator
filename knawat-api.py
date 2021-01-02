# Copyright (c) 2020 Fernando
# Url: https://github.com/fernandod1/
# License: MIT

import json
import requests
import xlwt 
from xlwt import Workbook

def existsjsonval(datajson, field):
    if field in datajson:
        value = datajson[field]
    else:
        value = ""
    return value

def existsindexinlist(list, i):
    if len(list) > i:
        return True
    else:
        return False

def checkexists(datajson, field1, index1, field2, index2, field3):
    value = ""
    if(field1!="" and index1!=""):
        if(existsindexinlist(datajson['product'][field1], index1)):
            if(field2!="" and index2!=""):
                #atributes
                if(existsindexinlist(datajson['product'][field1][index1][field2], index2)):
                    if(existsjsonval(datajson['product'][field1][index1][field2][index2], "en")!=""):
                        value = datajson['product'][field1][index1][field2][index2]['en']
                    elif(existsjsonval(datajson['product'][field1][index1][field2][index2], "tr")!=""):
                        value = datajson['product'][field1][index1][field2][index2]['tr'] 
                    else:
                        if (field2=="quantity"):
                            value = 0
                        else:
                            value = ""
            elif(field2!=""):
                value = datajson['product'][field1][index1][field2]
                #variations
    return value

def getdatajson(consumerKey, consumerSecret, sku):
    datajson = ''
    payload = {
    "consumerKey": consumerKey,
    "consumerSecret": consumerSecret,
    'limit':'1',
    'page':'1'
    }
    response = requests.post('https://mp.knawat.io/api/token', json=payload)  
    if response.status_code == 200:
        jsonformat = response.json()
        payload1 = {
            'limit':'21',
            'page':'1'
        }
        headers = {
            'Authorization' : 'Bearer ' + jsonformat['channel']['token']
        }
        response = requests.get('https://mp.knawat.io/api/catalog/products/'+sku, json=payload1 , headers=headers)
        if response.status_code == 200:
            datajson = json.loads(response.content.decode('utf-8'))
    return datajson


def processdata(i, data, datajson, CONVERSION_VALUE_SGD_TO_USD):
    object1 = {}
    object1["num"] = i
    object1["product"] = existsjsonval(datajson['product']['name'], "en")
    object1["sku"] = existsjsonval(datajson['product'], "sku")
    object1["cost1"] = float(str(round(checkexists(datajson, 'variations', 0, 'cost_price', '', '')*CONVERSION_VALUE_SGD_TO_USD, 2)))
    object1["cost2"] = float(str(round(checkexists(datajson, 'variations', 0, 'cost_price', '', ''), 2)))
    object1["sale1"] = float(str(round(checkexists(datajson, 'variations', 0, 'sale_price', '', '')*CONVERSION_VALUE_SGD_TO_USD, 2)))
    object1["sale2"] = float(str(round(checkexists(datajson, 'variations', 0, 'sale_price', '', ''), 2)))
    object1["weight"] = checkexists(datajson, 'variations', 0, 'weight', '', '')
    object1["days"] = "7-12 day(s)"
    object1["op34"] = checkexists(datajson, 'attributes', 0, 'options', 0, 'en')
    object1["op36"] = checkexists(datajson, 'attributes', 0, 'options', 1, 'en')
    object1["op38"] = checkexists(datajson, 'attributes', 0, 'options', 2, 'en')
    object1["op40"] = checkexists(datajson, 'attributes', 0, 'options', 3, 'en')
    object1["op42"] = checkexists(datajson, 'attributes', 0, 'options', 4, 'en')
    object1["op44"] = checkexists(datajson, 'attributes', 0, 'options', 5, 'en')
    object1["op46"] = checkexists(datajson, 'attributes', 0, 'options', 6, 'en')
    object1["op48"] = checkexists(datajson, 'attributes', 0, 'options', 7, 'en')
    object1["stock1"] = checkexists(datajson, 'variations', 0, 'quantity', '', '')
    object1["stock2"] = checkexists(datajson, 'variations', 1, 'quantity', '', '')
    object1["stock3"] = checkexists(datajson, 'variations', 2, 'quantity', '', '')
    object1["stock4"] = checkexists(datajson, 'variations', 3, 'quantity', '', '')
    object1["stock5"] = checkexists(datajson, 'variations', 4, 'quantity', '', '')
    object1["stock6"] = checkexists(datajson, 'variations', 5, 'quantity', '', '')
    object1["stock7"] = checkexists(datajson, 'variations', 6, 'quantity', '', '')
    object1["stock8"] = checkexists(datajson, 'variations', 7, 'quantity', '', '')
    object1["opmodelheight"] = checkexists(datajson, 'attributes', 1, 'options', 0, 'en')
    object1["opchest"] = checkexists(datajson, 'attributes', 2, 'options', 0, 'en')
    object1["opwaist"] = checkexists(datajson, 'attributes', 3, 'options', 0, 'en')
    object1["ophip"] = checkexists(datajson, 'attributes', 4, 'options', 0, 'en')
    object1["opcolour"] = checkexists(datajson, 'attributes', 5, 'options', 0, 'en')
    object1["opsex"] = checkexists(datajson, 'attributes', 6, 'options', 0, 'en')
    object1["opmaterials"] = checkexists(datajson, 'attributes', 7, 'options', 0, 'en')
    data.append(object1)
    return data

def readskus(filename):
    content = []
    try:
        f=open(filename, "r")
        if f.mode == 'r':
            content = f.read().split(" ")
    except:
        print(f"Error: can not open to read filename {filename}.")
    finally:
        f.close()
    return content

def fill_excel(d,i,sheet):    
    if i==0:
        style = xlwt.easyxf('font: bold 1')
        sheet.write(0, 0, 'S/N', style) 
        sheet.write(0, 1, 'Description', style) 
        sheet.write(0, 2, 'SKU', style) 
        sheet.write(0, 3, 'Cost(USD)', style) 
        sheet.write(0, 4, 'Cost(SGD)', style) 
        sheet.write(0, 5, 'Price(USD)', style) 
        sheet.write(0, 6, 'Price(SGD)', style) 
        sheet.write(0, 7, 'Weight', style)
        sheet.write(0, 8, 'Handling Time', style)                
        sheet.write(0, 9, 'Size1', style)
        sheet.write(0, 10, 'Size2', style)
        sheet.write(0, 11, 'Size3', style)
        sheet.write(0, 12, 'Size4', style)
        sheet.write(0, 13, 'Size5', style)
        sheet.write(0, 14, 'Size6', style)
        sheet.write(0, 15, 'Size7', style)
        sheet.write(0, 16, 'Size8', style)
        sheet.write(0, 17, 'Stock1', style)
        sheet.write(0, 18, 'Stock2', style)
        sheet.write(0, 19, 'Stock3', style)
        sheet.write(0, 20, 'Stock4', style)
        sheet.write(0, 21, 'Stock5', style)
        sheet.write(0, 22, 'Stock6', style)
        sheet.write(0, 23, 'Stock7', style)
        sheet.write(0, 24, 'Stock8', style)
        sheet.write(0, 25, 'Model\'s Height', style)
        sheet.write(0, 26, 'Model\'s Chest', style)
        sheet.write(0, 27, 'Model\'s Waist', style)
        sheet.write(0, 28, 'Model\'s Hip', style)
        sheet.write(0, 29, 'Colour', style)
        sheet.write(0, 30, 'Sex', style)
        sheet.write(0, 31, 'Material', style)   
    row = i+1
    col = 0
    itemplus = ""
    for key in d:        
        itemplus=""+itemplus+""+str(d[key])+""
        sheet.write(row, col, itemplus)
        itemplus = ""
        col += 1

def createspreadsheet(data, EXCEL_FILE):        
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("Sheet 1")
    for i in range(len(data)):
        fill_excel(data[i],i,sheet)
    workbook.save(EXCEL_FILE)

# -------------------------- Main program ------------------------------ #

def main():
    data=[]
    APICONSUMERKEY = ""
    APICONSUMERSECRET = ""
    CONVERSION_VALUE_SGD_TO_USD = 0.74
    SKUSFILENAME = '/path/to/skus.txt'
    EXCEL_FILE = '/path/to/products-shop.xls'
    i=0
    try:
        skus = readskus(SKUSFILENAME)    
        while i<len(skus):
            datajson = getdatajson(APICONSUMERKEY, APICONSUMERSECRET, skus[i])
            data = processdata(i, data, datajson, CONVERSION_VALUE_SGD_TO_USD)
            print(f"Processing SKU: {skus[i]}")
            i += 1
        createspreadsheet(data, EXCEL_FILE)
        print(f"File {EXCEL_FILE} generated.")
    except:
        print("An exception occurred executing main program.")


if __name__ == "__main__":

    main()


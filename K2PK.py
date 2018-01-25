#!/usr/bin/env python3
# coding=utf-8

# TODO: Put all CSS into single external stylesheet [partially done]
# TODO: Split html write headers / intro etc into seperate project


import csv
import decimal
import json
import os
import re
import subprocess
import sys
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
from datetime import datetime
from pprint import pprint

import requests

from currency_converter import CurrencyConverter
from json2html import *
from K2PKConfig import *
from mysql.connector import Error, MySQLConnection

# NOTE: Set precistion to cope with nano, pico & giga multipliers.
ctx = decimal.Context()
ctx.prec = 12


file_name = sys.argv[1]
projectName, ext = file_name.split(".")
print(projectName)

# Make baseline barcodes and web directories
try:
    os.makedirs('./assets/barcodes')
except OSError:
    pass

try:
    os.makedirs('./assets/web')
except OSError:
    pass

invalidate_BOM_Cost = False

try:
    distribConfig = read_distributors_config()
    preferred = (distribConfig['preferred'])
except:
    KeyError
    print('No preferred distributors in config.ini')
    pass


def float_to_str(f):
    d1 = ctx.create_decimal(repr(f))
    return format(d1, 'f')


def convert_units(num):
    factors = ["G", "M", "K", "k", "R", "", ".", "m", "u", "n", "p"]
    conversion = {
        'G': '1000000000',
        'M': '1000000',
        'K': '1000',
        'k': '1000',
        'R': '1',
        '.': '1',
        '': '1',
        'm': '0.001',
        "u": '0.000001',
        'n': '0.000000001',
        'p': '0.000000000001'}
    val = ""
    mult = ""

    for i in range(len(num)):
        if num[i] == ".":
            mult = num[i]
        if num[i] in factors:
            mult = num[i]
            val = val + "."
        else:
            if num[i].isdigit():
                val = val + (num[i])
            else:
                print("Invalid multiplier")
                return("0")
                break
    if val.endswith("."):
        val = val[:-1]
    m = float(conversion[mult])
    v = float(val)
    r = float_to_str(m * v)

    r = r.rstrip("0")
    r = r.rstrip(".")
    return(r)


def partStatus(partID, parameter):
    dbconfig = read_db_config()
    try:
        conn = MySQLConnection(**dbconfig)
        cursor = conn.cursor()
        sql = "SELECT R.stringValue FROM PartParameter R WHERE (R.name = '{}') AND (R.part_id = {})".format(
            parameter, partID)
        cursor.execute(sql)
        partStatus = cursor.fetchall()

        if partStatus == []:
            part = "Unknown"
        else:
            part = str(partStatus[0])[2:-3]
        return (part)

    except UnicodeEncodeError as err:
        print(err)

    finally:
        cursor.close()
        conn.close()


def getDistrib(partID):

    dbconfig = read_db_config()
    try:
        conn = MySQLConnection(**dbconfig)
        cursor = conn.cursor()
        sql = """SELECT D.name, PD.sku, D.skuurl FROM Distributor D
				LEFT JOIN PartDistributor PD on D.id = PD.distributor_id
				WHERE PD.part_id = {}""".format(partID)
        cursor.execute(sql)
        distributors = cursor.fetchall()
        unique = []
        d = []
        distributor = []

        for distributor in distributors:
            if distributor[0] not in unique and distributor[0] in preferred:
                unique.append(distributor[0])
                d.append(distributor)
        return(d)

    except UnicodeEncodeError as err:
        print(err)

    finally:
        cursor.close()
        conn.close()


def octopartLookup(partIn, bean):

    try:
        octoConfig = read_octopart_config()
        apikey = (octoConfig['apikey'])
    except:
        KeyError
        print('No Octopart API key in config.ini')
        return (2)

    try:
        currencyConfig = read_currency_config()
        locale = (currencyConfig['currency'])
    except:
        KeyError
        print("No currency configured in config.ini")
        return(4)

    # Get currency rates from European Central Bank
    # Fall back on cached cached rates
    try:
        c = CurrencyConverter(
            'http://www.ecb.europa.eu/stats/eurofxref/eurofxref.zip')
    except:
        URLError
        c = CurrencyConverter()
        return(8)

    web = str("./assets/web/" + partIn + ".html")

    webpage = open(web, "w")

    # Replace spaces in part name with asterisk for 'wildcard' search

    Part = partIn.replace(" ", "*")

    aside = open("./assets/web/tmp.html", "w")

    htmlHeader = """
    <!DOCTYPE html>
    <meta charset="utf-8">
    <html lang="en">
    <head>
      <meta charset="utf-8">
      <meta http-equiv="X-UA-Compatible" content="IE=edge">
      <title>Octopart Lookup</title>
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <meta name="Description" lang="en" content="Kicad2PartKeepr">
      <meta name="author" content="jpateman@gmail.com">
      <meta name="robots" content="index, follow">

      <!-- icons -->
      <link rel="apple-touch-icon" href="assets/img/apple-touch-icon.png">
      <link rel="shortcut icon" href="favicon.ico">

      <!-- Override CSS file - add your own CSS rules -->
      <link rel="stylesheet" href="../css/octopart.css">
    </head>

    <body>
      <div class="header">
          <h1 class="header-heading">Kicad2PartKeepr</h1>
      </div>
      <div class="nav-bar">
        <div class="container">
          <ul class="nav">
          </ul>
        </div>
      </div>
      """

    webpage.write(htmlHeader)

    ##################
    bean = False
    # FIXME: Treat 'bean' devices separately - use 'search' API not 'match'
    # FIXME: Search only for Non_PK parts in first instance.
    ##################

    if bean:
        #
        url = "http://octopart.com/api/v3/parts/search"
        url += '?apikey=' + apikey
        url += '&q="' + Part + '"'
        url += '&include[]=descriptions'
        # url += '&include[]=imagesets'
        # url += '&include[]=specs'
        # url += '&include[]=datasheets'
        url += '&country=GB'
        data = urllib.request.urlopen(url).read()
        response = json.loads(data)
        pprint(response)
    else:
        url = "http://octopart.com/api/v3/parts/match"
        url += '?apikey=' + apikey
        url += '&queries=[{"mpn":"' + Part + '"}]'
        url += '&include[]=descriptions'
        url += '&include[]=imagesets'
        url += '&include[]=specs'
        url += '&include[]=datasheets'
        url += '&country=GB'

    data = urllib.request.urlopen(url).read()
    response = json.loads(data)

    loop = False

    for result in response['results']:
        for item in result['items']:
            if loop:
                break
            loop = True

            partNum = item['mpn']

            try:
                description = str(item['descriptions'][0].get('value', None))
            except:
                IndexError
                description = ""

            webpage.write(
                "<div class='content'><div class='main'><br><h2>" +
                partNum +
                "</h2><h4>" +
                description +
                "</h4><br><br>")

    # Get image (if present). Also need to get attribution for Octopart licensing
            try:
                image = str(item['imagesets'][0]
                            ['medium_image'].get('url', None))
            except:
                IndexError
                image = ""

            try:
                credit = str(item['imagesets'][0].get('credit_string', None))
            except:
                IndexError
                credit = ""

            aside.write("<div class='aside'><img src='" + image +
                        "'><div class ='crdt'>Image credit: " + credit + "</div><br><br><table class='table table-striped'><thead>")
            aside.write("<th>Characteristic</th><th>Value</th></thead><tbody>")

            for spec in item['specs']:
                parm = item['specs'][spec]['metadata']['name']
                try:
                    if type(item['specs'][spec]['display_value']):
                        val = str(item['specs'][spec]['display_value'])
                except:
                    IndexError
                    val = "Not Listed by Manufacturer"

                parameter = (("{:30} ").format(parm))
                value = (("{:40}").format(val))
#                print(parameter, "  :  ", value)
                print(("| {:30} : {:124} |").format(parameter, value))
                aside.write(
                    "<tr><td>" +
                    parameter +
                    "</td><td>" +
                    value +
                    "</td></tr>")
            print(('{:_<162}').format(""))

            aside.write("</tbody></table><table>")
            aside.write(
                "<thead><th>Datasheets</th><th>Date</th><th>Pages</th></thead><tbody>")

            for datasheet in item['datasheets']:
                try:
                    if (datasheet['metadata']['last_updated']):
                        dateUpdated = (datasheet['metadata']
                                       ['last_updated'])[:10]
                    else:
                        dateUpdated = "Unknown"
                except:
                    IndexError
                    dateUpdated = "Unknown"

                if datasheet['attribution']['sources'] is None:
                    source = "Unknown"
                else:
                    source = datasheet['attribution']['sources'][0]['name']

                numPages = str(datasheet['metadata']['num_pages'])
                documents = (
                    ("| {:30} {:11} {:12} {:7} {:7} {:1} {:84} |").format(
                        source,
                        " Updated: ",
                        dateUpdated,
                        "Pages: ",
                        numPages, "", datasheet['url']))
                print(documents)
                aside.write(
                    "<tr><td><a href=" +
                    datasheet['url'] +
                    "> " +
                    source +
                    " </a></td><td>" +
                    dateUpdated +
                    "</td><td>" +
                    numPages +
                    "</td></tr>")
            if loop:
                webpage.write("<table class='table table-striped'>")
            else:
                webpage.write("<p> No Octopart results found </p>")

    # Header row here

    webpage.write("<thead><th>Seller</th><th>SKU</th><th>Stock</th><th>MOQ</th><th>Package</th><th>Currency</th><th>1</th><th>10</th><th>100</th><th>1000</th><th>10000</th><thead><tbody>")
    count = 0

    for result in response['results']:

        for item in result['items']:
            if count == 0:
                print(('{:_<162}').format(""))
                print(("| {:24} | {:19} | {:>9} | {:>7} | {:11} | {:5} ").format(
                    "Seller", "SKU", "Stock", "MOQ", "Package", "Currency"), end="")
                print(
                    ("|  {:>10}|  {:>10}|  {:>10}|  {:>10}|  {:>10}|").format(
                        "1",
                        "10",
                        "100",
                        "1000",
                        "10000"))
                print(('{:-<162}').format(""), end="")
                count += 1

            # Breaks at 1, 10, 100, 1000, 10000
            for offer in item['offers']:
                loop = 0
                _seller = offer['seller']['name']
                _sku = (offer['sku'])[:19]
                _stock = offer['in_stock_quantity']
                _moq = str(offer['moq'])

                if _moq == "None":
                    _moq = '-'
                _package = str(offer['packaging'])

                if _package == "None":
                    _package = "-"

                print()

                print(("| {:24} | {:19} | {:>9} | {:>7} | {:11} |").format(
                    _seller, _sku, _stock, _moq, _package), end="")
                line = "<tr><td>" + _seller + "</td><td><a target='_blank' href=" + str(offer['product_url']) + ">" + str(
                    offer['sku']) + "</a></td><td>" + str(_stock) + "</td><td>" + str(_moq) + "</td><td>" + _package + "</td>"
                webpage.write(line)

                valid = False
                points = ['-', '-', '-', '-', '-']
                for currency in offer['prices']:
                    # Some Sellers don't have currency so use this to fill the line
                    valid = True

                    if currency == locale:
                        # Base currency is local
                        loop += 1
                        if loop == 1:
                            print((" {:3}      |").format(currency), end="")
                            webpage.write("<td>" + currency + "</td>")
                    else:
                        # Only try and convert first currency
                        loop += 1
                        if loop == 1:
                            print((" {:3}*     |").format(locale), end="")
                            webpage.write("<td>" + locale + "*</td>")

                    if loop == 1:

                        for breaks in offer['prices'][currency]:

                            _moqv = offer['moq']
                            if _moqv is None:
                                _moqv = 1
                            _moqv = int(_moqv)
                            i = 0

                            # Break 0 - 9
                            if breaks[0] < 10:
                                points[0] = round(
                                    c.convert(breaks[1], currency, locale), 2)
    #                        if _moqv >= breaks[0]:
                            # Propogate right
                            for i in range(0, 4):
                                points[i + 1] = points[i]

                            # Break 10 to 99
                            if breaks[0] >= 10 and breaks[0] < 100:
                                points[1] = round(
                                    c.convert(breaks[1], currency, locale), 3)
    #                        if _moqv >= breaks[0]:
                            for i in range(1, 4):
                                points[i + 1] = points[i]

                            # Break 100 to 999
                            if breaks[0] >= 100 and breaks[0] < 1000:
                                points[2] = round(
                                    c.convert(breaks[1], currency, locale), 4)
    #                        if _moqv >= breaks[0]:
                            for i in range(2, 4):
                                points[i + 1] = points[i]

                            # Break 1000 to 9999
                            if breaks[0] >= 1000 and breaks[0] < 10000:
                                points[3] = round(
                                    c.convert(breaks[1], currency, locale), 5)
#                            if _moqv >= breaks[0]:
                            points[4] = points[3]

                            # Break 10000+
                            if breaks[0] >= 10000:
                                points[4] = round(
                                    c.convert(
                                        breaks[1],
                                        currency,
                                        locale),
                                    6)
                            else:
                                points[4] = points[3]

                        for i in range(0, 5):
                            print(("  {:>10.5}|").format(points[i]), end="")
                            webpage.write("<td>" + str(points[i]) + "</td>")
                    webpage.write("</tr>")
                if not valid:
                    print("          |", end="")
                    webpage.write("<td></td>")
                    for i in range(0, 5):
                        print(("  {:>10.5}|").format(points[i]), end="")
                        webpage.write("<td>" + str(points[i]) + "</td>")
                    webpage.write("</tr>")
                    valid = False
    webpage.write("</tbody></table></div>")
    print()
    print(('{:=<162}').format(""))

    aside.close()
    side = open("./assets/web/tmp.html", "r")

    aside = side.read()

    webpage.write(aside + "</tbody></table>")
    webpage.write(
        "</div></div><div class='footer'>&copy; Copyright 2017</div></body></html>")
    return

################################################################################
#   Further setup or web configuration here
#
#
################################################################################


compliance = {
    'Compliant': 'assets/img/ROHS_GREEN.png',
    'Non-Compliant': 'assets/img/ROHS_RED.png',
    'Unknown': 'assets/img/ROHS_BLACK.png'
}

manufacturing = {
    'Obsolete': 'assets/img/FACTORY_RED.png',
    'Not Recommended for New Designs': 'assets/img/FACTORY_YELLOW.png',
    'Unknown': 'assets/img/FACTORY_BLUE.png',
    'Active': 'assets/img/FACTORY_GREEN.png',
    'Not Listed by Manufacturer': 'assets/img/FACTORY_PURPLE.png'
}

dateBOM = datetime.now().strftime('%Y-%m-%d')
timeBOM = datetime.now().strftime('%H:%M:%S')

run = 0
web = open("assets/web/temp.html", "w")
picklist = open("assets/web/picklist.html", "w")
labels = open("assets/web/labels.html", "w")
accounting = open("assets/web/accounting.html", "w")
missing = open("assets/web/missing.csv", "w")
under = open("assets/web/under.csv", "w")
under.write('Name,Description,Quantity,Stock,Min Stock\n')

htmlHeader = """
<!DOCTYPE html PUBLIC '-//W3C//DTD HTML 4.01//EN'>
<meta charset="utf-8">
<html>
<head>
<link rel="stylesheet" href="assets/css/web.css">
<title>Kicad 2 PartKeepr</title>
</head>
<body>
    <div class="header">
            <h1 class="header-heading">Kicad2PartKeepr</h1>
    </div>
    <div class="nav-bar">
            <ul class="nav">
                <li><a href='assets/web/labels.html'>Labels</a></li>
                <li><a href='assets/web/picklist.html'>Pick List</a></li>
                <li><a href='assets/web/missing.csv'download='missing.csv'>Missing parts list</a></li>
                <li><a href='assets/web/under.csv'download='under.csv'>Understock list</a></li>
                <li><a href='assets/web/""" + projectName + """_PK.csv'> PartKeepr import  </a></li>
            </ul>

    </div>

    <h2> </h2>
    <h2><br>&nbsp&nbspProject name: """ + projectName + \
    "</h2><h3><br>&nbsp&nbspDate:\t" + dateBOM + \
    "</h3>\n<h3><br>&nbsp&nbspTime:\t" + timeBOM + "</h3><h3></h3>"


htmlBodyHeader = """
<br><br>
<table class = 'main'>
"""

picklistHeader = """
<!DOCTYPE html PUBLIC '-//W3C//DTD HTML 4.01//EN'>
<meta charset="utf-8">
<html>
<head>
<title>Picklist</title>
<link rel="stylesheet" href="../css/picklist.css">
</head>
<body>
<h2> Picklist </h2>\n<h2> Project name: """

pick2 = projectName + "</h2><h3>" + dateBOM + "</h3>\n<h3>" + timeBOM + "</h3>"

picklist.write(picklistHeader + pick2 + "<table class = 'main'>")


label_header = """
<!DOCTYPE html PUBLIC '-//W3C//DTD HTML 4.01//EN'>
<meta charset="utf-8">
<html>
<head>
<title>Kicad to PartKeepr</title>
<link rel="stylesheet" href="../css/labels.css">
</head>
<body>
<table style="width:100%" table border='0' cellpadding='35' cellspacing='0'>
"""

labels.write(label_header)
label_cnt = 0


htmlAccountingHeader = """
<table class = "accounting">
"""


resistors = ["R_1206", "R_0805", "R_0603", "R_0402"]
capacitors = ["C_1206", "C_0805", "C_0603", "C_0402"]


def punctuate(value):
    if "." in value:
        multiplier = (value.strip()[-1])
        new_string = "_" + (re.sub("\.", multiplier, value))[:-1]
    else:
        new_string = "_" + value
    return(new_string)


def get_choice(possible):
    print("More than one component in the PartKeepr database meets the criteria:")
    i = 1
    for name, description, stockLevel, minStockLevel, averagePrice, partNum, storage_locn, PKid, Manufacturer in possible:
        print(i, " : ", name, " : ", description)
#    subprocess.call(['/usr/local/bin/pyparts', 'specs', name])
        i = i + 1
    print("Choose which component to add to BOM (or 0 to defer)")

    while True:
        choice = int(input('>'))
        if choice == 0:
            return (possible)
        if choice < 0 or choice > len(possible):
            continue
        break

    i = 1
    for name, description, stockLevel, minStockLevel, averagePrice, partNum, storage_locn, PKid, Manufacturer in possible:
        possible = (name, description, stockLevel, minStockLevel,
                    averagePrice, partNum, storage_locn, PKid, Manufacturer)
        if i == choice:
            possible = (name, description, stockLevel, minStockLevel,
                        averagePrice, partNum, storage_locn, PKid, Manufacturer)
            print("Selected :")
            print(possible[0], " : ", possible[1])
            return [possible]
        i = i + 1


def find_part(part_num):

    dbconfig = read_db_config()

    try:
        conn = MySQLConnection(**dbconfig)
        cursor = conn.cursor()

        bean = False

        if (part_num[:6]) in resistors:
            quality = "Resistance"
            variant = "Resistance Tolerance"
            bean = True
        if (part_num[:6]) in capacitors:
            quality = "Capacitance"
            variant = "Dielectric Characteristic"
            bean = True

        if (bean):
            component = part_num.split('_')

            if (len(component)) <= 2:
                print("Insufficient parameters (Needs 3 or 4) e.g. R_0805_100K(_±5%)")
                return ("0")

            c_case = component[1]
            c_value = convert_units(component[2])

            if (len(component)) == 4:
                c_characteristics = component[3]

                # A fully specified 'bean'
                sql = """SELECT P.name, P.description, P.stockLevel, P.minStockLevel, P.averagePrice, P.internalPartNumber, S.name, P.id, M.name
                        FROM Part P
                        JOIN PartParameter R ON R.part_id = P.id
                        JOIN StorageLocation S ON  S.id = P.storageLocation_id
                        LEFT JOIN PartManufacturer PM on PM.part_id = P.id
                        LEFT JOIN Manufacturer M on M.id = PM.manufacturer_id
                        WHERE
                        (R.name = 'Case/Package' AND R.stringValue='{}') OR
                        (R.name = '{}' AND R.normalizedValue = '{}') OR
                        (R.name = '{}' AND R.stringValue = '%{}')
                        GROUP BY P.id, M.id, S.id
                        HAVING
                        COUNT(DISTINCT R.name)=3""".format(
                    c_case, quality, c_value, variant, c_characteristics)
            else:
                # A partially specified 'bean'
                sql = """SELECT P.name, P.description, P.stockLevel, P.minStockLevel, P.averagePrice, P.internalPartNumber, S.name, P.id, M.name
                        FROM Part P
                        JOIN PartParameter R ON R.part_id = P.id
                        JOIN StorageLocation S ON  S.id = P.storageLocation_id
                        LEFT JOIN PartManufacturer PM on PM.part_id = P.id
                        LEFT JOIN Manufacturer M on M.id = PM.manufacturer_id
                        WHERE
                        (R.name = 'Case/Package' AND R.stringValue='{}') OR
                        (R.name = '{}' AND R.normalizedValue = '{}')
                        GROUP BY P.id, M.id, S.id
                        HAVING
                        COUNT(DISTINCT R.name)=2""".format(
                    c_case, quality, c_value)
        else:

            sql = """SELECT P.name, P.description, P.stockLevel, P.minStockLevel, P.averagePrice, P.internalPartNumber, S.name, P.id, M.name
                 FROM Part P
                 JOIN StorageLocation S ON  S.id = P.storageLocation_id
                 LEFT JOIN PartManufacturer PM on PM.part_id = P.id
                 LEFT JOIN Manufacturer M on M.id = PM.manufacturer_id
                 WHERE P.name LIKE '%{}%'""".format(part_num)

        cursor.execute(sql)
        components = cursor.fetchall()
        return (components, bean)

    except UnicodeEncodeError as err:
        print(err)

    finally:
        cursor.close()
        conn.close()


###############################################################################
#
# Main part of program follows
#
#
###############################################################################


with open(file_name, newline='', encoding='utf-8') as csvfile:
    reader = csv.DictReader(csvfile, delimiter=',')
    headers = reader.fieldnames

    filename, file_extension = os.path.splitext(file_name)

    outfile = open("./assets/web/" + filename + '_PK.csv',
                   'w', newline='\n', encoding='utf-8')
    writeCSV = csv.writer(outfile, delimiter=',')

# Initialise accounting values
    countParts = 0
    count_BOMLine = 0
    count_NPKP = 0
    count_PKP = 0
    count_LowStockLines = 0
    count_PWP = 0
    bomCost = 0

    for row in reader:
        if not row:
            break
        new_string = ""
        part = row['Part#']
        value = row['Value']
        footprint = row['Footprint']
        datasheet = row['Datasheet']
        characteristics = row['Characteristics']
        references = row['References']
        quantity = row['Quantity Per PCB']
        # Need sufficient info to process. Some .csv reprocessing adds in lines
        # of NULL placeholders where there was a blank line.
        if part == "" and value == "" and footprint == "":
            break

        count_BOMLine = count_BOMLine + 1

        if footprint in resistors:
            if value.endswith("Ω"):  # Remove trailing 'Ω' (Ohms)
                value = (value[:-1])
            new_string = punctuate(value)

        if footprint in capacitors:
            if value.endswith("F"):  # Remove trailing 'F' (Farads)
                value = (value[:-1])
            new_string = punctuate(value)

        if characteristics is None:
            if characteristics != "-":
                new_string = new_string + "_" + str(characteristics)

        if part == "-":
            part = (str(footprint) + new_string)

        if references is None:
            break

        component_info, species = find_part(part)

        n_components = len(component_info)

        quantity = int(quantity)

# Print to screen - these could all do with neatening up...
        print(('{:=<162}').format(""))
        print(("| BOM Line number : {:3} {:136} |").format(count_BOMLine, ""))
    #    print("| BOM Line number : ", count_BOMLine)

        print(('{:_<162}').format(""))
        print(
            ("| {:100} | {:13.13}         |                  | Req =   {:5}|").format(
                references,
                part,
                quantity))
        print(('{:_<162}').format(""))

        uniqueNames = []

        if n_components == 0:
            #        print("|                                                                         |                                                         |")
            print(("| {:100} | {:21} | {:16} | {:12} |").format(
                "No matching parts in database", "", "", ""))
            print(('{:_<162}').format(""))

            octopartLookup(part, species)
            print('\n\n')

        else:
            for (
                name,
                description,
                stockLevel,
                minStockLevel,
                averagePrice,
                partNum,
                storage_locn,
                PKid,
                    Manufacturer) in component_info:
                ROHS = partStatus(PKid, 'RoHS')
                Lifecycle = partStatus(PKid, 'Lifecycle Status')
                # Can get rid of loop as never reset now
                print(
                    ("| {:100} | Location = {:10} | Part no = {:6} | Stock = {:5}|").format(
                        description, storage_locn, partNum, stockLevel))
                print(('{:_<162}').format(""))
                print(
                    ("| Manufacturing status: {} {:<136}|").format(
                        "", Lifecycle))
                print(("| RoHS: {}{:<153}|").format("", ROHS))
                print(("| Name: {}{:<153}|").format("", name))
                getDistrib(PKid)
                print(('{:_<162}').format(""))
                octopartLookup(name, species)
                print('\n\n')

# More than one matching component exists - prompt user to choose
        if len(component_info) >= 2:
            component_info = get_choice(component_info)
            for (
                name,
                description,
                stockLevel,
                minStockLevel,
                averagePrice,
                partNum,
                storage_locn,
                PKid,
                    Manufacturer) in component_info:
                ROHS = partStatus(PKid, 'RoHS')
                Lifecycle = partStatus(PKid, 'Lifecycle Status')

        if n_components != 0 and (quantity > stockLevel):
            count_LowStockLines = count_LowStockLines + 1
            background = 'rgba(60, 0, 0, 0.15)'  # Insufficient stock : pinkish
            under.write(
                name +
                ',' +
                description +
                ',' +
                str(quantity) +
                ',' +
                str(stockLevel) +
                ',' +
                str(minStockLevel) +
                '\n')
        else:
            background = 'rgba(0, 60, 0, 0.15)'  # Adequate stock : greenish

        countParts = countParts + quantity
        quantity = str(quantity)


# Print header row with white background  - should move this somewwhere
# else....
        if run == 0:
            web.write(
                "<tr style = 'background-color : white; font-weight : bold;'>")
            web.write("<th style = 'width : 7%;'>References</th><th style = 'width : 7%;'>Part</th><th>Description</th><th>Stock</th><th>Manufacturer</th><th>Distributor</th><th>Qty</th><th>Each</th><th>Line</th>")
            picklist.write("<tr style = 'font-weight : bold;'>")
            picklist.write(
                "<th>References</th><th>Part</th><th>Description</th><th>Stock</th><th>Part Number</th><th>Location</th><th>Qty</th><th>Pick</th>")
            run = 1

# No PK components fit search criteria. Deal with here and drop before loop.
# Not ideal but simpler.
        if n_components == 0:
            count_NPKP = count_NPKP + 1
            averagePrice = 0

            background = 'rgba(0, 60, 60, 0.3)'  # Green Blue background
            web.write("<tr style = 'background-color : " + background + ";'>")
            web.write(
                "<td style = 'font-weight : bold;'>" +
                references +
                "</td>")
            web.write(
                "<td ><a href = 'assets/web/" +
                part +
                ".html'>" +
                part +
                "</a></td>")
            web.write("<td> Non PartKeepr component</td>")
            web.write("<td style = 'font-weight : bold;'>NA</td>")
            web.write(
                "<td>NA</td>")
            web.write(
                "<td>&nbsp&nbsp&nbspNA</td>")
            web.write(
                "<td style = 'font-weight : bold;'> " +
                quantity +
                "</td>")
            web.write("<td></td>")
            web.write("<td></td></tr>")

            picklist.write('\n')

            missing.write(part + ',' + quantity + ',' + references + '\n')
            name = "-"

        if n_components > 1:  # Multiple component fit search criteria - set brown background
            background = 'rgba(60, 60, 0, 0.4)'

        i = 0
        for (
            name,
            description,
            stockLevel,
            minStockLevel,
            averagePrice,
            partNum,
            storage_locn,
            PKid,
                Manufacturer) in component_info:
            web.write("<tr style = 'background-color : " + background + ";'>")
            picklist.write('\n')
            picklist.write("<tr style = 'background-color : white;'>")
            if i == 0:  # 1st line where multiple components fit search showing RefDes
                web.write(
                    "<td style = 'font-weight : bold;'>" +
                    references +
                    "</td>")
                picklist.write("<td style = 'font-weight : bold;'>" +
                               references + "</td>")

                web.write("<td ><a href = 'assets/web/" + name +
                          ".html'>" + name + "</a></td>")
                picklist.write("<td>" + name + "</td>")
                i = i + 1
                count_PKP = count_PKP + 1
            else:  # 2nd and subsequent lines where multiple components fit search showing RefDes
                web.write(
                    "<td colspan='2' style = 'font-weight : bold;'> *** ATTENTION *** Multiple sources available *** Use only ONE line *** </td>")
                picklist.write(
                    "<td colspan='2' style = 'font-weight : bold;'> *** ATTENTION *** Multiple sources available *** Use only ONE line *** </td>")
                invalidate_BOM_Cost = True
            lineCost = float(averagePrice) * int(quantity)
            if lineCost == 0:
                count_PWP += 1

            rohsIcon = compliance[ROHS]
            lifecycleIcon = manufacturing[Lifecycle]

            web.write(
                "<td>" +
                description +
                "  <img align = 'right' style='padding-left: 5px;' src = '" +
                rohsIcon +
                "'alt='' title='ROHS: " +
                ROHS +
                "'/><img align = 'right' style='padding-left: 5px;' src = '" +
                lifecycleIcon +
                "'alt='' title='Lifecycle: " +
                Lifecycle +
                "'/></td>")
            web.write("<td style = 'font-weight : bold;'>" +
                      str(stockLevel) + "</td>")

            picklist.write("<td>" + description + "</td>")
            picklist.write("<td style = 'font-weight : bold;'>" +
                           str(stockLevel) + "</td>")

# Part number exists, therefore generate bar code
# Requires Zint >1.4 - doesn't seem to like to write to another directory.
# Ugly hack - write to current directory and move into place.
            if partNum != "":
                part_no = ""
                part_no = (partNum[1:])
                subprocess.call(['/usr/local/bin/zint',
                                 '--filetype=png',
                                 '-w',
                                 '10',
                                 '--height',
                                 '20',
                                 '-o',
                                 part_no,
                                 '-d',
                                 partNum])
                os.rename(part_no + '.png',
                          'assets/barcodes/' + part_no + '.png')
                if Manufacturer:
                    web.write("<td>" + Manufacturer + "</td>")
                else:
                    web.write("<td>NA</td>")
                picklist.write(
                    "<td style = 'background-color : white' align='center'><img src = '../barcodes/" +
                    part_no +
                    ".png' ></td>")
            else:
                # No Part number
                web.write(
                    "<td> NA </td>")
                picklist.write(
                    "<td style = 'background-color : white' align = 'center'> NA </td>")

# Storage location exists, therefore generate bar code. Ugly hack - my location codes start with
# a '$' which causes problems. Name the file without the leading character.
            if storage_locn != "":
                locn_trim = ""
                locn_trim = (storage_locn[1:])
                subprocess.call(['/usr/local/bin/zint',
                                 '--filetype=png',
                                 '-w',
                                 '10',
                                 '--height',
                                 '20',
                                 '-o',
                                 locn_trim,
                                 '-d',
                                 storage_locn])
                os.rename(
                    locn_trim +
                    '.png',
                    'assets/barcodes/' +
                    locn_trim +
                    '.png')

                web.write("<td>")
                distributors = getDistrib(PKid)
                web.write("<table style='width: 100%; border: 0px;'>")
                for distributor in distributors:
                    web.write("<tr><td style='width: 50%; border: 0px;'>")
                    web.write(distributor[0])
                    web.write("</td><td style='width: 50%; border: 0px;'>")
                    if distributor[2]:
                        web.write("<a href =")
                        web.write(distributor[2])
                        web.write(distributor[1])
                        web.write(">")
                        web.write(distributor[1])
                        web.write("</a></td></tr>")
                    else:
                        web.write(distributor[1])
                        web.write("</td></tr>")
                web.write("</table>")
                picklist.write(
                    "<td style = 'background-color : white' align='center'><img src = '../barcodes/" +
                    locn_trim +
                    ".png'></td>")
            else:
                # No storage location
                web.write(
                    "<td>NA</td>")
                picklist.write("<td align = 'center'> NA </td>")

            avPriceFMT = str(('£{:0,.2f}').format(averagePrice))
            linePriceFMT = str(('£{:0,.2f}').format(lineCost))
            bomCost = bomCost + lineCost

            web.write(
                "<td style = 'font-weight : bold;'> " +
                quantity +
                "</td>")
            web.write("<td> " +
                      avPriceFMT + "</td>")
            web.write("<td> " +
                      linePriceFMT + "</td></tr>")
            web.write('\n')

            picklist.write("<td style = 'font-weight : bold;'> " +
                           quantity + "</td>")

            picklist.write("<td></td></tr>")
            picklist.write("\n")

# Make labels for packets (need extra barcodes here)
            subprocess.call(['/usr/local/bin/zint',
                             '--filetype=png',
                             '-w',
                             '10',
                             '--height',
                             '20',
                             '-o',
                             name,
                             '-d',
                             name])
            os.rename(name + '.png', 'assets/barcodes/' + name + '.png')
            subprocess.call(['/usr/local/bin/zint',
                             '--filetype=png',
                             '-w',
                             '10',
                             '--height',
                             '20',
                             '-o',
                             quantity,
                             '-d',
                             quantity])
            os.rename(quantity + '.png',
                      'assets/barcodes/' + quantity + '.png')

# Write out label webpage too
            labels.write("<td><p class = 'heavy'>" + description[:64].upper())
            labels.write("<p>")
            labels.write("<p>" + name)
            labels.write("<p><img src = '../barcodes/" +
                         name + ".png' style='height:40px;'>")
            labels.write("<p>")
            labels.write(
                "<div class = 'left'>Part number: #" +
                part_no +
                "</div>")
            labels.write(
                "<div class = 'right'>Location:" +
                storage_locn +
                "</div>")
            labels.write("<div class = 'left'><img src = '../barcodes/" +
                         part_no + ".png' style='height:40px;'></div>")
            labels.write("<div class = 'right'><img src = '../barcodes/" +
                         locn_trim + ".png' style='height:40px;'></div>")
            labels.write(
                "<div class = 'left'>Quantity: " +
                quantity +
                "</div>")
            labels.write("<p><img src = '../barcodes/" + quantity +
                         ".png' style='height:40px;'></div>")
            labels.write("<p class = 'heavy'>" + file_name.upper())
            labels.write("<p>")
            labels.write("<p class = 'heavy'>" + references)
            labels.write("</td>")

            label_cnt = label_cnt + 1
            if label_cnt == 3:
                labels.write("</tr><tr>")
                label_cnt = 0

# Prevent variables from being recycled
            storage_locn = ""
            partNum = ""
            part_no = ""

        writeCSV.writerow([references, name, quantity])
        references = ""
        name = ""
        quantity = ""

# Write out footer for webpage
    web.write(("</table></body></html>"))

# Write out footer for picklist
    picklist.write(("</table></body></html>"))

# Write out footer for labels
    labels.write(("</tr></table></body></html>"))

# Now script has run, construct table with part counts 7 costs etc.
    bomCostDisp = str(('£{:0,.2f}').format(bomCost))

    accounting.write("<tr>")
    accounting.write("<td> Total parts </td>")
    accounting.write("<td>" + str(countParts) + "</td>")
    accounting.write("</tr><tr>")
    accounting.write("<td> BOM Lines </td>")
    accounting.write("<td>" + str(count_BOMLine) + "</td>")
    accounting.write("</tr><tr>")
    accounting.write("<td> Non-PartKeepr Parts </td>")
    accounting.write("<td>" + str(count_NPKP) + "</td>")
    accounting.write("</tr><tr>")
    accounting.write("<td>PartKeepr Parts</td>")
    accounting.write("<td>" + str(count_PKP) + "</td>")
    accounting.write("</tr><tr>")
    accounting.write("<td>Parts without pricing info</td>")
    accounting.write("<td>" + str(count_PWP) + "</td>")
    accounting.write("</tr><tr>")
    accounting.write("<td>Low Stock</td>")
    accounting.write("<td>" + str(count_LowStockLines) + "</td>")
    accounting.write("</tr><tr>")
    accounting.write("<td>BOM Cost (Based on PartKeepr inventory prices)</td>")
    if not invalidate_BOM_Cost:
        accounting.write("<td>" + bomCostDisp + "</td>")
    else:
        accounting.write("<td>BOM price not calculated</td>")
    accounting.write(
        ("</tr></table><p>"))

# Assemble webpage
web = open("assets/web/temp.html", "r")
web_out = open("webpage.html", "w")


accounting = open("assets/web/accounting.html", "r")
accounting = accounting.read()

htmlBody = web.read()

web_out.write(htmlHeader)
web_out.write(htmlAccountingHeader + accounting + "<br>")
web_out.write(htmlBodyHeader + htmlBody)

# Open webpage in default browser
webbrowser.open('file://' + os.path.realpath('webpage.html'))

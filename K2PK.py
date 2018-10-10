#!/usr/local/bin/python3
# -*- coding: utf-8 -*-

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
import http.server
import socketserver
from datetime import datetime
# from pprint import pprint
import requests
from math import log10, ceil
import numpy as np
import numpy.ma as ma
from fpdf import FPDF
from currency_converter import CurrencyConverter
from json2html import *
from K2PKConfig import *
from mysql.connector import Error, MySQLConnection

# NOTE: Set precision to cope with nano, pico & giga multipliers.
ctx = decimal.Context()
ctx.prec = 12

# Set colourscheme (this could go into preferences.ini)
# Colour seems to be most effective when used against a white background

adequate = 'rgb(253,246,227)'
# adequate = 'rgba(0, 60, 0, 0.15)'

lowstock = '#f6f6f6'
# lowstock = 'rgba(255, 255, 255, 0)'

nopkstock = '#c5c5c5'
# nopkstock = 'rgba(0, 60, 60, 0.3)'

multistock = '#e5e5e5'
# multistock = 'rgba(255, 255, 255, 0)'

minPriceCol = 'rgb(133,153,0)'

try:
    currencyConfig = read_currency_config()
    baseCurrency = (currencyConfig['currency'])
except KeyError:
    print("No currency configured in config.ini")

assert sys.version_info >= (3, 4)

file_name = sys.argv[1]
projectName, ext = file_name.split(".")
print(projectName)

numBoards = 0

try:
    while numBoards < 1:
        qty = input("How many boards? (Enter 1 or more) > ")
        numBoards = int(qty)
    print("Calculations for ", numBoards, " board(s)")
except ValueError:
    print("Integer values only, >= 1. Quitting now")
    raise SystemExit

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
except KeyError:
    print('No preferred distributors in config.ini')
    pass

# Initialise empty cost and coverage matrix
prefCount = preferred.count(",") + 1
costMatrix = [0] * prefCount
coverageMatrix = [0] * prefCount
countMatrix = [0] * prefCount
voidMatrix = [0] * prefCount


def float_to_str(f):
    d1 = ctx.create_decimal(repr(f))
    return format(d1, 'f')


def convert_units(num):
    '''
    Converts metric multipliers values into a decimal float.

    Takes one input eg 12.5m and returns the decimal (0.0125) as a string. Also supports
    using the multiplier as the decimal marker e.g. 4k7 
    '''

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
        'p': '0.000000000001'
    }
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
                return "0"
    if val.endswith("."):
        val = val[:-1]
    m = float(conversion[mult])
    v = float(val)
    r = float_to_str(m * v)

    r = r.rstrip("0")
    r = r.rstrip(".")
    return r


def limit(num, minimum=10, maximum=11):
    '''
    Limits input 'num' between minimum and maximum values.

    Default minimum value is 10 and maximum value is 11.
    '''

    return max(min(num, maximum), minimum)


def partStatus(partID, parameter):
    dbconfig = read_db_config()
    try:
        conn = MySQLConnection(**dbconfig)
        cursor = conn.cursor()
        sql = "SELECT DISTINCT R.stringValue FROM PartParameter R WHERE (R.name = '{}') AND (R.part_id = {})".format(
            parameter, partID)
        cursor.execute(sql)
        partStatus = cursor.fetchall()

        if partStatus == []:
            part = "Unknown"
        else:
            part = str(partStatus[0])[2:-3]
        return part

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
        distrbs = cursor.fetchall()
        unique = []
        d = []
        distributor = []

        for distributor in distrbs:
            if distributor[0] not in unique and distributor[0] in preferred:
                unique.append(distributor[0])
                d.append(distributor)
        return d

    except UnicodeEncodeError as err:
        print(err)

    finally:
        cursor.close()
        conn.close()


def labelsetup():
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    rows = 4
    cols = 3
    margin = 4  # In mm
    labelWidth = (210 - 2 * margin) / cols
    labelHeight = (297 - 2 * margin) / rows
    pdf.add_page()
    return (labelWidth, labelHeight, pdf)


def picksetup(BOMname, dateBOM, timeBOM):
    pdf2 = FPDF(orientation='L', unit='mm', format='A4')
    margin = 10  # In mm
    pdf2.add_page()
    pdf2.set_auto_page_break(1, 4.0)
    pdf2.set_font('Courier', 'B', 9)
    pdf2.multi_cell(80, 10, BOMname, align="L", border=0)
    pdf2.set_auto_page_break(1, 4.0)
    pdf2.set_font('Courier', 'B', 9)

    pdf2.set_xy(90, 10)
    pdf2.multi_cell(30, 10, dateBOM, align="L", border=0)

    pdf2.set_xy(120, 10)
    pdf2.multi_cell(30, 10, timeBOM, align="L", border=0)

    pdf2.set_font('Courier', 'B', 7)
    pdf2.set_xy(5, 20)
    pdf2.multi_cell(10, 10, "Line", border=1)

    pdf2.set_xy(15, 20)
    pdf2.multi_cell(40, 10, "Ref", border=1)

    pdf2.set_xy(55, 20)
    pdf2.multi_cell(95, 10, "Part", border=1)

    pdf2.set_xy(150, 20)
    pdf2.multi_cell(10, 10, "Stock", border=1)

    pdf2.set_xy(160, 20)
    pdf2.multi_cell(50, 10, "P/N", border=1)

    pdf2.set_xy(210, 20)
    pdf2.multi_cell(50, 10, "Location", border=1)

    pdf2.set_xy(260, 20)
    pdf2.multi_cell(10, 10, "Qty", border=1)

    pdf2.set_xy(270, 20)
    pdf2.multi_cell(10, 10, "Pick", border=1)
    return (pdf2)


def makepick(line, pdf2, pos):

    index = ((pos - 1) % 16) + 1

    pdf2.set_font('Courier', 'B', 6)
    pdf2.set_xy(5, 20 + 10 * index)  # Line Number
    pdf2.multi_cell(10, 10, str(pos), align="C", border=1)

    pdf2.set_xy(15, 20 + 10 * index)  # Blank RefDes box
    pdf2.multi_cell(40, 10, "", align="L", border=1)

    pdf2.set_xy(15, 20 + 10 * index)  # RefDes
    pdf2.multi_cell(40, 5, line[6], align="L", border=0)

    pdf2.set_xy(55, 20 + 10 * index)  # Blank Part box
    pdf2.multi_cell(95, 10, '', align="L", border=1)

    pdf2.set_font('Courier', 'B', 8)
    pdf2.set_xy(55, 20 + 10 * index)  # Part name
    pdf2.multi_cell(95, 5, line[1], align="L", border=0)

    pdf2.set_font('Courier', '', 6)
    pdf2.set_xy(55, 24 + 10 * index)  # Part Description
    pdf2.multi_cell(95, 5, line[0][:73], align="L", border=0)

    pdf2.set_xy(150, 20 + 10 * index)
    pdf2.multi_cell(10, 10, str(line[5]), align="C", border=1)  # Stock

    pdf2.set_xy(160, 20 + 10 * index)
    pdf2.multi_cell(50, 10, '', align="C", border=1)  # Blank cell

    pdf2.set_xy(160, 23.5 + 10 * index)
    pdf2.multi_cell(50, 10, line[2], align="C", border=0)  # PartNum

    pdf2.set_xy(172, 21 + 10 * index)
    if line[2] != "":
        pdf2.image('assets/barcodes/' + line[2] + '.png', h=6)  # PartNum BC

    pdf2.set_xy(210, 20 + 10 * index)
    pdf2.multi_cell(50, 10, '', align="C", border=1)  # Blank cell

    pdf2.set_xy(210, 23.5 + 10 * index)
    pdf2.multi_cell(50, 10, line[3], align="C", border=0)  # Location

    pdf2.set_xy(223, 21 + 10 * index)
    if line[3] != "":
        pdf2.image(
            'assets/barcodes/' + line[3][1:] + '.png', h=6)  # Location BC

    pdf2.set_font('Courier', 'B', 8)
    pdf2.set_xy(260, 20 + 10 * index)
    pdf2.multi_cell(10, 10, line[4], align="C", border=1)  # Qty

    pdf2.set_xy(270, 20 + 10 * index)
    pdf2.multi_cell(10, 10, "", align="L", border=1)

    pdf2.set_xy(273, 23 + 10 * index)
    if line[3] != "":
        pdf2.multi_cell(4, 4, "", align="L", border=1)

    if index % 16 == 0:
        pdf2.add_page()


def makelabel(label, labelCol, labelRow, lblwidth, lblheight, pdf):
    '''
    Take label info and make a label at position defined by row & column
    '''

    lineHeight = 3
    intMargin = 7

    labelx = int((lblwidth * (labelCol % 3)) + intMargin)
    labely = int((lblheight * (labelRow % 4)) + intMargin)

    pdf.set_auto_page_break(1, 4.0)

    pdf.set_font('Courier', 'B', 9)
    pdf.set_xy(labelx, labely)
    pdf.multi_cell(
        lblwidth - intMargin, lineHeight, label[0], align="L", border=0)
    pdf.set_font('Courier', '', 8)
    pdf.set_xy(labelx, labely + 10)
    pdf.cell(lblwidth, lineHeight, label[1], align="L", border=0)
    pdf.image('assets/barcodes/' + label[1] + '.png', labelx, labely + 13, 62,
              10)
    pdf.set_xy(labelx, labely + 25)
    pdf.cell(lblwidth, lineHeight, 'Part no: ' + label[2], align="L", border=0)

    pdf.set_xy(labelx + 32, labely + 25)
    pdf.cell(
        lblwidth, lineHeight, 'Location: ' + label[3], align="L", border=0)
    pdf.image('assets/barcodes/' + label[2] + '.png', labelx, labely + 28, 28,
              10)
    pdf.image('assets/barcodes/' + label[3][1:] + '.png', labelx + 30,
              labely + 28, 32, 10)

    pdf.set_xy(labelx, labely + 40)
    pdf.cell(
        lblwidth, lineHeight, 'Quantity: ' + label[4], align="L", border=0)
    pdf.image('assets/barcodes/' + label[4] + '.png', labelx, labely + 43, 20,
              10)

    pdf.set_font('Courier', 'B', 8)
    pdf.set_xy(labelx + 25, labely + 46)
    pdf.multi_cell(35, lineHeight, label[9], align="L", border=0)

    pdf.set_xy(labelx, labely + 56)
    pdf.multi_cell(
        lblwidth - intMargin, lineHeight, label[6], align="L", border=0)

    if (labelCol == 2) & ((labelRow + 1) % 4 == 0):
        pdf.add_page()


def getTable(partID, q, bcolour, row):
    '''
    There are 8 columns of manuf data /availability and pricing starts at col 9
    Pricing in quanties of 1, 10, 100 - so use log function
    background colour already set so ignored.
    '''

    index = int(log10(q)) + 9
    tbl = ""
    minPrice = 999999
    classtype = ''
    pricingExists = False

    fn = "./assets/web/" + str(partID) + ".csv"
    # If file is empty st_size should = 0 BUT file always contains exactly 1 byte ...
    if os.stat(fn).st_size == 1:  # File (almost) empty ...
        tbl = "<td colspan = " + str(
            len(preferred)
        ) + " class ='lineno' '><b>No data found from preferred providers</b></td>"
        return tbl, voidMatrix, voidMatrix, voidMatrix

    try:
        minData = np.genfromtxt(fn, delimiter=",")

        if np.ndim(minData) == 1:  # If only one line, nanmin fails
            minPrice = minData[index]
        else:
            minPrice = np.nanmin(minData[:, index])

    except (UserWarning, ValueError, IndexError) as error:
        print(
            "ATTENTION ", error
        )  # Just fails when empty file or any other error, returning no data
        tbl = "<td colspan = " + str(
            len(preferred)
        ) + " class ='lineno'><b>No data found from preferred providers</b></td>"
        return tbl, voidMatrix, voidMatrix, voidMatrix

    csvFiles = open(fn, "r")
    compPrefPrices = list(csv.reader(csvFiles))

    line = ""
    line2 = ""

    n = len(preferred)
    _costRow = [0] * n
    _coverageRow = [0] * n
    _countRow = [0] * n

    # line += "<form>"
    for d, dist in enumerate(preferred):

        line += "<td"
        terminated = False
        low = 0
        magnitude = 0
        i = 0
        for _comp in compPrefPrices:
            price = ""
            priceLine = ""
            try:
                if _comp[0] in dist:
                    try:
                        price = str("{0:2.2f}".format(float(_comp[index])))
                        dispPrice = price
                    except:
                        ValueError
                        price = "-"  # DEBUG "-"

                    try:
                        priceLine = str("{0:2.2f}".format(
                            q * float(_comp[index])))
                        calcPL = priceLine
                    except:
                        ValueError
                        priceLine = "-"  # DEBUG "-"
                        calcPL = "0.0"

                    if i == 0:  # 1st row only being considered
                        try:
                            _costRow[d] = q * float(_comp[index])
                        except:
                            ValueError
                            _costRow[d] = 0.0

                        _coverageRow[d] = 1
                        _countRow[d] = q

                        try:
                            _moq = int(_comp[3])
                        except:
                            ValueError
                            _moq = 999999

                        if bcolour == 'rgb(238, 232, 213)':
                            classtype = 'ambig'
                        else:
                            classtype = 'mid'

                        if _comp[index].strip() == str(minPrice):
                            pricingExists = True
                            classtype = 'min'
                            line += " class = '" + classtype + "'>"
                            line += " <input id = '" + str(d) + "-"+str(row)+"' type='radio' name='" + str(row) + "' value='" + \
                                calcPL + "' checked >"
                        else:
                            pricingExists = True
                            line += " class = '" + classtype + "'>"
                            line += " <input id='" + str(d) + "-"+str(row)+"' type='radio' name='" + str(row) + "' value='" + \
                                calcPL + "'>"
                            line += "<label for=" + str(d) + "></label>"

                        line += "&nbsp; <b><a href = '" + _comp[4] + "'>" + _comp[1] + "</a></b><br>"

                        if price == "-":
                            line += "<p class ='null' style= 'padding:5px;'>  Ea:"
                            line += "<span style='float: right; text-align: right;'>"
                            line += "<b >" + price + "</b> "
                            line += _comp[8]
                            price = "0"
                        else:
                            line += "<p style= 'padding:5px;'>  Ea:"
                            line += "<span style='float: right; text-align: right;'>"
                            line += "<b>" + price + "</b> "
                            line += _comp[8]
                        line += "</span>"

                        if priceLine == "-":
                            line += "<p class ='null' style= 'padding:5px;'>  Line:"
                            line += "<span style='float: right; text-align: right;'>"
                            line += "<b >" + priceLine + "</b> "
                            line += _comp[8]
                            priceline = "0"
                        else:
                            line += "<p style= 'padding:5px;'>  Line:"
                            line += "<span style='float: right; text-align: right;'>"
                            line += "<b>" + priceLine + "</b> "
                            line += _comp[8]
                        line += "</span>"

                        line += "<p style= 'padding:5px;'> MOQ: "
                        line += "<span style='float: right; text-align: right;'><b>" + _comp[3] + "</b>"

                        if int(q) >= _moq:  # MOQ satisfied
                            line += " &nbsp; <span class = 'icon'>🔹&nbsp; </span></span><p>"
                        else:
                            line += " &nbsp; <span class = 'icon'>🔺&nbsp; </span></span><p>"

                        line += "<p style= 'padding:5px;'> Stock:"
                        line += "<span style='float: right; text-align: right;'><b>"
                        line += _comp[2] + "</b>"
                        if int(q) <= int(_comp[2]):  # Stock satisfied
                            line += " &nbsp; <span class = 'icon'>🔹&nbsp; </span></span><p>"
                        else:
                            line += " &nbsp; <span class = 'icon'>🔸&nbsp; </span></span><p>"

                        P1 = ""
                        P2 = ""

                        magnitude = 10**(index - 9)

                        if _moq == 999999:
                            low = q
                            next = 10 * magnitude
                            column = int(log10(low) + 9)
                        elif _moq > magnitude:
                            low = _moq
                            next = 10**(ceil(log10(low)))
                            column = int(log10(low) + 10)
                        else:
                            low = magnitude
                            next = 10 * magnitude
                            column = int(log10(low) + 9)

                        try:
                            if float(_comp[column]) > 1:
                                P1 = str("{0:2.2f}".format(
                                    float(_comp[column])))
                            else:
                                P1 = str("{0:3.3f}".format(
                                    float(_comp[column])))
                        except:
                            ValueError
                            P1 = "-"

                        line += "<p style='text-align: left; padding:5px;'>" + str(
                            low) + " +"
                        line += "<span style='float: right; text-align: right;'><b>" + P1 + "</b>&nbsp; " + _comp[8] + "</span>"

                        if column <= 12:
                            try:
                                if float(_comp[column]) > 1:
                                    P2 = str("{0:2.2f}".format(
                                        float(_comp[column + 1])))
                                else:
                                    P2 = str("{0:3.3f}".format(
                                        float(_comp[column + 1])))
                            except:
                                ValueError
                                P2 = "-"
                            line += "<p style='text-align: left; padding:5px;'>" + str(
                                next) + " +"
                            line += "<span style='float: right; text-align: right;'><b>" + P2 + "</b>&nbsp; " + _comp[8] + "</span>"

                    else:  # Nasty kludge - relly need to iterate through these to get best deal
                        if i == 1:
                            line += "<br><br><br><p style= 'padding:5px;'><b> Alternatives</b><br>"
                        try:
                            if _comp[index].strip() == str(minPrice):
                                if classtype != 'min':
                                    line += "<div class = 'min'>"

                            price = str("{0:2.2f}".format(float(_comp[index])))
                            priceLine = str("{0:2.2f}".format(
                                q * float(_comp[index])))
                            line += "<p style= 'padding:5px;'><b><a href = '" + _comp[4] + "' > " + _comp[1] + " </b></a><br>"
                            line += "<p style= 'padding:5px;'> Ea:  <b>"
                            line += price + "</b> " + _comp[8]
                        except:
                            ValueError
                            line += "<p style= 'padding:5px;'><b><a href = '" + _comp[4] + "' > " + _comp[1] + " </b></a><br>"
                            line += " Pricing N/A"

                    i += 1

                    # FIXME This needs to count number of instances
                    if i <= 3:
                        terminated = True
                    else:
                        terminated = False
            except IndexError:
                pass

        if not terminated:
            line += ">"
            _costRow[d] = 0
            _countRow[d] = 0
            _coverageRow[d] = 0

        if pricingExists:
            line += "<p style='padding:5px;'>"
        pricingExists = False

        line += "</td>"

    # line += "</form>"
    tbl += line
    return tbl, _costRow, _coverageRow, _countRow


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
        return (4)

    # Get currency rates from European Central Bank
    # Fall back on cached cached rates
    try:
        c = CurrencyConverter(
            'http://www.ecb.europa.eu/stats/eurofxref/eurofxref.zip')
    except:
        URLError
        c = CurrencyConverter()
        return (8)

    # Remove invalid characters

    partIn = partIn.replace("/", "-")

    path = partIn.replace(" ", "")

    web = str("./assets/web/" + path + ".html")

    Part = partIn

    webpage = open(web, "w")

    combo = False
    if " " in partIn:
        # Possible Manufacturer/Partnumber combo. The Octopart mpn search does not include manufacturer
        # Split on space and assume that left part is Manufacturer and right is partnumber.
        # Mark as comboPart.
        combo = True
        comboManf, comboPart = partIn.split(" ")

    aside = open("./assets/web/tmp.html", "w")

    htmlHeader = """
    <!DOCTYPE html>
    <html lang = 'en'>
    <meta charset="utf-8">

    <head>
      <html lang="en">
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

    ##################

    if bean:
        #
        url = "http://octopart.com/api/v3/parts/search"
        url += '?apikey=' + apikey
        url += '&q="' + Part + '"'
        url += '&include[]=descriptions'
        url += '&include[]=imagesets'
        # url += '&include[]=specs'
        # url += '&include[]=datasheets'
        url += '&country=GB'

    elif combo:
        #
        url = "http://octopart.com/api/v3/parts/match"
        url += '?apikey=' + apikey
        url += '&queries=[{"brand":"' + comboManf + \
            '","mpn":"' + comboPart + '"}]'
        url += '&include[]=descriptions'
        url += '&include[]=imagesets'
        url += '&include[]=specs'
        url += '&include[]=datasheets'
        url += '&country=GB'

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
    response = json.loads(data.decode('utf8'))

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

            try:
                brand = str(item['brand']['name'])
            except:
                IndexError
                brand = ""

            # Get image (if present). Also need to get attribution for Octopart licensing
            try:
                image = str(item['imagesets'][0]['medium_image'].get(
                    'url', None))
            except:
                IndexError
                image = ""

            try:
                credit = str(item['imagesets'][0].get('credit_string', None))
            except:
                IndexError
                credit = ""

            webpage.write(
                "<div class='content' id = 'thumbnail'><table class = 'table2'><tr><td style = 'width:100px;'><img src='"
                + image + "' alt='thumbnail'></td><td><h2>" + brand + " " + partNum + "</h2><h4>" +
                description + "</h4></td></tr><tr><td style = 'color:#aaa;'>Image: " + credit +"</td><td></td></tr></table></div>")



            specfile = open("./assets/web/" + path, 'w')
            specfile.write(image)

            aside.write(
                "<div class = 'aside'><table class='table table-striped'><thead>")
            aside.write("<th>Characteristic</th><th>Value</th></thead><tbody>")


            for spec in item['specs']:
                parm = item['specs'][spec]['metadata']['name']
                try:
                    if type(item['specs'][spec]['display_value']):
                        val = str(item['specs'][spec]['display_value'])
                except:
                    IndexError
                    val = "Not Listed by Manufacturer"

                parameter = (("{:34} ").format(parm))
                value = (("{:40}").format(val))

                print(("| {:30} : {:120} |").format(parameter, value))
                aside.write("<tr><td>" + parameter + "</td><td>" + value +
                            "</td></tr>")
            print(('{:_<162}').format(""))

            aside.write("</tbody></table><table class='table table-striped'>")
            aside.write(
                "<thead><th>Datasheets</th><th>Date</th><th>Pages</th></thead><tbody>"
            )

            for d, datasheet in enumerate(item['datasheets']):
                if d == 1:
                    specfile.write(',' + datasheet['url'])
                try:
                    if (datasheet['metadata']['last_updated']):
                        dateUpdated = (
                            datasheet['metadata']['last_updated'])[:10]
                    else:
                        dateUpdated = "Unknown"
                except:
                    IndexError
                    dateUpdated = "Unknown"

                if datasheet['attribution']['sources'] is None:
                    source = "Unknown"
                else:
                    source = datasheet['attribution']['sources'][0]['name']

                try:
                    numPages = str(datasheet['metadata']['num_pages'])
                except:
                    TypeError
                    numPages = "-"

                documents = ((
                    "| {:30.30} {:11} {:12} {:7} {:7} {:1} {:84.84} |").format(
                        source, " Updated: ", dateUpdated, "Pages: ", numPages,
                        "", datasheet['url']))
                print(documents)
                aside.write("<tr><td><a href='" + datasheet['url'] + "'> " +
                            source + " </a></td><td>" + dateUpdated +
                            "</td><td>" + numPages + "</td></tr>")
            # if loop:
            #     webpage.write("<table class='table table-striped'>")
            # else:
            #     webpage.write("<p> No Octopart results found </>")

    # Header row here

    webpage.write(
        "<div class ='main'><table><thead><th>Seller</th><th>SKU</th><th>Stock</th><th>MOQ</th><th>Package</th><th>Currency</th><th>1</th><th>10</th><th>100</th><th>1000</th><th>10000</th></thead><tbody>"
    )
    count = 0

    for result in response['results']:

        stockfile = open("./assets/web/" + path + '.csv', 'w')

        for item in result['items']:
            if count == 0:
                print(('{:_<162}').format(""))
                print(
                    ("| {:24} | {:19} | {:>9} | {:>7} | {:11} | {:5} ").format(
                        "Seller", "SKU", "Stock", "MOQ", "Package",
                        "Currency"),
                    end="")
                print(
                    ("|  {:>10}|  {:>10}|  {:>10}|  {:>10}|  {:>10}|").format(
                        "1", "10", "100", "1000", "10000"))
                print(('{:-<162}').format(""), end="")

                count += 1

            # Breaks at 1, 10, 100, 1000, 10000
            for offer in item['offers']:
                loop = 0
                _seller = offer['seller']['name']
                _sku = (offer['sku'])[:19]
                _stock = offer['in_stock_quantity']
                _moq = str(offer['moq'])
                _productURL = str(offer['product_url'])
                _onOrderQuant = offer['on_order_quantity']
                _onOrderETA = offer['on_order_eta']
                _factoryLead = offer['factory_lead_days']
                _package = str(offer['packaging'])
                _currency = str(offer['prices'])

                if _moq == "None":
                    _moq = '-'

                if _package == "None":
                    _package = "-"

                if not _factoryLead or _factoryLead == "None":
                    _factoryLead = "-"
                else:
                    _factoryLead = int(int(_factoryLead) / 7)

                if _seller in preferred:
                    data = str(_seller) + ", " + str(_sku) + ", " + \
                        str(_stock) + ", " + str(_moq) + ", " + str(_productURL) + ", " +\
                        str(_factoryLead) + ", " + str(_onOrderETA) + ", " + str(_onOrderQuant) +\
                        ", " + str(locale)
                    stockfile.write(data)

                print()

                print(
                    ("| {:24.24} | {:19} | {:>9} | {:>7} | {:11} |").format(
                        _seller, _sku, _stock, _moq, _package),
                    end="")
                line = "<tr><td>" + _seller + "</td><td><a target='_blank' href=" + str(
                    offer['product_url']) + ">" + str(
                        offer['sku']) + "</a></td><td>" + str(
                            _stock) + "</td><td>" + str(
                                _moq) + "</td><td>" + _package + "</td>"
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
                                    c.convert(breaks[1], currency, locale), 6)
                            else:
                                points[4] = points[3]

                        for i in range(0, 5):
                            print(("  {:>10.5}|").format(points[i]), end="")
                            if points[i] == '-':
                                webpage.write("<td>-</td>")
                            elif float(points[i]) >= 1:
                                pp = str("{0:.2f}".format(points[i]))
                                webpage.write("<td>" + pp + "</td>")
                            else:
                                webpage.write("<td>" + str(points[i]) +
                                              "</td>")

                            if _seller in preferred:
                                stockfile.write(", " + str(points[i]))
                    webpage.write("</tr>")
                    if _seller in preferred:
                        stockfile.write("\n")
                if not valid:
                    print("          |", end="")
                    webpage.write("<td></td>")
                    for i in range(0, 5):
                        print(("  {:>10.5}|").format(points[i]), end="")
                        webpage.write("<td>" + str(points[i]) + "</td>")
                        if _seller in preferred:
                            stockfile.write(", " + str(points[i]))
                    webpage.write("</tr>")
                    stockfile.write("\n")
                    valid = False
    try:
        if _seller in preferred:
            stockfile.write("\n")
    except:
        UnboundLocalError
        stockfile.write("\n")

    webpage.write("</tbody></table></div>")
    print()
    print(('{:=<162}').format(""))

    aside.close()
    side = open("./assets/web/tmp.html", "r")

    aside = side.read()

    webpage.write(aside + "</tbody></table>")
    webpage.write("</div></body></html>")
    return


################################################################################
#   Further setup or web configuration here
#
#
################################################################################

compliance = {
    'Compliant': 'assets/img/rohs-logo.png',
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

lblwidth, lblheight, pdf = labelsetup()
pdf2 = picksetup(projectName, dateBOM, timeBOM)

labelRow = 0
labelCol = 0

run = 0

web = open("assets/web/temp.html", "w")
accounting = open("assets/web/accounting.html", "w")
missing = open("assets/web/missing.csv", "w")
under = open("assets/web/under.csv", "w")
under.write('Name,Description,Quantity,Stock,Min Stock\n')

htmlHeader = """
<!DOCTYPE html>
<meta charset="utf-8">
<html lang = 'en'>

<head>
<link rel="stylesheet" href="assets/css/web.css">
<title>Kicad 2 PartKeepr</title>

<script src="https://code.jquery.com/jquery-3.3.1.min.js" integrity="sha256-FgpCb/KJQlLNfOu91ta32o/NMZxltwRo8QtmkMRdAu8=" crossorigin="anonymous"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/2.4.0/Chart.min.js"></script>

</head>

<body>
    <div class="header">
            <h1 class="header-heading">Kicad2PartKeepr</h1>
    </div>
    <div class="nav-bar">
            <ul class="nav">
                <li><a href='labels.pdf' download='labels.pdf'>Labels</a></li>
                <li><a href='picklist.pdf' download='picklist.pdf'>Pick List</a></li>
                <li><a href='assets/web/missing.csv' download='missing.csv'>Missing parts list</a></li>
                <li><a href='assets/web/under.csv' download='under.csv'>Understock list</a></li>
                <li><a href='assets/web/""" + projectName + """_PK.csv'> PartKeepr import  </a></li>
            </ul>
    </div>

    <h2>Project name: """ + projectName + \
    "</h2><h3>Date:&nbsp; " + dateBOM + \
    "</h3><h3>Time:&nbsp; " + timeBOM + "</h3>"

htmlBodyHeader = '''
<table class='main' id='main'>
'''

labelCol = 0

htmlAccountingHeader = '''
<table class = 'accounting'>
'''

resistors = ["R_1206", "R_0805", "R_0603", "R_0402"]
capacitors = ["C_1206", "C_0805", "C_0603", "C_0402"]


def punctuate(value):
    if "." in value:
        multiplier = (value.strip()[-1])
        new_string = "_" + (re.sub("\.", multiplier, value))[:-1]
    else:
        new_string = "_" + value
    return (new_string)


def get_choice(possible):
    print(
        "More than one component in the PartKeepr database meets the criteria:"
    )
    i = 1
    for name, description, stockLevel, minStockLevel, averagePrice, partNum, storage_locn, PKid, Manufacturer in possible:
        #        print(i, " : ", name, " : ", description,
        #              " [Location] ", storage_locn, " [Part Number] ", partNum, " [Stock] ", stockLevel)
        print((
            "{:3} {:25} {:50.50} Location {:10} Manf {:12} HPN {:5} Stock {:5} Price(av) {:3s} {:6}"
        ).format(i, name, description, storage_locn, Manufacturer, partNum,
                 stockLevel, baseCurrency, averagePrice))
        #    subprocess.call(['/usr/local/bin/pyparts', 'specs', name])
        i = i + 1
    print("Choose which component to add to BOM (or 0 to defer)")

    while True:
        choice = int(input('> '))
        if choice == 0:
            return (possible)
        if choice < 0 or choice > len(possible):
            continue
        break

    i = 1
    for name, description, stockLevel, minStockLevel, averagePrice, partNum, storage_locn, PKid, Manufacturer in possible:
        possible = (name, description, stockLevel, minStockLevel, averagePrice,
                    partNum, storage_locn, PKid, Manufacturer)
        if i == choice:
            possible = (name, description, stockLevel, minStockLevel,
                        averagePrice, partNum, storage_locn, PKid,
                        Manufacturer)
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
                print(
                    "Insufficient parameters (Needs 3 or 4) e.g. R_0805_100K(_±5%)"
                )
                return ("0")

            c_case = component[1]
            c_value = convert_units(component[2])

            if (len(component)) == 4:
                c_characteristics = component[3]

                # A fully specified 'bean'
                sql = """SELECT DISTINCT P.name, P.description, P.stockLevel, P.minStockLevel, P.averagePrice, P.internalPartNumber, S.name, P.id, M.name
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
                sql = """SELECT DISTINCT P.name, P.description, P.stockLevel, P.minStockLevel, P.averagePrice, P.internalPartNumber, S.name, P.id, M.name
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

            sql = """SELECT DISTINCT P.name, P.description, P.stockLevel, P.minStockLevel, P.averagePrice, P.internalPartNumber, S.name, P.id, M.name
                 FROM Part P
                 JOIN StorageLocation S ON  S.id = P.storageLocation_id
                 LEFT JOIN PartManufacturer PM on PM.part_id = P.id
                 LEFT JOIN Manufacturer M on M.id = PM.manufacturer_id
                 WHERE P.name LIKE '%{}%'
                 GROUP BY P.id, M.id, S.id""".format(part_num)

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

    outfile = open(
        "./assets/web/" + filename + '_PK.csv',
        'w',
        newline='',
        encoding='utf-8')

    writeCSV = csv.writer(outfile, delimiter=',', lineterminator='\n')

    labelpdf = labelsetup()

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

        # Print to screen
        print(('{:=<162}').format(""))
        print(("| BOM Line number : {:3} {:136} |").format(count_BOMLine, ""))

        print(('{:_<162}').format(""))
        print(("| {:100} | {:13.13}         |                  | Req =   {:5}|"
               ).format(references, part, quantity))
        print(('{:_<162}').format(""))

        uniqueNames = []

        if n_components == 0:
            print(("| {:100} | {:21} | {:16} | {:12} |").format(
                "No matching parts in database", "", "", ""))
            print(('{:_<162}').format(""))
            octopartLookup(part, species)
            print('\n\n')

        else:
            for (name, description, stockLevel, minStockLevel, averagePrice,
                 partNum, storage_locn, PKid, Manufacturer) in component_info:
                ROHS = partStatus(PKid, 'RoHS')
                Lifecycle = partStatus(PKid, 'Lifecycle Status')
                # Can get rid of loop as never reset now
                print((
                    "| {:100} | Location = {:10} | Part no = {:6} | Stock = {:5}|"
                ).format(description, storage_locn, partNum, stockLevel))
                print(('{:_<162}').format(""))
                print(("| Manufacturing status: {} {:<136}|").format(
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
            for (name, description, stockLevel, minStockLevel, averagePrice,
                 partNum, storage_locn, PKid, Manufacturer) in component_info:
                ROHS = partStatus(PKid, 'RoHS')
                Lifecycle = partStatus(PKid, 'Lifecycle Status')
        if n_components != 0 and (quantity > stockLevel):
            count_LowStockLines = count_LowStockLines + 1
            background = lowstock  # Insufficient stock : pinkish
            under.write(name + ',' + description + ',' + str(quantity) + ',' +
                        str(stockLevel) + ',' + str(minStockLevel) + '\n')
        else:
            background = adequate  # Adequate stock : greenish

        countParts = countParts + quantity
        quantity = str(quantity)

        if run == 0:
            preferred = preferred.split(",")
            web.write("<th class = 'thmain' colspan = '6'> Component details</th>")
            web.write("<th class = 'thmain' colspan ='" + str(len(preferred) + 2) +
                      "'> Supplier details</th>")
            web.write(
                "<th class = 'stock' colspan ='5'> PartKeepr Stock</th></tr>")
            web.write("""
            <tr>

            <th class = 'thmain' style = 'width: 2%;'>No</th>
            <th class = 'thmain' style = 'width: 5%;'>Part</th>
            <th class = 'thmain' style = 'width : 13%;'>Description</th>
            <th class = 'thmain' style = 'width : 5%;'>References</th>
            <th class = 'thmain' style = 'width : 4%;'>RoHS</th>
            <th class = 'thmain' style = 'width : 4%;'>Lifecycle</th>
            """)

            i = 0

            while i < len(preferred):
                web.write("<th class = 'thmain' style = 'width : 10%;'>" + preferred[i] + "</th>")
                i += 1

            web.write("""<th class = 'thmain' style = 'width : 3%;'>Exclude</th>
            <th class = 'thmain' style = 'width : 3%;'>Qty</th>

            <th class = 'stock' style = 'width : 3%;'>Each</th>
            <th class = 'stock' style = 'width : 3%;'>Line</th>
            <th class = 'stock' style = 'width : 3%;'>Stock</th>
            <th class = 'stock' style = 'width : 3%;'>PAR</th>
            <th class = 'stock' style = 'width : 3%;'>Net</th>
            </tr></th>""")

            run = 1

# No PK components fit search criteria. Deal with here and drop before loop.
# Not ideal but simpler.
        if n_components == 0:
            count_NPKP = count_NPKP + 1
            averagePrice = 0
            background = nopkstock
            nopk = [
                '*** NON-PARTKEEPR PART ***', part, "", "", "", "", references,
                "", "", ""
            ]
            web.write("<tr style = 'background-color : " + background + ";'>")
            web.write("<td class = 'lineno'>" + str(count_BOMLine) + " </td>")
            web.write(
                "<td style = 'background-color : #fff; vertical-align:middle;'><img src = 'assets/img/noimg.png' alt='none'></td>"
            )
            web.write("<td class = 'partname' ><b>" + part + "</b>")
            web.write("<p class = 'descr'>Non PartKeepr component</td>")
            web.write("<td class = 'refdes'" + references + "</td>")
            web.write("<td class = 'ROHS'>NA</td>")
            web.write("<td class = 'ROHS'>NA</td>")
            for i, d in enumerate(preferred):
                web.write("<td></td>")
            web.write("<td class = 'stck'>-</td>")
            web.write("<td class = 'stck'> " + quantity + "</td>")
            web.write("<td class = 'stck'>-</td>")
            web.write("<td class = 'stck'>-</td>")
            web.write("<td class = 'stck'>-</td>")
            web.write("<td class = 'stck'>-</td>")
            web.write("<td class = 'stck'>-</td></tr>")

            makepick(nopk, pdf2, count_BOMLine)

            missing.write(part + ',' + quantity + ',' + references + '\n')
            name = "-"

        if n_components > 1:  # Multiple component fit search criteria - set brown background -
            # TODO Make colors configurable in ini
            background = multistock
            columns = str(len(preferred) + 13)
            if len(component_info) == 1:
                # More than one line exists but has been disambiguated at runtime
                web.write(
                    "<tr><td colspan='" + columns +
                    "' style = 'font-weight : bold; background-color : " +
                    background + ";'><nbsp><nbsp><nbsp><nbsp><nbsp>" +
                    str(n_components) +
                    " components meet the selection criteria but the following line was selected at runtime</td></tr>"
                )
            else:
                web.write(
                    "<tr style = 'background-color : " + background +
                    ";'><td colspan='" + columns +
                    "' style = 'font-weight : bold;'> <nbsp><nbsp><nbsp><nbsp><nbsp>The following "
                    + str(n_components) +
                    " components meet the selection criteria *** Use only ONE line *** </td></tr>"
                )

        i = 0
        for (name, description, stockLevel, minStockLevel, averagePrice,
             partNum, storage_locn, PKid, Manufacturer) in component_info:

            if n_components > 1:
                web.write("<tr style = 'background-color : " + background + ";'>")
            else:
                web.write("<tr>")

            web.write("<td class = 'lineno'>" + str(count_BOMLine) + " </td>")

            if i == 0:  # 1st line where multiple components fit search showing RefDes

                name_safe = name.replace("/", "-")
                name_safe = name.replace(" ", "")
                try:
                    f = open("./assets/web/" + name_safe, 'r')
                    imageref = f.read()
                except:
                    FileNotFoundError

                try:
                    imageref, datasheet = imageref.split(',')
                except ValueError:
                    datasheet = ''

                web.write("<td class = 'center'>")
                if imageref:
                    web.write("<img src = '" + imageref + "' alt='thumbnail' >")
                else:
                    web.write(
                        "<img src = 'assets/img/noimg.png' alt='none'>")

                imageref = ""

                web.write("</td>")

                web.write("<td class = 'partname'><a href = 'assets/web/" + name_safe +
                          ".html'>" + name + "</a>")

                if Manufacturer:
                    web.write("<p class='manf'>" + Manufacturer + "</p>")

                i = i + 1
                count_PKP = count_PKP + 1

            else:  # 2nd and subsequent lines where multiple components fit search showing RefDes

                web.write("<td>")

                name_safe = name.replace("/", "-")
                f = open("./assets/web/" + name_safe, 'r')
                imageref = f.read()

                if imageref:
                    web.write("<img src = '" + imageref + "' alt='thumbnail'>")

                web.write("<td><a href = 'assets/web/" + name_safe +
                          ".html'>" + name + "</a>")

                if Manufacturer:
                    web.write("<p class = 'manf'" + Manufacturer + "</>")

                invalidate_BOM_Cost = True

            lineCost = float(averagePrice) * int(quantity)
            if lineCost == 0:
                count_PWP += 1

            web.write("<p class = 'descr'>" + description + "</p>")

            if datasheet:
                web.write("<a href='" + datasheet +
                          "'><img class = 'partname' src = 'assets/img/pdf.png' alt='datasheet' ></a>")

            web.write("</td>")

            web.write("<td class = 'refdes'>" + references + "</td>")
            # TODO Probably could be more elegant

            if ROHS == "Compliant":
                ROHSClass = 'compliant'
                _cIndicator = "C_cp"
            elif ROHS == "Non-Compliant":
                ROHSClass = 'uncompliant'
                _cIndicator = "C_nc"
            else:
                ROHS = "Unknown"
                ROHSClass = 'unknown'
                _cIndicator = "C_nk"

            web.write("<td id ='" + _cIndicator + str(count_BOMLine) +
                      "' class = '" + ROHSClass + "'>" +
                      ROHS)

            if ROHS == "Compliant":
                web.write("<br><br><img src = 'assets/img/rohs-logo.png' alt='compliant'>")

            web.write("</td>")

            if Lifecycle == "Active":
                LifeClass = 'active'
                _pIndicator = "P_ac"
            elif Lifecycle == "EOL":
                LifeClass = 'eol'
                _pIndicator = "P_el"
            elif Lifecycle == "NA":
                Lifecycle = "Unknown"
                LifeClass = 'unknown'
                _pIndicator = "P_nk"
            else:
                LifeClass = 'unknown'
                _pIndicator = "P_nk"

            web.write("<td id = '" + _pIndicator + str(count_BOMLine) +
                      "' class = '" + LifeClass + "'>" + Lifecycle +
                      "</td>")

            # Part number exists, therefore generate bar code
            # Requires Zint >1.4 - doesn't seem to like to write to another directory.
            # Ugly hack - write to current directory and move into place.
            if partNum:
                part_no = partNum
                # part_no = (partNum[1:])
                subprocess.call([
                    '/usr/local/bin/zint', '--filetype=png', '--notext', '-w',
                    '10', '--height', '20', '-o', part_no, '-d', partNum
                ])
                os.rename(part_no + '.png',
                          'assets/barcodes/' + part_no + '.png')

# Storage location exists, therefore generate bar code. Ugly hack - my location codes start with
# a '$' which causes problems. Name the file without the leading character.
            if storage_locn != "":
                locn_trim = ""
                locn_trim = (storage_locn[1:])
                subprocess.call([
                    '/usr/local/bin/zint', '--filetype=png', '--notext', '-w',
                    '10', '--height', '20', '-o', locn_trim, '-d', storage_locn
                ])
                os.rename(locn_trim + '.png',
                          'assets/barcodes/' + locn_trim + '.png')

                table, priceRow, coverage, items = getTable(
                    name_safe, int(quantity), background, count_BOMLine)

                web.write(str(table))

                costMatrix = [sum(x) for x in zip(priceRow, costMatrix)]
                coverageMatrix = [
                    sum(x) for x in zip(coverage, coverageMatrix)
                ]
                countMatrix = [sum(x) for x in zip(items, countMatrix)]

                # NOTE Do I need to think about building matrix by APPENDING lines.

                avPriceFMT = str(('{:3s} {:0,.2f}').format(
                    baseCurrency, averagePrice))
                linePriceFMT = str(('{:3s} {:0,.2f}').format(
                    baseCurrency, lineCost))
                bomCost = bomCost + lineCost

                averagePrice = str('{:0,.2f}').format(averagePrice)
                lineCost = str('{:0,.2f}').format(lineCost)
                web.write("<td class = 'lineno'><input type='checkbox' name='ln"+str(count_BOMLine)+"' value='"+str(count_BOMLine)+"'></td>")
                web.write("<td class = 'stck'> " + quantity + "</td>")
                web.write("<td class = 'stck'><b>" + averagePrice + "</b> " + baseCurrency + "</td>")
                web.write("<td class = 'stck'><b>" + lineCost + "</b> " + baseCurrency + "</td>")

                if int(quantity) >= stockLevel:
                    qtyCol = '#dc322f'
                else:
                    qtyCol = "#859900"

                if (int(stockLevel) - int(quantity)) <= int(minStockLevel):
                    minLevelCol = '#dc322f'
                elif (int(stockLevel) -
                      int(quantity)) <= int(minStockLevel) * 1.2:
                    minLevelCol = "#cb4b16"
                else:
                    minLevelCol = "#859900"

                web.write(
                    "<td class = 'stck' style = 'font-weight : bold; text-align: right;'>"
                    + str(stockLevel) + "</td>")

                web.write(
                    "<td class = 'stck' style = 'font-weight : bold; text-align: right;'>"
                    + str(minStockLevel) + "</td>")

                web.write(
                    "<td class = 'stck' style = 'font-weight : bold; text-align: right; color:"
                    + minLevelCol + "'>" +
                    str(int(stockLevel) - int(quantity)) + "</td>")

            else:
                # No storage location
                web.write("<td class = 'stck'>NA</td>")

            web.write('</tr>\n')

            # Make labels for packets (need extra barcodes here)

            name = name.replace("/", "-")  # Deal with / in part names
            subprocess.call([
                '/usr/local/bin/zint', '--filetype=png', '--notext', '-w',
                '10', '--height', '20', '-o', name, '-d', name
            ])

            os.rename(name + '.png', 'assets/barcodes/' + name + '.png')
            subprocess.call([
                '/usr/local/bin/zint', '--filetype=png', '--notext', '-w',
                '10', '--height', '20', '-o', quantity, '-d', quantity
            ])
            os.rename(quantity + '.png',
                      'assets/barcodes/' + quantity + '.png')

            # Write out label pdf
            lb2 = []
            lb2.append(description[:86].upper())
            lb2.append(name)
            lb2.append(part_no)
            lb2.append(storage_locn)
            lb2.append(quantity)
            lb2.append(stockLevel)
            lb2.append(references)
            lb2.append(dateBOM)
            lb2.append(timeBOM)
            lb2.append(filename)

            makelabel(lb2, labelCol, labelRow, lblwidth, lblheight, pdf)
            makepick(lb2, pdf2, count_BOMLine)

            labelCol = labelCol + 1
            if labelCol == 3:
                labelRow += 1
                labelCol = 0

# Prevent variables from being recycled
            storage_locn = ""
            partNum = ""
            part_no = ""
        qty = str(int(quantity) / numBoards)
        writeCSV.writerow([references, name, qty])
        references = ""
        name = ""
        quantity = ""

# Write out footer for webpage
    web.write("<tr id = 'supplierTotal'>")
    web.write("<td style = 'border: none;'></td>" * 4)
    web.write("<td colspan ='2'>Total for Supplier</td>")
    for p, d in enumerate(preferred):
        web.write("<td style='text-align: right;'><b><span id='t" + str(p) +
                  "'></span></b><span style='text-align: right;'> " +
                  baseCurrency + "</span></td>")
    web.write("<td style = 'border: none;'></td>" * 7)
    web.write("</tr>")

    web.write("<tr id = 'linesCoverage'>")
    web.write("<td style = 'border: none;'></td>" * 4)
    web.write("<td colspan ='2'>Coverage by Lines</td>")
    for p, d in enumerate(preferred):
        coveragePC = str(int(round((coverageMatrix[p] / count_BOMLine) * 100)))
        web.write("<td style='text-align: right;'><span class ='value'>" +
                  coveragePC + "</span>%</td>")
    web.write("<td style = 'border: none;'></td>" * 7)
    web.write("</tr>")

    web.write("<tr id = 'itemsCoverage'>")
    web.write("<td style = 'border: none;'></td>" * 4)
    web.write("<td colspan ='2'>Coverage by Items</td>")
    for p, d in enumerate(preferred):
        countPC = str(int(round((countMatrix[p] / countParts) * 100)))
        web.write("<td style='text-align: right;'>" + countPC + "%</td>")
    web.write("<td style = 'border: none;'></td>" * 7)
    web.write("</tr>")

    web.write("<tr id = 'linesSelected'>")
    web.write("<td style = 'border: none;'></td>" * 4)
    web.write("<td colspan ='2'>Lines selected</td>")
    for p, d in enumerate(preferred):
        web.write("<td style='text-align: right;'><span id ='cz" + str(p) +
                  "'></span></td>")
    web.write("<td style = 'border: none;'></td>" * 7)
    web.write("</tr>")

    web.write("<tr id = 'linesPercent'>")
    web.write("<td style = 'border: none;'></td>" * 4)
    web.write("<td colspan ='2'>Percent selected (by lines)</td>")
    for p, d in enumerate(preferred):
        web.write("<td style='text-align: right;'><span id ='cc" + str(p) +
                  "'></span>%</td>")
    web.write("<td style = 'border: none;'></td>" * 7)
    web.write("</tr>")

    web.write("<tr id = 'pricingSelected'>")
    web.write("<td style = 'border: none;'></td>" * 4)
    web.write("<td colspan ='2'>Total for Selected</td>")
    for p, d in enumerate(preferred):
        web.write("<td style='text-align: right;'><b><span id='c" + str(p) +
                  "'></span></b><span style='text-align: right;'> " +
                  baseCurrency + "</span></td>")
    web.write("<td style = 'border: none;'></td>" * 7)
    web.write("</tr>")

    # I am really struggling with using js arrays .....
    script = '''
    <script>
    $(document).ready(function() {

        bomlines = $('#BOMLines').text();
        bomlines = Number(bomlines);

        numboards = $('#numboards').text();
        numboards = Number(numboards);

        excluded = $(':checkbox:checked').length;
        excluded = Number(excluded);

        bomlines = bomlines - excluded;

        providers = Number($("#providers th").length) - 15;

        var compliance_cp = $("[id*=C_cp]").length;
        var compliance_nc = $("[id*=C_nc]").length;
        var compliance_nk = $("[id*=C_nk]").length;

        var production_ip = $("[id*=P_ac]").length;
        var production_el = $("[id*=P_el]").length;
        var production_nk = $("[id*=P_nk]").length;

        line0 = $('#linesCoverage').find("td:eq(5)").find('span.value').text();
        line1 = $('#linesCoverage').find("td:eq(6)").find('span.value').text();
        line2 = $('#linesCoverage').find("td:eq(7)").find('span.value').text();
        line3 = $('#linesCoverage').find("td:eq(8)").find('span.value').text();
        line4 = $('#linesCoverage').find("td:eq(9)").find('span.value').text();

        var total = 0;
        var totalDisp = 0;
        var totalPerBoard = 0;


        var t0 = 0;
        var t1 = 0;
        var t2 = 0;
        var t3 = 0;
        var t4 = 0;

        var c0 = 0;
        var c1 = 0;
        var c2 = 0;
        var c3 = 0;
        var c4 = 0;

        var cz0 = $("[id*=0-]:radio:checked:not(:hidden)").length;
        var cz1 = $("[id*=1-]:radio:checked:not(:hidden)").length;
        var cz2 = $("[id*=2-]:radio:checked:not(:hidden)").length;
        var cz3 = $("[id*=3-]:radio:checked:not(:hidden)").length;
        var cz4 = $("[id*=4-]:radio:checked:not(:hidden)").length;

        var cc0 = Math.round(Number(cz0) / bomlines * 100);
        var cc1 = Math.round(Number(cz1) / bomlines * 100);
        var cc2 = Math.round(Number(cz2) / bomlines * 100);
        var cc2 = Math.round(Number(cz2) / bomlines * 100);
        var cc3 = Math.round(Number(cz3) / bomlines * 100);
        var cc4 = Math.round(Number(cz4) / bomlines * 100);

        $(":radio:checked").each(function() {
            total += Number(this.value);
            total = Math.round(total * 100) / 100;
            totalDisp = total.toFixed(2);
            totalPerBoard = (total/numboards).toFixed(2);
        });


        $("[id*=0-]:radio:checked").each(function() {
            c0 += Number(this.value);
            c0 = Math.round(c0 * 100) / 100;
        });

        $("[id*=0-]:radio").each(function() {
            t0 += Number(this.value);
            t0 = Math.round(t0 * 100) / 100;
        });

        $("[id*=1-]:radio:checked").each(function() {
            c1 += Number(this.value);
            c1 = Math.round(c1 * 100) / 100;
        });

        $("[id*=1-]:radio").each(function() {
            t1 += Number(this.value);
            t1 = Math.round(t1 * 100) / 100;
        });

        $("[id*=2-]:radio:checked").each(function() {
            c2 += Number(this.value);
            c2 = Math.round(c2 * 100) / 100;
        });

        $("[id*=2-]:radio").each(function() {
            t2 += Number(this.value);
            t2 = Math.round(t2 * 100) / 100;
        });

        $("[id*=3-]:radio:checked").each(function() {
            c3 += Number(this.value);
            c3 = Math.round(c3 * 100) / 100;
        });

        $("[id*=3-]:radio:checked").each(function() {
            t3 += Number(this.value);
            t3 = Math.round(t3 * 100) / 100;
        });

        $("[id*=4-]:radio:checked").each(function() {
            c4 += Number(this.value);
            c4 = Math.round(c4 * 100) / 100;
        });

        $("[id*=4-]:radio:checked").each(function() {
            t4 += Number(this.value);
            t4 = Math.round(t4 * 100) / 100;
        });

        var t0Disp = t0.toFixed(2);
        var t1Disp = t1.toFixed(2);
        var t2Disp = t2.toFixed(2);
        var t3Disp = t3.toFixed(2);
        var t4Disp = t4.toFixed(2);

        var c0Disp = c0.toFixed(2);
        var c1Disp = c1.toFixed(2);
        var c2Disp = c2.toFixed(2);
        var c3Disp = c3.toFixed(2);
        var c4Disp = c4.toFixed(2);

        $("#total").text(totalDisp);

        $("#totalPerBoard").text(totalPerBoard);

        $("#c0").text(c0Disp);
        $("#c1").text(c1Disp);
        $("#c2").text(c2Disp);
        $("#c3").text(c3Disp);
        $("#c4").text(c4Disp);

        $("#cz0").text(cz0);
        $("#cz1").text(cz1);
        $("#cz2").text(cz2);
        $("#cz3").text(cz3);
        $("#cz4").text(cz4);

        $("#cc0").text(cc0);
        $("#cc1").text(cc1);
        $("#cc2").text(cc2);
        $("#cc3").text(cc3);
        $("#cc4").text(cc4);

        $("#t0").text(t0Disp);
        $("#t1").text(t1Disp);
        $("#t2").text(t2Disp);
        $("#t3").text(t3Disp);
        $("#t4").text(t4Disp);



        var options =
        {
            responsive: false,
            maintainAspectRatio: false,
            scales: {
                yAxes: [{
                    ticks: {
                        beginAtZero:true
                    }
                }]
            }
        };

        var ctx = document.getElementById('pricing').getContext('2d');
        var datapoints = [t0, t1, t2, t3, t4];
        var chart = new Chart(ctx, {
            // The type of chart we want to create
            type: 'bar',

            // The data for our dataset
            data: {
                labels: ["Newark", "Farnell", "DigiKey", "RS Components", "Mouser"],
                datasets: [{
                    label: "Pricing",
                    backgroundColor: [
                            'rgb(133, 153, 0)',
                            'rgb(42, 161, 152)',
                            'rgb(38, 139, 210)',
                            'rgb(108, 113, 196)',
                            'rgb(211, 54, 130)',
                            'rgb(220, 50, 47)'
                        ],
                    borderColor: 'rgb(88, 110, 117)',
                    data: datapoints,
                }]
            },
            options: {
                legend: {
                    display: false
                },
                title: {
                    display: true,
                    text: 'Cost of BOM by provider'
                }
            }
        });

        var ctx2 = document.getElementById('coverage').getContext('2d');
        var coverage = [line0, line1, line2, line3, line4];
        var chart = new Chart(ctx2, {
            type: 'bar',
            data: {
                labels: ["Newark", "Farnell", "DigiKey", "RS Components", "Mouser"],
                datasets: [{
                    label: "Coverage",
                    backgroundColor: [
                            'rgb(133, 153, 0)',
                            'rgb(42, 161, 152)',
                            'rgb(38, 139, 210)',
                            'rgb(108, 113, 196)',
                            'rgb(211, 54, 130)',
                            'rgb(220, 50, 47)'
                    ],
                    borderColor: 'rgb(88, 110, 117)',
                    data: coverage,
            }]
            },
            options: {
                legend: {
                    display: false
                },
                title: {
                    display: true,
                    text: 'Coverage of BOM by Lines (%)'
                }
            }
        });

        var ctx3 = document.getElementById('compliance').getContext('2d');
        var compliant = [compliance_cp, compliance_nc, compliance_nk]
        var data1 = {
            datasets: [{
                backgroundColor: [
                '#2aa198',
                '#586e75',
                '#073642'],
                data: compliant,
        }],
            labels: [
            'Compliant',
            'Non-Compliant',
            'Unknown'
        ]
        };

        var compliance = new Chart(ctx3, {
            type: 'doughnut',
            data: data1,
            options: {
                title: {
                    display: true,
                    text: 'ROHS Compliance'
                }
            }
        });

        var ctx4 = document.getElementById('lifecycle').getContext('2d');
        var life = [production_ip, production_el, production_nk]

        var data2 = {
            datasets: [{
                backgroundColor: [
                '#2aa198',
                '#d33682',
                '#073642'],
                data: life
        }],
            labels: [
            'Active',
            'EOL',
            'Unknown'
        ]
        };

        var lifecycle = new Chart(ctx4, {
            type: 'doughnut',
            data: data2,
            options: {
                title: {
                    display: true,
                    text: 'Production Status'
                }
            }
        });

    });


    $(":radio, :checkbox").on("change", function() {

        var t0 = 0;
        var t1 = 0;
        var t2 = 0;
        var t3 = 0;
        var t4 = 0;

        var c0 = 0;
        var c1 = 0;
        var c2 = 0;
        var c3 = 0;
        var c4 = 0;

        $('#main tr').filter(':has(:checkbox)').find('radio,.td,.min,.mid,.radbut,.min a,.min p,.mid a,.mid p, .ambig, .ambig a, .ambig p, .null p').removeClass("deselected");
        $('#main tr').filter(':has(:checkbox)').find('.icon').show();
        $('#main tr').filter(':has(:checkbox)').find(':radio').show();
        $('#main tr').filter(':has(:checkbox:checked)').find('radio,.td,.min,.mid,.radbut,.min a,.min p,.mid a,.mid p,.ambig,.ambig a,.ambig p,.null p').addClass("deselected");
        $('#main tr').filter(':has(:checkbox:checked)').find('.icon').hide();
        $('#main tr').filter(':has(:checkbox:checked)').find(':radio').hide();

        var compliance_cp = $("[id*=C_cp]").length;
        var compliance_nc = $("[id*=C_nc]").length;
        var compliance_nk = $("[id*=C_nk]").length;

        var production_ip = $("[id*=P_ac]").length;
        var production_el = $("[id*=P_el]").length;
        var production_nk = $("[id*=P_nk]").length;

        line0 = $('#linesCoverage').find("td:eq(5)").find('span.value').text();
        line1 = $('#linesCoverage').find("td:eq(6)").find('span.value').text();
        line2 = $('#linesCoverage').find("td:eq(7)").find('span.value').text();
        line3 = $('#linesCoverage').find("td:eq(8)").find('span.value').text();
        line4 = $('#linesCoverage').find("td:eq(9)").find('span.value').text();


        bomlines = $('#BOMLines').text();
        bomlines = Number(bomlines);

        numboards = $('#numboards').text();
        numboards = Number(numboards);

        excluded = $(':checkbox:checked').length;
        excluded = Number(excluded);

        bomlines = bomlines - excluded;

        var total = 0;
        var totalDisp = 0;
        var totalPerBoard = 0;



        var cz0 = $("[id*=0-]:radio:checked:not(:hidden)").length;
        var cz1 = $("[id*=1-]:radio:checked:not(:hidden)").length;
        var cz2 = $("[id*=2-]:radio:checked:not(:hidden)").length;
        var cz3 = $("[id*=3-]:radio:checked:not(:hidden)").length;
        var cz4 = $("[id*=4-]:radio:checked:not(:hidden)").length;

        var cc0 = Math.round(Number(cz0) / bomlines * 100);
        var cc1 = Math.round(Number(cz1) / bomlines * 100);
        var cc2 = Math.round(Number(cz2) / bomlines * 100);
        var cc2 = Math.round(Number(cz2) / bomlines * 100);
        var cc3 = Math.round(Number(cz3) / bomlines * 100);
        var cc4 = Math.round(Number(cz4) / bomlines * 100);






        $(":radio:checked:not(:hidden)").each(function() {
            total += Number(this.value);
            total = Math.round(total * 100) / 100;
            totalDisp = total.toFixed(2);
            totalPerBoard = (total/numboards).toFixed(2);
        });

        $("[id*=0-]:radio:checked:not(:hidden)").each(function() {
            c0 += Number(this.value);
            c0 = Math.round(c0 * 100) / 100;
        });

        $("[id*=0-]:radio:not(:hidden)").each(function() {
            t0 += Number(this.value);
            t0 = Math.round(t0 * 100) / 100;
        });

        $("[id*=1-]:radio:checked:not(:hidden)").each(function() {
            c1 += Number(this.value);
            c1 = Math.round(c1 * 100) / 100;
        });

        $("[id*=1-]:radio:not(:hidden)").each(function() {
            t1 += Number(this.value);
            t1 = Math.round(t1 * 100) / 100;
        });

        $("[id*=2-]:radio:checked:not(:hidden)").each(function() {
            c2 += Number(this.value);
            c2 = Math.round(c2 * 100) / 100;
        });

        $("[id*=2-]:radio:not(:hidden)").each(function() {
            t2 += Number(this.value);
            t2 = Math.round(t2 * 100) / 100;
        });

        $("[id*=3-]:radio:checked:not(:hidden)").each(function() {
            c3 += Number(this.value);
            c3 = Math.round(c3 * 100) / 100;
        });

        $("[id*=3-]:radio:checked:not(:hidden)").each(function() {
            t3 += Number(this.value);
            t3 = Math.round(t3 * 100) / 100;
        });

        $("[id*=4-]:radio:checked:not(:hidden)").each(function() {
            c4 += Number(this.value);
            c4 = Math.round(c4 * 100) / 100;
        });

        $("[id*=4-]:radio:checked:not(:hidden)").each(function() {
            t4 += Number(this.value);
            t4 = Math.round(t4 * 100) / 100;
        });

        var t0Disp = t0.toFixed(2);
        var t1Disp = t1.toFixed(2);
        var t2Disp = t2.toFixed(2);
        var t3Disp = t3.toFixed(2);
        var t4Disp = t4.toFixed(2);

        var c0Disp = c0.toFixed(2);
        var c1Disp = c1.toFixed(2);
        var c2Disp = c2.toFixed(2);
        var c3Disp = c3.toFixed(2);
        var c4Disp = c4.toFixed(2);

        $("#total").text(totalDisp);

        $("#totalPerBoard").text(totalPerBoard);

        $("#c0").text(c0Disp);
        $("#c1").text(c1Disp);
        $("#c2").text(c2Disp);
        $("#c3").text(c3Disp);
        $("#c4").text(c4Disp);

        $("#cz0").text(cz0);
        $("#cz1").text(cz1);
        $("#cz2").text(cz2);
        $("#cz3").text(cz3);
        $("#cz4").text(cz4);

        $("#cc0").text(cc0);
        $("#cc1").text(cc1);
        $("#cc2").text(cc2);
        $("#cc3").text(cc3);
        $("#cc4").text(cc4);

        $("#t0").text(t0Disp);
        $("#t1").text(t1Disp);
        $("#t2").text(t2Disp);
        $("#t3").text(t3Disp);
        $("#t4").text(t4Disp);

        var options = {};

        var ctx = document.getElementById('pricing').getContext('2d');
        var datapoints = [t0, t1, t2, t3, t4];
        var chart = new Chart(ctx, {
            // The type of chart we want to create
            type: 'bar',

            // The data for our dataset
            data: {
                labels: ["Newark", "Farnell", "DigiKey", "RS Components", "Mouser"],
                datasets: [{
                    label: "Pricing",
                    backgroundColor: [
                            'rgb(133, 153, 0)',
                            'rgb(42, 161, 152)',
                            'rgb(38, 139, 210)',
                            'rgb(108, 113, 196)',
                            'rgb(211, 54, 130)',
                            'rgb(220, 50, 47)'
                        ],
                    borderColor: 'rgb(88, 110, 117)',
                    data: datapoints,
                }]
            },
            options: {
                legend: {
                    display: false
                },
                title: {
                    display: true,
                    text: 'Cost of BOM by provider'
                }
            }
        });

        var ctx2 = document.getElementById('coverage').getContext('2d');
        var coverage = [line0, line1, line2, line3, line4];
        var chart = new Chart(ctx2, {
            type: 'bar',
            data: {
                labels: ["Newark", "Farnell", "DigiKey", "RS Components", "Mouser"],
                datasets: [{
                    label: "Coverage",
                    backgroundColor: [
                            'rgb(133, 153, 0)',
                            'rgb(42, 161, 152)',
                            'rgb(38, 139, 210)',
                            'rgb(108, 113, 196)',
                            'rgb(211, 54, 130)',
                            'rgb(220, 50, 47)'
                    ],
                    borderColor: 'rgb(88, 110, 117)',
                    data: coverage,
            }]
            },
            options: {
                legend: {
                    display: false
                },
                title: {
                    display: true,
                    text: 'Coverage of BOM by Lines (%)'
                }
            }
        });
    });
    </script>'''

    footer = '''
<br><p>
<a href="http://jigsaw.w3.org/css-validator/check/referer">
    <img style="border:0;width:88px;height:31px"
        src="http://jigsaw.w3.org/css-validator/images/vcss"
        alt="Valid CSS!" />
</a>
</p>'''
    web.write("</table></body>" + script + "</html>")

    # Now script has run, construct table with part counts & costs etc.
    bomCostDisp = str(('{:3s} {:0,.2f}').format(baseCurrency, bomCost))
    bomCostBoardDisp = str(('{:3s} {:0,.2f}').format(baseCurrency, bomCost/numBoards))
    currency = str(('{:3s}').format(baseCurrency))

    accounting.write("<tr class = 'accounting'>")
    accounting.write("<td colspan = 2> Total number of boards </td>")
    accounting.write("<td style='text-align:right;' id = 'numboards'>" + str(numBoards) +
                     "</td>")

    accounting.write("</tr> <tr class = 'accounting'>")
    accounting.write("<td colspan = 2> Total parts </td>")
    accounting.write("<td style='text-align:right;'>" + str(countParts) +
                     "</td>")

    accounting.write("</tr> <tr class = 'accounting'>")
    accounting.write("<td colspan = 2>BOM Lines </td>")
    accounting.write("<td style='text-align:right;'><span id='BOMLines'>" +
                     str(count_BOMLine) + "</span></td>")

    accounting.write("</tr> <tr class = 'accounting'>")
    accounting.write("<td colspan = 2>  Non-PartKeepr Parts  </td>")
    accounting.write("<td style='text-align:right;' id='npkParts'>" + str(count_NPKP) +
                     "</td>")

    accounting.write("</tr> <tr class = 'accounting'>")
    accounting.write("<td colspan = 2> PartKeepr Parts</td>")
    accounting.write("<td style='text-align:right;' id='pkParts'>" + str(count_PKP) +
                     "</td>")

    accounting.write("</tr> <tr class = 'accounting'>")
    accounting.write("<td colspan = 2> Parts without pricing info </td>")
    accounting.write("<td style='text-align:right;'>" + str(count_PWP) +
                     "</td>")

    accounting.write("</tr> <tr class = 'accounting'>")
    accounting.write("<td colspan = 2> Low Stock </td>")
    accounting.write("<td style='text-align:right;'>" +
                     str(count_LowStockLines) + "</td></tr>")

    accounting.write("<tr class = 'accounting'><td></td><td></td><td></td></tr>")
    accounting.write("<tr class = 'accounting'><td></td>")
    accounting.write("<td style='text-align:right; font-weight: bold;'> Total </td>")
    accounting.write("<td style='text-align:right; font-weight: bold;'> Per board</td>")

    accounting.write("</tr> <tr class = 'accounting'>")
    accounting.write("<td> Inventory prices </td>")
    if not invalidate_BOM_Cost:
        accounting.write("<td style='text-align:right'>" + bomCostDisp +
                         " </td>")
        accounting.write("<td style='text-align:right' >" + bomCostBoardDisp +
                         "</td>")
    else:
        accounting.write("<td class = 'accounting'>BOM price not calculated </td> ")

    accounting.write("</tr> <tr class = 'accounting'>")
    accounting.write("<td>Price from selected </td><td style='text-align:right'>" +
                     currency + " <span id='total'></span></td>")
    accounting.write("<td style='text-align:right'>" + currency + " <span id='totalPerBoard'></span></td></table>")

    accounting.write(
        "<div class = 'canvas'><canvas id='pricing'></canvas></div>")
    accounting.write(
        "<div class = 'canvas'><canvas id='coverage'></canvas></div>")
    accounting.write(
        "<div class = 'canvas2'><canvas id='compliance'></canvas></div>")
    accounting.write(
        "<div class = 'canvas2'><canvas id='lifecycle'></canvas></div>")

# Assemble webpage

pdf.output('labels.pdf', 'F')
pdf2.output('picklist.pdf', 'F')

web = open("assets/web/temp.html", "r")
web_out = open("webpage.html", "w")

accounting = open("assets/web/accounting.html", "r")
accounting = accounting.read()

htmlBody = web.read()

web_out.write(htmlHeader)
web_out.write(htmlAccountingHeader + accounting)
web_out.write(htmlBodyHeader + htmlBody)


print('Starting server...')

PORT = 8109

Handler = http.server.SimpleHTTPRequestHandler
httpd = socketserver.TCPServer(("", PORT), Handler, bind_and_activate=False)
httpd.allow_reuse_address = True
httpd.server_bind()
httpd.server_activate()
print('Running server...')
print('Type Ctr + C to halt')
try:
    webbrowser.open('http://localhost:'+str(PORT)+'/webpage.html')
    httpd.serve_forever()
except KeyboardInterrupt:
    httpd.socket.close()
    httpd.shutdown()
    httpd.server_close()

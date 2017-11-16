#!/usr/bin/env python3
#coding=utf-8

from mysql.connector import MySQLConnection, Error
from python_mysql_dbconfig import read_db_config
import sys
import csv
import re
import os
import subprocess
import webbrowser
import decimal
from datetime import datetime

ctx = decimal.Context()
ctx.prec = 12

file_name = sys.argv[1]
projectName, ext = file_name.split(".")
print(projectName)

try:
    os.makedirs('./KP/barcodes')
except OSError:
    pass


def float_to_str(f):
    d1 = ctx.create_decimal(repr(f))
    return format(d1, 'f')


def convert_units(num):
    factors = ["G", "M", "K", "k", "R", "", ".", "m", "u", "n", "p"]
    conversion = {'G': '1000000000', 'M': '1000000', 'K': '1000', 'k': '1000', 'R': '1', '.': '1', '': '1', 'm': '0.001', "u": '0.000001', 'n': '0.000000001', 'p': '0.000000000001'}
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


dateBOM = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

run = 0
web = open("./KP/temp.html", "w")


labels = open("./KP/labels.html", "w")
accounting = open("./KP/accounting.html", "w")
missing = open("./KP/missing.tsv", "w")
under = open("./KP/under.tsv", "w")
under.write('name\tpartNum\tQuantity\tStock\tMin Stock\n')

htmlHeader = """
<!DOCTYPE html PUBLIC '-//W3C//DTD HTML 4.01//EN'>
<meta charset="utf-8">
<html>
<head>
<style>
body {
    background-color: #ffffff
}

h1 {
    color: #ffffff;
    text-align: right;
    font-family: Arial, Helvetica, sans-serif;
    background : -webkit-linear-gradient(left, #ffffff, #667399);
    padding:20px 40px 20px 40px;
    text-shadow: 2px 2px #33394d;
}

h2, h3 {
    font-family: Arial, Helvetica, sans-serif;
    padding:0px 0px 0px 0px;
}


p {
    font-family: Arial, Helvetica, sans-serif;
    font-weight : bold;
    font-size : 12pt;
    padding   : 0px 40px 0px 0px;
    text-align  : right;
}

.main, th, td {
    font-family: Arial, Helvetica, sans-serif;
    font-size : 12pt;
    border :    1px solid black;
    padding :   5px;
    spacing : 0px;
    border-collapse: collapse;
}
.accounting {
    font-family: Arial, Helvetica, sans-serif;
    font-size : 12pt;
    font-weight : bold;
    border :    1px solid black;
    padding :   5px;
    spacing : 0px;
    border-collapse: collapse;
    width : 35%;
    float   : left;
}
</style>
<title>Kicad 2 PartKeepr</title>
</head>
"""

htmlIntro = "<body><h1> KiCad 2 PartKeepr </h1><br><h2> Project name: " + projectName + "</h2><h3>" + dateBOM + "</h3>"

htmlBodyHeader = """
<br><br>
<table class = "main">
"""


label_header = """
<!DOCTYPE html PUBLIC '-//W3C//DTD HTML 4.01//EN'>
<meta charset="utf-8">
<html>
<head>
<title>Kicad to PartKeepr</title>
</head>
<body>
<style>
p {
font-family: "Courier New", Courier, monospace;
font-size : 10pt;
margin-top: 6px;
margin-bottom: 6px;
}
.right {
font-family: "Courier New", Courier, monospace;
font-size : 10pt;
float:right;
text-align: right;
margin-top: 6px;
margin-bottom: 6px;
width: 50%;
display:inline;
}
.left {
font-family: "Courier New", Courier, monospace;
font-size : 10pt;
float:left;
text-align: left;
margin-top: 6px;
margin-bottom: 6px;
width: 50%;
display:inline;
}
.heavy {
font-size: 12pt;
font-family: "Courier New", Courier, monospace;
margin-top: 8px;
margin-bottom: 8px;
font-weight: bold;
text-align: left;
}
</style>
<table style="width:100%" table border='0' cellpadding='35' cellspacing='0'>
"""

labels.write(label_header)
label_cnt = 0


htmlAccountingHeader = """
<table class = "accounting">
"""


htmlLinks = """<br><p><a href='labels.html'> Labels  </a>
<br><p><a href='missing.tsv'> Missing parts list  </a>
<br><p><a href='under.tsv'> Understock items list </a>
<br><p><a href= """
htmlLinks += (projectName + "_PK.csv> PartKeepr import  </a>")


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
    print ("More than one component in the PartKeepr database meets the criteria:")
    i = 1
    for name, description, stockLevel, minStockLevel, averagePrice, partNum, storage_locn in possible:
        print (i, " : ", name, " : ", description)
        i = i + 1
    print ("Choose which component to add to BOM (or 0 to defer)")

    while True:
        choice = int(input('>'))
        if choice == 0:
            return (possible)
        if choice < 0 or choice > len(possible):
            continue
        break

    i = 1
    for name, description, stockLevel, minStockLevel, averagePrice, partNum, storage_locn in possible:
        possible = (name, description, stockLevel, minStockLevel, averagePrice, partNum, storage_locn)
        if i == choice:
            possible = (name, description, stockLevel, minStockLevel, averagePrice, partNum, storage_locn)
            print ("Selected :")
            print (possible[0], " : ", possible[1])
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
                print ("Insufficient parameters (Needs 3 or 4) e.g. R_0805_100K_±5%")
                return ("0")

            c_case = component[1]
            c_value = convert_units(component[2])

            if (len(component)) == 4:
                c_characteristics = component[3]

                # A fully specified 'bean'
                sql = """SELECT P.name, P.description, P.stockLevel, P.minStockLevel, P.averagePrice, P.internalPartNumber, S.name
                        FROM Part P
                        JOIN PartParameter R ON R.part_id = P.id
                        JOIN StorageLocation S ON  S.id = P.storageLocation_id
                        WHERE
                        (R.name = 'Case/Package' AND R.stringValue='{}') OR
                        (R.name = '{}' AND R.normalizedValue = '{}') OR
                        (R.name = '{}' AND R.stringValue = '%{}')
                        GROUP BY P.id
                        HAVING
                        COUNT(DISTINCT R.name)=3""".format(c_case, quality, c_value, variant, c_characteristics)
            else:
                # A partially specified 'bean'
                sql = """SELECT P.name, P.description, P.stockLevel, P.minStockLevel, P.averagePrice, P.internalPartNumber, S.name
                        FROM Part P
                        JOIN PartParameter R ON R.part_id = P.id
                        JOIN StorageLocation S ON  S.id = P.storageLocation_id
                        WHERE
                        (R.name = 'Case/Package' AND R.stringValue='{}') OR
                        (R.name = '{}' AND R.normalizedValue = '{}')
                        GROUP BY P.id
                        HAVING
                        COUNT(DISTINCT R.name)=2""".format(c_case, quality, c_value)
        else:

            sql = """SELECT P.name, P.description, P.stockLevel, P.minStockLevel, P.averagePrice, P.internalPartNumber, S.name
                 FROM Part P
                 JOIN StorageLocation S ON  S.id = P.storageLocation_id
                 WHERE P.name LIKE '%{}%'""".format(part_num)

        cursor.execute(sql)
        components = cursor.fetchall()
        return (components)

    except UnicodeEncodeError as err:
        print(err)

    finally:
        cursor.close()
        conn.close()


with open(file_name, newline='', encoding='utf-8') as csvfile:
    reader = csv.DictReader(csvfile, delimiter=',')
    headers = reader.fieldnames

    filename, file_extension = os.path.splitext(file_name)

    outfile = open("./KP/" + filename + '_PK.csv', 'w', newline='\n', encoding='utf-8')
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

        component_info = find_part(part)

        n_components = len(component_info)

        quantity = int(quantity)

# Print to screen - these could all do with neatening up...
        print(('{:_<141}').format(""))
        print (("|{:80} | {:13.13}         |                  | Req =   {:5}|").format(references, part, quantity))
        print(('{:_<141}').format(""))

        if n_components == 0:
            print("|No matching parts in database                                                    |                                                         |")
            print(('{:_<141}').format(""))
            print('\n')

        else:
            for (name, description, stockLevel, minStockLevel, averagePrice, partNum, storage_locn) in component_info:
                print(("|{:80} | Location = {:10} | Part no = {:6} | Stock = {:5}|").format(description, storage_locn, partNum, stockLevel))
                print(('{:_<141}').format(""))
            print('\n')


# More than one matching component exists - prompt user to choose
        if len(component_info) >= 2:
            component_info = get_choice(component_info)

        if quantity > stockLevel and n_components != 0:
            count_LowStockLines = count_LowStockLines + 1
            background = 'rgba(60, 0, 0, 0.2)'  # Insufficient stock : pinkish
            under.write(name + '\t' + partNum + '\t' + str(quantity) + '\t' + str(stockLevel) + '\t' + str(minStockLevel) + '\n')
        else:
            background = 'rgba(0, 60, 0, 0.2)'  # Adequate stock : greenish

        countParts = countParts + quantity
        quantity = str(quantity)


# Print header row with white background  - should move this somewwhere else....
        if run == 0:
            web.write("<tr style = 'background-color : white; font-weight : bold;'>")
            web.write("<th>References</th><th>Part</th><th>Description</th><th>Stock</th><th>Part Number</th><th>Location</th><th>Qty</th><th>Each</th><th>Line</th>")
            run = 1

# No PK components fit search criteria. Deal with here and drop before loop.
# Not ideal but simpler.
        if n_components == 0:
            count_NPKP = count_NPKP + 1
            averagePrice = 0

            background = 'rgba(0, 60, 60, 0.4)' # Green Blue background
            web.write("<tr style = 'background-color : "+background+";'>")
            web.write("<td style = 'font-weight : bold'>"+references+"</td>")
            web.write("<td >"+part+"</td>")
            web.write("<td > Non PartKeepr component</td>")
            web.write("<td style = 'font-weight : bold'>NA</td>")
            web.write("<td style = 'background-color : white' align='center'>NA</td>")
            web.write("<td style = 'background-color : white' align='center'>NA</td>")
            web.write("<td style = 'font-weight : bold; background-color : white;'> " + quantity + "</td>")
            web.write("<td style = 'background-color : white;'> £-:-- </td>")
            web.write("<td style = 'background-color : white;'> £-:-- </td></tr>")
            missing.write(references + '\t' + part + '\t' + quantity + '\n')
            name = "-"

        if n_components > 1:  # Multiple component fit search criteria - set brown background
            background = 'rgba(60, 60, 0, 0.4)'


        i = 0
        for (name, description, stockLevel, minStockLevel, averagePrice, partNum, storage_locn) in component_info:
            web.write("<tr style = 'background-color : "+background+";'>")
            if i == 0:  # 1st line where multiple components fit search showing RefDes
                web.write("<td style = 'font-weight : bold'>"+references+"</td>")
                web.write("<td >"+part+"</td>")

                i = i + 1
                count_PKP = count_PKP + 1
            else:  # 2nd and subsequent lines where multiple components fit search showing RefDes
                web.write("<td colspan='2' style = 'font-weight : bold;'> *** ATTENTION *** Multiple sources available *** Use only ONE line *** </td>")

            lineCost = float(averagePrice) * int(quantity)
            if lineCost == 0:
                count_PWP += 1
            web.write("<td >"+description+"</td>")
            web.write("<td style = 'font-weight : bold'>"+str(stockLevel)+"</td>")

# Part number exists, therefore generate bar code
# Requires Zint >1.4 - doesn't seem to like to write to another directory.
# Ugly hack - write to current directory and move into place.
            if partNum != "":
                part_no = ""
                part_no = (partNum[1:])
                subprocess.call(['/usr/local/bin/zint', '--filetype=png', '-w', '10', '--height', '20', '-o', part_no, '-d', partNum])
                os.rename (part_no+'.png', 'KP/barcodes/'+part_no+'.png')
                web.write("<td style = 'background-color : white' align='center'><img src = barcodes/"+part_no+".png ></td>")
            else:
                # No Part number
                web.write("<td style = 'background-color : white' align = 'center'> NA </td>")

# Storage location exists, therefore generate bar code. Ugly hack - my location codes start with
# a '$' which causes problems. Name the file without the leading character.
            if storage_locn != "":
                locn_trim = ""
                locn_trim = (storage_locn[1:])
                web.write("<td style = 'background-color : white' align='center'><img src = barcodes/"+locn_trim+".png ></td>")
                subprocess.call(['/usr/local/bin/zint', '--filetype=png', '-w', '10', '--height', '20', '-o', locn_trim, '-d', storage_locn])
                os.rename (locn_trim+'.png', 'KP/barcodes/'+locn_trim+'.png')
            else:
                # No storage location
                web.write("<td style = 'background-color : white' align = 'center'> NA </td>")
            avPriceFMT = str(('£{:0,.2f}').format(averagePrice))
            linePriceFMT = str(('£{:0,.2f}').format(lineCost))
            bomCost = bomCost + lineCost
            web.write("<td style = 'font-weight : bold; background-color : white;'> "+quantity+"</td>")
            web.write("<td style = ' background-color : white;'> " + avPriceFMT + "</td>")
            web.write("<td style = ' background-color : white;'> " + linePriceFMT + "</td></tr>")
            web.write('\n')

# Make labels for packets (need extra barcodes here)
            subprocess.call(['/usr/local/bin/zint', '--filetype=png', '-w', '10', '--height', '20', '-o', name, '-d', name])
            os.rename (name+'.png', 'KP/barcodes/'+name+'.png')
            subprocess.call(['/usr/local/bin/zint', '--filetype=png', '-w', '10', '--height', '20', '-o', quantity, '-d', quantity])
            os.rename (quantity+'.png', 'KP/barcodes/'+quantity+'.png')

# Write out label webpage too
            labels.write("<td><p class = 'heavy'>" + description[:64].upper())
            labels.write("<p>")
            labels.write("<p>" + name)
            labels.write("<p><img src = barcodes/" + name + ".png style='height:40px;'>")
            labels.write("<p>")
            labels.write("<div class = 'left'>Part number:" + part_no + "</div>")
            labels.write("<div class = 'right'>Location:" + storage_locn + "</div>")
            labels.write("<div class = 'left'><img src = barcodes/" + part_no + ".png style='height:40px;'></div>")
            labels.write("<div class = 'right'><img src = barcodes/" + locn_trim + ".png style='height:40px;'></div>")
            labels.write("<div class = 'left'>Quantity: " + quantity + "</div>")
            labels.write("<p><img src = barcodes/" + quantity + ".png style='height:40px;'></div>")
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
    accounting.write("<td>BOM Cost</td>")
    accounting.write("<td>" + bomCostDisp + "</td>")
    accounting.write(("</tr></table><p>"))

#Assemble webpage
web = open("./KP/temp.html", "r")
web_out = open("./KP/webpage.html", "w")

accounting = open("./KP/accounting.html", "r")
accounting = accounting.read()

htmlBody = web.read()

web_out.write(htmlHeader + htmlIntro)
web_out.write(htmlAccountingHeader + accounting + htmlLinks + "<br><br><br>")
web_out.write(htmlBodyHeader + htmlBody)

# Open webpage in default browser
webbrowser.open('file://' + os.path.realpath('./KP/webpage.html'))

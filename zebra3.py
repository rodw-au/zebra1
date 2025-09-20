import os, sys
import platform
import ctypes
from ctypes.wintypes import BYTE, DWORD, LPCWSTR
# import win32print
import csv
import configparser
from configparser import ConfigParser
import requests
import socket
from xml.dom import minidom
from tkinter import *
import tkinter.filedialog
import tkinter.messagebox
import string

# Global variables
g_printer = ""
g_filename = ""
g_apiurl = ""
g_apikey = ""
g_zpl = ""
g_csvname = ""
g_lpripaddr = '192.168.1.47'
g_lprport = 9100
g_uselpr = 1
g_sku = ''
g_name = ''
g_warehouse = ''
g_upc = ''
g_quantity = '1'
g_shippingweight = ''
g_shippingwidth = ''
g_cubicweight = ''
g_shippinglength = ''
g_shippingheight = ''
g_defaultprice = ''
g_promotionprice = ''
g_misc10 = ''
g_location = ''
config = ''

def clearglobals():
    global g_sku, g_name, g_warehouse, g_upc, g_shippingweight, g_shippingwidth
    global g_cubicweight, g_shippinglength, g_shippingheight, g_defaultprice
    global g_promotionprice, g_misc10, g_location, g_quantity
    g_sku = ''
    g_name = ''
    g_warehouse = ''
    g_upc = ''
    g_quantity = '1'
    g_shippingweight = ''
    g_shippingwidth = ''
    g_cubicweight = ''
    g_shippinglength = ''
    g_shippingheight = ''
    g_defaultprice = ''
    g_promotionprice = ''
    g_misc10 = ''
    g_location = ''

def printglobals():
    global g_sku, g_name, g_warehouse, g_upc, g_shippingweight, g_shippingwidth
    global g_cubicweight, g_shippinglength, g_shippingheight, g_defaultprice
    global g_promotionprice, g_misc10
    print("g_sku =", g_sku.strip())
    print("g_name =", g_name.strip())
    print("g_warehouse =", g_warehouse.strip())
    print("g_upc =", g_upc.strip())
    print("g_shippingweight =", g_shippingweight.strip())
    print("g_shippingwidth =", g_shippingwidth.strip())
    print("g_cubicweight =", g_cubicweight.strip())
    print("g_shippinglength =", g_shippinglength.strip())
    print("g_shippingheight =", g_shippingheight.strip())
    print("g_defaultprice =", g_defaultprice.strip())
    print("g_promotionprice =", g_promotionprice.strip())
    print("g_misc10 =", g_misc10)

def printable(input_str):
    return ''.join(filter(lambda x: x in string.printable, input_str))

headers = {
    'NETOAPI_ACTION': "GetItem",
    'NETOAPI_KEY': "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx",
    'Content-Type': "application/xml"
}

payload = """<?xml version="1.0" encoding="utf-8"?>
<GetItem>
    <Filter>
        <SKU>[SKU]</SKU>
        <OutputSelector>SKU</OutputSelector>
        <OutputSelector>Name</OutputSelector>
        <OutputSelector>DefaultPrice</OutputSelector>
        <OutputSelector>ShippingLength</OutputSelector>
        <OutputSelector>ShippingWidth</OutputSelector>
        <OutputSelector>ShippingHeight</OutputSelector>
        <OutputSelector>ShippingWeight</OutputSelector>
        <OutputSelector>CubicWeight</OutputSelector>
        <OutputSelector>UPC</OutputSelector>
        <OutputSelector>WarehouseLocations</OutputSelector>
        <OutputSelector>Misc10</OutputSelector>
    </Filter>
</GetItem>"""

def parsexml(tpayload, sku):
    global g_name, g_sku, g_shippingweight, g_shippingwidth, g_cubicweight
    global g_shippinglength, g_shippingheight, g_defaultprice, g_promotionprice
    global g_upc, g_apiurl, g_warehouse, g_misc10, g_location
    response = requests.request("POST", g_apiurl, data=tpayload, headers=headers)
    xmldoc = minidom.parseString(response.text)
    itemlist = xmldoc.getElementsByTagName('Item')
    for node in xmldoc.getElementsByTagName('Item'):
        try:
            name = node.getElementsByTagName('Name')[0]
        except:
            print("SKU <%s> Not found, aborted" % sku)
            return
        tsku = node.getElementsByTagName('SKU')[0]
        shippingweight = node.getElementsByTagName('ShippingWeight')[0]
        shippingwidth = node.getElementsByTagName('ShippingWidth')[0]
        cubicweight = node.getElementsByTagName('CubicWeight')[0]
        shippinglength = node.getElementsByTagName('ShippingLength')[0]
        shippingheight = node.getElementsByTagName('ShippingHeight')[0]
        defaultprice = node.getElementsByTagName('DefaultPrice')[0]
        upc = node.getElementsByTagName('UPC')[0]
        warehouse = node.getElementsByTagName('WarehouseLocations')[0]
        misc10 = node.getElementsByTagName('Misc10')[0]
        g_sku = tsku.firstChild.data.strip()
        g_name = name.firstChild.data.strip()
        g_shippingweight = shippingweight.firstChild.data.strip() if shippingweight.firstChild else '0.0'
        g_shippingwidth = shippingwidth.firstChild.data.strip() if shippingwidth.firstChild else '0.0'
        g_shippinglength = shippinglength.firstChild.data.strip() if shippinglength.firstChild else '0.0'
        g_shippingheight = shippingheight.firstChild.data.strip() if shippingheight.firstChild else '0.0'
        g_defaultprice = str(defaultprice.firstChild.data).strip() if defaultprice.firstChild else "0.00"
        g_upc = upc.firstChild.data.strip() if upc.firstChild else ''
        g_warehouse = warehouse.firstChild.data.strip() if warehouse.firstChild else ''
        g_misc10 = misc10.firstChild.data.strip() if misc10.firstChild else ''

def readConfig():
    global g_zpl, g_filename, g_printer, g_apiurl, g_apikey, g_csvname
    global g_uselpr, g_lpripaddr, g_lprport, config, root
    config = ConfigParser()
    config.read('zebra1.ini')
    g_printer = config.get('PRINTER', 'queue', fallback='')
    g_filename = config.get('PRINTER', 'label', fallback='')
    g_csvname = config.get('PRINTER', 'lastcsv', fallback='')
    g_apiurl = config.get('API', 'URL', fallback='')
    g_apikey = config.get('API', 'KEY', fallback='')
    g_lpripaddr = config.get('LPR', 'ipaddr', fallback='192.168.1.47')
    g_lprport = config.get('LPR', 'port', fallback='9100')
    g_uselpr = config.get('LPR', 'uselpr', fallback='1')
    print("g_filename =", g_filename)
    if g_filename:
        try:
            with open(g_filename, "r") as file:
                g_zpl = file.read()
        except:
            tkinter.messagebox.showwarning("Warning", f"Cannot open label format {g_filename}, check settings before printing", parent=root)
    else:
        tkinter.messagebox.showwarning("Warning", "No Label format selected in Setup", parent=root)

def writeConfig():
    global g_zpl, g_filename, g_printer, g_apiurl, g_apikey, g_csvname
    global config, g_uselpr, g_lpripaddr, g_lprport
    config.set('API', 'key', g_apikey)
    config.set('API', 'url', g_apiurl)
    config.set('PRINTER', 'lastcsv', g_csvname)
    config.set('PRINTER', 'label', g_filename)
    config.set('PRINTER', 'queue', g_printer)
    config.set('LPR', 'ipaddr', g_lpripaddr)
    config.set('LPR', 'port', g_lprport)
    config.set('LPR', 'uselpr', g_uselpr)
    try:
        with open('zebra1.ini', 'w') as configfile:
            config.write(configfile)
    except:
        tkinter.messagebox.showwarning("Warning", "Cannot write to INI file", parent=root)

def printCSV():
    global g_zpl, g_sku, g_name, g_warehouse, g_misc10, g_upc, g_defaultprice
    global g_quantity, g_csvname, g_shippingweight, g_location, root
    print("printCSV: root defined:", root is not None)  # Debug
    clearglobals()
    label = ''
    csvname = g_csvname
    if not csvname:
        tkinter.messagebox.showwarning("Warning", "No CSV file selected", parent=root)
        return
    # Check ZPL template for required placeholders
    required_placeholders = ['[SKU]', '[NAME]', '[MISC10]', '[UPC]', '[PRICE]', '[QTY]', '[WEIGHT]', '[WAREHOUSE]', '[LOCATION]']
    missing_placeholders = [ph for ph in required_placeholders if ph not in g_zpl]
    if missing_placeholders:
        tkinter.messagebox.showerror(
            "Error",
            f"ZPL template missing required placeholders: {', '.join(missing_placeholders)}. Please update {g_filename}.",
            parent=root
        )
        return
    column_mapping = {
        "NAME": "g_name",
        "UPC": "g_upc",
        "PRICE": "g_defaultprice",
        "MISC10": "g_misc10",
        "WAREHOUSE": "g_warehouse",
        "LOCATION": "g_location",
        "QUANTITY": "g_quantity",
        "SKU": "g_sku"
    }
    try:
        with open(csvname, 'r', encoding='utf-8') as csvfile:
            reader = csv.reader(csvfile, delimiter=',')
            headers = next(reader, None)
            if headers is None:
                tkinter.messagebox.showerror("Error", "CSV file is empty", parent=root)
                return
            headers = [h.upper() for h in headers]
            print("Headers:", headers)
            header_indices = {}
            found_columns = []
            for col_name in column_mapping:
                try:
                    header_indices[col_name] = headers.index(col_name)
                    found_columns.append(col_name)
                except ValueError:
                    header_indices[col_name] = None
            if not found_columns:
                tkinter.messagebox.showerror(
                    "Error",
                    f"No valid columns found in CSV. Expected: {', '.join(column_mapping.keys())}",
                    parent=root
                )
                return
            # Check that all CSV headers have corresponding ZPL placeholders
            missing_zpl_fields = [col for col in found_columns if f'[{col}]' not in g_zpl]
            if missing_zpl_fields:
                tkinter.messagebox.showerror(
                    "Warning",
                    f"CSV columns not found in ZPL template: {', '.join(missing_zpl_fields)}. Please update {g_filename}.",
                    parent=root
                )
            for row in reader:
                if not row:  # Skip empty rows
                    print("Skipping empty row")
                    continue
                print("row =", row)
                clearglobals()
                if header_indices["NAME"] is not None and len(row) > header_indices["NAME"]:
                    g_name = row[header_indices["NAME"]][:100] if len(row[header_indices["NAME"]]) > 100 else row[header_indices["NAME"]]
                if header_indices["UPC"] is not None and len(row) > header_indices["UPC"]:
                    g_upc = row[header_indices["UPC"]]
                if header_indices["PRICE"] is not None and len(row) > header_indices["PRICE"]:
                    try:
                        g_defaultprice = formatprice(row[header_indices["PRICE"]])
                    except ValueError:
                        g_defaultprice = ""
                        tkinter.messagebox.showwarning("Warning", f"Invalid price in row: {row}", parent=root)
                if header_indices["MISC10"] is not None and len(row) > header_indices["MISC10"]:
                    g_misc10 = row[header_indices["MISC10"]]
                if header_indices["WAREHOUSE"] is not None and len(row) > header_indices["WAREHOUSE"]:
                    g_warehouse = row[header_indices["WAREHOUSE"]]
                if header_indices["SKU"] is not None and len(row) > header_indices["SKU"]:
                    g_sku = row[header_indices["SKU"]]
                if header_indices["LOCATION"] is not None and len(row) > header_indices["LOCATION"]:
                    g_location = row[header_indices["LOCATION"]]
                if header_indices["QUANTITY"] is not None and len(row) > header_indices["QUANTITY"]:
                    g_quantity = row[header_indices["QUANTITY"]]
                label += (FormatLabel(g_zpl, g_sku, g_name, g_misc10, g_upc, g_defaultprice,
                                     g_quantity, g_shippingweight, g_warehouse, g_location) + "\n")
        if not label:  # Check if any labels were generated
            tkinter.messagebox.showwarning("Warning", "No valid data rows processed in CSV", parent=root)
        else:
            tkinter.messagebox.showinfo("Success", "CSV processed successfully", parent=root)
            print("Generated labels:\n", label)
            if g_uselpr:
                printlprlabel(label)
            else:
                printlabel(label)
    except Exception as e:
        tkinter.messagebox.showerror("Error", f"Failed to process CSV: {e}", parent=root)

def FormatLabel(label, sku, name, misc10, upc, price, quantity, weight, warehouse, location):
    lbl1 = label.replace('[SKU]', printable(sku))
    lbl2 = lbl1.replace('[NAME]', printable(name))
    lbl3 = lbl2.replace('[MISC10]', printable(misc10))
    lbl4 = lbl3.replace('[UPC]', printable(upc))
    lbl5 = lbl4.replace('[PRICE]', printable(price))
    lbl6 = lbl5.replace('[QTY]', printable(quantity))
    lbl7 = lbl6.replace('[WEIGHT]', printable(weight))
    lbl8 = lbl7.replace('[WAREHOUSE]', printable(warehouse))
    lbl9 = lbl8.replace('[LOCATION]', printable(location))
    return lbl9

def printlabel(lblfmt):
    global g_printer
    # Windows-specific code (unused on Linux with g_uselpr = 1)
    # if platform.system() == "Windows":
    #     p = win32print.OpenPrinter(g_printer)
    #     try:
    #         job = win32print.StartDocPrinter(p, 1, ("Neto Product Labels", None, "RAW"))
    #         try:
    #             win32print.StartPagePrinter(p)
    #             win32print.WritePrinter(p, lblfmt)
    #             win32print.EndPagePrinter(p)
    #         finally:
    #             win32print.EndDocPrinter(p)
    #     finally:
    #         win32print.ClosePrinter(p)
    pass  # Stub for non-LPR printing

def printlprlabel(lblfmt):
    global g_lpripaddr, g_lprport
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.connect((g_lpripaddr, int(g_lprport)))
    s.send(str(lblfmt).encode())
    s.close()

def printApiLabels(sku, qty):
    global g_sku, g_quantity, g_name, g_misc10, g_upc, g_defaultprice
    global g_shippingweight, g_warehouse, g_zpl, g_location, headers, g_apikey, payload, root
    g_sku = sku.get()
    g_quantity = qty.get()
    prod = g_sku
    headers['NETOAPI_KEY'] = g_apikey
    if prod:
        theXML = payload.replace('[SKU]', prod)
        parsexml(theXML, prod)
    else:
        print("ERROR: No SKU")
        tkinter.messagebox.showwarning("Warning", "SKU not found", parent=root)
        return
    lbl = FormatLabel(g_zpl, g_sku, g_name, g_misc10, g_upc, g_defaultprice,
                      g_quantity, g_shippingweight, g_warehouse, g_location)
    print(lbl)
    printglobals()
    if g_uselpr:
        printlprlabel(lbl)
    else:
        printlabel(lbl)
    print("sku =", g_sku)
    print("Name =", g_name)
    print("quantity =", g_quantity)

def findFile(tkname):
    global g_filename, g_zpl, root
    root.filename = tkinter.filedialog.askopenfilename(
        initialdir=".",
        title="Select file",
        filetypes=(("ZPL files", "*.zpl"), ("all files", "*.*"))
    )
    g_filename = root.filename
    tkname.set(g_filename)
    print("g_filename =", g_filename)
    try:
        with open(g_filename, 'r') as file:
            g_zpl = file.read()
    except:
        tkinter.messagebox.showwarning("Warning", "Cannot open label format, fix Setup before printing", parent=root)

def printApi():
    global root
    clearglobals()
    apiwin = Toplevel(root)
    skuvar = StringVar(apiwin)
    qtyvar = StringVar(apiwin)
    productvar = StringVar(apiwin)
    productvar.set(g_name)
    lbl_sku = Label(apiwin, text="neto SKU:")
    lbl_qty = Label(apiwin, text="Quantity to print:")
    entry_prod = Entry(apiwin, textvariable=productvar, width=60, state='disabled')
    entry_sku = Entry(apiwin, textvariable=skuvar, width=20)
    entry_qty = Spinbox(apiwin, textvariable=qtyvar, width=6, from_=1, to=1000)
    printAPI = Button(apiwin, text="Get from API and Print", command=lambda: printApiLabels(skuvar, qtyvar))
    closeWin = Button(apiwin, text="Close", command=apiwin.destroy)
    lbl_sku.grid(row=0, sticky=E)
    entry_sku.grid(row=0, column=1, sticky=W)
    lbl_qty.grid(row=1, sticky=E)
    entry_qty.grid(row=1, column=1)
    entry_prod.grid(columnspan=2)
    printAPI.grid(columnspan=2)
    closeWin.grid(columnspan=2)

def printLabels():
    global g_csvname
    if len(g_csvname) > 0:
        printCSV()

def doNothing():
    printlabel(g_zpl)

def openPrinter(printer):
    # Windows-specific code (unused on Linux with g_uselpr = 1)
    # if platform.system() == "Windows":
    #     p = win32print.OpenPrinter(printer)
    #     job = win32print.StartDocPrinter(p, 1, ("Neto Product Labels", None, "RAW"))
    #     win32print.StartPagePrinter(p)
    # return p
    pass

def closePrinter(p):
    # Windows-specific code (unused on Linux with g_uselpr = 1)
    # if platform.system() == "Windows":
    #     win32print.EndPagePrinter(p)
    #     win32print.ClosePrinter(p)
    pass

def sendToPrinter(p, plbl):
    # Windows-specific code (unused on Linux with g_uselpr = 1)
    # if platform.system() == "Windows":
    #     win32print.WritePrinter(p, plbl)
    pass

def formatprice(price):
    thePrice = price.strip(' \'\n\r\t')
    pos = thePrice.find('.')
    if pos > 0:
        s = thePrice[pos:]
        if len(s) == 2:
            thePrice += '0'
        else:
            thePrice = thePrice[:pos + 3]
    else:
        thePrice += ".00"
    return thePrice

def getCSVname():
    global g_csvname, root
    print("getCSVname: root defined:", root is not None)  # Debug
    root.csvname = tkinter.filedialog.askopenfilename(
        initialdir=".",
        title="Select file",
        filetypes=(("CSV files", "*.csv"), ("all files", "*.*"))
    )
    g_csvname = root.csvname
    if g_csvname:
        tkinter.messagebox.showinfo("Success", f"File selected: {g_csvname}", parent=root)
    else:
        tkinter.messagebox.showwarning("Warning", "No CSV file selected", parent=root)

def callback(sv):
    global g_printer
    g_printer = sv.get()

def on_closing():
    global root
    writeConfig()
    root.destroy()

def setup_window():
    global g_printer, g_filename, g_apiurl, g_apikey, g_lpripaddr, g_lprport, g_uselpr, root
    options = []
    Name = "none"
    # if platform.system() == "Windows":
    #     if not g_printer:
    #         g_printer = win32print.GetDefaultPrinter()
    #     printers = win32print.EnumPrinters(2, Name, 1)
    #     for printer in printers:
    #         pos = printer[1].find(',')
    #         options.append(printer[1][:pos])
    top = Toplevel(root)
    top.title("Setup")
    lbl_prn = Label(top, text="Choose Printer:")
    lbl_zpl = Label(top, text="Choose Label:")
    lbl_fname = Label(top, textvariable=g_filename)
    lbl_url = Label(top, text="Enter API URL:")
    lbl_key = Label(top, text="Enter API Key:")
    lbl_ipaddr = Label(top, text="Enter LPR IP Address:")
    lbl_port = Label(top, text="Enter LPR Port:")
    lbl_uselpr = Label(top, text="Use LPR:")
    lbl_uselprhlp = Label(top, text="(0 = no, 1 = yes):")
    tkvar = StringVar(top)
    keyvar = StringVar(top)
    urlvar = StringVar(top)
    filevar = StringVar(top)
    ipaddrvar = StringVar(top)
    portvar = StringVar(top)
    uselprvar = StringVar(top)
    tkvar.set(g_printer)
    keyvar.set(g_apikey)
    urlvar.set(g_apiurl)
    filevar.set(g_filename)
    ipaddrvar.set(g_lpripaddr)
    portvar.set(g_lprport)
    uselprvar.set(g_uselpr)
    tkvar.trace('w', lambda *args: callback(tkvar))
    # if platform.system() == "Windows":
    #     sel_printer = OptionMenu(top, tkvar, *options)
    btn_close = Button(top, text="Close", command=top.destroy)
    entry_file = Entry(top, textvariable=filevar, width=60)
    entry_url = Entry(top, textvariable=urlvar, width=50, show="*")
    entry_key = Entry(top, textvariable=keyvar, width=50, show="*")
    entry_ipaddr = Entry(top, textvariable=ipaddrvar, width=16)
    entry_port = Entry(top, textvariable=portvar, width=5)
    entry_uselpr = Entry(top, textvariable=uselprvar, width=5)
    btn_browse = Button(top, text="Browse", command=lambda: findFile(filevar))
    lbl_prn.grid(row=0, sticky=E)
    # if platform.system() == "Windows":
    #     sel_printer.grid(row=0, column=1, sticky=W)
    lbl_zpl.grid(row=1, sticky=E)
    entry_file.grid(row=1, column=1)
    lbl_fname.grid(row=1, column=1)
    btn_browse.grid(row=1, column=1, sticky=E)
    lbl_url.grid(row=2, sticky=E)
    entry_url.grid(row=2, column=1, sticky=W)
    lbl_key.grid(row=3, sticky=E)
    entry_key.grid(row=3, column=1, sticky=W)
    lbl_ipaddr.grid(row=4, sticky=E)
    entry_ipaddr.grid(row=4, column=1, sticky=W)
    lbl_port.grid(row=5, sticky=E)
    entry_port.grid(row=5, column=1, sticky=W)
    lbl_uselpr.grid(row=6, sticky=E)
    entry_uselpr.grid(row=6, column=1, sticky=W)
    lbl_uselprhlp.grid(row=6, column=1, sticky=W, padx=(40, 40))
    btn_close.grid(columnspan=2)
    if platform.system() == "Linux":
        top.focus_force()
        top.lift()
        top.attributes('-topmost', True)
        top.attributes('-topmost', False)
        top.update()
        top.deiconify()

# Initialize main window
root = None
try:
    root = Tk()
    root.title("Printer Selection App")
except Exception as e:
    print(f"Error creating Tkinter window: {e}")
    exit(1)

# Initialize other variables
g_printer = StringVar()
filename = StringVar()
csvname = StringVar()

try:
    readConfig()
except Exception as e:
    tkinter.messagebox.showerror("Error", f"Failed to read config: {e}", parent=root)

# Main Menu
menu = Menu(root)
root.config(menu=menu)
subMenu = Menu(menu, tearoff=0)
menu.add_cascade(label="File", menu=subMenu)
subMenu.add_command(label="Setup", command=setup_window)
subMenu.add_separator()
subMenu.add_command(label="Exit", command=on_closing)

# Toolbar
toolbar = Frame(root, bg="blue")
insertButton = Button(toolbar, text="Setup", command=setup_window)
insertButton.pack(side=LEFT, padx=2, pady=2)
csvButton = Button(toolbar, text="Choose CSV File", command=getCSVname)
csvButton.pack(side=LEFT, padx=2, pady=2)
printButton = Button(toolbar, text="Print CSV", command=printLabels)
printButton.pack(side=LEFT, padx=2, pady=2)
apiButton = Button(toolbar, text="API Print", command=printApi)
apiButton.pack(side=LEFT, padx=2, pady=2)
toolbar.pack(side=TOP, fill=X)

# Status Bar
status = Label(root, textvariable=g_printer, bd=1, relief=SUNKEN, anchor=W)
status.pack(side=BOTTOM, fill=X)

# Handle Linux focus issues
if platform.system() == "Linux":
    root.focus_force()
    root.lift()
    root.attributes('-topmost', True)
    root.attributes('-topmost', False)
    root.update()
    root.deiconify()

root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()

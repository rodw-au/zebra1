import os, sys
import platform
import ctypes
from ctypes.wintypes import BYTE, DWORD, LPCWSTR
import win32print
import csv
import configparser
from configparser import ConfigParser
import requests
import socket
from xml.dom import minidom
from tkinter import *
import tkinter.filedialog
import tkinter.messagebox
import logging

import string

g_printer = ""
g_filename = ""
g_apiurl=''
g_apikey = ''
g_zpl = ""
g_csvname= ""
g_lpripaddr = '192.168.1.47'
g_lprport = 9100 
g_uselpr = 1


g_sku = ''
g_name = ''
#g_shop = ''
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
config = ''

  
def clearglobals():
  global g_sku
  global g_name
  global g_warehouse
  global g_upc
  global g_shippingweight
  global g_shippingwidth
  global g_cubicweight
  global g_shippinglength
  global g_shippingheight
  global g_defaultprice
  global g_promotionprice
  global g_misc10
  global g_location
  global g_quantity

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
  global g_sku
  global g_name
  global g_warehouse
  global g_upc
  global g_shippingweight
  global g_shippingwidth
  global g_cubicweight
  global g_shippinglength
  global g_shippingheight
  global g_defaultprice
  global g_promotionprice
  global g_misc10

  print ("g_sku = ", g_sku.strip())
  print ("g_name = ", g_name.strip())
  print ("g_warehouse = ", g_warehouse.strip())
  print ("g_upc = ", g_upc.strip())
  print ("g_shippingweight = ", g_shippingweight.strip())
  print ("g_shippingwidth = ",g_shippingwidth.strip())
  print ("g_cubicweight = ", g_cubicweight.strip())
  print ("g_shippinglength = ",g_shippinglength.strip())
  print ("g_shippingheight = ", g_shippingheight.strip())
  print ("g_defaultprice = ", g_defaultprice.strip())
  print ("g_promotionprice = ", g_promotionprice.strip())
  print ("g_misc10 = ", g_misc10)
  
def printable(input):
  str = ''.join(filter(lambda x:x in string.printable, input))
  return str

headers = {
    'NETOAPI_ACTION': "GetItem",
    'NETOAPI_KEY':    "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx",   
    'Content-Type':   "application/xml"
    }

payload = """  <?xml version=\"1.0\" encoding=\"utf-8\"?>
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
  #Not very pythonic using global variables but if we didn't, we'd have too many parameters to pass to procedures
  #we have to tell Python we want touse the global variables, not local ones
  global g_name
  global g_sku
  global g_shippingweight
  global g_shippingwidth
  global g_cubicweight
  global g_shippinglength
  global g_shippingheight
  global g_defaultprice
  global g_promotionprice
  global g_upc
  global g_apiurl
  global g_warehouse
  global g_misc10
  global g_location
  global g_quantity


  response = requests.request("POST", g_apiurl, data=tpayload, headers=headers)
  xmldoc = minidom.parseString(response.text)
  itemlist = xmldoc.getElementsByTagName('Item')
  for node in xmldoc.getElementsByTagName('Item'):  # visit every node <Item />
    try:
      name = node.getElementsByTagName('Name')[0]   # check the first element exists to confirm the SKU was found
    except:
      print ("SKU <%s>Not found, aborted"%(sku))
      return
      #sys.exit(1)                                   # Quit if SKU is not found in neto
    tsku           =  node.getElementsByTagName('SKU')[0]
    shippingweight =  node.getElementsByTagName('ShippingWeight')[0]
    shippingwidth  =  node.getElementsByTagName('ShippingWidth')[0]
    cubicweight    =  node.getElementsByTagName('CubicWeight')[0]
    shippinglength =  node.getElementsByTagName('ShippingLength')[0]
    shippingheight =  node.getElementsByTagName('ShippingHeight')[0]
    defaultprice   =  node.getElementsByTagName('DefaultPrice')[0]
    upc            =  node.getElementsByTagName('UPC')[0]
    warehouse      =  node.getElementsByTagName('WarehouseLocations')[0]
    misc10         =  node.getElementsByTagName('Misc10')[0]
    #promotionprice =  node.getElementsByTagName('PromotionPrice')[0]
    

    g_sku = tsku.firstChild.data
    g_name = name.firstChild.data
    g_sku = sku.strip()
    g_name = g_name.strip()
    if(shippingweight.firstChild):
      g_shippingweight = shippingweight.firstChild.data
      g_shippingweight = g_shippingweight.strip()
    else:
      g_shippingweight='0.0'
    if(shippingwidth.firstChild):
      g_shippingwidth = shippingwidth.firstChild.data
      g_shippingwidth = g_shippingwidth.strip()
    else:
      g_shippingwidth = '0.0'
    if(shippinglength.firstChild):
      g_shippinglength = shippinglength.firstChild.data
      g_shippingwidth = g_shippingwidth.strip()
    else:
      g_shippinglength = '0.0'
    if (shippingheight.firstChild):
      g_shippingheight = shippingheight.firstChild.data
      g_shippingheight = g_shippingheight.strip()
    else:
      g_shippinglength = '0.0'
    if(defaultprice.firstChild):
      g_defaultprice = str(defaultprice.firstChild.data)
      #defaultprice = defaultprice.strip()
    else:
      g_defaultprice = "0.00"
    if (upc.firstChild):
      g_upc = upc.firstChild.data
      g_upc = g_upc.strip()
    else:
      g_upc = ''
    if (warehouse.firstChild):
      g_warehouse = warehouse.firstChild.data
      g_warehouse = g_warehouse.strip()
    else:
      g_warehoue = ''
    #if (promotionprice.firstChild):
    #  g_promotionprice = promotionprice.firstChild.data


def readConfig():
  global g_zpl
  global g_filename
  global g_printer
  global g_apiurl
  global g_apikey
  global g_csvname
  global g_uselpr
  global g_lpripaddr
  global g_lprport
  global config

  config = ConfigParser()
  #parser.read_file()
  config.read_file(open('zebra1.ini'))
  config.sections()

  g_printer  = config.get('PRINTER', 'queue')  
  g_filename = config.get('PRINTER', 'label')  
  g_csvname = config.get('PRINTER', 'lastcsv')  
  g_apiurl = config.get('API', 'URL')  
  g_apikey = config.get('API', 'KEY')
  g_lpripaddr = config.get('LPR', 'ipaddr')
  g_lprport =  config.get('LPR', 'port') 
  g_uselpr = config.get('LPR', 'uselpr')
  print(g_filename)
  if g_filename:
    try:
      file = open(g_filename,"r")
      g_zpl = file.read()
      #print("zpl = ",g_zpl)
      file.close()
    except:
      msg = "Cannot open label format " + g_filename + " , check settings before printing"
      tkinter.messagebox.showwarning("Warning",msg)
  else:
    tkinter.messagebox.showwarning("Warning","No Label format selected in Setup")

def writeConfig():
  global g_zpl
  global g_filename
  global g_printer
  global g_apiurl
  global g_apikey
  global g_csvname
  global config
  global g_uselpr
  global g_lpripaddr
  global g_lprport
  #config.add_section('API')
  config.set('API', 'key', g_apikey)
  config.set('API', 'url', g_apiurl)
  #config.add_section('PRINTER') 
  config.set('PRINTER', 'lastcsv', g_csvname) 
  config.set('PRINTER', 'label', g_filename) 
  config.set('PRINTER', 'queue', g_printer)   
  config.set('LPR', 'ipaddr', g_lpripaddr)
  config.set('LPR', 'port', g_lprport)
  config.set('LPR', 'uselpr', g_uselpr)  
  try:
    with open('zebra1.ini', 'wb') as configfile:
      config.write(configfile)  
  except:
    print ("Warning Cannot write to INI file")
    
def printCSV():
  #Name,UPC/EAN,Price (B),Store Location,Pick Zone,SKU*
  global root
  global g_zpl
  global g_sku
  global g_name
  global g_warehouse
  global g_misc10
  global g_upc
  global g_upc
  global g_defaultprice
  global g_quantity
  global g_csvname
  global g_shippingweight
  global g_location
  global g_quantity
  islpr = 0
  
  clearglobals()
  label = ''
  csvname = g_csvname
  if not csvname:
    tkinter.messagebox.showwarning("Warning", "No CSV file selected", parent=root)
    return
  # Check ZPL template for required placeholders
  required_placeholders = ['[SKU]', '[NAME]', '[MISC10]', '[UPC]', '[PRICE]', '[QUANTITY]', '[WEIGHT]', '[WAREHOUSE]', '[LOCATION]','[QTY]']
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
      # Check that all CSV headers exist in the ZPL 
      missing_zpl_fields = [col for col in found_columns if f'[{col}]' not in g_zpl]
      if missing_zpl_fields:
        tkinter.messagebox.showerror(
          "Error"
           f"CSV columns not found in ZPL template: {', '.join(missing_zpl_fields)}. Please update {g_filename}.",
           parent=root
        )
        return
      for row in reader:
        if not row:  # Skip empty rows
          print("Skipping empty row")
          continue
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
        return
      print("Success, CSV processed with no errors")
      islpr = int(g_uselpr)
      #print("g_uselpr = ", g_uselpr, type(g_uselpr),"islpr = ", islpr, type(islpr))
      if islpr:
        #print("About to call printlprlabel()")
        printlprlabel(label)
      else:
        #print("About to call printlabel()")
        printlabel(label)
  except Exception as e:
    tkinter.messagebox.showerror("Error", f"Failed to process CSV: {e}", parent=root)

def FormatLabel(label, sku, name, misc10, upc, price, quantity, weight, warehouse, location):
  lbl1 = label.replace('[SKU]',  printable(sku))
  lbl2 = lbl1.replace('[NAME]',  printable(name))
  lbl3 = lbl2.replace('[MISC10]', printable(misc10))
  lbl4 = lbl3.replace('[UPC]',   printable(upc))
  lbl5 = lbl4.replace('[PRICE]', printable(price))
  lbl6 = lbl5.replace('[QTY]',   printable(quantity))
  lbl7 = lbl6.replace('[WEIGHT]', printable(weight))
  lbl8 = lbl7.replace('[WAREHOUSE]', printable(warehouse))
  lbl9 = lbl8.replace('[LOCATION]', printable(location))
  return lbl9

def printlabel(lblfmt):
  global g_printer
  global g_csv
  printer_name = g_printer
  if platform.system() == "Windows":	
    print("About to print to ",printer_name)
    p = win32print.OpenPrinter (printer_name)
    print("Printer ",printer_name, " opened sucessfully")
    try: 
      print("About to StartDocPrinter for ",printer_name)
      job = win32print.StartDocPrinter (p, 1, ("Neto Product Labels", None, "RAW"))
      try:
        print("About to StartPagePrinter for ",printer_name)
        win32print.StartPagePrinter (p)
        print("About to WritePrinter for ",printer_name)
        win32print.WritePrinter (p, lblfmt.encode('utf-8'))
        win32print.EndPagePrinter (p)
      finally:
        win32print.EndDocPrinter(p)      
    finally:    
      win32print.ClosePrinter (p)

def printlprlabel(lblfmt):
  global g_lpripaddr
  global g_lprport
  
  s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
  s.connect((g_lpripaddr, int(g_lprport)))
  s.send(lblfmt.encode('utf_8'))
  s.close()
    
  
def printApiLabels(sku,qty):
  global g_sku
  global g_quantity
  global g_name
  global g_misc10
  global g_upc
  global g_defaultprice
  global g_shippingweight
  global g_warehouse
  global g_zpl
  global g_location
  global headers
  global g_apikey
  global payload
  g_sku=sku.get()
  g_quantity=qty.get()
  prod = g_sku
  headers['NETOAPI_KEY'] = g_apikey
  if(prod):
    theXML = payload.replace('[SKU]', prod)
    parsexml(theXML,prod)
  else:
    print ("ERRROR: No SKU")
    tkMessageBox.showwarning("Warning","SKU not found")
    return
  lbl = FormatLabel(g_zpl, g_sku, g_name, g_misc10, g_upc, g_defaultprice, g_quantity, g_shippingweight, g_warehouse,g_location)
  print (lbl)
  printglobals()
  if g_uselpr:
    printlprlabel(lbl)
  else:
    printlabel(lbl)

  print ("sku = " + g_sku)
  print ("Name = " + g_name)
  print ("quantity = " + g_quantity)    

def findFile(tkname):
  global g_filename
  global g_zpl
  root.filename = tkinter.filedialog.askopenfilename(initialdir = ".",title = "Select file",filetypes = (("ZPL files","*.zpl"),("all files","*.*")))
  g_filename = root.filename
  tkname.set(g_filename)
  print ("g_filename = ",g_filename)
  try:
    file = open(g_filename,'r')
    g_zpl = file.read()
    file.close()
  except:
     tkinter.messagebox.showwarning("Warning","Cannot open label format, fix Setup before printing") 
  

def printApi():
  clearglobals()
  global g_name
  apiwin = Toplevel(root)
  #apiwin.attributes('-topmost', True)
  skuvar = StringVar(apiwin) 
  qtyvar = StringVar(apiwin)
  productvar = StringVar(apiwin)
  productvar.set(g_name)
  lbl_sku = Label(apiwin, text="neto SKU:")
  lbl_qty = Label(apiwin, text="Quantity to print:")
  entry_prod = Entry(apiwin, textvar=productvar,width=60, state='disabled')
  #entry_prod = Entry(apiwin, textvar=productvar,width=60)
  entry_sku = Entry(apiwin,text=skuvar,width=20)
  entry_qty = Spinbox(apiwin,text=qtyvar,width=6,from_=1,to=1000)
  printAPI = Button(apiwin,text="Get from API and Print",command=lambda:printApiLabels(skuvar,qtyvar))
  #rintAPI.bind('<Button-1>', onclick)
  #oot.bind('<Return>', onclick)
  closeWin = Button(apiwin,text="Close",command=apiwin.destroy)
  lbl_sku.grid(row = 0,sticky=E)
  entry_sku.grid(row=0,column=1,sticky=W)
  lbl_qty.grid(row = 1,sticky=E)
  entry_qty.grid(row=1,column=1)
  entry_prod.grid(columnspan=2)
  printAPI.grid(columnspan=2)
  closeWin.grid(columnspan=2)

def printLabels():
  global g_csvname
  if(len(g_csvname) > 0):
    printCSV()
  
  
def doNothing():
  printlabel(g_zpl)
  
def openPrinter(printer):
  #printer_name ="My printer name"
  #printer_name = printer
  if platform.system() == "Windows":	  
    p = win32print.OpenPrinter (printer)
    job = win32print.StartDocPrinter (p, 1, ("Neto Product Labels", None, "RAW"))
    win32print.StartPagePrinter (p)
  return p;
  
def closePrinter(p):
  if platform.system() == "Windows":		
    win32print.EndPagePrinter (p)
    win32print.ClosePrinter (p)
  
def sendToPrinter(p, plbl):
  if platform.system() == "Windows":		
    win32print.WritePrinter (p, plbl)
  
def formatprice(price):
  thePrice = price.strip(' \'\n\r\t') 
  pos = thePrice.find('.')
  if(pos > 0):
    # check position
    s = price[price.find('.'):]
    if len(s) == 2:
      thePrice += thePrice +'0'
    else: 
      thePrice = thePrice[:pos+3]
  else:
    thePrice = thePrice + ".00"
  return thePrice  
  
def getCSVname():
  global g_csvname
  root.csvname = tkinter.filedialog.askopenfilename(initialdir = ".",title = "Select file",filetypes = (("CSV files","*.csv"),("all files","*.*")))
  g_csvname = root.csvname



def callback(sv):
  global g_printer
  g_printer = sv.get()

import configparser

def save_config():
  global g_printer, g_filename, g_apiurl, g_apikey, g_lpripaddr, g_lprport, g_uselpr, g_csvname
  print ('csv = ',g_csvname)
  config = configparser.ConfigParser()
  config['PRINTER'] = {
      'queue': g_printer if g_printer else '',
      'label': g_filename if g_filename else '',
      'lastcsv': g_csvname if g_csvname else ''
  }
  config['API'] = {
      'URL': g_apiurl if g_apiurl else '',
      'KEY': g_apikey if g_apikey else ''
  }
  config['LPR'] = {
      'ipaddr': g_lpripaddr if g_lpripaddr else '',
      'port': g_lprport if g_lprport else '',
      'uselpr': str(g_uselpr)
  }
  try:
    with open('zebra1.ini', 'w') as configfile:
      config.write(configfile)
      logging.debug("Configuration saved to zebra1.ini")
  except Exception as e:
    logging.error(f"Error saving configuration to zebra1.ini: {e}")
    tkinter.messagebox.showerror("Error", f"Failed to save configuration: {e}", parent=root)
 
def on_closing():
  #writeConfig()
  root.destroy()

def callback_var(var, global_name, window, *args):
  global g_printer, g_filename, g_apiurl, g_apikey, g_lpripaddr, g_lprport, g_uselpr
  value = var.get()
  globals()[global_name] = value
  logging.debug(f"Updated {global_name} to {value}")
  window.focus_force()  # Restore focus to Setup window after update
  
def setup_window():
  global g_printer, g_filename, g_apiurl, g_apikey, g_lpripaddr, g_lprport, g_uselpr, root
  options = []
  Name = "none"
  
  if platform.system() == "Windows":
    try:
      if not g_printer:
        g_printer = win32print.GetDefaultPrinter()
        logging.debug(f"Default printer set to: {g_printer}")
      printers = win32print.EnumPrinters(2, Name, 1)
      for printer in printers:
        pos = printer[1].find(',')
        options.append(printer[1][:pos] if pos != -1 else printer[1])
      logging.debug(f"Found Windows printers: {options}")
    except Exception as e:
      logging.error(f"Error enumerating Windows printers: {e}")
      options = [g_printer] if g_printer else []
  elif platform.system() == "Linux":
    try:
      import cups
      conn = cups.Connection()
      printers = conn.getPrinters()
      options = list(printers.keys())
      logging.debug(f"Found CUPS printers: {options}")
    except ImportError as e:
      logging.error(f"Cannot import pycups: {e}")
      options = [g_printer] if g_printer else []
    except cups.IPPError as e:
      logging.error(f"Error connecting to CUPS: {e}")
      options = [g_printer] if g_printer else []
    except Exception as e:
      logging.error(f"Unexpected error enumerating CUPS printers: {e}")
      options = [g_printer] if g_printer else []
  try:
    window = Toplevel(root)
    window.title("Setup")
    window.geometry("600x250")
    window.resizable(True, True)
    window.transient(root)  # Make window modal relative to root
    window.grab_set()  # Grab focus to keep window on top
    window.columnconfigure(1, weight=1)
    window.rowconfigure(0, weight=1)
    window.rowconfigure(1, weight=1)
    window.rowconfigure(2, weight=1)
    window.rowconfigure(3, weight=1)
    window.rowconfigure(4, weight=1)
    window.rowconfigure(5, weight=1)
    window.rowconfigure(6, weight=1)
    window.rowconfigure(7, weight=1)
    logging.debug("Initialized Toplevel window")
    lbl_prn = Label(window, text="Choose Printer:")
    lbl_zpl = Label(window, text="Choose Label:")
    lbl_fname = Label(window, textvariable=g_filename)
    lbl_url = Label(window, text="Enter API URL:")
    lbl_key = Label(window, text="Enter API Key:")
    lbl_ipaddr = Label(window, text="Enter LPR IP Address:")
    lbl_port = Label(window, text="Enter LPR Port:")
    lbl_uselpr = Label(window, text="Use LPR:")
    lbl_uselprhlp = Label(window, text="(0 = USB/CUPS, 1 = LPR):")
    logging.debug("Created labels")
    tkvar = StringVar(window)
    keyvar = StringVar(window)
    urlvar = StringVar(window)
    filevar = StringVar(window)
    ipaddrvar = StringVar(window)
    portvar = StringVar(window)
    uselprvar = StringVar(window)
    tkvar.set(g_printer)
    keyvar.set(g_apikey)
    urlvar.set(g_apiurl)
    filevar.set(g_filename)
    ipaddrvar.set(g_lpripaddr)
    portvar.set(g_lprport)
    uselprvar.set(g_uselpr)
    logging.debug("Initialized StringVars")
    tkvar.trace('w', lambda name, index, mode: callback_var(tkvar, 'g_printer', window))
    keyvar.trace('w', lambda name, index, mode: callback_var(keyvar, 'g_apikey', window))
    urlvar.trace('w', lambda name, index, mode: callback_var(urlvar, 'g_apiurl', window))
    filevar.trace('w', lambda name, index, mode: callback_var(filevar, 'g_filename', window))
    ipaddrvar.trace('w', lambda name, index, mode: callback_var(ipaddrvar, 'g_lpripaddr', window))
    portvar.trace('w', lambda name, index, mode: callback_var(portvar, 'g_lprport', window))
    uselprvar.trace('w', lambda name, index, mode: callback_var(uselprvar, 'g_uselpr', window))
    #g_uselpr = int(uselpr)
    logging.debug("Set up trace callbacks")
    if (platform.system() == "Windows" or platform.system() == "Linux") and options:
      sel_printer = OptionMenu(window, tkvar, *options)
      sel_printer.config(width=50)
      logging.debug(f"Using OptionMenu for {platform.system()} printers")
    else:
      sel_printer = Entry(window, textvariable=tkvar, width=50)
      logging.debug("Using Entry widget for no printers or unsupported platform")
    btn_close = Button(window, text="Close", command=lambda: [save_config(), window.destroy()])
    entry_file = Entry(window, textvariable=filevar, width=50)
    entry_url = Entry(window, textvariable=urlvar, width=50, show="*")
    entry_key = Entry(window, textvariable=keyvar, width=50, show="*")
    entry_ipaddr = Entry(window, textvariable=ipaddrvar, width=20)
    entry_port = Entry(window, textvariable=portvar, width=10)
    entry_uselpr = Entry(window, textvariable=uselprvar, width=10)
    btn_browse = Button(window, text="Browse", command=lambda: findFile(filevar))
    logging.debug("Created input widgets")
    lbl_prn.grid(row=0, column=0, sticky="e", padx=10, pady=5)
    sel_printer.grid(row=0, column=1, sticky="w", padx=10, pady=5)
    lbl_zpl.grid(row=1, column=0, sticky="e", padx=10, pady=5)
    entry_file.grid(row=1, column=1, sticky="w", padx=10, pady=5)
    btn_browse.grid(row=1, column=2, sticky="e", padx=10, pady=5)
    lbl_url.grid(row=2, column=0, sticky="e", padx=10, pady=5)
    entry_url.grid(row=2, column=1, sticky="w", padx=10, pady=5)
    lbl_key.grid(row=3, column=0, sticky="e", padx=10, pady=5)
    entry_key.grid(row=3, column=1, sticky="w", padx=10, pady=5)
    lbl_ipaddr.grid(row=4, column=0, sticky="e", padx=10, pady=5)
    entry_ipaddr.grid(row=4, column=1, sticky="w", padx=10, pady=5)
    lbl_port.grid(row=5, column=0, sticky="e", padx=10, pady=5)
    entry_port.grid(row=5, column=1, sticky="w", padx=10, pady=5)
    lbl_uselpr.grid(row=6, column=0, sticky="e", padx=10, pady=5)
    entry_uselpr.grid(row=6, column=1, sticky="w", padx=10, pady=5)
    lbl_uselprhlp.grid(row=6, column=1, sticky="w", padx=(60, 10), pady=5)
    btn_close.grid(row=7, column=0, columnspan=3, pady=15)
    logging.debug("Gridded all widgets")
    window.focus_force()  # Ensure window retains focus
    window.update()
    logging.debug("Setup window updated and created successfully")
  except Exception as e:
    logging.error(f"Error creating setup window: {e}")
    tkinter.messagebox.showerror("Error", f"Failed to create setup window: {e}", parent=root)
    
root = Tk()
try:
    readConfig()
except Exception as e:
    tkinter.messagebox.showerror("Error", f"Failed to read config: {e}", parent=root)

#printlabel(g_zpl)
filename = StringVar() 
csvname = StringVar() 
# --- Main Menu ---
menu = Menu(root)
root.config(menu=menu)
subMenu= Menu(menu)
menu.add_cascade(label = "File",menu=subMenu)
#subMenu.add_command(label="New Project",command=doNothing)
subMenu.add_command(label="Setup",command=setup_window)
subMenu.add_separator()
subMenu.add_command(label="Exit",command = on_closing)
#editMenu = Menu(menu)
#menu.add_cascade(label = "Edit", menu=editMenu)
#editMenu.add_command(label="Redo", command = doNothing)

# --- The Toolbar ----
toolbar = Frame(root,bg="blue")
insertButton = Button(toolbar,text="Setup",command=setup_window)
insertButton.pack(side = LEFT, padx=2, pady=2)
csvButton = Button(toolbar,text="Choose CSV File",command=getCSVname)
csvButton.pack(side = LEFT, padx=2, pady=2)
printButton = Button(toolbar,text="Print CSV",command=printLabels)
printButton.pack(side = LEFT, padx=2, pady=2)
apiButton = Button(toolbar,text="API Print",command=printApi)
apiButton.pack(side = LEFT, padx=2, pady=2)
toolbar.pack(side=TOP,fill=X)

# --- The Statusbar ----
p = Label(root, textvariable=g_printer)
status = StringVar()
status = Label(root,textvariable=g_printer,bd=1,relief=SUNKEN, anchor=W)

status.pack(side=BOTTOM,fill=X)
p.pack(side=BOTTOM)

root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()

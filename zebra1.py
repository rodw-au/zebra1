import os, sys
import platform
import ctypes
from ctypes.wintypes import BYTE, DWORD, LPCWSTR
#import win32print
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
    tkMessageBox.showwarning("Warning","Cannot write to INI file")
    
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
  clearglobals()
  ctr = 0;
  label = ''
  csvname = g_csvname
  if not csvname:
    tkinter.messagebox.showwarning("Warning", "No CSV file selected", parent=root)
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
  label = ""
  with open(csvname, 'r', encoding='utf-8') as csvfile:	  
    reader = csv.reader(csvfile,delimiter = ",")
    print(reader)
    headers = next(reader, None)
    if headers is None:
      tkinter.messagebox.showerror("Error", "CSV file is empty", parent=root)
      return
    headers = [h.upper() for h in headers]  # Convert headers to uppercase
    print(headers)
    # Map headers to indices
    header_indices = {}
    found_columns = []
    for col_name in column_mapping:
      try:
        header_indices[col_name] = headers.index(col_name)
        print( header_indices[col_name])
        found_columns.append(col_name)
      except ValueError:
        header_indices[col_name] = None

      # Warn if no expected columns are found
      print(found_columns)
      if not found_columns:
        tkinter.messagebox.showerror(
          "Error",
          f"No valid columns found in CSV. Expected: {', '.join(column_mapping.keys())}",
          parent=root
        )
        return
    for row in reader:
      print("row = ", row)
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
        g_sku = row[header_indices["LOCATION"]]
      if header_indices["QUANTITY"] is not None and len(row) > header_indices["QUANTITY"]:
        g_sku = row[header_indices["QUANTITY"]]              
      label += ((FormatLabel(g_zpl, g_sku, g_name, g_misc10, g_upc, g_defaultprice, g_quantity, g_shippingweight, g_warehouse,g_location)) + "\n")	
  exit()
  if g_uselpr:
    printlprlabel(label)
  else:
    printlabel(label)  # print the whole batch in one call

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
    p = win32print.OpenPrinter (printer_name)
    try: 
      job = win32print.StartDocPrinter (p, 1, ("Neto Product Labels", None, "RAW"))
      try:
        win32print.StartPagePrinter (p)
        win32print.WritePrinter (p, lblfmt)
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
  s.send(str(lblfmt))
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

 
def on_closing():
  #writeConfig()
  root.destroy()

  
def setup_window():
  global g_printer
  global g_filename
  global g_apiurl
  global g_apikey
  global g_lpripaddr
  global g_lprport
  global g_uselpr
  global root
  
  options=[]
  Name = "none"
  if platform.system() == "Windows": 
    if(len(g_printer) == 0):
      g_printer = win32print.GetDefaultPrinter()
    printers = win32print.EnumPrinters(2,Name,1)
    for printer in printers:
      pos = printer[1].find(',')
      options.append(printer[1][:pos])
  window = tkinter.Toplevel(root)
  #window.attributes('-topmost', True)
  lbl_prn = Label(window, text="Choose Printer:")
  lbl_zpl = Label(window, text="Choose Label:")
  lbl_fname = Label(window, textvariable=g_filename)
  lbl_url = Label(window, text="Enter API URL:")
  lbl_key = Label(window, text="Enter API Key:")
  lbl_ipaddr = Label(window, text="Enter LPR IP Address:")
  lbl_port = Label(window, text="Enter LPR Port:")
  lbl_uselpr  = Label(window, text="Use LPR:")
  lbl_uselprhlp  = Label(window, text="(0 = no, 1 = yes):")
  tkvar = StringVar(window)   
  keyvar = StringVar(window)   
  urlvar = StringVar(window)
  filevar = StringVar(window)
  ipaddrvar = StringVar(window)
  portvar = StringVar(window)
  uselprvar = StringVar(window)
  tkvar.set(g_printer) # set the default option
  keyvar.set(g_apikey)
  urlvar.set(g_apiurl)
  filevar.set(g_filename)
  ipaddrvar.set(g_lpripaddr)
  portvar.set(g_lprport)
  uselprvar.set(g_uselpr)
  tkvar.trace('w', lambda *args: callback(tkvar))
  if platform.system() == "Windows":
    sel_printer = OptionMenu(window, tkvar, *options)    
  btn_close = Button(window, text="Close", command=window.destroy)
  #btn_close.bind('<Return>',click)
  entry_file = Entry(window,text=filevar,width=60 )
  entry_url = Entry(window,text=urlvar,width=50,show="*")
  entry_key = Entry(window,text=keyvar,width=50, show="*") 
  entry_ipaddr = Entry(window,text=ipaddrvar,width=16)
  entry_port = Entry(window,text=portvar,width=5)
  entry_uselpr = Entry(window,text=uselprvar,width=5)
  btn_browse = Button(window, text="Browse", command=lambda: findFile(filevar),)
  lbl_prn.grid(row = 0,sticky=E)
  if platform.system() == "Windows":
    sel_printer.grid(row=0,column=1,sticky=W)
  lbl_zpl.grid(row = 1,sticky=E)
  entry_file.grid(row=1,column=1)
  lbl_fname.grid(row=1,column=1)
  btn_browse.grid(row=1,column=1,sticky=E)
  lbl_url.grid(row=2,sticky=E)
  entry_url.grid(row=2,column=1,sticky=W)
  lbl_key.grid(row=3,sticky=E)
  entry_key.grid(row=3,column=1,sticky=W)
  lbl_ipaddr.grid(row=4, sticky=E)
  entry_ipaddr.grid(row=4,column = 1,sticky=W)
  lbl_port.grid(row=5, sticky=E)
  entry_port.grid(row=5,column = 1,sticky=W)
  lbl_uselpr.grid(row=6, sticky=E)
  entry_uselpr.grid(row=6,column = 1,sticky=W)
  lbl_uselprhlp.grid(row=6,column=1,sticky = W, padx=(40, 40))
  
  btn_close.grid(columnspan=2)
"""
  if platform.system() == "Linux":
    window.focus_force()
    window.lift()
    window.attributes('-topmost', True)
    window.attributes('-topmost', False)
    window.update()
    window.deiconify()
"""
root = Tk()
try:
    readConfig()
except Exception as e:
    tkinter.messagebox.showerror("Error", f"Failed to read config: {e}", parent=window)

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

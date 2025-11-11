import requests
import json

url = "https://api.starshipit.com/api/products?search_term=570&page_number=1&page_size=50&skip_records=0&sort_column=Sku&sort_direction=Ascending"
#url = "https://api.starshipit.com/api/products"

payload = {}
#payload={'search_term':'570', 'page_number':1,'page_size':50,'skip_records':0, 'sort_column':'Sku','sort_direction':'Ascending'}
# ==== PUT YOUR KEYS HERE ====02cb101d29694e2ebab329dca72d3a04
API_KEY = "02cb101d29694e2ebab329dca72d3a04"               # from Settings > API
SUB_KEY = "fbc49db0117c47328f97b7a5fb00a83c"      # also on the same page
headers = {
  'Content-Type': 'application/json',
  'StarShipIT-Api-Key': API_KEY,
  'Ocp-Apim-Subscription-Key': SUB_KEY
}

response = requests.request("GET", url, headers=headers, data=payload)
data = response.json()
print(response.text)
print('---------- data ')
print(data)
print('---------- type')
print(type(data))
print('----------')
sku= data['data']['products'][0]['sku']
title= data['data']['products'][0]['title']
upc= data['data']['products'][0]['barcode']
location = data['data']['products'][0]['bin_location']
print("sku = ", sku)
print("title = ", title)
print("barcode = ", upc)
print("location = ", location)

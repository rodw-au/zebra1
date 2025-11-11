# starshipit_products.py
import requests
import json
import time

# ==== PUT YOUR KEYS HERE ====
API_KEY = "02cb101d29694e2ebab329dca72d3a04"               # from Settings > API
SUB_KEY = "fbc49db0117c47328f97b7a5fb00a83c "      # also on the same page
# ============================

BASE_URL = "https://api.starshipit.com/api/products"

headers = {
    "Content-Type": "application/json",
    "StarShipIT-Api-Key": API_KEY,
    "Ocp-Apim-Subscription-Key": SUB_KEY
}

def get_all_products():
    page = 1
    all_products = []
    
    while True:
        params = {"page": page, "pageSize": 100}   # max pageSize = 100
        #auth = (API_KEY, "")                       # basic auth, password empty
        
        print(f"Fetching page {page}...")
        response = requests.get(BASE_URL, headers=headers, params=params)
        print (response.status_code)
        if response.status_code != 200:
            print(f"Error {response.status_code}: {response.text}")
            break
        
        data = response.json()
        products = data.get("Items", [])
        
        if not products:
            print("No more products.")
            break
            
        all_products.extend(products)
        print(f"   got {len(products)} products (total so far: {len(all_products)})")
        
        # Starshipit returns TotalCount & PageCount
        if page >= data.get("PageCount", 1):
            break
            
        page += 1
        time.sleep(0.2)  # be kind to the API
    
    return all_products

if __name__ == "__main__":
    products = get_all_products()
    print(products)
    # Pretty print first 3
    print("\nFirst 3 products:")
    for p in products[:3]:
        print(json.dumps(p, indent=2))
    
    # Save everything
    with open("starshipit_products.json", "w", encoding="utf-8") as f:
        json.dump(products, f, indent=2)
    
    print(f"\nSaved {len(products)} products to starshipit_products.json")

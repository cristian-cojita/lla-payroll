import uuid
import json

# List of brands
brands = [
    "Advanta", "BFG", "BKT", "Bridgestone", "Carlise", "Continental", "Cooper", "Deestone",
    "Delinte", "Dunlop", "Eldorado", "Falken", "Firestone", "Fuzion", "General", "Goodyear",
    "GT Radial", "Hankook", "Harvest King", "Hercules", "Hi Run", "Iron Man", "Kelly", "Kenda",
    "Kumho", "Mastercraft", "Maxxis", "Michelin", "Multimile", "Nexen", "nitto", "Nokian", "Ohtsu",
    "Power King", "pirelli", "Retread", "Roadmaser", "Sailun", "Samsun", "Starfire", "Sumitomo",
    "Toyo", "Trailer King", "Tube", "Uniroyal", "Yokohama", "Misc", "Milestar", "Firestone/HD"
]

# Convert list into JSON format with BrandId (GUID) and brandName
brands_json = [{"BrandId": str(uuid.uuid4()), "brandName": brand} for brand in brands]

# Convert to JSON string
json_string = json.dumps(brands_json, indent=2)
json_string
    
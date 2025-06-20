import json
import requests
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv

# Function to fetch product data from IndiaMART
def fetch_indiamart_data(product, category):
    url = f"https://dir.indiamart.com/search.mp?ss={product.replace(' ', '+')}"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
    }
    
    response = requests.get(url, headers=headers)
    if response.status_code != 200:
        messagebox.showerror("Error", "Failed to fetch data!")
        return []
    
    soup = BeautifulSoup(response.text, "html.parser")
    product_list = []

    # Find all product containers
    products = soup.find_all("div", class_="cardbody")
    
    for product in products:
        try:
            # Extract company name
            company_name = product.select_one("div.companyname")
            company_name = company_name.text.strip() if company_name else "Not Available"

            # Extract product name
            product_name = product.select_one("span.elps.elps2")
            product_name = product_name.text.strip() if product_name else "Not Available"

            # Extract price
            price_tag = product.select_one("p.price")
            price = price_tag.text.strip() if price_tag else "Not Available"

            # Extract location
            location = product.select_one("div.newLocationUi")
            location = location.text.strip() if location else "Not Available"

            # Extract rating
            rating_tag = product.select_one("div.ratingPopup.lft")
            rating = rating_tag.text.strip().replace("★", "") if rating_tag else "Not Available"

            # Extract IndiaMART Member Since
            member_since_tag = product.select_one("div.\\31h21.pdinb.wid3.mb20.verT span.on.color7")
            if member_since_tag and "IndiaMART Member Since" in member_since_tag.text:
                member_since = member_since_tag.find_next("span").text.strip()
            else:
                member_since = "Not Available"

            # Extract phone number
            phone_tag = product.select_one("span.duet")
            phone_number = phone_tag.text.strip() if phone_tag else "Not Available"

            # Create product data dictionary
            product_data = {
                "category": category,
                "company_name": company_name,
                "product_name": product_name,
                "price": price,
                "location": location,
                "rating": rating,
                "member_since": member_since,  # Add IndiaMART Member Since
                "phone_number": phone_number   # Add phone number
            }
            product_list.append(product_data)

        except Exception as e:
            print(f"Error extracting product data: {e}")
            continue

    return product_list

# Function to save data to JSON
def save_to_json(data, filename="indiamart_data.json"):
    with open(filename, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)
    print(f"Data saved in {filename}")

# Function to load data from JSON
def load_from_json(filename="indiamart_data.json"):
    with open(filename, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data

# Function to normalize price
def normalize_price(price):
    if price == "Not Available":
        return float('inf')  # Treat as highest possible price (lowest priority)
    
    # Remove currency symbols and commas
    price_str = price.replace("₹", "").replace(",", "").strip()
    
    # Check if the remaining string is a valid number
    if price_str.replace(".", "", 1).isdigit():
        return float(price_str)
    else:
        return float('inf')  # Treat as highest possible price (lowest priority)

# Function to normalize rating
def normalize_rating(rating):
    if rating == "Not Available":
        return 0.0  # Default to 0 if rating is not available
    
    try:
        # Extract the numeric part (e.g., "4.5/5" → "4.5")
        numeric_part = rating.split("/")[0]
        return float(numeric_part)
    except (ValueError, IndexError):
        print(f"Warning: Unable to parse rating '{rating}'. Defaulting to 0.0.")
        return 0.0

# Function to extract the last word from a string
def get_last_word(text):
    words = text.split()
    return words[-1] if words else ""

# Ranking Algorithm
def rank_products(products, location_preference):
    ranked_products = []

    for product in products:
        # Normalize rating
        rating = normalize_rating(product["rating"])

        # Normalize price
        price = normalize_price(product["price"])

        # Extract the last word of the location
        location_last_word = get_last_word(product["location"])

        # Location preference (1 if last word matches user preference, 0 otherwise)
        location_match = 1 if location_last_word.lower() == location_preference.lower() else 0

        # Calculate score (weights: rating=50%, price=30%, location=20%)
        score = (0.5 * rating) + (0.3 * (1 / price if price != 0 else 0)) + (0.2 * location_match)

        # Add score to product data
        product["score"] = score
        ranked_products.append(product)

    # Sort products by score in descending order
    ranked_products.sort(key=lambda x: x["score"], reverse=True)

    return ranked_products

# List of Categories
categories = [
    "Automotive", "Lighting", "Machinery", "Appliances", 
    "Electronics", "Energy", "Power & Gas", "Construction", "Agriculture"
]

# GUI Application
class IndiaMartScraperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("IndiaMART Scraper")
        self.root.geometry("1200x700")
        self.root.configure(bg="#f0f0f0")

        # Apply custom styling
        self.style = ttk.Style()
        self.style.configure("TFrame", background="#f0f0f0")
        self.style.configure("TLabel", background="#f0f0f0", font=("Arial", 12))
        self.style.configure("TButton", font=("Arial", 12), padding=5)
        self.style.configure("Treeview", font=("Arial", 11), rowheight=25)
        self.style.configure("Treeview.Heading", font=("Arial", 12, "bold"))
        self.style.map("Treeview", background=[("selected", "#347083")])  # Selected row color

        # Main Frame
        self.main_frame = ttk.Frame(root, padding="20")
        self.main_frame.grid(row=0, column=0, sticky="nsew")

        # Category Selection
        self.category_label = ttk.Label(self.main_frame, text="Select Category:")
        self.category_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.category_combobox = ttk.Combobox(self.main_frame, values=categories, state="readonly", font=("Arial", 12))
        self.category_combobox.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        self.category_combobox.current(0)

        # Product Name Entry
        self.product_label = ttk.Label(self.main_frame, text="Enter Product Name:")
        self.product_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.product_entry = ttk.Entry(self.main_frame, width=40, font=("Arial", 12))
        self.product_entry.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

        # Location Preference Entry
        self.location_label = ttk.Label(self.main_frame, text="Preferred Location:")
        self.location_label.grid(row=2, column=0, padx=10, pady=10, sticky="w")
        self.location_entry = ttk.Entry(self.main_frame, width=40, font=("Arial", 12))
        self.location_entry.grid(row=2, column=1, padx=10, pady=10, sticky="ew")

        # Fetch Data Button
        self.fetch_button = ttk.Button(self.main_frame, text="Fetch Data", command=self.fetch_data)
        self.fetch_button.grid(row=3, column=0, columnspan=2, pady=20)

        # Results Display (Table)
        self.tree_frame = ttk.Frame(self.main_frame)
        self.tree_frame.grid(row=4, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

        self.tree = ttk.Treeview(
            self.tree_frame,
            columns=("Company", "Product", "Price", "Location", "Rating", "Member Since", "Phone"),
            show="headings",
            selectmode="browse"
        )
        self.tree.heading("Company", text="Company")
        self.tree.heading("Product", text="Product")
        self.tree.heading("Price", text="Price")
        self.tree.heading("Location", text="Location")
        self.tree.heading("Rating", text="Rating")
        self.tree.heading("Member Since", text="Member Since")
        self.tree.heading("Phone", text="Phone")

        # Set column widths
        self.tree.column("Company", width=150, anchor="w")
        self.tree.column("Product", width=150, anchor="w")
        self.tree.column("Price", width=100, anchor="w")
        self.tree.column("Location", width=250, anchor="w")  # Wider column for location
        self.tree.column("Rating", width=80, anchor="w")
        self.tree.column("Member Since", width=100, anchor="w")
        self.tree.column("Phone", width=100, anchor="w")

        self.tree.pack(side="left", fill="both", expand=True)

        # Scrollbar for the table
        scrollbar = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)

        # Tooltip for location
        self.tooltip = ttk.Label(self.tree_frame, background="#ffffe0", relief="solid", borderwidth=1, font=("Arial", 10))
        self.tooltip.pack_forget()  # Hide initially

        # Bind hover event to show tooltip
        self.tree.bind("<Motion>", self.show_tooltip)

        # Export Buttons
        self.export_frame = ttk.Frame(self.main_frame)
        self.export_frame.grid(row=5, column=0, columnspan=2, pady=10)

        self.export_json_button = ttk.Button(self.export_frame, text="Export to JSON", command=self.export_json)
        self.export_json_button.pack(side="left", padx=10)

        self.export_csv_button = ttk.Button(self.export_frame, text="Export to CSV", command=self.export_csv)
        self.export_csv_button.pack(side="left", padx=10)

        # Configure grid weights for resizing
        self.main_frame.columnconfigure(1, weight=1)
        self.main_frame.rowconfigure(4, weight=1)

    def show_tooltip(self, event):
        # Show tooltip for location column
        region = self.tree.identify("region", event.x, event.y)
        if region == "cell":
            column = self.tree.identify_column(event.x)
            if column == "#4":  # Location column
                row = self.tree.identify_row(event.y)
                item = self.tree.item(row)
                location = item["values"][3]  # Location is the 4th column
                self.tooltip.config(text=location)
                self.tooltip.place(x=event.x + 20, y=event.y + 20)
            else:
                self.tooltip.place_forget()
        else:
            self.tooltip.place_forget()

    def fetch_data(self):
        # Get user inputs
        category = self.category_combobox.get()
        product = self.product_entry.get()
        location_preference = self.location_entry.get()

        if not product or not location_preference:
            messagebox.showwarning("Input Error", "Please enter a product name and location preference.")
            return

        # Fetch data
        data = fetch_indiamart_data(product, category)
        save_to_json(data)

        # Load data from JSON
        data = load_from_json()

        # Rank products
        ranked_products = rank_products(data, location_preference)

        # Clear existing data in the table
        for row in self.tree.get_children():
            self.tree.delete(row)

        # Display results in the table
        for product in ranked_products:
            self.tree.insert("", "end", values=(
                product["company_name"],
                product["product_name"],
                product["price"],
                product["location"],
                product["rating"],
                product["member_since"],
                product["phone_number"]
            ))

        messagebox.showinfo("Success", "Data fetched and ranked successfully!")

    def export_json(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if file_path:
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump([self.tree.item(row)["values"] for row in self.tree.get_children()], f, indent=4, ensure_ascii=False)
            messagebox.showinfo("Success", "Data exported to JSON successfully!")

    def export_csv(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if file_path:
            with open(file_path, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(["Company", "Product", "Price", "Location", "Rating", "Member Since", "Phone"])
                for row in self.tree.get_children():
                    writer.writerow(self.tree.item(row)["values"])
            messagebox.showinfo("Success", "Data exported to CSV successfully!")

# Main program
if __name__ == "__main__":
    root = tk.Tk()
    app = IndiaMartScraperApp(root)
    root.mainloop()
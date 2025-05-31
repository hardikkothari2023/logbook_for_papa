import pandas as pd
import os

FILE_NAME = "contacts.xlsx"
ADMIN_PASSWORD = "papa123"  # 🔐 Change this to your preferred password

def init_file():
    if not os.path.exists(FILE_NAME):
        df = pd.DataFrame(columns=["Name", "Phone", "Address", "City"])
        df.to_excel(FILE_NAME, index=False)

def load_data():
    if not os.path.exists(FILE_NAME):
        df = pd.DataFrame(columns=["Name", "Phone", "Address", "City"])
        df.to_excel(FILE_NAME, index=False)
    return pd.read_excel(FILE_NAME)


def save_data(df):
    df.to_excel(FILE_NAME, index=False)

def add_contact():
    print("\n--- ➕ Add New Contact ---")
    name = input("Name       : ").strip()
    phone = input("Phone      : ").strip()
    address = input("Address    : ").strip()
    city = input("City       : ").strip()

    if name == "" or phone == "":
        print("❌ Name and Phone are required!")
        return

    df = load_data()
    if phone in df["Phone"].astype(str).values:
        print("⚠️ Contact with this phone number already exists.")
        return

    new_contact = pd.DataFrame([[name, phone, address, city]], columns=df.columns)
    df = pd.concat([df, new_contact], ignore_index=True)
    save_data(df)
    print("✅ Contact added successfully!")

def delete_contact():
    print("\n--- 🗑️ Delete Contact ---")
    key = input("Enter Name or Phone to delete: ").strip()

    df = load_data()
    mask = (df["Name"].str.contains(key, case=False, na=False)) | (df["Phone"].astype(str) == key)
    found = df[mask]

    if found.empty:
        print("❌ No contact found.")
    else:
        print("\nFound Contact(s):")
        print(found)
        confirm = input("Are you sure to delete these? (y/n): ").lower()
        if confirm == "y":
            df = df[~mask]
            save_data(df)
            print("✅ Contact(s) deleted.")
        else:
            print("❌ Deletion cancelled.")

def search_contact():
    print("\n--- 🔍 Search Contact ---")
    query = input("Search anything (Name, Phone, Address, City): ").strip()
    df = load_data()
    result = df[df.apply(lambda row: query.lower() in str(row).lower(), axis=1)]

    if result.empty:
        print("😕 No match found.")
    else:
        print("\n🔎 Found Contact(s):")
        print(result.to_string(index=False))

def edit_contact():
    print("\n--- 📝 Edit Contact ---")
    phone = input("Enter Phone number to edit: ").strip()
    df = load_data()
    match = df[df["Phone"].astype(str) == phone]

    if match.empty:
        print("❌ No contact with that phone number.")
        return

    idx = match.index[0]
    print("\nExisting Contact:")
    print(df.loc[idx])

    print("\nLeave field blank to keep current value.")

    name = input(f"New Name    [{df.loc[idx, 'Name']}]: ").strip()
    address = input(f"New Address [{df.loc[idx, 'Address']}]: ").strip()
    city = input(f"New City    [{df.loc[idx, 'City']}]: ").strip()

    if name:
        df.loc[idx, "Name"] = name
    if address:
        df.loc[idx, "Address"] = address
    if city:
        df.loc[idx, "City"] = city

    save_data(df)
    print("✅ Contact updated successfully!")

def view_all():
    print("\n📘 All Contacts in Excel:\n")
    df = load_data()
    if df.empty:
        print("📭 No contacts found.")
    else:
        print(df.to_string(index=False))

def export_contacts():
    df = load_data()
    file_name = input("Enter filename (without extension): ").strip()
    format_ = input("Export as CSV or Excel? ").lower()

    if format_ == "csv":
        df.to_csv(f"{file_name}.csv", index=False)
        print(f"✅ Exported to {file_name}.csv")
    elif format_ == "excel":
        df.to_excel(f"{file_name}.xlsx", index=False)
        print(f"✅ Exported to {file_name}.xlsx")
    else:
        print("❌ Invalid format.")

def login():
    print("🔐 Admin Login Required")
    attempts = 3
    while attempts > 0:
        password = input("Enter Admin Password: ").strip()
        if password == ADMIN_PASSWORD:
            print("✅ Login Successful!\n")
            return True
        else:
            attempts -= 1
            print(f"❌ Incorrect Password. Attempts left: {attempts}")
    print("🚫 Too many wrong attempts. Exiting.")
    return False

def main_menu():
    init_file()
    while True:
        print("\n===============================")
        print("📒 PHONE BOOK DIARY for Papa")
        print("===============================")
        print("1. Add Contact")
        print("2. Delete Contact")
        print("3. Search Contact")
        print("4. Edit Contact")
        print("5. View All Contacts")
        print("6. Export Contacts")
        print("7. Exit")
        choice = input("Choose an option (1-7): ")

        if choice == "1":
            add_contact()
        elif choice == "2":
            delete_contact()
        elif choice == "3":
            search_contact()
        elif choice == "4":
            edit_contact()
        elif choice == "5":
            view_all()
        elif choice == "6":
            export_contacts()
        elif choice == "7":
            print("👋 Thank you, Papa! Exiting now.")
            break
        else:
            print("❌ Invalid option. Please choose 1-7.")

if __name__ == "__main__":
    if login():
        main_menu()

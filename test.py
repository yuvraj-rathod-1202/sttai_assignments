import pandas as pd
import json

# -------- CONFIG --------
INPUT_FILE = r"C:\Users\ratho\Downloads\payments - 14 Feb 26 - 16 Feb 26.xlsx"
OUTPUT_FILE = r"C:\Users\ratho\Downloads\output.xlsx"
AFTERBURN_SPECIAL_PRICE = 153.54
# ------------------------


def safe_json_load(value):
    """Safely load JSON. Returns dict or empty dict."""
    try:
        if pd.isna(value):
            return {}
        return json.loads(value)
    except Exception:
        return {}


def parse_order_items(order_items_str):
    """Parse order_items which is a JSON string inside JSON."""
    try:
        if not order_items_str:
            return []
        return json.loads(order_items_str)
    except Exception:
        return []


def extract_rows(df):
    afterburn_rows = []
    frostflame_rows = []

    for _, row in df.iterrows():
        if(row.get("status") != "captured"):
            continue
        notes_data = safe_json_load(row["notes"])

        timestamp = notes_data.get("created_at", "")
        name = notes_data.get("customer_name", "")
        email = notes_data.get("customer_email", "")
        phone = notes_data.get("customer_phone", "")
        backprint = notes_data.get("backprint_names", "")

        payment_id = row.get("id", "")
        order_id = row.get("order_id", "")

        # ------------------------
        # FORMAT 1 → Has order_items
        # ------------------------
        if "order_items" in notes_data:
            items = parse_order_items(notes_data.get("order_items"))
            team_member = False
            
            for item in items:
                product = item.get("product", "")
                size = item.get("size", "")
                price = item.get("price", "")
                if product == "AFTERBURN" and price <= AFTERBURN_SPECIAL_PRICE+1:
                    team_member = True

            for item in items:
                product = item.get("product", "")
                size = item.get("size", "")
                price = item.get("price", "")

                output_row = [
                    timestamp,
                    name,
                    email,
                    phone,
                    product,
                    size,
                    backprint,
                    price,
                    team_member,
                    payment_id,
                    order_id
                ]

                if product == "AFTERBURN":
                    afterburn_rows.append(output_row)
                elif product == "Frostflame":
                    frostflame_rows.append(output_row)

        # ------------------------
        # FORMAT 2 → No order_items
        # ------------------------
        else:
            size = notes_data.get("size_breakdown", "")
            price = row.get("amount", 0)
            team_member = True

            output_row = [
                timestamp,
                name,
                email,
                phone,
                "",  # Product will be set below
                size,
                backprint,
                price,
                team_member,
                payment_id,
                order_id
            ]

                
            if float(price) <= AFTERBURN_SPECIAL_PRICE+1:
                output_row[4] = "AFTERBURN"
                afterburn_rows.append(output_row)

            # Frostflame condition
            elif float(price) > 600:
                output_row[4] = "AFTERBURN"  # Assuming it's AFTERBURN if price is high but no order_items
                output_row[7] = 150
                afterburn_rows.append(output_row)
                new_output_row = output_row.copy()
                new_output_row[4] = "Frostflame"
                new_output_row[7] = 450
                frostflame_rows.append(new_output_row)

    return afterburn_rows, frostflame_rows


def main():
    df = pd.read_excel(INPUT_FILE)

    afterburn_data, frostflame_data = extract_rows(df)

    columns = [
        "Timestamp",
        "Name",
        "Email",
        "Phone",
        "Product",
        "Size",
        "Back Print",
        "Price",
        "Team Member",
        "Payment ID",
        "Order ID"
    ]

    afterburn_df = pd.DataFrame(afterburn_data, columns=columns)
    frostflame_df = pd.DataFrame(frostflame_data, columns=columns)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        afterburn_df.to_excel(writer, sheet_name="AFTERBURN", index=False)
        frostflame_df.to_excel(writer, sheet_name="Frostflame", index=False)

    print("✅ Output file created:", OUTPUT_FILE)


if __name__ == "__main__":
    main()
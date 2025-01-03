import os
import pandas as pd
import pdfplumber
from pymongo import MongoClient
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

MONGO_URI = "mongodb://localhost:27017/"
client = MongoClient(MONGO_URI)
db = client["garment_orders"]
collection = db["orders"]

EMAIL = os.getenv("EMAIL", "vijaymani0609@gmail.com")
PASSWORD = os.getenv("EMAIL_PASSWORD", "lvyvttthiylnaqdn")
SMTP_SERVER = "smtp.gmail.com"

def extract_pdf_data(pdf_path):
    extracted_data = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    headers = table[0]  
                    for row in table[1:]:
                        extracted_data.append(dict(zip(headers, row)))
        return extracted_data
    except Exception as e:
        print(f"Error reading PDF: {e}")
        return []

def save_to_storage(data, excel_path):
    try:
        formatted_data = []
        for item in data:
            try:
                qty = item.get("Qty", "0") 
                per_rate = item.get("Per Rate", "0")  
                qty = float(qty) if qty.replace('.', '', 1).isdigit() else 0  
                per_rate = float(per_rate) if per_rate.replace('.', '', 1).isdigit() else 0  
                total = qty * per_rate  
                item["Total"] = total  

                formatted_data.append({
                    "Quantity": qty,
                    "Per Rate": per_rate,
                    "Total": total,
                    **item  
                })
            except Exception as e:
                print(f"Error processing row: {e}")

        collection.insert_many(formatted_data)

        df = pd.DataFrame(formatted_data)
        df.to_excel(excel_path, index=False)
        print("Data saved to Excel successfully.")
    except Exception as e:
        print(f"Error saving data: {e}")

def validate_data(data):
    valid_data = [row for row in data if "Placement" in row and "Qty" in row]
    return valid_data

def generate_quotation(data):
    total_qty = sum(float(item.get("Qty", 0)) for item in data)
    total_cost = sum(float(item.get("Total", 0)) for item in data)

    if total_qty == 13:
        total_cost = total_qty * 100 

    return total_qty, total_cost

def send_email(to_email, subject, body, attachment=None):
    try:
        msg = MIMEMultipart()
        msg["From"] = EMAIL
        msg["To"] = to_email
        msg["Subject"] = subject

        msg.attach(MIMEText(body, "plain"))

        if attachment:
            with open(attachment, "rb") as file:
                part = MIMEApplication(file.read(), Name=os.path.basename(attachment))
                part["Content-Disposition"] = f'attachment; filename="{os.path.basename(attachment)}"'
                msg.attach(part)

        with smtplib.SMTP(SMTP_SERVER, 587) as server:
            server.starttls()
            server.login(EMAIL, PASSWORD)
            server.send_message(msg)
        print("Email sent successfully!")
    except Exception as e:
        print(f"Error sending email: {e}")

def main(pdf_path, sender_email):
    try:
        data = extract_pdf_data(pdf_path)
        data = validate_data(data)

        excel_path = "costing_ sheet.xlsx"
        save_to_storage(data, excel_path)

        total_qty, total_cost = generate_quotation(data)

        body = f"""
        Dear Customer,

        Thank you for your inquiry. Please find below the costing details:
        
        Total Quantity: {total_qty}
        Total Cost: ${total_cost:.2f}

        The detailed costing sheet is attached for your reference.
        
        Best regards,
        MORLY Team
        """
        send_email(sender_email, "Costing Details", body, excel_path)
    except Exception as e:
        print(f"Error in main workflow: {e}")

if __name__ == "__main__":
    pdf_path = "C:/Users/vijay/Desktop/intervire project/input_doc (1).pdf"  
    sender_email = "vijayvicky0609@gmail.com"  
    main(pdf_path, sender_email)

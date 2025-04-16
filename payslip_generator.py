import os
import pandas as pd
from fpdf import FPDF
import yagmail
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

SENDER_EMAIL = os.getenv("eversinceshipe@gmail.com")
EMAIL_PASSWORD = os.getenv("megy apgu xapm vpnm")

# Ensure payslips folder exists
os.makedirs("payslips", exist_ok=True)

# Read employee data
def read_employee_data(filepath):
    try:
        df = pd.read_excel(filepath)
        print("Detected columns:", df.columns.tolist())
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

# Calculate net salary
def calculate_net_salary(row):
    return row['Basic Salary'] + row['Allowances'] - row['Deductions']

# # Generate PDF payslip
def generate_payslip(row):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    pdf.cell(200, 10, txt="Employee Payslip", ln=True, align="C")
    pdf.ln(10)

    fields = [
        f"Employee ID: {row['Employee ID']}",
        f"Name: {row['Name']}",
        f"Basic Salary: ${row['Basic Salary']}",
        f"Allowances: ${row['Allowances']}",
        f"Deductions: ${row['Deductions']}",
        f"Net Salary: ${row['Net Salary']}"
    ]

    for field in fields:
        pdf.cell(200, 10, txt=field, ln=True)

    filename = f"payslips/{row['Employee ID']}.pdf"
    pdf.output(filename)
    return filename

# # Send email with payslip
def send_email_with_payslip(to_email, name, pdf_path):
    try:
        yag = yagmail.SMTP("eversinceshipe@gmail.com", "megy apgu xapm vpnm")
        subject = "Your Payslip for This Month"
        body = f"Hi {name},\n\nAttached is your payslip for this month.\n\nRegards,\nHR Department"
        yag.send(to=to_email, subject=subject, contents=body, attachments=pdf_path)
        print(f"Email sent to {to_email}")
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")

# # Main logic
def main():
    df = read_employee_data("employees.xlsx")
    if df is None:
        return
    for index, row in df.iterrows():
        try:
            row['Net Salary'] = calculate_net_salary(row)
            pdf_path = generate_payslip(row)
            send_email_with_payslip(row['Email'], row['Name'], pdf_path)
        except Exception as e:
            print(f"Error processing {row['Name']}: {e}")

if __name__ == "__main__":
    main()

import pandas as pd
from fpdf import FPDF
import smtplib
from email.message import EmailMessage
import os

# Load employee data from Excel
df = pd.read_excel("employee_data.xlsx")

# Normalize column names (safe access)
df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')

# Debug: Show actual column names
print("✅ Loaded columns from Excel:", df.columns.tolist())

# Optional: Rename columns if necessary
df.rename(columns={
    'deduction': 'deductions',
    'total_deductions': 'deductions'
}, inplace=True)

# Function to generate PDF payslip
def generate_payslip(emp):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    # Safely extract values with defaults
    basic = emp.get('basic_salary', 0)
    allowances = emp.get('allowances', 0)
    deductions = emp.get('deductions', 0)
    net_pay = basic + allowances - deductions

    pdf.cell(0, 10, "DESIGN ART WORKS", ln=True, align="C")
    pdf.cell(0, 10, "Payslip for April 2025", ln=True, align="C")

    pdf.ln(10)
    pdf.cell(0, 10, "Employee Details", ln=True)
    pdf.cell(60, 10, f"Employee ID: {emp.get('employee_id', 'N/A')}", ln=True)
    pdf.cell(60, 10, f"Name: {emp.get('name', 'N/A')}", ln=True)
    pdf.cell(60, 10, f"Email: {emp.get('email', 'N/A')}", ln=True)

    pdf.ln(10)
    pdf.cell(0, 10, "Salary Breakdown", ln=True)

    pdf.cell(60, 10, "Basic Salary:", align="L")
    pdf.cell(60, 10, f"${basic:.2f}", ln=True, align="R")

    pdf.cell(60, 10, "Allowances:", align="L")
    pdf.cell(60, 10, f"${allowances:.2f}", ln=True, align="R")

    pdf.cell(60, 10, "Deductions:", align="L")
    pdf.cell(60, 10, f"${deductions:.2f}", ln=True, align="R")

    pdf.cell(60, 10, "Net Salary:", align="L")
    pdf.cell(60, 10, f"${net_pay:.2f}", ln=True, align="R")

    pdf.ln(10)
    pdf.cell(0, 10, "This is a system-generated payslip. For any queries, contact HR.", ln=True, align="C")

    filename = f"{emp.get('name', 'employee').replace(' ', '_')}_Payslip.pdf"
    pdf.output(filename)
    return filename

# Function to send email with PDF attachment
def send_email(to_email, pdf_filename, sender_email, sender_password):
    msg = EmailMessage()
    msg['Subject'] = 'Your Monthly Payslip'
    msg['From'] = sender_email
    msg['To'] = to_email
    msg.set_content('Please find your attached payslip.')

    with open(pdf_filename, 'rb') as f:
        file_data = f.read()
        msg.add_attachment(file_data, maintype='application', subtype='pdf', filename=pdf_filename)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(sender_email, sender_password)
        smtp.send_message(msg)

# Prompt for sender credentials
sender_email = input("Enter sender email (Gmail): ")
sender_password = input("Enter app password: ")  # Use Gmail App Password

# Process each employee
for _, row in df.iterrows():
    pdf_file = None  # Prevent NameError in case of exceptions
    try:
        pdf_file = generate_payslip(row)
        send_email(row.get('email', ''), pdf_file, sender_email, sender_password)
        print(f"✅ Payslip sent to {row.get('name', 'Unknown')} at {row.get('email', 'No email')}")
    except Exception as e:
        print(f"❌ Failed to send payslip to {row.get('email', 'Unknown')}: {e}")
    finally:
        if pdf_file and os.path.exists(pdf_file):
            os.remove(pdf_file)

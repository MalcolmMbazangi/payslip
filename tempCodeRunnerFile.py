import pandas as pd
from fpdf import FPDF
import smtplib
from email.message import EmailMessage
import os

# Step 1: Load Employee Data from Excel
df = pd.read_excel("payslip_data_Excel.xlsx")  # Change to your actual file name

# Step 2: Generate Payslip PDF
def generate_payslip(emp):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    net_pay = emp['basic'] + emp['allowances'] - emp['deductions']
  
    # Company Name and Payslip Title
    pdf.cell(0, 10, "DESIGN ART WORKS", ln=True, align="C")
    pdf.cell(0, 10, "Payslip for April 2025", ln=True, align="C")

    # Employee Details
    pdf.ln(10)
    pdf.cell(0, 10, "Employee Details", ln=True, align="L")
    pdf.cell(60, 10, f"Employee ID: {emp['employee_id']}", ln=True, align="L")
    pdf.cell(60, 10, f"Name: {emp['name']}", ln=True, align="L")
    pdf.cell(60, 10, f"Email: {emp['email']}", ln=True, align="L")

    # Salary Breakdown Table
    pdf.ln(10)
    pdf.cell(0, 10, "Salary Breakdown", ln=True, align="L")

    pdf.cell(60, 10, "Basic Salary:", align="L")
    pdf.cell(60, 10, f"${emp['basic']:.2f}", ln=True, align="R")

    pdf.cell(60, 10, "Allowances:", align="L")
    pdf.cell(60, 10, f"${emp['allowances']:.2f}", ln=True, align="R")

    pdf.cell(60, 10, "Deductions:", align="L")
    pdf.cell(60, 10, f"${emp['deductions']:.2f}", ln=True, align="R")

    pdf.cell(60, 10, "Net Salary:", align="L")
    pdf.cell(60, 10, f"${net_pay:.2f}", ln=True, align="R")

    # Footer
    pdf.ln(10)
    pdf.cell(0, 10, "This is a system-generated payslip. For any queries, contact HR.", ln=True, align="C")

    filename = f"{emp['name'].replace(' ', '_')}_Payslip.pdf"
    pdf.output(filename)
    return filename

# Step 3: Send Email with PDF
def send_email(to_email, pdf_filename, sender_email, sender_password):
    msg = EmailMessage()
    msg['Subject'] = 'Your Monthly Payslip'
    msg['From'] = sender_email
    msg['To'] = to_email

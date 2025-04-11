from fpdf import FPDF
import smtplib
from email.message import EmailMessage
import os

# Step 1: Employee Data
employees = [
    {"employee_id": "63-000005", "name": "Malcolm Mbazangi", "email": "malcolmmbazangi@gmail.com"},
    {"employee_id": "63-000006", "name": "Welma Kayanga", "email": "welmakayanga7@gmail.com"},
    {"employee_id": "63-000007", "name": "Tafadzwa Kanhoodza", "email": "kanhohodza5369@gmail.com"},
    {"employee_id": "63-000008", "name": "Kirsty Matyukira", "email": "kirstyservie03@gmail.com"},
    {"employee_id": "63-000009", "name": "Nomsa Sibanda", "email": "snomsa2301@gmail.com"},
    {"employee_id": "63-000010", "name": "Lorraine Tomu", "email": "lorrainethom93@gmail.com"},
    {"employee_id": "63-000011", "name": "Tadiwa Sipambeni", "email": "tadiwanashesipambeni@gmail.com"},
    {"employee_id": "63-000012", "name": "Sharon Mwandura", "email": "sharonemwandura@gmail.com"},
    {"employee_id": "63-000013", "name": "Tapiwa Gombarume", "email": "tapiwaroy55@gmail.com"},
    {"employee_id": "63-000014", "name": "Precious Tendayi", "email": "precioustendayi36@gmail.com"}
]

# Step 2: Generate Payslip PDF
def generate_payslip(emp, basic, allowances, deductions):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    net_pay = basic + allowances - deductions
  
    # Company Name and Payslip Title
    pdf.cell(0, 10, "DESIGN ART WORKS", ln=True, align="C") # Assuming "Company Name" is the company's name
    pdf.cell(0, 10, "Payslip for April 2025", ln=True, align="C")

    # Employee Details
    pdf.ln(10) # Add some space
    pdf.cell(0, 10, "Employee Details", ln=True, align="L")
    pdf.cell(60, 10, f"Employee ID: {emp['employee_id']}", ln=True, align="L")
    pdf.cell(60, 10, f"Name: {emp['name']}", ln=True, align="L")
    pdf.cell(60, 10, f"Email: {emp['email']}", ln=True, align="L")

    # Salary Breakdown Table
    pdf.ln(10)
    pdf.cell(0, 10, "Salary Breakdown", ln=True, align="L")

    pdf.cell(60, 10, "Basic Salary:", align="L")
    pdf.cell(60, 10, f"${basic:.2f}", ln=True, align="R")

    pdf.cell(60, 10, "Allowances:", align="L")
    pdf.cell(60, 10, f"${allowances:.2f}", ln=True, align="R")

    pdf.cell(60, 10, "Deductions:", align="L")
    pdf.cell(60, 10, f"${deductions:.2f}", ln=True, align="R")

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
    msg.set_content('Please find your attached payslip.')

    with open(pdf_filename, 'rb') as f:
        file_data = f.read()
        msg.add_attachment(file_data, maintype='application', subtype='pdf', filename=pdf_filename)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(sender_email, sender_password)
        smtp.send_message(msg)

# Step 4: Input Salary Details and Send
sender_email = input("Enter sender email (Gmail): ")
sender_password = input("Enter app password: ")  # Use Gmail App Passwords

# Salary data from the document
salary_data = [
    {"basic": 3100.00, "allowances": 300.00, "deductions": 180.00}, # Malcolm Mbazangi
    {"basic": 80, "allowances": 80, "deductions": 40},
    {"basic": 58, "allowances": 58, "deductions": 70},
    {"basic": 92, "allowances": 92, "deductions": 38},
    {"basic": 150, "allowances": 150, "deductions": 100},
    {"basic": 68, "allowances": 68, "deductions": 37},
    {"basic": 122, "allowances": 122, "deductions": 100},
    {"basic": 34, "allowances": 34, "deductions": 50},
    {"basic": 87, "allowances": 87, "deductions": 80},
    {"basic": 45, "allowances": 45, "deductions": 90}
]

for i, emp in enumerate(employees):
    basic = salary_data[i]["basic"]
    allowances = salary_data[i]["allowances"]
    deductions = salary_data[i]["deductions"]

    pdf_file = generate_payslip(emp, basic, allowances, deductions)
    send_email(emp['email'], pdf_file, sender_email, sender_password)
    print(f"Payslip sent to {emp['email']}")

    os.remove(pdf_file)  # Optional: delete file after sending
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import pandas as pd
from TextOnImage import generate_payslip_overlay
import smtplib
from email.message import EmailMessage
import os


def send_email_with_payslip(to_email, employee_name, pdf_path):
    msg = EmailMessage()
    msg['Subject'] = f"Payslip for {employee_name}"
    msg['From'] = "add_your_mail"
    msg['To'] = to_email
    msg.set_content(f"Dear {employee_name},\n\nPlease find attached your payslip.\n\nRegards,\nHR Team")

    # Attach PDF
    with open(pdf_path, 'rb') as f:
        file_data = f.read()
        file_name = os.path.basename(pdf_path)
        msg.add_attachment(file_data, maintype='application', subtype='pdf', filename=file_name)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login("shrithan.adipuram@gmail.com", "zvvz excu tfrj skwj")
        smtp.send_message(msg)

def check_and_run_function(*args):
    month = month_var.get()
    year = year_var.get()
    if month and year:
        selected = f"{month} {year}"
        generate_payslips(selected)


def upload_excel():
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not filepath:
        return
    
    try: 
        print("Im here")
        df_raw = pd.read_excel(filepath, engine='openpyxl', header=None)        
        print("Excel loaded.")

        header_row_index = df_raw[df_raw.apply(lambda row: row.astype(str).str.contains("S.No.").any(), axis=1)].index[0]       \

        headers = df_raw.iloc[header_row_index].values      

        global df_data 
        df_data = df_raw.iloc[header_row_index + 1:].copy()
        df_data.columns = headers
        df_data = df_data.dropna(subset=["S.No."]) 

        df_data.reset_index(drop=True, inplace=True)
        messagebox.showinfo("Success", f"Excel loaded with {len(df_raw)} rows")
        #print(df_data)
    except Exception as e:
        messagebox.showerror("Error", f"failed to read the excel file")
    
def generate_payslips(monthandyear):
    arr = []
    for i in range(len(df_data)):
        arr.append(df_data.iloc[i])
    #print(arr)

    for row in arr:
        sno = row["S.No."]
        emp_code = row["Employee Code"]
        emp_name = row["Employee Name"]
        emp_mail = row["Employee Mail"]
        designation = row["Designation"]
        department = row["Department"]
        bank_details = row["Bank Details"]

        ctc = round(row["CTC"])
        ctc_per_month = round(row["CTC per month for company"])
        bonus = round(row["Bonus"])
        calendar_days = round(row["Calendar Days"])
        lop_days = round(row["Loss Of Pay - No. of days (Absent Days)"])
        basic = round(row["Basic"])
        hra = round(row["HRA"])
        conveyance = round(row["Conveyance+Medical"])
        special_allowance = round(row["Special Allowance"])
        net = round(row["Net salary"])
        loss_of_pay = round(row["Loss Of Pay"])
        pt = round(row["PT"])
        tds = round(row["TDS"])
        balance_payable = round(row["Balance payable"])
        paid = round(row["Paid"])
        cum_arrears = round(row["Cum. Arrears"])


        if cum_arrears<0:
            arrears_paid = - cum_arrears
            arrears_withheld = 0
        elif cum_arrears>0:
            arrears_withheld = cum_arrears
            arrears_paid = 0
        else:
            arrears_paid = 0
            arrears_withheld = 0

        print(f"{emp_name} ({emp_code}) | Designation: {designation} | Net Salary: ₹{net}")

        generate_payslip_overlay(
            template_path="template 2.jpg",
            save_path=f"{monthandyear}_{emp_code}_{emp_name}.jpg",
            month=f"{monthandyear}",
            name=f"{emp_name}",
            designation=f"{designation}",
            department=f"{department}",
            calendar_days=f"{calendar_days}",
            ctc=f"{ctc_per_month}",
            emp_code=f"{emp_code}",
            email=f"{emp_mail}",
            lop=f"{lop_days}",
            present_days=f"{calendar_days-lop_days}",
            bank_details=f"{bank_details}",
            basic=f"{basic}",
            hra=f"{hra}",
            conveyance=f"{conveyance}",
            special_allowance=f"{special_allowance}",
            arrears=f"{arrears_paid}",
            bonus=f"{bonus}",
            gross=f"{net}",
            balance_payable=f"{paid}",
            provident_fund="0",
            employee_state_insurance="0",
            pt=f"{pt}",
            tds=f"{tds}",
            deduction="0",
            arrears_withheld=f"{arrears_withheld}",
            total_deduction=f"{pt+tds+arrears_withheld}"
        )

        send_email_with_payslip(to_email="add_your_mail", employee_name=emp_name, pdf_path=f"{monthandyear}_{emp_code}_{emp_name}.pdf")

root = tk.Tk()
root.title("Payslip Generator")
root.geometry("300x200")

label = tk.Label(root, text = "Upload excel file")
label.pack(pady=10)

def on_generate_click():
    dropdown_frame.pack(pady=10)

tk.Button(root, text="Generate Payslip", command=on_generate_click).pack(pady=20)

dropdown_frame = tk.Frame(root)

month_var = tk.StringVar()
month_dropdown = ttk.Combobox(dropdown_frame, textvariable=month_var, state="readonly")
month_dropdown["values"] = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                            "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
month_dropdown.pack(side=tk.LEFT, padx=5)

year_var = tk.StringVar()
year_dropdown = ttk.Combobox(dropdown_frame, textvariable=year_var, state="readonly")
year_dropdown["values"] = ["24", "25", "26", "27", "28", "29", "30"]
year_dropdown.pack(side=tk.LEFT, padx=5)

month_var.trace_add("write", check_and_run_function)
year_var.trace_add("write", check_and_run_function)


uplaod_btn = tk.Button(root, text="Browse", command=upload_excel)
uplaod_btn.pack()

root.mainloop()

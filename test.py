import tkinter as tk
from tkinter import messagebox
import subprocess
import tkinter as tk
from tkinter import ttk
from ttkbootstrap.constants import *
from PIL import Image, ImageTk  # Import the necessary PIL modules
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import logging
import time
from tkinter import scrolledtext
import pandas as pd
from datetime import datetime, timedelta
import calendar
import tkinter as tk
from tkinter import filedialog, Text
from tkinter import ttk, PhotoImage
from ttkbootstrap.constants import *
from tkinter import ttk
from PIL import Image, ImageTk

# Initialize the email_count variable
email_count = 0

smtp_password = "Beta@12345"  # Default value

try:
    with open("pass.txt", "r") as file:
        smtp_password = file.readline().strip()
        print(f"SMTP password read from file: {smtp_password}")
except FileNotFoundError:
    print("Error: The 'pass.txt' file was not found.")
# Count the entire rows without including the top row
# ... (previous code)
def open_logs_file():
    try:
        log_excel_file = "email_log.xlsx"
        if os.path.exists(log_excel_file):
            os.system(f'start excel "{log_excel_file}"')
        else:
            print(f"Log file '{log_excel_file}' not found.")
    except Exception as e:
        print(f"Error opening logs file: {str(e)}")


# Function to send emails
def send_emails():
    global email_count  # Define a global variable to keep track of the email count
    email_count = 0  # Initialize the email count to zero
    # Get the PDF folder directory and Excel file directory from the input boxes
    pdf_folder = pdf_folder_entry.get()
    excel_file = excel_file_entry.get()

    # Create a list to store log data for conversion to DataFrame
    log_data = []

    # Get the list of PDF files in the folder
    os.makedirs(pdf_folder, exist_ok=True)
    pdf_files = os.listdir(pdf_folder)

    # Read data from Excel file
    df = pd.read_excel(excel_file)

    # Get the list of people with their IDs and corresponding emails from the DataFrame
    people = []
    for index, row in df.iterrows():
        person = {
            "id": row[0],  # Assuming ID is in the first column
            "email": row[1]  # Assuming email is in the second column
        }
        people.append(person)

    # Display the sending output in the text box
    output_text.config(state=tk.NORMAL)  # Enable text box for editing
    output_text.delete(1.0, tk.END)  # Clear previous content
    output_text.update()  # Refresh the GUI

    # Iterate over each person

    for person in people:
        # Find the PDF file matching the person's ID
        pdf_file = next((f for f in pdf_files if str(person["id"]) in f), None)

        if pdf_file:
          

            # Dynamically generate the subject with the previous month
            # Define the current date and last month
            today = datetime.now()
            last_month = today - timedelta(days=30)  # Assuming an average month has 30 days
            last_month_name = calendar.month_name[last_month.month]
            last_month_year = last_month.year
            
            # Create the email message
            message = MIMEMultipart()
            message["From"] = "billing.reports@ntc.org.pk"
            message["To"] = person["email"]
            message["Subject"] = f"NTC Bill For the Month of {last_month_name}-{last_month_year}"

            # Create the email body with the defined variables
            body_message = f"""
                        <html>
    <head>
        <style>
        body {{
            font-family: Arial, Helvetica, sans-serif;
        }}
        .container {{
           
            padding: 10px;
        }}
        .header {{
            background-color: #333;
            color: #fff;
            padding: 10px;
            text-align: center;
        }}
        .section {{
            margin-top: 10px;
            padding: 5px;
            background-color: #fff;
            border: 1px solid #ccc;
            
        }}
        table {{
            width: 100%; /* 4 columns take up the full width */
            border-collapse: collapse;
           
        }}
        table, th, td {{
            border: 1px solid #ccc;
            
        }}
        th, td {{
            padding: 10px;
            text-align: left;
            border-radius: 30px;
        
            
        }}

        h5 {{
            text-align: center;
            background-color: #333;
            color: white;
        }}


        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>NTC Bill for {last_month_name}-{last_month_year}</h1>
            </div>
            <div class="section">
            <p>Dear Customer,</p>
                    <p>Please find attached bill for the Month of <b>{last_month_name}-{last_month_year}</b>.</p>
                    <p><b>PLEASE DO NOT REPLY, THIS IS A SYSTEM GENERATED EMAIL</b></p>
            </div>
            <div class="section">
                    <h4 style="color: blue">Bill can be paid via online using the following methods:  </h4>
                    
                </div>
            <table>
                <tr>
                    <td class="section">
                        <h5>Internet Banking:</h5>
                        <ol>
                            <li>Login to your Internet Banking</li>
                            <li>Select 1BILL</li>
                            <li>Select Invoice/Vouchers</li>
                            <li>Enter Invoice/Voucher No.</li>
                            <li>View details and Pay Bill</li>
                        </ol>
                    </td>
                    <td class="section">
                        <h5>ATM:</h5>
                        <ol>
                            <li>Enter ATM card PIN</li>
                            <li>Select 1BILL</li>
                            <li>Select Invoice/Vouchers</li>
                            <li>Enter Invoice/Voucher No.</li>
                            <li>View details and Pay Bill</li>
                        </ol>
                    </td>
                
                    <td class="section">
                        <h5>Mobile Banking:</h5>
                        <ol>
                            <li>Login to your Mobile Banking App</li>
                            <li>Select 1BILL</li>
                            <li>Select Invoice/Vouchers</li>
                            <li>Enter Invoice/Voucher No.</li>
                            <li>View details and Pay Bill</li>
                        </ol>
                    </td>
                    <td class="section">
                        <h5>Over the Counter (OTC):</h5>
                        <ol>
                            <li>Visit the nearest OTC </li>
                            <li>Inform the representative that you want to pay with 1BILL</li>
                            <li>Show Bill on OTC with reference number (Invoice/Voucher No.)</li>
                            <li>Give cash as requested by the representative.</li>
                            <li>Check details and Pay Bill</li>
                        </ol>
                    </td>
                   
                </tr>
            </table>

            <div class="section">
                <p"><strong>Note:</strong></p>
                <p>You may also visit <a href="https://app.kuickpay.com/PaymentsBillPayment">here</a> for Bank-wise Bill Payment Method.</p>
            </div>
        </div>
    </body>
                        """


            # Set the MIME type to 'html' for the email body
            message.attach(MIMEText(body_message, "html"))



            # Attach the PDF file to the email message
            with open(os.path.join(pdf_folder, pdf_file), "rb") as f:
                attachment = MIMEBase("application", "octet-stream")
                attachment.set_payload(f.read())
                encoders.encode_base64(attachment)
                attachment.add_header("Content-Disposition", f"attachment; filename={pdf_file}")
                message.attach(attachment)

            # Send the email message
            try:
                smtp_server = smtplib.SMTP("mail.ntc.net.pk", 587)
                smtp_server.login("billing.reports@ntc.org.pk", smtp_password)
                smtp_server.sendmail("billing.reports@ntc.org.pk", person["email"], message.as_string())
                logging.info(f"Email sent to {person['email']} - Mail send success")
                log_data.append({
                    "ID": person["id"],
                    "Email": person["email"],
                    "Status": "Success",
                    "Attachment": pdf_file,
                    "Timestamp": time.strftime('%Y-%m-%d %H:%M:%S')
                })
                # Increment the email count
                email_count += 1
                counter_label.config(text=f"Total Email's Send: {email_count}")
                output_text.insert(tk.END, f"Email sent to {person['email']} - Mail send success\n")
                
                output_text.update()  # Refresh the GUI
            except smtplib.SMTPRecipientsRefused as e:
                logging.error(f"Failed to send email to {person['email']}. Recipient email not found. - Mail not send")
                log_data.append({
                    "ID": person["id"],
                    "Email": person["email"],
                    "Status": "Recipient not found",
                    "Attachment": pdf_file if pdf_file else "No PDF found",
                    "Timestamp": time.strftime('%Y-%m-%d %H:%M:%S')
                })
                output_text.insert(tk.END, f"Failed to send email to {person['email']}. Recipient email not found. - Mail not send\n")
                output_text.update()  # Refresh the GUI
            except smtplib.SMTPException as e:
                logging.error(f"Failed to send email to {person['email']}. Error: {str(e)} - Mail not send")
                log_data.append({
                    "ID": person["id"],
                    "Email": person["email"],
                    "Status": f"Error: {str(e)}",
                    "Attachment": pdf_file,
                    "Timestamp": time.strftime('%Y-%Y-%d %H:%M:%S')
                })
                output_text.insert(tk.END, f"Failed to send email to {person['email']}. Error: {str(e)} - Mail not send\n")
                output_text.update()  # Refresh the GUI
            finally:
                smtp_server.quit()

            # Pause for 10 seconds before sending the next email
            time.sleep(10)
        else:
            logging.warning(f"No PDF file found for person ID {person['id']} - Mail not send")
            log_data.append({
                "ID": person["id"],
                "Email": person["email"],
                "Status": "No PDF found",
                "Attachment": "No PDF found",
                "Timestamp": time.strftime('%Y-%m-%d %H:%M:%S')
            })
            output_text.insert(tk.END, f"No PDF file found for person ID {person['id']} - Mail not send\n")
            output_text.update()  # Refresh the GUI

    # Convert the log data to a DataFrame
    log_df = pd.DataFrame(log_data)

    # Save the log data to an Excel file
    log_excel_file = "email_log.xlsx"
    log_df.to_excel(log_excel_file, index=False)

    output_text.insert(tk.END, "Email sending process completed.\n")
    output_text.config(state=tk.DISABLED)  # Disable text box for editing

    # Send email with log file attached to recipients
    send_email_with_attachment(log_excel_file)

# Function to send email with attachment
def send_email_with_attachment(log_file):
    sender_email = "billing.reports@ntc.org.pk"  # Replace with your email address
    sender_password = smtp_password  # Replace with your email password
    recipient_emails = ["arslan.manzoor@ntc.org.pk","muhammad.safeer@ntc.org.pk"
                        ,"sidra.shahzad@ntc.org.pk",
                        "muhammad-ali@ntc.org.pk"
                        ]
    

    # Create the email message
    message = MIMEMultipart()
    message["From"] = sender_email
    message["Bcc"] = ", ".join(recipient_emails)  # Comma-separated list of recipients
    message["Subject"] = "AUTO GENERATED Email Log File"
    # email_body = "Please find attached the spreadsheet containing the list of customers whose bills have been successfully sent."
    body_message = "Please find attached the Excel sheet containing the list of customers whose bills have been successfully sent.\n THIS IS SYSTEM GENERATED EMAIL"
    message.attach(MIMEText(body_message, "plain"))

    # Attach the log file to the email message
    with open(log_file, "rb") as f:
        attachment = MIMEBase("application", "octet-stream")
        attachment.set_payload(f.read())
        encoders.encode_base64(attachment)
        attachment.add_header("Content-Disposition", f"attachment; filename={os.path.basename(log_file)}")
        message.attach(attachment)

    # Send the email message
    try:
        smtp_server = smtplib.SMTP("mail.ntc.net.pk", 587)
        smtp_server.starttls()
        smtp_server.login(sender_email, sender_password)
        smtp_server.sendmail(sender_email, recipient_emails, message.as_string())
        smtp_server.quit()
        print("Email with log file sent successfully.")
    except Exception as e:
        print(f"Failed to send email with log file: {str(e)}")

def browse_pdf_folder():
    folder_path = filedialog.askdirectory()
    pdf_folder_entry.delete(0, tk.END)
    pdf_folder_entry.insert(0, folder_path)

# Function to browse for Excel file directory
def browse_excel_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    excel_file_entry.delete(0, tk.END)
    excel_file_entry.insert(0, file_path)

# # Function to change the SMTP email password
# def change_password():
#     current_password = current_password_entry.get()
#     new_password = new_password_entry.get()

#     if current_password == smtp_password:  # Match with the stored SMTP password
#         # global smtp_password  # Declare smtp_password as global to modify it
#         smtp_password = new_password  # Update the stored SMTP password
#         messagebox.showinfo("Success", "Password changed successfully!")
#     else:
#         messagebox.showerror("Error", "Please check your current password.")

# Create a GUI window
root = tk.Tk()
root.title("Auto Email Send")

# Set the window size
root.geometry("600x600")

# Create and set up a title label with a logo on the left
title_frame = ttk.Frame(root)
title_frame.pack()
title_label = ttk.Label(title_frame, text="NTC Auto Email Send", style="TLabel")
title_label.pack(side=tk.LEFT)
title_label.configure(font=("Olivia", 16), foreground="#1a75ff")

# Create and set up input boxes with browse buttons
pdf_folder_label = ttk.Label(root, text="PDF Folder Directory:")
pdf_folder_label.pack()
pdf_folder_frame = ttk.Frame(root)
pdf_folder_frame.pack()
pdf_folder_entry = ttk.Entry(pdf_folder_frame, width=50)
pdf_folder_entry.pack(side=tk.LEFT)
pdf_folder_browse_button = ttk.Button(pdf_folder_frame, text="Browse",command=browse_pdf_folder, bootstyle= (SUCCESS,OUTLINE ))
pdf_folder_browse_button.pack(side=tk.RIGHT,padx=15)  # Add the Browse button to the right

excel_file_label = ttk.Label(root, text="Excel File Directory:")
excel_file_label.pack()
excel_file_frame = ttk.Frame(root)
excel_file_frame.pack()
excel_file_entry = ttk.Entry(excel_file_frame, width=50)
excel_file_entry.pack(side=tk.LEFT)

excel_file_browse_button = ttk.Button(excel_file_frame, text="Browse", command=browse_excel_file,bootstyle= (SUCCESS,OUTLINE ))
excel_file_browse_button.pack(side=tk.RIGHT,padx=15)  # Add the Browse button to the right


# Create and set up the email sending button
send_button = ttk.Button(root, text="Send Emails", command=send_emails,bootstyle= (PRIMARY,OUTLINE))
send_button.pack(pady=15)
# Create a label for the counter
counter_label = tk.Label(root, text=f"Total Email Sended: {email_count}")
counter_label.pack(pady=5)




output_text = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=70, height=10)
output_text.pack()

# Disable the Text widget for editing
output_text.config(state=tk.DISABLED)

 # Add a button to open the logs file
open_logs_button = ttk.Button(root, text="Open Logs File", command=open_logs_file,
                               bootstyle= (PRIMARY,OUTLINE))
open_logs_button.pack(pady=15)


# # Create and set up input boxes for current password and new password
# current_password_label = ttk.Label(root, text="Current Password:")
# current_password_label.pack()
# current_password_entry = ttk.Entry(root, show="*")  # Show '*' for password entry
# current_password_entry.pack()

# new_password_label = ttk.Label(root, text="New Password:")
# new_password_label.pack()
# new_password_entry = ttk.Entry(root, show="*")  # Show '*' for password entry
# new_password_entry.pack()

# # Create and set up the change password button
# change_password_button = ttk.Button(root, text="Change Password", command=change_password, bootstyle=(WARNING, OUTLINE))
# change_password_button.pack(pady=10)

# Center the entire content
root.update_idletasks()  # Ensure geometry changes are applied before centering
width = root.winfo_width()
height = root.winfo_height()
x_offset = (root.winfo_screenwidth() - width) // 2
y_offset = (root.winfo_screenheight() - height) // 2
root.geometry(f"{width}x{height}+{x_offset}+{y_offset}")

# Run the GUI application
root.mainloop()

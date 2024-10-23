from flask import Flask, render_template, request, redirect, url_for, session, flash
import time
import pandas as pd
import os
import win32com.client as win32
from datetime import datetime
import pythoncom
import csv

app = Flask(__name__)
app.secret_key = 'your_secret_key'

# List of valid users for login
users = {
    "user1": "invinc1",
    "user2": "invinc2",
    "user3": "invinc3",
    "user4": "invinc4",
    "user5": "invinc5"
}

# File path for storing records
RECORDS_FILE = 'login_records.csv'  # Adjust if you want to move the records file

# Load login records from the CSV file
def load_login_records():
    if os.path.exists(RECORDS_FILE):
        with open(RECORDS_FILE, mode='r') as file:
            reader = csv.reader(file)
            return [row[0] for row in reader]  # Return list of records
    return []

# Save login record to a CSV file
def save_login_record(record):
    with open(RECORDS_FILE, mode='a', newline='') as file:
        writer = csv.writer(file)
        writer.writerow([record])

# Function to clear login records
def clear_login_records():
    if os.path.exists(RECORDS_FILE):
        os.remove(RECORDS_FILE)

# Initialize records from file on startup
login_records = load_login_records()

# Read email data from Excel file
def read_excel(sheet_name):
    # Example for reading Excel file
    df = pd.read_excel(r'Data\MasterContact.xlsx', sheet_name=sheet_name)
    # Updated path
    return df

# Function to send emails in batches
def send_email_batch(sheet_name, field):
    df = read_excel(sheet_name)
    total_emails = len(df)

    for i in range(0, total_emails, 10):
        batch = df.iloc[i:i + 10]  # Get the current batch of up to 10 emails

        for index, row in batch.iterrows():
            email = row[field]  # Get the recipient's email
            name = row.get('Name', '')  # Get the name from the row, assuming 'Name' column exists
            send_individual_email(email, name)  # Send individual email using only the email address and name

        # Wait for 10 minutes after each batch, except for the last one
        if i + 10 < total_emails:
            time.sleep(600)  # Wait for 10 minutes

    flash("All emails sent successfully!")  # Notify user after all emails are sent


    for i in range(0, total_emails, 10):
        batch = df.iloc[i:i + 10]  # Get the current batch of up to 10 emails

        for index, row in batch.iterrows():
            email = row[field]  # Get the recipient's email
            send_individual_email(email)  # Send individual email using only the email address

        # Wait for 10 minutes after each batch, except for the last one
        if i + 10 < total_emails:
            time.sleep(600)  # Wait for 10 minutes

    flash("All emails sent successfully!")  # Notify user after all emails are sent


# Function to send an individual email using a pre-built msg template
def send_individual_email(email, name=None):
    try:
        # Initialize the COM library
        pythoncom.CoInitialize()

        # Start Outlook
        outlook = win32.Dispatch('outlook.application')

        # Load the pre-saved .msg file (draft email)
        msg_path = r'C:\Automailer Content\Content.msg'  # Updated path
        mail = outlook.CreateItemFromTemplate(msg_path)

        # Set the recipient email
        mail.To = email

        # Replace [Name] in the body
        if name:
            body = mail.Body.replace("[Name]", name)  # Replace placeholder with the actual name
            mail.Body = body

        # Send the email
        mail.Send()

        flash(f"Email sent to {email}")

    except Exception as e:
        flash(f"Failed to send email to {email}: {str(e)}")

    finally:
        # Uninitialize the COM library (optional, usually handled by Python)
        pythoncom.CoUninitialize()
    flash("Email Ready to be sent!")
    # Provide the list of sheets for the dropdown
    sheet_names = ['Overseas', 'Malaysia', 'Singapore', 'Others']
    return render_template('index.html', sheet_names=sheet_names)  # Pass the sheet names to the template

@app.route('/send_email', methods=['POST'])
def send_email_route():
    sheet_name = request.form.get('sheet_name')  # Access the sheet_name
    field = request.form.get('field')  # Access the field

    # Optional: Validate that all values are present
    if not sheet_name or not field:
        flash("Sheet name and field are required!")
        return redirect(url_for('index'))

    # Call send_email_batch function with the appropriate parameters
    send_email_batch(sheet_name=sheet_name, field=field)

    return redirect(url_for('index'))  # Redirect after processing

@app.route('/test_email', methods=['POST'])
def test_email():
    email = request.form['test_email']  # Get the email from the form
    send_individual_email(email)  # No need to pass 'name' anymore since we are not modifying the body
    flash("Test email sent successfully!")
    return redirect(url_for('index'))

@app.route('/logout')
def logout():
    if 'username' in session:
        record = f"{session['username']} Logout on {datetime.now()}"
        login_records.append(record)
        save_login_record(record)
        session.pop('username')
    return redirect(url_for('login'))

@app.route('/view_record', methods=['POST'])
def view_record():
    code = request.form['verification_code']
    if code == "RECORD123":
        return render_template('records.html', records=login_records)  # Path to the records template
    else:
        flash("Invalid verification code")
        return redirect(url_for('index'))

@app.route('/delete_records', methods=['POST'])
def delete_records():
    clear_login_records()
    global login_records
    login_records = []
    flash("Login records deleted successfully")
    return redirect(url_for('index'))

@app.route('/')
def index():
    sheet_names = ['Overseas', 'Malaysia', 'Singapore', 'Others']
    return render_template('index.html', sheet_names=sheet_names)


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)


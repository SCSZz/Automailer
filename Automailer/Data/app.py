from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify
import time
import pandas as pd
import os
import win32com.client as win32
from datetime import datetime
import pythoncom
import csv
import threading

app = Flask(__name__)
app.secret_key = 'your_secret_key'


RECORDS_FILE = os.path.join('Data', 'login_records.csv')
EMAIL_RECORDS_FILE = os.path.join('Data', 'email_records.csv')

# Global variable to hold email records
email_records = []
email_sending = False  # New variable to track if emails are being sent

# Initialize threading variables
paused = False
stop_flag = False  # New flag to stop the email sending
pause_condition = threading.Condition()


def load_user_credentials():
    # Path to the Excel file
    credentials_file = r'Z:\Data\UserCredentials.xlsx'

    try:
        df = pd.read_excel(credentials_file)
        credentials = dict(zip(df['Username'], df['Password']))  # Create a dictionary of {username: password}
        return credentials
    except Exception as e:
        print(f"Error loading credentials: {str(e)}")
        return {}

# Load login records from the CSV file
def load_login_records():
    if os.path.exists(RECORDS_FILE):
        with open(RECORDS_FILE, mode='r') as file:
            reader = csv.reader(file)
            return [row[0] for row in reader]  # Return list of records
    return []

# Load email records from the CSV file
def load_email_records():
    email_records_file = os.path.join('Data', 'email_records.csv')  # Specify the path for your email records
    if os.path.exists(email_records_file):
        with open(email_records_file, mode='r') as file:
            reader = csv.reader(file)
            return [row[0] for row in reader]  # Return list of email records
    return []

# Save email record to a CSV file
def save_email_record(record):
    email_records_file = os.path.join('Data', 'email_records.csv')
    with open(email_records_file, mode='a', newline='') as file:
        writer = csv.writer(file)
        writer.writerow([record])

# Function to clear email records
def clear_email_records_file():
    if os.path.exists(EMAIL_RECORDS_FILE):
        os.remove(EMAIL_RECORDS_FILE)

# Initialize login records and email records from file on startup
login_records = load_login_records()
email_records = load_email_records()  # Load email records into memory

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
    df = pd.read_excel(r'Z:\MasterContact.xlsx', sheet_name=sheet_name)  # Updated path
    return df

# Function to send emails in batches
def send_email_batch(sheet_name, field):
    global stop_flag, email_sending
    email_sending = True  # Set sending status to True
    df = read_excel(sheet_name)
    total_emails = len(df)

    for i in range(0, total_emails, 10):
        with pause_condition:
            while paused:
                pause_condition.wait()
            if stop_flag:
                flash("Email sending stopped.")
                break  # Exit the loop if stopped

        batch = df.iloc[i:i + 10]
        for index, row in batch.iterrows():
            email = row[field]
            name = row.get('Name', '')
            send_individual_email(email, name)

        if i + 10 < total_emails:
            for _ in range(600):
                with pause_condition:
                    if paused:
                        pause_condition.wait()
                    if stop_flag:
                        flash("Email sending stopped.")
                        return
                time.sleep(1)

    email_sending = False  # Set sending status to False when done
    flash("All emails sent successfully!")



# Update the send_individual_email function
def send_individual_email(email, name=None):
    try:
        pythoncom.CoInitialize()
        outlook = win32.Dispatch('outlook.application')
        msg_path = r'C:\Automailer Content\Content.msg'
        mail = outlook.CreateItemFromTemplate(msg_path)
        mail.To = email

        if name:
            body = mail.HTMLBody.replace("[Name]", name)
            mail.HTMLBody = body

        mail.Send()

        # Save the email sent record
        record = f"Email sent to {email} on {datetime.now()}"
        timestamp = datetime.now()  # Get the current timestamp
        email_record = f"Email sent to {email} on {timestamp}"  # Format the message
        email_records.append(email_record)  # Add the record to the list
        print(record)  # You may want to log this somewhere else

    except Exception as e:
        print(f"Failed to send email to {email}: {str(e)}")  # Replace flash with print or log

    finally:
        pythoncom.CoUninitialize()

@app.route('/send_email', methods=['POST'])
def send_email_route():
    global stop_flag
    stop_flag = False  # Reset stop flag before starting a new batch

    try:
        sheet_name = request.form.get('sheet_name')
        field = request.form.get('field')

        if not sheet_name or not field:
            flash("Sheet name and field are required!")
            return redirect(url_for('index'))

        # Start email sending in a new thread
        email_thread = threading.Thread(target=send_email_batch, args=(sheet_name, field))
        email_thread.start()

        return redirect(url_for('index'))

    except Exception as e:
        flash(f"An error occurred: {str(e)}")
        return redirect(url_for('index'))

@app.route('/stop_email', methods=['POST'])
def stop_email():
    global stop_flag
    stop_flag = True  # Set the stop flag to True to stop sending emails
    flash("Email sending stopped.")
    return redirect(url_for('index'))

@app.route('/pause_email', methods=['POST'])
def pause_email():
    global paused
    paused = True
    flash("Email sending paused.")
    return redirect(url_for('index'))

@app.route('/clear_email_records', methods=['POST'])
def clear_email_records():
    global email_records
    email_records.clear()  # Clear the email records

    # Define your sheet_names here or load them as needed
    sheet_names = ['Overseas', 'Malaysia', 'Singapore', 'Others']

    # Render template with both the email records and sheet names
    return render_template('index.html', email_records=email_records, sheet_names=sheet_names)

@app.route('/send_email', methods=['POST'])
def send_email():
    email = request.form['email']  # Get the email from form
    send_individual_email(email)  # Send the email
    return render_template('index.html', email_records=email_records)  # Render template with updated records

@app.route('/resume_email', methods=['POST'])
def resume_email():
    global paused
    with pause_condition:
        paused = False
        pause_condition.notify_all()  # Notify the thread to resume
    flash("Email sending resumed.")
    return redirect(url_for('index'))

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
    flash("Logged out successfully!")
    return redirect(url_for('login'))

@app.route('/view_record', methods=['POST'])
def view_record():
    # Read the verification codes from Excel
    df = pd.read_excel(r'Z:/Data/UserCredentials.xlsx')
    new_code = df.iloc[1, 3]  # Column D, row 2 (index 1, 3)

    code = request.form['verification_code']

    # Check if the entered code matches either the old or new code
    if code == new_code:
        return render_template('records.html', records=login_records)
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


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']

        users = load_user_credentials()  # Load credentials from Excel

        if username in users and users[username] == password:
            session['username'] = username
            record = f"{username} Login on {datetime.now()}"
            login_records.append(record)
            save_login_record(record)
            flash("Login successful!")
            return redirect(url_for('index'))
        else:
            flash("Invalid username or password")

    return render_template('login.html')

# Function to check if user is logged in
def is_logged_in():
    return 'username' in session

@app.route('/')
def index():
    if not is_logged_in():
        flash("Please log in to access the application.")
        return redirect(url_for('login'))

    sheet_names = ['Overseas', 'Malaysia', 'Singapore', 'Others']
    return render_template('index.html', sheet_names=sheet_names, email_records=email_records, email_sending=email_sending)

@app.route('/get_email_records', methods=['GET'])
def get_email_records():
    # Return email records in JSON format
    return jsonify(email_records)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True, use_reloader=False)
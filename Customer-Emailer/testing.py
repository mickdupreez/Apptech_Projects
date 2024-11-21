import csv  # For reading and processing CSV files containing customer data
import subprocess  # To run AppleScript for sending emails via macOS Mail
import re  # For using regular expressions to clean and format email addresses
import time  # To introduce delays between sending emails (rate limiting)
import os  # For interacting with the file system (checking if files exist)
from datetime import datetime, timedelta  # For working with dates and times (tracking when emails were sent)
import logging  # For logging errors to a file

# Setup logging to capture any errors that occur during the script's execution
logging.basicConfig(filename='email_errors.log', level=logging.ERROR, format='%(asctime)s %(message)s')

# Paths to the files that the script will use
csv_file = "workbench_listings_workorders.csv"  # CSV file containing customer data
geelong_template_file = "geelong_email_template.txt"  # Email template for Geelong
mt_waverley_template_file = "mt_waverley_email_template.txt"  # Email template for Mt Waverley
sent_emails_file = "sent_emails.txt"  # File to track emails that have already been sent
confirmation_file = "email_confirmation_list.txt"  # File to list all emails for confirmation

# Rate limiting settings
PAUSE_TIME = 3  # Time (in seconds) between sending individual emails
BATCH_SIZE = 50  # Number of emails to send in each batch before pausing

def clean_email(email):
    """
    Clean the email by removing trailing commas, numbers, 'SVC', and extra whitespace.
    """
    # Strip whitespace
    email = email.strip()
    # Remove trailing ', SVC' (case-insensitive)
    email = re.sub(r',\s*SVC$', '', email, flags=re.IGNORECASE)
    # Remove trailing commas
    email = email.rstrip(',')
    # Remove trailing numbers (digits at the end)
    email = re.sub(r'\d+$', '', email)
    # Validate email format (basic check)
    if '@' not in email or not re.match(r"[^@]+@[^@]+\.[^@]+", email):
        return None
    return email

def send_email_via_macos_mail(name, email, subject, body):
    """
    Sends an email using macOS Mail via AppleScript.
    """
    email = clean_email(email)  # Ensure email is cleaned before sending
    if not email:
        logging.error(f"Invalid email address for {name}: {email}")
        return  # Skip sending invalid emails

    # Prepare an AppleScript to send an email via the macOS Mail app
    applescript = f'''
    tell application "Mail"
        set newMessage to make new outgoing message with properties {{subject:"{subject}", content:"{body}", visible:true}}
        tell newMessage
            make new to recipient at end of to recipients with properties {{address:"{email}"}}
            send
        end tell
    end tell
    '''
    process = subprocess.run(['osascript', '-e', applescript], capture_output=True)
    
    if process.returncode == 0:
        print(f"Email successfully sent to {email}")
    else:
        logging.error(f"Failed to send email to {email}. Error: {process.stderr.decode('utf-8')}")
        print(f"Failed to send email to {email}.")

def read_email_template(file_path):
    """
    Read an email template and validate presence of required placeholders.
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as template_file:
            template = template_file.read()
            if '{name}' not in template:
                logging.error(f"Template {file_path} is missing the required placeholder {{name}}.")
                print(f"Error: The email template {file_path} is missing the required {{name}} placeholder.")
                return None
            return template
    except FileNotFoundError:
        logging.error(f"Email template file {file_path} was not found.")
        print(f"Error: The email template file {file_path} was not found.")
        return None

def select_email_template():
    """
    Prompt the user to select an email template based on location.
    """
    while True:
        location = input("Enter the location (G for Geelong or M for Mt Waverley): ").strip().upper()
        if location == "G":
            return geelong_template_file
        elif location == "M":
            return mt_waverley_template_file
        else:
            print("Invalid location. Please choose either 'G' for Geelong or 'M' for Mt Waverley.")

def load_sent_emails(file_path):
    """
    Load the list of sent emails from a file and parse their timestamps.
    """
    sent_emails = {}
    if os.path.exists(file_path):
        with open(file_path, 'r', encoding='utf-8') as file:
            for line in file:
                parts = line.strip().split(',')
                if len(parts) == 2:
                    email, timestamp = parts
                    email = clean_email(email)  # Clean the email address
                    if email:
                        try:
                            sent_emails[email] = datetime.strptime(timestamp, "%Y-%m-%d")
                        except ValueError:
                            logging.error(f"Invalid date format in sent_emails.txt for {email}. Skipping entry.")
                else:
                    logging.error(f"Incorrect format in sent_emails.txt: {line}. Skipping entry.")
    return sent_emails

def save_sent_email(file_path, email):
    """
    Save the email to the sent emails log file after cleaning.
    """
    email = clean_email(email)  # Ensure email is cleaned before saving
    if not email:
        logging.error(f"Attempted to save invalid email: {email}")
        return
    with open(file_path, 'a', encoding='utf-8') as file:
        file.write(f"{email},{datetime.now().strftime('%Y-%m-%d')}\n")

def prepare_email_confirmation_list(csv_file):
    """
    Prepare an email confirmation list from the CSV file.
    """
    email_list = []
    with open(csv_file, newline='', encoding='utf-8') as file:
        reader = csv.DictReader(file)
        for row in reader:
            name = row.get('Customer', '').strip()
            email = clean_email(row.get('Hook In, Out', '').strip())
            if email:
                email_list.append((name, email))

    with open(confirmation_file, 'w', encoding='utf-8') as file:
        for name, email in email_list:
            file.write(f"{name}, {email}\n")
    
    print(f"Email confirmation list saved to {confirmation_file}. Please review and confirm before sending emails.")

def send_emails_from_csv(csv_file):
    """
    Main function to send emails from a CSV file.
    """
    prepare_email_confirmation_list(csv_file)
    confirm = input("Do you want to proceed with sending emails? (y/n): ").strip().lower()
    if confirm != 'y':
        print("Email sending canceled.")
        return
    
    email_template_file = select_email_template()
    email_template = read_email_template(email_template_file)
    
    if email_template is None:
        return

    sent_emails = load_sent_emails(sent_emails_file)

    try:
        with open(csv_file, newline='', encoding='utf-8') as file:
            reader = csv.DictReader(file)
            processed = 0

            for row in reader:
                name = row.get('Customer', '').strip()
                email = clean_email(row.get('Hook In, Out', '').strip())
                
                if not email:
                    logging.error(f"Skipping {name} due to invalid email: {email}")
                    continue
                
                last_sent_date = sent_emails.get(email)
                if last_sent_date and last_sent_date > datetime.now() - timedelta(days=30):
                    continue
                
                subject = f"Hello {name}"
                body = email_template.format(name=name)
                
                send_email_via_macos_mail(name, email, subject, body)
                save_sent_email(sent_emails_file, email)
                sent_emails[email] = datetime.now()

                processed += 1
                if processed % BATCH_SIZE == 0:
                    print("Batch complete, taking a break to avoid triggering spam filters.")
                    time.sleep(60)
                
                time.sleep(PAUSE_TIME)

    except FileNotFoundError:
        logging.error(f"Error: The file {csv_file} was not found.")
        print(f"Error: The file {csv_file} was not found.")
    except KeyError as e:
        logging.error(f"Missing expected column {str(e)} in the CSV file.")
        print(f"Error: Missing expected column {str(e)} in the CSV file.")

# Start the email sending process by reading the CSV file
send_emails_from_csv(csv_file)

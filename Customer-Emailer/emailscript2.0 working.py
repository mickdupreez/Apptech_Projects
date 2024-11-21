import csv  # For reading and processing CSV files containing customer data
import subprocess  # To run AppleScript for sending emails via macOS Mail
import re  # For using regular expressions to clean and format email addresses
import time  # To introduce delays between sending emails (rate limiting)
import os  # For interacting with the file system (checking if files exist)
from datetime import datetime, timedelta  # For working with dates and times (tracking when emails were sent)
import logging  # For logging errors to a file

# Setup logging to capture any errors that occur during the script's execution
logging.basicConfig(filename='email_errors.log', level=logging.ERROR, format='%(asctime)s %(message)s')

# Define the path to the CSV file containing customer data
csv_file = "workbench_listings_workorders.csv"  # Input file with customer information
# Define the paths to email templates for specific locations
geelong_template_file = "geelong_email_template.txt"  # Email template for Geelong location
mt_waverley_template_file = "mt_waverley_email_template.txt"  # Email template for Mt Waverley location
# Define the file that tracks which emails have already been sent
sent_emails_file = "sent_emails.txt"  # Tracks emails to avoid duplicates
# Define the file that contains the list of emails prepared for confirmation
confirmation_file = "email_confirmation_list.txt"  # Lists emails for manual review before sending

# Set the time delay (in seconds) between sending individual emails to avoid being flagged as spam
PAUSE_TIME = 3  # Wait 3 seconds between sending emails
# Set the batch size for emails to send before taking a longer break
BATCH_SIZE = 50  # Send 50 emails, then pause for a longer duration

def clean_email(email):
    """
    Cleans email addresses by removing unwanted characters, patterns, or formats.
    """
    email = email.strip()  # Remove leading and trailing spaces
    email = re.sub(r',\s*SVC$', '', email, flags=re.IGNORECASE)  # Remove ', SVC' at the end (case-insensitive)
    email = email.rstrip(',')  # Remove any trailing commas
    email = re.sub(r'\d+$', '', email)  # Remove any trailing numbers
    # Check if the email contains '@' and matches a basic email pattern
    if '@' not in email or not re.match(r"[^@]+@[^@]+\.[^@]+", email):
        return None  # Return None if email is invalid
    return email  # Return the cleaned email

def send_email_via_macos_mail(name, email, subject, body):
    """
    Sends an email via macOS Mail using AppleScript.
    """
    email = clean_email(email)  # Ensure the email is cleaned before sending
    if not email:  # Skip sending if the email is invalid
        logging.error(f"Invalid email address for {name}: {email}")  # Log the error
        return

    # Create the AppleScript to send the email
    applescript = f'''
    tell application "Mail"
        set newMessage to make new outgoing message with properties {{subject:"{subject}", content:"{body}", visible:true}}
        tell newMessage
            make new to recipient at end of to recipients with properties {{address:"{email}"}}
            send
        end tell
    end tell
    '''
    # Run the AppleScript using the subprocess module
    process = subprocess.run(['osascript', '-e', applescript], capture_output=True)
    
    if process.returncode == 0:  # Check if the AppleScript ran successfully
        print(f"Email successfully sent to {email}")  # Notify success
    else:  # Log and print an error if the script failed
        logging.error(f"Failed to send email to {email}. Error: {process.stderr.decode('utf-8')}")
        print(f"Failed to send email to {email}.")

def read_email_template(file_path):
    """
    Reads an email template from a file and validates its content.
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as template_file:  # Open the template file
            template = template_file.read()  # Read the content
            # Check if the placeholder '{name}' exists in the template
            if '{name}' not in template:
                logging.error(f"Template {file_path} is missing the required placeholder {{name}}.")  # Log the error
                print(f"Error: The email template {file_path} is missing the required {{name}} placeholder.")  # Notify the user
                return None
            return template  # Return the template content if valid
    except FileNotFoundError:  # Handle missing file errors
        logging.error(f"Email template file {file_path} was not found.")  # Log the error
        print(f"Error: The email template file {file_path} was not found.")  # Notify the user
        return None

def select_email_template():
    """
    Prompts the user to select the appropriate email template based on the location.
    """
    while True:  # Loop until a valid input is provided
        location = input("Enter the location (G for Geelong or M for Mt Waverley): ").strip().upper()  # Ask for location
        if location == "G":  # If user selects Geelong
            return geelong_template_file  # Return Geelong template file path
        elif location == "M":  # If user selects Mt Waverley
            return mt_waverley_template_file  # Return Mt Waverley template file path
        else:  # If input is invalid, prompt again
            print("Invalid location. Please choose either 'G' for Geelong or 'M' for Mt Waverley.")

def load_sent_emails(file_path):
    """
    Reads the list of previously sent emails and their timestamps from a file.
    """
    sent_emails = {}  # Dictionary to store emails and their last sent date
    if os.path.exists(file_path):  # Check if the file exists
        with open(file_path, 'r', encoding='utf-8') as file:  # Open the file
            for line in file:  # Read each line in the file
                parts = line.strip().split(',')  # Split the line into email and timestamp
                if len(parts) == 2:  # Ensure the line contains both email and timestamp
                    email, timestamp = parts
                    email = clean_email(email)  # Clean the email
                    if email:  # If email is valid
                        try:
                            sent_emails[email] = datetime.strptime(timestamp, "%Y-%m-%d")  # Parse the timestamp
                        except ValueError:  # Handle invalid timestamp formats
                            logging.error(f"Invalid date format in sent_emails.txt for {email}. Skipping entry.")
                else:  # Log if the line format is incorrect
                    logging.error(f"Incorrect format in sent_emails.txt: {line}. Skipping entry.")
    return sent_emails  # Return the dictionary of sent emails

def save_sent_email(file_path, email):
    """
    Saves a newly sent email to the log file.
    """
    email = clean_email(email)  # Ensure the email is cleaned
    if not email:  # Skip saving if the email is invalid
        logging.error(f"Attempted to save invalid email: {email}")  # Log the error
        return
    with open(file_path, 'a', encoding='utf-8') as file:  # Open the file in append mode
        file.write(f"{email},{datetime.now().strftime('%Y-%m-%d')}\n")  # Append the email and timestamp

def prepare_email_confirmation_list(csv_file):
    """
    Generates a list of emails for review before sending them.
    """
    email_list = []  # List to store prepared emails
    with open(csv_file, newline='', encoding='utf-8') as file:  # Open the CSV file
        reader = csv.DictReader(file)  # Read the CSV data as a dictionary
        for row in reader:  # Process each row
            name = row.get('Customer', '').strip()  # Get and clean the customer name
            email = clean_email(row.get('Hook In, Out', '').strip())  # Get and clean the email
            if email:  # If the email is valid
                email_list.append((name, email))  # Add the email and name to the list

    with open(confirmation_file, 'w', encoding='utf-8') as file:  # Open the confirmation file for writing
        for name, email in email_list:  # Write each email and name to the file
            file.write(f"{name}, {email}\n")
    
    print(f"Email confirmation list saved to {confirmation_file}. Please review and confirm before sending emails.")

def send_emails_from_csv(csv_file):
    """
    Main function that orchestrates the email sending process from the CSV file.
    """
    prepare_email_confirmation_list(csv_file)  # Generate the email confirmation list
    confirm = input("Do you want to proceed with sending emails? (y/n): ").strip().lower()  # Ask for confirmation
    if confirm != 'y':  # Cancel if the user does not confirm
        print("Email sending canceled.")  # Notify the user
        return
    
    email_template_file = select_email_template()  # Select the appropriate email template
    email_template = read_email_template(email_template_file)  # Read the selected template
    
    if email_template is None:  # If the template is invalid, stop execution
        return

    sent_emails = load_sent_emails(sent_emails_file)  # Load the list of sent emails

    try:
        with open(csv_file, newline='', encoding='utf-8') as file:  # Open the CSV file
            reader = csv.DictReader(file)  # Read the CSV data as a dictionary
            processed = 0  # Counter for processed emails

            for row in reader:  # Process each row in the CSV
                name = row.get('Customer', '').strip()  # Extract and clean the customer name
                email = clean_email(row.get('Hook In, Out', '').strip())  # Extract and clean the email
                
                if not email:  # Skip invalid emails
                    logging.error(f"Skipping {name} due to invalid email: {email}")
                    continue
                
                last_sent_date = sent_emails.get(email)  # Check the last sent date for this email
                if last_sent_date and last_sent_date > datetime.now() - timedelta(days=30):  # Skip if sent recently
                    continue
                
                subject = f"Hello {name}"  # Create the email subject
                body = email_template.format(name=name)  # Customize the email body with the name
                
                send_email_via_macos_mail(name, email, subject, body)  # Send the email
                save_sent_email(sent_emails_file, email)  # Log the sent email
                sent_emails[email] = datetime.now()  # Update the sent emails list

                processed += 1  # Increment the processed counter
                if processed % BATCH_SIZE == 0:  # If a batch is complete
                    print("Batch complete, taking a break to avoid triggering spam filters.")  # Notify the user
                    time.sleep(60)  # Pause for 60 seconds
                
                time.sleep(PAUSE_TIME)  # Pause between emails

    except FileNotFoundError:  # Handle missing file errors
        logging.error(f"Error: The file {csv_file} was not found.")
        print(f"Error: The file {csv_file} was not found.")
    except KeyError as e:  # Handle missing columns in the CSV file
        logging.error(f"Missing expected column {str(e)} in the CSV file.")
        print(f"Error: Missing expected column {str(e)} in the CSV file.")

# Start the process by invoking the main function
send_emails_from_csv(csv_file)

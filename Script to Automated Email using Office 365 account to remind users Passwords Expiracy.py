import smtplib  # Module for sending emails
# For constructing the email body in HTML format
from email.mime.text import MIMEText
# For constructing multi-part emails
from email.mime.multipart import MIMEMultipart
import ssl  # For secure connection
from datetime import datetime, timedelta  # For handling date and time
# Module for connecting to Active Directory
from ldap3 import Server, Connection, ALL

# Configuration variables
smtp_server = "smtp.office365.com"  # Office 365 SMTP server
smtp_port = 587  # SMTP port for Office 365 (uses TLS)
smtp_user = "IT_TEAM@example.com"  # Office 365 account email for sending
smtp_password = "password"  # Office 365 account password

expire_in_days = 5  # Number of days before password expiration to notify the user
from_email = "IT_TEAM@example.com"  # From address in the email
logging_enabled = True  # Enable or disable logging
# Log file name
log_file = f"PasswordChangeNotification_{datetime.utcnow().strftime('%Y%m%dT%H%M%S')}.csv"
testing_enabled = False  # Disable or Enable sending emails in test mode
test_recipient = "recipient@example.com"  # Test recipient email address

# Function to send emails


def send_email(to_email, subject, body):
    # Create a multi-part message
    msg = MIMEMultipart()
    msg['From'] = smtp_user
    msg['To'] = to_email
    msg['Subject'] = subject

    # Attach the email body as HTML content
    msg.attach(MIMEText(body, 'html'))

    # Create a secure SSL context
    context = ssl.create_default_context()

    # Connect to the SMTP server and send the email
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls(context=context)  # Secure the connection using TLS
            # Login to the email account
            server.login(smtp_user, smtp_password)
            server.sendmail(smtp_user, to_email,
                            msg.as_string())  # Send the email
        print(f"Email sent to {to_email}")
    except Exception as e:
        print(f"Error sending email: {e}")

# Function to log events to a CSV file


def log_event(log_message):
    if logging_enabled:  # Check if logging is enabled
        with open(log_file, 'a') as f:
            f.write(log_message + "\n")  # Write the log message to the file

# Function to fetch users from Active Directory


def get_ad_users():
    # Connect to the Active Directory server
    server = Server('ldap://your_ad_server', get_info=ALL)
    # Bind the connection with credentials
    conn = Connection(server, 'username', 'password', auto_bind=True)

    # Search for users with non-expired passwords in Active Directory
    conn.search('dc=yourdomain,dc=com',
                '(&(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(pwdLastSet>=0))',
                attributes=['cn', 'mail', 'pwdLastSet'])

    users = conn.entries  # Fetch the user entries
    return users

# Main function to process users and send password expiry notifications


def process_users():
    users = get_ad_users()  # Fetch users from Active Directory
    for user in users:
        name = user.cn  # Get user's name
        email = user.mail  # Get user's email address
        pwd_last_set = user.pwdLastSet.value  # Get the last time the password was set
        # Default max password age in AD is usually 90 days
        max_password_age = timedelta(days=90)

        # Calculate the password expiration date
        expires_on = pwd_last_set + max_password_age
        # Days until password expiration
        days_to_expire = (expires_on - datetime.now()).days

        # Only process if the password is set to expire within the defined period
        if days_to_expire <= expire_in_days and days_to_expire >= 0:
            # Determine the expiration message based on days remaining
            message_days = f"in {days_to_expire} days" if days_to_expire > 0 else "today"
            # Email subject line
            subject = f"Your password will expire {message_days}"

            # Email body (HTML format)
            body = f"""
            <p>Dear User {name},<br></p>
            <p><strong>Polish:</strong><br>
            Uprzejmie przypominamy, że za {message_days} wygaśnie Twoje hasło. Pamiętaj, aby je zmienić.<br>
            Jeśli potrzebujesz pomocy możesz:<br>
                - skorzystać z artykułu pomocy technicznej dostępnego tutaj: <a href='https://link/to/article/'>Jak zmienić hasło?</a><br>
                - skontaktować się z Helpdeskiem IT: <a href='mailto:helpdesk@example.com'>helpdesk@example.com</a><br>
            </p>
            <p><strong>English:</strong><br>
            We kindly inform you that your password will expire {message_days}. Please remember to change it.<br>
            If you need help:<br>
                - you can read a dedicated article: <a href='https://link/to/article/'>How to change your password?</a><br>
                - contact our Helpdesk: <a href='mailto:helpdesk@example.com'>helpdesk@example.com</a><br>
            </p>
            <p>Best regards,<br>IT Helpdesk</p>"""

            # If testing is enabled, send the email to the test recipient
            if testing_enabled:
                email = test_recipient

            # If the user has no valid email, default to a known email
            if not email:
                email = "it@example.com"

            # Log the action if logging is enabled
            log_event(
                f"{datetime.now().strftime('%Y-%m-%d')},{name},{email},{days_to_expire},{expires_on}")

            # Send the email notification
            send_email(email, subject, body)


# Main script execution
if __name__ == "__main__":
    process_users()  # Start processing users

# Script-to-Automated-Email-using-Office-365-account(SMTP)-to-remind-users-Passwords-Expiracy
Script Description: Automated Email for Password Expiry Notification Using Office 365 in Python
This Python script is designed to automatically send email reminders to Active Directory (AD) users whose passwords are about to expire. The script connects to the Active Directory via LDAP, SMTP retrieves users whose passwords are nearing expiration, and sends personalized notifications using an Office 365 email account.

Main Features:
Active Directory Integration (LDAP): Uses LDAP to query AD and retrieve users’ data, including their password expiration date.
Customizable Reminder Period: Allows customization of how many days before the password expiration users should receive a reminder.
Office 365 SMTP Integration: Sends email using Office 365’s SMTP server 
Email Notifications: Sends HTML-formatted email reminders in  English/Polish, with instructions on how to change the password and contact IT support.
Logging: Optionally logs details of the notifications sent, including usernames, email addresses, and password expiry details.
Test Mode: Includes a test mode to verify email functionality by sending notifications to a designated test recipient before sending to actual users.

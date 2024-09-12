**Email Automation PowerShell Script**
Introduction and Overview
> This repository contains a PowerShell script designed to automate the creation and sending of emails. The script utilizes Windows Forms to provide a graphical user interface (GUI) for user input.

**Purpose of the Application**
The application aims to:

> Simplify the process of sending emails with predefined templates.
Ensure consistency in email formatting and content.
Save time by automating repetitive email tasks.
GUI Components and Design
The script creates a form with various input fields, including:

**Fields**
> To
CC
BCC
Subject
Body
Attachment

***Additionally, it includes a ComboBox for selecting predefined email templates. TextBox controls are dynamically added for each email field, and a footer label is included to display copyright information. The form also includes buttons for various actions.***

**Buttons**
> Add Attachment: Attach a file to the email.
Upload File: Upload a file directly to SharePoint.
Remove Attachment: Remove an attached file from the email.
Send Email: Send the composed email.
Preview Email: Open a pop-up window in Outlook to preview the email.
Upload to Refresh Pivots: Refresh large Excel files directly, avoiding manual refresh and potential crashes.

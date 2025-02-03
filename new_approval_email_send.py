import re
import win32com.client as win32
import win32timezone
from datetime import datetime
import win32com.client
import pandas as pd
import openpyxl
import logging
import os
import glob
import sys
from dateutil.parser import parse

# Get the current date
current_date = datetime.now().strftime("%Y%m%d")

def create_folder(base_path='D:\\TRQ_sheet\\Latest_emails\\'):
    """
    Creates a folder named with the current date under the specified base path.
    If the folder already exists, it prints a message instead.

    :param base_path: The root directory where the date folder will be created.
    """
    # Get the current date in "YYYYMMDD" format
    #current_date = datetime.now().strftime("%Y%m%d")

    global current_date
    
    # Define the folder path
    log_folder = os.path.join(base_path, current_date)
    
    # Check if the folder already exists
    if os.path.exists(log_folder):
        print(f"Folder '{log_folder}' already exists.")
    else:
        # Create the folder
        os.makedirs(log_folder)
        print(f"Folder '{log_folder}' created successfully.")
    
    return log_folder


def get_valid_time():
    """
    Function to prompt the user for a valid time in HH:MM format or exit the program.
    
    Returns:
        str: A valid time in HH:MM format.
    """
    while True:  # Keep the program running until the user decides to exit
        s_time = input("Enter start time (HH:MM) or type 'exit' to quit: ")
        if s_time.lower() == 'exit':
            print("Exiting the program. Goodbye!")
            return None  # Graceful exit
        
        # Validate the time format using a regular expression
        if re.match(r'^[0-2][0-9]:[0-5][0-9]$', s_time):
            return s_time  # Return the valid time
        else:
            print("Invalid time format! Please provide time in HH:MM format (e.g., 13:45).")



def get_client_names():
    """
    Function to prompt the user for client names, ensuring at least one name is provided.
    
    Returns:
        list: A list of client names.
    """
    while True:
        print('_______________________________________________________')
        client_names_input = input("Please provide client names (comma-separated) or type 'exit' to quit: ")
        if client_names_input.lower() == 'exit':
            print("Exiting the program. Goodbye!")
            return None  # Gracefully return None to indicate exit
        
        # Split the input into a list of names and validate
        client_names = [name.strip() for name in client_names_input.split(',') if name.strip()]
        if client_names:
            return client_names  # Return the list of valid client names
        else:
            print("No names provided! Please enter at least one client name.")


#Get the TRQ excel sheet file path
def get_latest_excel_file(base_path):
    """
    Get the full path of the latest Excel file in the specified directory.

    Parameters:
        base_path (str): The directory to search for Excel files.

    Returns:
        str: The full path of the latest Excel file found.
    """
    # Use glob to find all Excel files in the directory
    excel_files = glob.glob(os.path.join(base_path, "*.xls*"))
    #print(excel_files)
    
    # Check if any Excel files exist
    if not excel_files:
        raise FileNotFoundError("No Excel files found in the specified directory.")
    
    # Find the latest modified file
    latest_file = max(excel_files, key=os.path.getmtime)

    full_path = os.path.join(base_path, latest_file)
    #print(f"File Path is :",full_path)
    
    return full_path


def get_emails_after_time(start_time):
    """
    Fetch emails from the Inbox after a specific time on today's date.

    Parameters:
        start_time (str): Start time in "HH:MM" format.

    Returns:
        list: Filtered emails received after the specified time.
    """
    # Initialize Outlook
    outlook = win32.Dispatch('outlook.application')
    namespace = outlook.GetNamespace("MAPI")

    # Access the Inbox folder
    inbox = namespace.GetDefaultFolder(6)  # 6 refers to the Inbox folder

    # Get today's date
    today = datetime.now().strftime("%d/%m/%Y")  # Format required by Outlook

    # Combine today's date with the provided start time
    print("********************************************************")
    start_datetime = f"{today} {start_time}"

    # Debugging output
    print(f"Start DateTime: {start_datetime}")

    # Fetch and filter emails
    emails = inbox.Items
    filter_query = f"[ReceivedTime] >= '{start_datetime}'"
    print(f"Filter Query: {filter_query}")  # Debugging
    print("********************************************************")

    filtered_emails = emails.Restrict(filter_query)

    # Check if any emails were retrieved
    if filtered_emails.Count == 0:
        print("No emails found after the specified time.")
        return []

    # Return filtered emails
    return filtered_emails


def get_client_data_as_html(client_name, excel_file_path):
    """
    Fetch the row from the Excel file where the client name matches
    and return it as an HTML table styled like an Excel sheet.

    Args:
        client_name (str): Name of the client to search for.
        excel_file_path (str): Path to the Excel file.

    Returns:
        str: HTML representation of the row styled like an Excel sheet.
    """
    # Load the Excel file using openpyxl
    wb = openpyxl.load_workbook(excel_file_path, data_only=True)
    sheet = wb.active

    # Extract the header row (first row) and limit columns dynamically
    header = [cell.value for cell in sheet[1] if cell.value is not None]
    num_columns = len(header)  # Determine the actual number of columns based on header

    # Search for the client name in the 'TRQ Holders Name' column
    client_row = None
    for row in sheet.iter_rows(min_row=2, max_col=num_columns):  # Exclude extra columns
        if row[1].value and client_name.lower() in str(row[1].value).lower():
            client_row = row
            break

    if client_row:
        # Start building the HTML table
        table_html = '<table border="1" style="border-collapse: collapse; width: 100%; font-family: Arial; font-size: 12px;">'

        # Add the header row
        table_html += '<tr style="background-color: #f2f2f2;">'
        for cell in header:
            table_html += f'<th style="padding: 8px; text-align: center;">{cell}</th>'
        table_html += '</tr>'

        # Add the client's data row
        table_html += '<tr>'
        for cell in client_row[:num_columns]:
            if cell.value is None:
                table_html += '<td style="padding: 8px; text-align: center;"></td>'
            elif isinstance(cell.value, (int, float)) and cell.number_format.endswith('%'):
                table_html += f'<td style="padding: 8px; text-align: center;">{cell.value * 100:.2f}%</td>'
            elif isinstance(cell.value, (int, float)):
                table_html += f'<td style="padding: 8px; text-align: center;">{cell.value:.2f}</td>'
            else:
                table_html += f'<td style="padding: 8px; text-align: center;">{cell.value}</td>'
        table_html += '</tr>'

        # Close the HTML table
        table_html += '</table>'

        return table_html
    else:
        # Return a message if no data is found for the client
        return "<p>No data found for the client in the Excel file.</p>"



def find_latest_reply_email(client_name, inbox):
    """
    Find the latest reply email from the client with the subject starting with
    'Re: Submission of AD Letter Request - {client_name}' or 'RE: Submission of AD Letter Request - {client_name}' and received today.

    Args:
        client_name (str): Client name to filter the emails.
        inbox (Outlook Folder): The inbox folder to search emails.

    Returns:
        MailItem or None: The latest reply email if found, otherwise None.
    """
    latest_email = None
    latest_received_time = None
    subject_filter = f"Submission of AD Letter Request - {client_name}"

    # Get today's date in the correct format
    today = datetime.now().date()

    for mail in inbox.Items:
        if mail.Subject.lower().startswith(f"re: {subject_filter.lower()}"):
            received_time = mail.ReceivedTime.date()  # Get the date part only

            # Check if the email is from today
            if received_time == today:
                received_time_full = mail.ReceivedTime  # Full datetime for comparison
                if latest_email is None or received_time_full > latest_received_time:
                    latest_email = mail
                    latest_received_time = received_time_full

    return latest_email


def get_client_data_and_send_email():
    #Create a latest Email reply folder date wise
    folder_path =create_folder()
    print(f"Folder is Created {folder_path}")

    # TRQ sheet file path
    base_path = r"D:\TRQ_sheet"

    while True:  # Keep the program running until the user decides to exit
        input_time = get_valid_time()
        print(f"You have entered {input_time}")


        # Get emails after the specified start time
        emails = get_emails_after_time(input_time)

        # Get the latest Excel file
        excel_full_path = get_latest_excel_file(base_path)

        # Prompt the user for client names
        print('_______________________________________________________')
        #client_names_input = input("Please provide client names (comma-separated) or type 'exit' to quit: ")
        #if client_names_input.lower() == 'exit':
        #    print("Exiting the program. Goodbye!")
        #    break  # Exit the loop and program

        client_names = get_client_names()
        if client_names:
            print(f"Client Name : {client_names}")
        else:
            print('No Client Name is provided')
            break

        #client_names = [name.strip() for name in client_names_input.split(",")]

        # Initialize Outlook Application
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)  # 6 refers to the Inbox folder

       # Check if any emails match the criteria
        if len(emails) > 0:  # Use len() if emails is a list
            email_found = False  # Flag to track if a relevant email is found
            for mail in emails:
                if mail.Subject.startswith("Submission of AD Letter Request"):
                    for client_name in client_names:
                        # Check if the client name is present in the email body
                        if client_name in mail.Body:
                            email_found = True
                            print(f"Found email for client name '{client_name}'.")
                            
                            # Extract contract details from the email body
                            email_body = mail.Body  # Use plain text body for parsing
        
                            # Use regular expressions to find Contract Name and Quantity
                            contract_name_match = re.search(r"Contract Name\s*[:-]\s*(.*)", email_body, re.IGNORECASE)
                            quantity_match = re.search(r"Quantity\s*[:-]\s*(.*)", email_body, re.IGNORECASE)
        
                            contract_name = contract_name_match.group(1).strip() if contract_name_match else "N/A"
                            quantity = quantity_match.group(1).strip() if quantity_match else "N/A"
        
                            # Get additional data from the Excel file
                            client_data_html = get_client_data_as_html(client_name, excel_full_path)
        
                            # Find the latest reply email
                            latest_reply = find_latest_reply_email(client_name, inbox)
        
                            # Create a reply to the email
                            reply = mail.Reply()
        
                            # Customize the reply HTML body
                            reply.HTMLBody = f"""
                            <p>Dear Sir,</p>
                            <p> </p>
                            <p>Kindly allow us to use your digital sign for signing the AD Bank letter, your DSC is available with us, details as mentioned below.</p>
                            <ul>
                                <li><strong>Contract Name:</strong> {contract_name}</li>
                                <li><strong>Quantity:</strong> {quantity}</li>
                            </ul>
                            <p>  </p>
                            <p>  </p>
                            <p>  </p>
                            {client_data_html}  <!-- Insert data from Excel -->
                            <p>Regards,<br>Trading Operations Department</p>
                            <p><img src="D:\\Python_projects\\Send_approval_email\\iibx_logo.png" alt="IIBX Logo" width="200" height="auto" /></p>  <!-- Logo -->
                            <p>Unit No. 1302A, Brigade International Financial Centre,<br>13th Floor, Building No. 14A, Block 14,<br>Zone 1, GIFT SEZ, GIFT CITY,<br>Gandhinagar, 382050, Gujarat<br>Direct : 079 6969 7118 </p>
                            <p>  </p>  <!-- Blank line -->
                            {mail.HTMLBody}  <!-- Include original email content -->
                            """  # Add reply content

                            #Get the current date time
                            current_date_time = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
                            
                            if latest_reply:
                                # Save the latest reply email to a temporary file
                                reply_attachment_path = os.path.join(folder_path, f"{client_name}_{current_date_time}.msg")
                                latest_reply.SaveAs(reply_attachment_path)
        
                                # Attach the latest reply email
                                reply.Attachments.Add(reply_attachment_path)
                            else:
                                print(f"No reply email found from '{client_name}' for today's date. Skipping attachment.")
        
                            # Change the "To" recipient and add CC recipients
                            reply.To = "vinod.ramachandran@iibx.co.in;chetan.pabari@iibx.co.in"
                            #reply.To = "siddharth.thorat@iibx.co.in" # Replace with the actual recipient
                            reply.CC = "trading.operations@iibx.co.in;surveillance@iibx.co.in;bd@iibx.co.in;cs.ops@iibx.co.in;"  # Replace with CC email IDs
        
                            # Display the email (do not send)
                            reply.Display()  # This will open the email in the Outlook editor
        
                            print(f"Draft email prepared for {mail.SenderName} and client name '{client_name}'.")
        
            # If no relevant email was found after processing all emails
            if not email_found:
                print(f"No emails found matching the subject 'Submission of AD Letter Request' or any of the provided client names.")
        else:
            print(f"No emails found in your inbox after the specified start time: {input_time}")


if __name__ == "__main__":
    #Fetch Client data and Reply on submission email
    get_client_data_and_send_email()
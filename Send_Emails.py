import urllib.parse
import base64
import os
from typing import Any
import msal
import requests
import pandas as pd


# Get the sensitive information from environment variables
def load_env_file(env_file: str) -> None:
    """
    Loads environment variables from a specified file.

    :param env_file: Path to the .env file.
    """
    with open(env_file, 'r') as file:
        for line in file:
            if line.strip() and not line.startswith('#'):
                key, value = line.strip().split('=', 1)
                os.environ[key] = value


# Load environment variables from the .env file
load_env_file('azure.env')
# Note - this file should be in the same directory as the script.  It is listed in the .gitignore file so it 
# is not uploaded to the repository.  You will need to create this file and add the following lines with no quotes:
# SECRET_VALUE=your_secret_value
# CLIENT_ID=your_client_id
# TENANT_ID=your_tenant_id

SECRET_VALUE = os.getenv('SECRET_VALUE')
CLIENT_ID = os.getenv('CLIENT_ID')
TENANT_ID = os.getenv('TENANT_ID')

# Check if the environment variables are loaded correctly
if not SECRET_VALUE or not CLIENT_ID or not TENANT_ID:
    raise ValueError("Missing required environment variables.")

# This is the sender email of the application.
FROM_EMAIL = 'seattle.lab@shanwil.com'
# FROM_EMAIL = 'oliver.hoopes@shanwil.com'


def get_access_token_graph() -> str or None:
    """
    Get the access token for MS Graph application
    :return:
    """
    config = {
        "authority": f"https://login.microsoftonline.com/{TENANT_ID}",
        "scope": ["https://graph.microsoft.com/.default"],
        "endpoint": "https://graph.microsoft.com/v1.0/users"
    }

    app = msal.ConfidentialClientApplication(
        authority=config['authority'],
        client_id=CLIENT_ID,
        client_credential=SECRET_VALUE
    )

    # The pattern to acquire a token looks like this.
    result = None

    # First, looks up a token from cache
    # Since we are looking for token for the current app, NOT for an end user,
    # notice we give account parameter as None.
    result = app.acquire_token_silent(config["scope"], account=None)

    if not result:
        print("No suitable token exists in cache. Let's get a new one from AAD.\n")
        result = app.acquire_token_for_client(scopes=config["scope"])

    if "access_token" in result:
        return result['access_token']
    else:
        return None


access_token = get_access_token_graph()
print(access_token)


def get_headersURL(from_email: str = 'seattle.lab@shanwil.com') -> tuple[str, dict[str, str]]:
    """
    Get the headers and URL for the MS Graph application
    :return:
    """

    user_id = from_email
    url = f'https://graph.microsoft.com/v1.0/users/{user_id}/sendMail'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    return url, headers


def get_attachments_email(attachment_paths: list[str]) -> list[dict[str, Any]]:
    """
    Formats the attachment pdfs for the email
    :param attachment_paths:
    :return:
    """
    attachments_email = []
    if attachment_paths is not None:
        for path in attachment_paths:
            file_name = os.path.basename(path)
            file = path
            extension = file.split('.')[-1]
            # Read the PDF file and encode it to Base64
            with open(file, 'rb') as pdf_file:
                pdf_content = pdf_file.read()
                pdf_base64 = base64.b64encode(pdf_content).decode('utf-8')

            attachments_email.append({
                '@odata.type': '#microsoft.graph.fileAttachment',
                'name': file_name,
                'contentType': extension,
                'contentBytes': pdf_base64
            })

    return attachments_email

def read_template_without_bom(template_path: str) -> str:
    """
    Reads an HTML template file, removing any BOM if present.

    :param template_path: Path to the email template file (HTML format).
    :return: The template content as a string.
    """
    with open(template_path, 'rb') as file:
        raw_content = file.read()
        encoding = 'windows-1252'
        if raw_content.startswith(b'\xff\xfe'):
            raw_content = raw_content[2:]  # Remove BOM
            encoding = 'utf-16'
        return raw_content.decode(encoding)


def fill_email_template(template_path: str, variables: dict[str, str], table: pd.DataFrame = None, links: dict[str, str] = None) -> str:
    """
    Fills an email template with given variables, inserts an HTML table from a DataFrame,
    and adds hyperlinks.

    :param template_path: Path to the email template file (HTML format).
    :param variables: A dictionary of variables to replace in the template.
    :param table: The pandas DataFrame or HTML table string to insert into the template.
    :param links: A dictionary of placeholders and their corresponding URLs.
    :return: The filled email content as a string.
    """
    if isinstance(table, pd.DataFrame):
        # Convert the DataFrame to an HTML table string
        html_table = dataframe_to_html_with_style(table)
    else:
        html_table = table

    # Read the template content
    email_content = read_template_without_bom(template_path)

    print("Original template content (first 500 characters):")
    print(email_content[:500])  # Print the first 500 characters for brevity

    # Replace placeholders with variable content
    for key, value in variables.items():
        placeholder = f'##{key}##'
        if placeholder in email_content:
            print(f"Placeholder {placeholder} found in template.")
        else:
            print(f"Placeholder {placeholder} not found in template.")

        email_content = email_content.replace(placeholder, str(value))

    # Check if variables have been replaced
    for key in variables.keys():
        placeholder = f'##{key}##'
        if placeholder in email_content:
            print(f"Replacement for {placeholder} failed.")
        else:
            print(f"Replacement for {placeholder} succeeded.")

    # Replace the table placeholder with the HTML table string
    table_placeholder = '##table_placeholder##'
    if table_placeholder in email_content:
        print("Placeholder ##table_placeholder## found in template.")
    else:
        print("Placeholder ##table_placeholder## not found in template.")
    if html_table:
        email_content = email_content.replace(table_placeholder, html_table)

    # Check if the table placeholder has been replaced
    if table_placeholder in email_content:
        print("Replacement for ##table_placeholder## failed.")
    else:
        print("Replacement for ##table_placeholder## succeeded.")

    # Replace link placeholders with HTML hyperlinks
    if links:
        for key, url in links.items():
            encoded_url = urllib.parse.quote(url, safe=':/')
            link_placeholder = f'##{key}##'
            if link_placeholder in email_content:
                print(f"Placeholder {link_placeholder} found in template.")
            else:
                print(f"Placeholder {link_placeholder} not found in template.")

            # Create HTML hyperlink
            html_link = f'<a href="{encoded_url}">{url}</a>'
            email_content = email_content.replace(link_placeholder, html_link)

    # print("Final email content (first 500 characters):")
    # print(email_content[:500])  # Print the first 500 characters for brevity

    return email_content


def dataframe_to_html_with_style(df: pd.DataFrame) -> str:
    """
    Converts a pandas DataFrame into an HTML table string with styling to match the template.

    :param df: The pandas DataFrame to convert.
    :return: A string containing the DataFrame as an HTML table with inline styling.
    """
    # Start the table and add the header row
    html = ['<table class="MsoNormalTable" border="0" cellspacing="0" cellpadding="0" style="border-collapse:collapse;mso-yfti-tbllook:1184;mso-padding-alt:0in 0in 0in 0in;width:100%;">']

    # Header row
    html.append('<tr style="height:.2in;">')
    for col in df.columns:
        html.append(f'<td style="border:solid #156082 1.0pt;background:#156082;padding:.75pt .75pt .75pt .75pt;"><b><span style="font-family:\'Calibri\',sans-serif;color:white;">{col}</span></b></td>')
    html.append('</tr>')

    # Data rows
    for i, row in df.iterrows():
        html.append('<tr style="height:.2in;">')
        for val in row:
            html.append(f'<td style="border:solid #156082 1.0pt;border-top:none;padding:.75pt .75pt .75pt .75pt;"><span style="font-family:\'Calibri\',sans-serif;color:black;">{val}</span></td>')
        html.append('</tr>')

    # Close the table
    html.append('</table>')

    # Join all HTML parts and return
    return ''.join(html)


def send_email(toRecipients: list[str], emailBody: str = None, attachment_paths: list[str] = None, subject: str = 'Email Subject Line', ccRecipients: list[str] = None, replyTo: list[str] = None) -> None:
    """
    Send an email with the MS Graph application
    :return:
    """

    if emailBody is None:
        emailBody = """
            <!DOCTYPE html>
            <html>
                <head>
                </head>
                <body>
                    <p>
                    THIS IS THE TEXT OF THE EMAIL. IT CAN BE AS LONG AS YOU WANT.
                    IT CAN CONTAIN HTML TAGS.
                    </p>
                </body>
            </html>
        """
    to_list = [{'emailAddress': {'address': email}} for email in toRecipients]
    if attachment_paths:
        # List of file paths to attachment PDF's - currently only supports PDF's but can be modified.
        attachments = get_attachments_email(attachment_paths)
        email = {
            'message': {
                'subject': subject,
                'body': {
                    'contentType': 'HTML',
                    'content': emailBody
                },
                'toRecipients': to_list,
                'attachments': attachments
            }
        }
    else:
        # No attachments
        email = {
            'message': {
                'subject': subject,
                'body': {
                    'contentType': 'HTML',
                    'content': emailBody
                },
                'toRecipients': to_list
            }
        }

    if ccRecipients is not None:
        cc_list = [{'emailAddress': {'address': email}} for email in ccRecipients]
        email['message']['ccRecipients'] = cc_list

    if replyTo is not None:
        reply_list = [{'emailAddress': {'address': email}} for email in replyTo]
        email['message']['replyTo'] = reply_list

    url, headers = get_headersURL(from_email=FROM_EMAIL)
    response = requests.post(url, headers=headers, json=email)

    if response.status_code == 202:
        print('Email with attachments sent successfully.\n')
    else:
        print(
            f'Error sending email. Status code: {response.status_code}, '
            f'Response content: {response.text}\n')


if __name__ == '__main__':
    """
    send_email() is the main function and automatically handles get_access_token_graph(), get_headersURL(), and get_attachments_email()
    """
    send_email()


def find_user_email_by_name(name: str) -> str:
    """
    Searches for a user by first and last name and retrieves their email address.

    :param access_token: A valid access token for Microsoft Graph API.
    :param first_name: The user's first name.
    :param last_name: The user's last name.
    :return: The email address of the user, or None if not found or if multiple matches are found.
    """
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

    name_parts = name.split(' ')
    first_name = name_parts[0]
    last_name = name_parts[-1]

    # This query assumes the displayName format is "First Last". Adjust the format as needed.
    query_filter = f"startswith(displayName,'{first_name} {last_name}') or startswith(givenName, '{first_name}') and endsswith(surname, '{last_name}')"
    url = f'https://graph.microsoft.com/v1.0/users?$filter={query_filter}'

    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        users_data = response.json()
        users = users_data.get('value', [])
        if len(users) == 1:  # Assuming only one user matches the criteria
            email_address = users[0].get('mail', None)  # 'userPrincipalName' can also be used
            return email_address
        else:
            # print(f"Error: Found {len(users)} users matching the criteria. Expected exactly 1. Users:\n{users}")
            # return None
            return users
    else:
        # print(f"Error searching for user. Status code: {response.status_code}, Response content: {response.text}")
        return None


def find_user_email_by_employee_id(employee_id: str) -> str:
    """
    Finds a user's email address by their employee ID using Microsoft Graph API.

    :param access_token: A valid access token for Microsoft Graph API.
    :param employee_id: The employee ID of the user.
    :return: The email address of the user, or None if not found.
    """
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    # Filter query for the employeeId
    query_filter = f"employeeId eq '{employee_id}'"
    url = f'https://graph.microsoft.com/v1.0/users?$filter={query_filter}'

    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        users_data = response.json()
        users = users_data.get('value', [])
        if len(users) == 1:  # Assuming only one user matches the criteria
            email_address = users[0].get('mail', None)
            return email_address
        else:
            print(f"Error: Found {len(users)} users matching the criteria. Expected exactly 1.")
            return None
    else:
        print(f"Error searching for user by employee ID. Status code: {response.status_code}, Response content: {response.text}")
        return None


def find_user_firstname_by_employee_id(employee_id: str) -> str:
    """
    Finds a user's first name by their employee ID using Microsoft Graph API.

    :param access_token: A valid access token for Microsoft Graph API.
    :param employee_id: The employee ID of the user.
    :return: The email address of the user, or None if not found.
    """
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    # Filter query for the employeeId
    query_filter = f"employeeId eq '{employee_id}'"
    url = f'https://graph.microsoft.com/v1.0/users?$filter={query_filter}'

    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        users_data = response.json()
        users = users_data.get('value', [])
        if len(users) == 1:  # Assuming only one user matches the criteria
            first_name = users[0].get('givenName', None)
            return first_name
        else:
            # print(f"Error: Found {len(users)} users matching the criteria. Expected exactly 1.")
            return None
    else:
        # print(f"Error searching for user by employee ID. Status code: {response.status_code}, Response content: {response.text}")
        return None


def pull_contact_by_employee_id(employee_id: str) -> str:
    """
    Finds a user's name and email address by their employee ID using Microsoft Graph API.

    :param access_token: A valid access token for Microsoft Graph API.
    :param employee_id: The employee ID of the user.
    :return: tuple of first name, last name, and email address of the user, or None if not found.
    """
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    # Filter query for the employeeId
    query_filter = f"employeeId eq '{employee_id}'"
    url = f'https://graph.microsoft.com/v1.0/users?$filter={query_filter}'

    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        users_data = response.json()
        users = users_data.get('value', [])
        if len(users) == 1:  # Assuming only one user matches the criteria
            first_name = users[0].get('givenName', None)
            last_name = users[0].get('surname', None)
            email_address = users[0].get('mail', None)
            return {'first_name': first_name, 'last_name': last_name, 'email': email_address}
        elif len(users) > 1:
            users_string = ', '.join([f'{user["givenName"]} {user["surname"]}' for user in users])
            # print(f"Error: Found {len(users)} users matching the criteria. Expected exactly 1.")
            return {'first_name': f'MULTIPLE RECORDS for {employee_id}', 'last_name': f'[{users_string}]', 'email': 'oth@shanwil.com'}
        elif len(users) == 0:
            return {'first_name': f'NO GRAPH RECORD for {employee_id}', 'last_name': f'NO GRAPH RECORD for {employee_id}', 'email': 'oth@shanwil.com'}
    else:
        # print(f"Error searching for user by employee ID. Status code: {response.status_code}, Response content: {response.text}")
        return {'first_name': f'Error searching for user by {employee_id}', 'last_name': f'Status code: {response.status_code}', 'email': 'oth@shanwil.com'}

# Send Emails with Microsoft Graph

This Python script allows you to send emails through a Microsoft Exchange account using Microsoft Graph API. It supports sending emails with attachments, filling email templates with dynamic content, and more.

## Prerequisites

- Python 3.x
- `requests` library
- `msal` library
- `pandas` library

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/send-emails.git
   cd send-emails
   ```

2. Set up a virtual environment and activate it:
   ```bash
   conda create --name email-env python=3.8
   conda activate email-env
   ```

3. Install the required libraries:
   ```bash
   pip install requests msal pandas
   ```

4. Create a `.env` file in the project root directory with your Microsoft Graph API credentials:
   ```bash
   CLIENT_ID=your-client-id
   TENANT_ID=your-tenant-id
   SECRET_VALUE=your-secret-value
   FROM_EMAIL=your-from-email@example.com
   ```

6. Ensure the `.env` file is included in your `.gitignore` to avoid exposing your credentials.

## Usage

### Sending an Email

To send an email, use the `send_email` function. Here’s an example:

```python
from Send_Emails import send_email

to_list = ['recipient@example.com']
cc_list = ['cc@example.com']
subject = 'Test Email'
body = '<h1>Hello, World!</h1>'
attachment_paths = ['path/to/attachment1.pdf', 'path/to/attachment2.pdf']
reply_to = ['replyto@example.com']

send_email(to_list, emailBody=body, attachment_paths=attachment_paths, subject=subject, ccRecipients=cc_list, replyTo=reply_to)
```

### Filling an Email Template

You can fill an email template with dynamic content using the `fill_email_template` function. Here’s an example:

```python
from Send_Emails import fill_email_template, dataframe_to_html_with_style
import pandas as pd

# Variables to replace in the template
email_vars = {'invoice_month': 'May 2024', 'due_string': 'Monday (6/3/2024)'}

# Create a DataFrame and convert it to an HTML table
data = {'Column1': [1, 2, 3], 'Column2': ['A', 'B', 'C']}
df = pd.DataFrame(data)
html_table = dataframe_to_html_with_style(df)

# Fill the template
template_path = 'path/to/template.html' # Save your email from Outlook as an HTML file!
email_body = fill_email_template(template_path, email_vars, html_table)

# Send the email
send_email(['recipient@example.com'], emailBody=email_body, subject='Test Email')
```

### Finding a User by Employee ID

To find a user's email address by their employee ID, use the `find_user_email_by_employee_id` function. Here’s an example:

```python
from Send_Emails import find_user_email_by_employee_id

employee_id = '12345'
email = find_user_email_by_employee_id(employee_id)
print(email)
```

## Functions

```python
load_env_file(env_file: str) -> None
```
Loads Microsoft Graph API credentials as environment variables from a specified file (azure.env).

```python
get_access_token_graph() -> str or None
```
Gets the access token for the MS Graph application.

```python
get_headersURL(from_email: str = 'seattle.lab@shanwil.com') -> tuple[str, dict[str, str]]
```
Gets the headers and URL for the MS Graph application.

```python
get_attachments_email(attachment_paths: list[str]) -> list[dict[str, Any]]
```
Formats the attachment PDFs for the email.

```python
read_template_without_bom(template_path: str) -> str
```
Reads an HTML template file, removing any BOM if present.

```python
fill_email_template(template_path: str, variables: dict[str, str], table: pd.DataFrame = None, links: dict[str, str] = None) -> str
```
Fills an email template with given variables, inserts an HTML table from a DataFrame, and adds hyperlinks.

```python
dataframe_to_html_with_style(df: pd.DataFrame) -> str
```
Converts a pandas DataFrame into an HTML table string with styling to match the template.

```python
send_email(toRecipients: list[str], emailBody: str = None, attachment_paths: list[str] = None, subject: str = 'Email Subject Line', ccRecipients: list[str] = None, replyTo: list[str] = None) -> None
```

Sends an email with the MS Graph application.

```python
find_user_email_by_name(name: str) -> str
```

Searches for a user by first and last name and retrieves their email address.

```python
find_user_email_by_employee_id(employee_id: str) -> str
```
Finds a user's email address by their employee ID using Microsoft Graph API.

```python
find_user_firstname_by_employee_id(employee_id: str) -> str
```
Finds a user's first name by their employee ID using Microsoft Graph API.

```python
pull_contact_by_employee_id(employee_id: str) -> dict[str, str]
```
Finds a user's name and email address by their employee ID using Microsoft Graph API.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
"""

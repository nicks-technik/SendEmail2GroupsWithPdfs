# pip install --upgrade python-dotenv pandas google-api-python-client google-auth-httplib2 google-auth-oauthlib

# Import the required modules

# The os module is used to interact with the operating system
import os

# The Create_Service function is used to create a connection to the Google API
# The Create_Service function takes the path to the client secret file, the name of the API, the version of the API, and the scopes of the API as arguments
# The Create_Service function returns a service object that is used to interact with the Google API
from Google import Create_Service

# The base64 module is used to encode and decode data
# The email module is used to create email messages
# The mimetypes module is used to determine the type of a file
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import mimetypes

# The pandas module is used to read and write CSV files
import pandas as pd
from dotenv import load_dotenv

file_attachments = []

# CLIENT_SECRET_FILE = "240527_client_secret_thorstenh_sendEmail.json"
API_NAME = "gmail"
API_VERSION = "v1"
SCOPES = ["https://mail.google.com/"]


def loadEnvVars():
    """
    Loads the environment variables from the .env file

    This function is called once at the beginning of the program.
    It loads the environment variables from the .env file and assigns
    them to the global variables defined above the function.

    Parameters
    ----------
    None

    Returns
    -------
    None
    """
    # Load environment variables from .env file
    load_dotenv()

    # Access environment variables using os.getenv('VARIABLE_NAME')
    global CLIENT_SECRET_FILE
    global CSV_FILE
    global PDF_SUBFOLDER_PATH
    global MAIL_SUBJECT
    global HTML_BODY_FILE
    global html_body

    # Load the Client Secret File
    CLIENT_SECRET_FILE = os.getenv("ENV_SECURITY_JSON")
    print(f"Client Secret File: {CLIENT_SECRET_FILE}")

    # Load the CSV Group File
    CSV_FILE = os.getenv("ENV_CSV_FILE")
    print(f"CSV Group File: {CSV_FILE}")

    # Load the PDF Subfolder Path
    PDF_SUBFOLDER_PATH = os.getenv("ENV_PDF_SUBFOLDER_PATH")
    print(f"PDF Subfolder Path: {PDF_SUBFOLDER_PATH}")

    # Load the HTML Body File
    MAIL_SUBJECT = os.getenv("ENV_MAIL_SUBJECT")
    print(f"Mail Subject: {MAIL_SUBJECT}")

    # Load the HTML Body File
    HTML_BODY_FILE = os.getenv("ENV_MAIL_TEXT_FILE")
    print(f"HTML Body File: {HTML_BODY_FILE}")

    # Load the HTML Body
    html_body = loadHTMLBodyFile()
    print(html_body)

    print()


def loadHTMLBodyFile():
    """
    Loads the HTML body file.

    This function reads the HTML body file specified by the
    HTML_BODY_FILE environment variable and returns the
    contents of the file.

    If the file is not found, it prints an error message and
    exits with an error code.

    Parameters
    ----------
    None

    Returns
    -------
    str
        The contents of the HTML body file.

    Raises
    ------
    FileNotFoundError
        If the HTML body file is not found.
    """
    try:
        with open(HTML_BODY_FILE, "r") as f:
            return f.read()
    except FileNotFoundError:
        print(f"Error: HTML body file '{html_body_file_path}' not found.")
        # html_body = "<p>An error occurred while loading the HTML body.</p>"
        exit(1)


def loadGroupsCsv():
    """
    Loads the CSV file and prints the first few rows.
    df contains than the CSV file as a pandas DataFrame

    Args:
        None

    Returns:
        None

    Raises:
        FileNotFoundError: If the CSV file is not found
    """
    global df

    try:
        # Read the CSV file using pandas.read_csv
        # df = pd.read_csv(PDF_SUBFOLDER_PATH + "/" + CSV_FILE, delimiter=";")
        df = pd.read_csv(CSV_FILE, delimiter=";")

        # Print the first few rows of the DataFrame (optional)
        print(df.head())

    except FileNotFoundError:
        print(f"Error: CSV file '{CSV_FILE}' not found.")


def listFiles():
    """
    Lists the names of files in the given subfolder

    Parameters
    ----------
    None

    Returns
    -------
    None

    Notes
    -----
    This function lists the names of files in the given subfolder.
    It does not only list the files but also sends an email with the files
    that belong to the same group.
    """
    global file_attachments
    global PDF_SUBFOLDER_PATH

    file_attachments = []
    last_one = ""

    # Get all files in the subfolder
    files = os.listdir(PDF_SUBFOLDER_PATH)

    # Sort the list of files (ascending order)
    files.sort()
    print(files)
    print()

    # Print each sorted file
    for filename in files:
        print(f"filename: {filename}")
        # Construct the full path
        # full_path = os.path.join(PDF_SUBFOLDER_PATH, filename)

        if not filename.endswith(".pdf"):
            continue

        group_name = " ".join(filename.split(" ")[:2])
        print(f"group_name: {group_name}, last_one: {last_one}")

        # Check if the group name has changed or first run
        if (group_name != last_one) and (last_one != ""):
            print(f"file_attachmentsif: {file_attachments}")
            sendEmail(last_one)
            last_one = group_name
            file_attachments = []
            file_attachments.append(os.path.join(PDF_SUBFOLDER_PATH, filename))
        else:
            last_one = group_name
            file_attachments.append(os.path.join(PDF_SUBFOLDER_PATH, filename))

    if len(file_attachments) > 0:
        print(f"file_attachmentslen: {file_attachments}")
        sendEmail(group_name)


def getEmailFromGroup(groupname):
    """
    Gets the email address for the given group name from the csv file

    Args:
        groupname (str): The name of the group to search for

    Returns:
        str: The email address associated with the group, or "Not found" if not found
    """
    try:
        # .loc creates a label-based selection by index
        # df["Name"] == groupname creates a boolean mask
        # Select the Email column and choose the first value (should be the only one)
        email = df.loc[df["Name"] == groupname, "Email"].values[0]
        print(f"Email for '{groupname}': {email}")
        return email
    except IndexError:
        print(f"Name '{groupname}' not found in the DataFrame.")
        email = "Not found"
        return email


def sendEmail(groupname):
    """
    Sends an email with the files in the given subfolder

    Args:
        groupname (str): The name of the group to search for

    Returns:
        None
    """
    # Get the email address associated with the group from the CSV file
    emailaddress = getEmailFromGroup(groupname)

    # Create a mime message which is a container for the email
    mimeMessage = MIMEMultipart()

    # Set the recipient of the email
    mimeMessage["to"] = emailaddress
    # mimeMessage["to"] = "email1@EmailDomain.com; email2@EmailDomain.com"

    # Set the subject of the email
    mimeMessage["subject"] = MAIL_SUBJECT

    # Create the HTML body of the email
    # html_body = """
    # <p style="font-family: Arial, sans-serif; font-size: 16px;">Hi [Group Name],</p>
    # <p>Here are the files attached to this email.</p>
    # <p><i>Please note: This email is automatically generated.</i></p>
    # """
    # Replace the placeholder with the group name
    # html_body = html_body.replace("[Group Name]", group_name)  # Replace placeholder

    # Attach the HTML body
    mimeMessage.attach(MIMEText(html_body, "html"))

    # Print the attachments
    print(f"file_attachments: {file_attachments}")

    # Attach the files
    for attachment in file_attachments:
        print(f"attachment: {attachment}")
        # Get the content type of the file
        content_type, encoding = mimetypes.guess_type(attachment)
        # Split the content type into main and sub type
        main_type, sub_type = content_type.split("/", 1)
        # Get the file name
        file_name = os.path.basename(attachment)

        # Open the file in binary mode
        f = open(attachment, "rb")

        # Create a mime base object
        attachment_converted = MIMEBase(main_type, sub_type)
        # Set the payload of the mime base object
        attachment_converted.set_payload(f.read())
        # Add the content disposition header
        attachment_converted.add_header(
            "Content-Disposition", "attachment", filename=file_name
        )
        # Encode the attachment
        encoders.encode_base64(attachment_converted)

        # Close the file
        f.close()

        # Attach the mime base object to the mime message
        mimeMessage.attach(attachment_converted)

    # Convert the mime message to a safe base64 string
    raw_string = base64.urlsafe_b64encode(mimeMessage.as_bytes()).decode()

    try:
        # Send the email
        message = (
            service.users()
            .messages()
            .send(userId="me", body={"raw": raw_string})
            .execute()
        )
        print(message)
        print("Email sent successfully (initial processing).")

        # Move the sent files to the old subfolder
        moveFile2Old()

    except Exception as e:
        print(f"Error sending email: {e}")


def moveFile2Old():
    """
    Moves the sent files to the old subfolder

    Parameters
    ----------
    None

    Returns
    -------
    None
    """
    # Move the sent files to the old subfolder
    for attachment in file_attachments:

        file_name = os.path.basename(attachment)
        print(f"file_namemoveFile2Old: {file_name}")

        # Get the directory path
        directory_path = os.path.dirname(attachment)
        print(f"directory_pathmoveFile2Old: {directory_path}")

        # Rename the file by moving it to the old subfolder
        os.rename(attachment, f"{PDF_SUBFOLDER_PATH}/old/{file_name}")


# Load the environment variables from .env file
loadEnvVars()

# Create the Gmail API service, inclusive the pickle token file
service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)

# Loads the CSV file and prints the first few rows
loadGroupsCsv()

# Lists the names of files and action on them in the given subfolder
listFiles()
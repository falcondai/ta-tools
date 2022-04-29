import os.path
from email.message import EmailMessage
import base64
import mimetypes
import glob
import csv

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.compose']


def create_message(sender, to, subject, body, attachment_path, attachment_name):
    """Create a message for an email.

    Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    body: The text of the email message.

    Returns:
    An object containing a base64url encoded email object.
    """
    message = EmailMessage()
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject
    message.set_content(body)

    if attachment_path is not None:
        # Add attachment.
        ctype, encoding = mimetypes.guess_type(attachment_path)
        maintype, subtype = ctype.split('/')
        with open(attachment_path, 'rb') as fp:
            message.add_attachment(
                fp.read(),
                maintype=maintype,
                subtype=subtype,
                filename=attachment_name,
                )

    return {
        'raw': base64.urlsafe_b64encode(message.as_bytes()).decode('ascii'),
        }


def create_draft(service, user_id, message_body):
    """Create and insert a draft email. Print the returned draft's message and id.

    Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    message_body: The body of the email message, including headers.

    Returns:
    Draft object, including draft id and message meta data.
    """
    try:
        draft = service.users().drafts().create(userId=user_id, body={'message': message_body}).execute()
        return draft
    except HttpError as error:
        print('An error occurred: %s' % error)
    return None


def send(service, user_id, message):
    """Send an email

    Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    message: The body of the email message, including headers.

    Returns:
    Message object, including message id and message meta data.
    """
    try:
        mail = service.users().messages().send(userId=user_id, body=message).execute()
        return mail
    except HttpError as error:
        print('An error occurred: %s' % error)
    return None


if __name__ == '__main__':
    # Parameters
    sheet_path = 'students.tsv'
    pdf_dir = r'Homework 2'
    email_subject = 'Graded homework 2 from TTIC 31250'
    solution_url = 'https://home.ttic.edu/~avrim/MLT22/soln2.pdf'
    attachment_name_fmt = '%s - HW2 graded.pdf'

    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        # Call the Gmail API
        service = build('gmail', 'v1', credentials=creds)

        # Match the PDFs and the students.
        pdf_paths = glob.glob(os.path.join(pdf_dir, '*.pdf'))
        name_pdfs = {}
        for pdf_path in pdf_paths:
            filename = os.path.basename(pdf_path)
            file_sn = filename.split(' - ')[0].strip()
            name_pdfs[file_sn] = pdf_path

        i = 0
        with open(sheet_path, 'r') as sheet_fp:
            reader = csv.reader(sheet_fp, delimiter='\t')
            for row in reader:
                # Tab-separated.
                student_email, student_name = row

                if student_name in name_pdfs:
                    pdf_path = name_pdfs[student_name]
                    attachment_name = attachment_name_fmt % student_name
                    msg = create_message(
                        sender='dai@ttic.edu',
                        to=student_email,
                        subject=email_subject,
                        body='Please see the attached annotated PDF. The solution is posted at %s.\n\nCheers,\nFalcon Dai''' % solution_url,
                        attachment_path=pdf_path,
                        attachment_name=attachment_name,
                        )
                    i += 1
                    print(i, 'Sent', student_name, student_email, attachment_name, pdf_path)
                    # create_draft(service, 'me', msg)
                    send(service, 'me', msg)
                else:
                    print('No PDF named %s!' % student_name)

    except HttpError as error:
        print(f'An error occurred: {error}')

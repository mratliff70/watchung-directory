"""
Generates PDF membership directory from data in Google Sheet
Includes member photos stored in Google Drive
"""
from __future__ import print_function
import os
import io
import datetime
# Import Google libraries
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
# Import PDF libraries
import pdfkit
from PyPDF2 import PdfFileReader, PdfFileWriter, PdfFileMerger
# Import AWS library
import boto3


#### Global variables

sheets_service = None
drive_service = None

s3resource = boto3.resource('s3')

s3bucketname = 'watchungairstream'
credentialsfile = 'credentials.json'
tokenfile = 'token.json'
photoplaceholderfile = 'photoplaceholder.png'

# Get ID of Google Sheet from environment variable
GOOGLE_SHEET_ID = os.getenv('GOOGLE_SHEET_ID')
# The range of cells containing member data
GOOGLE_SHEET_RANGE = os.getenv('GOOGLE_SHEET_RANGE')

# Get password for PDF
PDF_PASSWORD = os.getenv('PDF_PASSWORD')


def setGoogleService():

    global sheets_service
    global drive_service

    # Get Google credentials and token from S3 bucket
    #TODO:  We should really be getting these from AWS Secrets Vault
    s3resource.meta.client.download_file(s3bucketname, credentialsfile, credentialsfile)
    s3resource.meta.client.download_file(s3bucketname,tokenfile, tokenfile)

    # Setup the Sheets API
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly', 'https://www.googleapis.com/auth/drive.readonly']

    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(tokenfile):
        creds = Credentials.from_authorized_user_file(tokenfile, SCOPES)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                credentialsfile, SCOPES)
            creds = flow.run_console()
        # Save the credentials for the next run
        with open(tokenfile, 'w') as token:
            token.write(creds.to_json())

    sheets_service = build('sheets', 'v4', credentials=creds)

    drive_service = build('drive', 'v3', credentials=creds)



def makeHTML():
    """
    This function extracts data from the Google Sheet and formats it in HTML
    :return:
    """

    # Get the data
    result = sheets_service.spreadsheets().values().get(spreadsheetId=GOOGLE_SHEET_ID, range=GOOGLE_SHEET_RANGE).execute()

    #TODO: If no data was found, we should probably raise an exception!!
    values = result.get('values', [])
    if not values:
        print('No data found.')

    else:
        # Download photoplaceholder
        # 
        s3resource.meta.client.download_file(s3bucketname, photoplaceholderfile, photoplaceholderfile)

        outfile = open('./directory.html', 'w')
        outfile.write('<html><body>\n')

        # Set the font size
        outfile.write('<style> li { font-size: 20px; } </style>\n')

        # Start the list
        outfile.write('<p><ul>\n')

        for row in values:

            # If there is a value and value is 'Yes"
            if len(row) > 21 and row[21] and row[21] == 'Yes':

                # Print the names
                # If the second member's first name is present without a last name, then we assume the last name is the same
                if row[7] and not row[6]:
                    outfile.write('<li><b> %s & %s %s' % (row[1], row[7], row[0]))
                # If the second member's first and last name are present
                elif row[7] and row[6]:
                    outfile.write('<li><b> %s %s & %s %s' % (row[1], row[0], row[7], row[6]))
                # Otherwise, just print the first person's name
                else:
                    outfile.write('<li><b> %s %s</b>' % (row[1], row[0]))

                # Print WBCCI number on same line as names
                if row[19]:
                    outfile.write(' <font color="red">%s</font>\n' % row[19])

                outfile.write('</b><br>\n')

                # Print the address
                outfile.write('%s<br>\n%s, %s %s<br>\n' % (row[11], row[12], row[13], row[14]))

                # Print 1st person's info name: phone number, e-mail
                outfile.write('%s: %s \n' % (row[1], row[3]))

                # If the 1st person has an e-mail address
                if row[2]:
                    outfile.write('<a href="mailto:%s">%s</a>' % (row[2], row[2]))

                outfile.write('<br>\n')

                # Print 2nd person's info if there is any
                if row[7] and (row[8] or row[9]):
                    outfile.write('%s: \n' % row[7])

                    # If the 2nd person has a phone number
                    if row[9]:
                        outfile.write('%s ' % row[9])

                    # If the 2nd person has an e-mail address
                    if row[8]:
                        outfile.write('<a href="mailto:%s">%s</a>\n' % (row[8], row[8]))

                    outfile.write('<br>\n')

                # Insert picture if there is one
                if len(row) > 22 and row[22]:

                    request = drive_service.files().get_media(fileId=row[22])
                    local_file = "/data/" + row[22]
                    fh = io.FileIO(local_file, mode='wb')
                    downloader = MediaIoBaseDownload(fh, request)
                    done = False
                    while done is False:
                        status, done = downloader.next_chunk()

                    # Insert image into HTML
                    outfile.write('<br><br><img src="/data/' + row[22] + '"/>')

                # If there is no picture available, insert placeholder image
                else:

                    outfile.write('<br><br><img src="/data/logo.png"')

                outfile.write('<br><br><br><br><br><br></li>')

        # Close the list
        outfile.write('</ul></p>')

        # Print the current date and time
        timenow = datetime.datetime.now()
        outfile.write('<p>This directory was created on: %s</p>' % timenow)

        # close the page
        outfile.write('</body></html>')
        outfile.close()


def makePDF():
    """
    This function converts the HTML file into a PDF
    :return:
    """

    # Allow pdfkit to access member photos downloaded locally
    pdfkitoptions = {
        'enable-local-file-access': '',
    }

    #NOTE: Investigated using the "cover" option https://pypi.org/project/pdfkit/ but got error loading cover page.  Gave up.

    pdfkit.from_file('./directory.html', './directory.pdf', options=pdfkitoptions)

def addCoverpage():
    """
    Add coverpage to PDF
    :return:
    """

    # If GOOGLE_SHEET_RANGE begins with Members, then this is the current members directory
    if GOOGLE_SHEET_RANGE.startswith("Members"):
        coverpagefile = 'CoverPageMembers.pdf'
    # Otherwise this must be the past members directory
    else:
        coverpagefile = 'CoverPagePastMembers.pdf'

    pdf_merger = PdfFileMerger()

    s3resource.meta.client.download_file(s3bucketname, coverpagefile, coverpagefile)

    pdf_merger.append(coverpagefile)
    pdf_merger.append('./directory.pdf')

    with open('./directorywcover.pdf', 'wb') as fileobj:
        pdf_merger.write(fileobj)


def protectPDF():
    """
    This function encrypts the PDF and adds password protection
    :return:
    """

    in_file = open("./directorywcover.pdf", "rb")
    input_pdf = PdfFileReader(in_file)

    numpages = input_pdf.getNumPages()
    print ("Number of pages in PDF = %s" % numpages)

    output_pdf = PdfFileWriter()
    output_pdf.appendPagesFromReader(input_pdf)
    output_pdf.encrypt(PDF_PASSWORD)

    out_file = open("./secure_directory.pdf", "wb")

    output_pdf.write(out_file)

    out_file.close()
    in_file.close()

def uploadToS3():
    """
    This function uploads the PDF to S3
    :return:
    """

    sourcefile = './secure_directory.pdf'

    # If GOOGLE_SHEET_RANGE begins with Members, then this is the current members directory
    if GOOGLE_SHEET_RANGE.startswith("Members"):
        targetfile = 'WatchungMemberDirectory.pdf'
    # Otherwise this must be the past members directory
    else:
        targetfile = 'WatchungPastMemberDirectory.pdf'

    # Upload the PDF
    s3resource.meta.client.upload_file(sourcefile, s3bucketname, targetfile, ExtraArgs={'ACL':'public-read'})

    # Upload the token and credentials files
    s3resource.meta.client.upload_file(tokenfile, s3bucketname, tokenfile)
    s3resource.meta.client.upload_file(credentialsfile, s3bucketname, credentialsfile)



#######
# Main Code
#######

if __name__ == '__main__':

    setGoogleService()

    makeHTML()

    makePDF()

    addCoverpage()

    protectPDF()

    uploadToS3()

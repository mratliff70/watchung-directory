"""
Shows basic usage of the Sheets API. Prints values from a Google Spreadsheet.
"""
from __future__ import print_function
from apiclient.discovery import build
from httplib2 import Http
from oauth2client import file as oauth_file, client, tools
import pdfkit
from PyPDF2 import PdfFileReader, PdfFileWriter
import boto3
import os

# Setup the Sheets API
SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'
store = oauth_file.Storage('token.json')
creds = store.get()
if not creds or creds.invalid:
    flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
    creds = tools.run_flow(flow, store)
service = build('sheets', 'v4', http=creds.authorize(Http()))

# Get ID of Google Sheet from environment variable
GOOGLE_SHEET_ID = os.getenv('GOOGLE_SHEET_ID')

# Get password for PDF
PDF_PASSWORD = os.getenv('PDF_PASSWORD')

def makeHTML():
    """
    This function extracts data from the Google Sheet and formats it in HTML
    :return:
    """

    # The range of cells containing data
    RANGE_NAME = 'Members!A2:T100'

    # Get the data
    result = service.spreadsheets().values().get(spreadsheetId=GOOGLE_SHEET_ID,
                                                 range=RANGE_NAME).execute()

    #TODO: If no data was found, we should probably raise an exception!!
    values = result.get('values', [])
    if not values:
        print('No data found.')

    else:
        outfile = open('./directory.html', 'w')
        outfile.write('<html><body><ul>')

        #print('Firstname, Lastname:')
        for row in values:

            # Print the names
            # If the second member's first name is present without a last name, then we assume the last name is the same
            if row[7] and not row[6]:
                outfile.write('<li> %s & %s %s<br>\n' % (row[1], row[7], row[0]))
            # If the second member's first and last name are present
            elif row[7] and row[6]:
                outfile.write('<li> %s %s & %s %s<br>\n' % (row[1], row[0], row[7], row[6]))
            # Otherwise, just print the first person's name
            else:
                outfile.write('<li> %s %s<br>\n' % (row[1], row[0]))

            # Print the address
            outfile.write('%s<br>\n%s<br>\n%s<br>\n' % (row[11], row[12], row[13]))


            outfile.write('<br></li>')


        outfile.write('</ul></body></html>')
        outfile.close()


def makePDF():
    """
    This function converts the HTML file into a PDF
    :return:
    """

    pdfkit.from_file('./directory.html', './directory.pdf')


def protectPDF():
    """
    This function encrypts the PDF and adds password protection
    :return:
    """

    in_file = open("./directory.pdf", "rb")
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

    s3resource = boto3.resource('s3')

    sourcefile = './secure_directory.pdf'
    s3bucketname = 'watchungairstream'
    targetfile = 'WatchungMemberDirectory.pdf'

    s3resource.meta.client.upload_file(sourcefile, s3bucketname, targetfile, ExtraArgs={'ACL':'public-read'})


#######
# Main Code
#######

if __name__ == '__main__':

    makeHTML()

    makePDF()

    protectPDF()

    uploadToS3()
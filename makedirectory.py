"""
Generates PDF membership directory from data in Google Sheet
"""
from __future__ import print_function
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file as oauth_file, client, tools
import pdfkit
from PyPDF2 import PdfFileReader, PdfFileWriter, PdfFileMerger
import boto3
import os


#### Global variables

gservice = None

s3resource = boto3.resource('s3')

s3bucketname = 'watchungairstream'
credentialsfile = 'credentials.json'
tokenfile = 'token.json'
coverpagefile = '2018CoverPage.pdf'

# Get ID of Google Sheet from environment variable
GOOGLE_SHEET_ID = os.getenv('GOOGLE_SHEET_ID')

# Get password for PDF
PDF_PASSWORD = os.getenv('PDF_PASSWORD')


def setGoogleService():

    global gservice

    # Get Google credentials and token from S3 bucket

    s3resource.meta.client.download_file(s3bucketname, credentialsfile, credentialsfile)
    s3resource.meta.client.download_file(s3bucketname,tokenfile, tokenfile)

    # Setup the Sheets API
    SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'

    store = oauth_file.Storage(tokenfile)
    creds = store.get()

    # If credentials could not be loaded from token file
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets(credentialsfile, SCOPES)
        creds = tools.run_flow(flow, store)

    gservice = build('sheets', 'v4', http=creds.authorize(Http()))



def makeHTML():
    """
    This function extracts data from the Google Sheet and formats it in HTML
    :return:
    """

    # The range of cells containing member data
    RANGE_NAME = 'Members!A2:V100'


    # Get the data
    result = gservice.spreadsheets().values().get(spreadsheetId=GOOGLE_SHEET_ID,
                                                 range=RANGE_NAME).execute()

    #TODO: If no data was found, we should probably raise an exception!!
    values = result.get('values', [])
    if not values:
        print('No data found.')

    else:
        outfile = open('./directory.html', 'w')
        outfile.write('<html><body>\n')

        # Set the font size
        outfile.write('<style> li { font-size: 20px; } </style>\n')

        # Start the list
        outfile.write('<ul>\n')

        for row in values:

            if row[21] == 'Yes':

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


                outfile.write('<br></li>')


        outfile.write('</ul></body></html>')
        outfile.close()


def makePDF():
    """
    This function converts the HTML file into a PDF
    :return:
    """

    pdfkit.from_file('./directory.html', './directory.pdf')

def addCoverpage():
    """
    Add coverpage to PDF
    :return:
    """
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

    targetfile = 'WatchungMemberDirectory.pdf'

    # Upload the PDF
    s3resource.meta.client.upload_file(sourcefile, s3bucketname, targetfile, ExtraArgs={'ACL':'public-read'})

    # Upload the
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

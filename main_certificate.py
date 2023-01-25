
import os

os.system("pip install -r requirements.txt")

from openpyxl import Workbook, load_workbook

from certificate import *
from docx import Document
import csv
from docx2pdf import convert


mailerpath = "Data/Mail.xlsm"

# create output folder if not exist
try:
    os.makedirs("Output/Doc")
    os.makedirs("Output/PDF")
except OSError:
    pass


def get_participants(f):
    data = [] # create empty list
    with open(f, mode="r", encoding='utf-8') as file:
        csv_reader = csv.DictReader(file)
        for row in csv_reader:
            data.append(row) # append all results
    return data

offset = 2

def updatemailer(row, workbook, sheet, email, filepath, sub, body, status, cc=""):
    sheet.cell(row=row, column=1).value = email
    sheet.cell(row=row, column=2).value = cc
    sheet.cell(row=row, column=3).value = sub
    sheet.cell(row=row, column=4).value = body
    sheet.cell(row=row, column=5).value = filepath
    sheet.cell(row=row, column=6).value = status
    workbook.save(filename = mailerpath)

def getworkbook(filename):
    wb = load_workbook(filename=filename, read_only=False, keep_vba=True)
    sheet = wb.active
    return wb, sheet

def create_docx_files(filename, list_participate):

    wb, sheet = getworkbook(mailerpath)

    gdsc = "s"
    lead = gdsc
    facilitator = gdsc
    event = gdsc

    for index, participate in enumerate(list_participate):
        # use original file everytime
        doc = Document(filename)

        if participate["Pathway Completion Status"] == "No":
            continue

        name = participate["Student Name"]
        email = participate["Student Email"]

        sub = f"[{event}] Certificate of Completion"

        body = f''' Dear {name},

    Thank you for participating in {event}. 
    Your participation is valuable to us.
    Please find the attached certificate of completion.

Thanks and Regards,
{facilitator} 
{event} Facilitator,
and
{lead}
GDSC TMSL Lead'''

        replace_participant_name(doc, name)
        replace_gdsc_name(doc, gdsc)
        replace_lead_name(doc, lead)
        replace_facilitator_name(doc, facilitator)
        replace_event(doc, event)

        doc.save('Output/Doc/{}.docx'.format(name))

        # ! if your program working slowly, comment this two line and open other 2 line.
        print("Output/{}.pdf Creating".format(name))
        convert('Output/Doc/{}.docx'.format(name), 'Output/Pdf/{}.pdf'.format(name))

        filepath = os.path.abspath('Output/Pdf/{}.pdf'.format(name))

        updatemailer(row=(offset+index), workbook=wb,  sheet=sheet, email=email, filepath=filepath, sub=sub, body=body, status="Send")

    
# get certificate temple path
certificate_file = "Data/Event Certificate Template.docx"
# get participants path
participate_file = "Data/"+("ParticipantList.csv" if (input("Test Mode (Y/N): ").lower())[0]=="n" else "temp.csv")

# get participants
list_participate = get_participants(participate_file);

# process data
create_docx_files(certificate_file, list_participate)




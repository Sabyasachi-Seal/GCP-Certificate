#!/usr/bin/env python3
import re

# Based on this link
# https://stackoverflow.com/a/42829667/11970836
# This function replace data and keeps style
def docx_replace_regex(doc_obj, regex , replace):

    for p in doc_obj.paragraphs:
        if regex.search(p.text):
            inline = p.runs
            # Loop added to work with runs (strings with same style)
            for i in range(len(inline)):
                if regex.search(inline[i].text):
                    text = regex.sub(replace, inline[i].text)
                    inline[i].text = text

    for table in doc_obj.tables:
        for row in table.rows:
            for cell in row.cells:
                docx_replace_regex(cell, regex , replace)

# call docx_replace_regex due to inputs
def replace_info(doc, name, string):
    reg = re.compile(r""+string)
    replace = r""+name
    docx_replace_regex(doc, reg , replace)

def replace_participant_name(doc, name):
    string = "{Name Surname}"
    replace_info(doc, name, string)

def replace_gdsc_name(doc, gdsc):
    string = "{College Name}"
    replace_info(doc, gdsc, string)

def replace_lead_name(doc, lead):
    string = "{Lead Name}"
    replace_info(doc, lead, string)

def replace_facilitator_name(doc, facilitator):
    string = "{Facilitator Name}"
    replace_info(doc, facilitator, string)

def replace_event(doc, posi):
    string = "{Event Name}"
    replace_info(doc, posi+".", string)


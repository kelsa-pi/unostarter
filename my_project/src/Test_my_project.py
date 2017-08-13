# -*- coding: utf-8 -*-
# !/usr/bin/env python

import uno
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX

try:
    from pythonpath.unostarter import Office, Gui, Inspector
except:
    from unostarter import Office, Gui, Inspector


def Run_my_project(*args):
    # ---------- Office examples ----------
    office = Office()
    # get document
    document = office.getDocument()
    # access the document's text property
    text = document.Text
    # create a cursor
    cursor = text.createTextCursor()
    # insert the text into the document
    text.insertString(cursor, "Hello World", 0)

    # ---------- Gui examples ----------
    # a = Gui.SelectBox(message="Select one item", title="SelectBox", choices=['a', 'b', 'c'])
    # print(a)

    # b = Gui.MessageBox(message="Message", title="Title", messageType=QUERYBOX, messageButtons=393222)
    # print(b)

    # c = Gui.FolderPathBox(title='Get directory path')
    # print(c)

    # message box wizard
    # Gui.MBWizard()

    # ---------- Inspector examples ----------
    insp = Inspector()
    # insp object with MRI
    # insp.callMRI(document)

    # inspect object with inspect method
    # g = insp.inspect(document, item=['AllVersions','CharacterCount'])
    # print(g)

    # use print or message box to show result
    # Gui.MessageBox(message=str(g), title="Inspector", messageType=INFOBOX, messageButtons=65537)

    # show documentacion in browser
    # insp.showServiceDocs(text)


# Execute macro from LibreOffice - Tools - Macro
g_exportedScripts = Run_my_project,

# Execute macro from IDE
# Start the office from the command line eg:
# soffice "--accept=socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" --writer --norestore

if __name__ == "__main__":
    import os
    import sys

    sys.path.append(os.path.join(os.path.dirname(__file__), 'pythonpath'))

    Run_my_project()

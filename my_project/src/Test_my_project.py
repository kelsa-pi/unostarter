import uno
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX

try:
    from pythonpath.unostarter import Office, Gui, Inspector
except:
    from unostarter import Office, Gui, Inspector


def Run_my_project(*args):

    # ---------- Office example ----------
    office = Office()
    # get document
    document = office.getDocument()
    # access the document's text property
    text = document.Text
    # create a cursor
    cursor = text.createTextCursor()
    # insert the text into the document
    text.insertString(cursor, "Hello World", 0)

    # ---------- Gui example ----------
    a = Gui.SelectBox(message="Select one item", title="SelectBox", choices=['a', 'b', 'c'])
    print(a)

    b = Gui.MessageBox(message="Message", title="Title", messageType=QUERYBOX, messageButtons=393222)
    print(b)

    c = Gui.FolderPathBox(title='Get directory path')
    print(c)

    # message box wizard
    Gui.MBWizard()

    # ---------- Inspector example ----------
    inspector = Inspector()
    # inspect object with MRI
    inspector.callMRI(document)

    # inspect object with inspect method
    g = inspector.inspect(document, items=['AllVersions','CharacterCount'])
    # use print or message box to show result
    Gui.MessageBox(message=g, title="Inspector", messageType=INFOBOX, messageButtons=65537)

# Execute macro from LibreOffice - Tools - Macro
g_exportedScripts = Run_my_project,


# Execute macro from IDE
if __name__ == "__main__":
    import os
    import sys
    sys.path.append(os.path.join(os.path.dirname(__file__), 'pythonpath'))

    Run_my_project()

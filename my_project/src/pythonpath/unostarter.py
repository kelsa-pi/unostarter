# -*- coding: utf-8 -*-
#!/usr/bin/env python

# unostarter is helper for LibreOffice macro development
# Copyright (C) 2017  Sasa Kelecevic
#
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
#
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with this library; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA

import uno
import unohelper
from com.sun.star.awt import XActionListener
from com.sun.star.task import XJobExecutor
from com.sun.star.uno import RuntimeException

from com.sun.star.awt.MessageBoxType import \
    MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX
from com.sun.star.awt.MessageBoxButtons import \
    BUTTONS_OK, BUTTONS_OK_CANCEL, BUTTONS_YES_NO, \
    BUTTONS_YES_NO_CANCEL, BUTTONS_RETRY_CANCEL, BUTTONS_ABORT_IGNORE_RETRY
from com.sun.star.awt.MessageBoxButtons import \
    DEFAULT_BUTTON_OK, DEFAULT_BUTTON_CANCEL, DEFAULT_BUTTON_RETRY, \
    DEFAULT_BUTTON_YES, DEFAULT_BUTTON_NO, DEFAULT_BUTTON_IGNORE
from com.sun.star.beans.MethodConcept import ALL as METHOD_CONCEPT_ALL
from com.sun.star.beans.PropertyConcept import ALL as PROPERTY_CONCEPT_ALL

from com.sun.star.reflection.ParamMode import \
    IN as PARAM_MODE_IN, \
    OUT as PARAM_MODE_OUT, \
    INOUT as PARAM_MODE_INOUT

# change if needed
HOST = 'localhost'
PORT = 2002

__all__ = ['StartOffice', 'ConnectOffice', 'Office', 'Gui', 'Inspector']


def _mode_to_str(mode):
    ret = "[]"
    if mode == PARAM_MODE_INOUT:
        ret = "[inout]"
    elif mode == PARAM_MODE_OUT:
        ret = "[out]"
    elif mode == PARAM_MODE_IN:
        ret = "[in]"
    return ret


def _get_connection_url(host, port, pipe=None):
    if pipe:
        connection = 'pipe,name={}'.format(pipe)
    else:
        connection = 'socket,host={},port={}'.format(host, port)
    return 'uno:{};urp;StarOffice.ComponentContext'.format(connection)


def ConnectOffice(host=HOST, port=PORT, pipe=None, context=None):
    """Connect LibreOffice
    :param host: connect via socket, default 'localhost'
    :param port: connect via socket, default 2002
    :param pipe: connect via pipe, default None
    :param context: custom context, default None
    
    Start office:
    soffice "--accept=socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" --writer --norestore
    Connect office:
    "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"

    """
    conn = None
    if context is None:
        localContext = uno.getComponentContext()
        try:
            # LibreOffice is started as an OS process, remote connection
            url = _get_connection_url(host, port, pipe)
            remote = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext)
            conn = remote.resolve(url)
        except:
            # Connection inside the Office
            conn = localContext
    else:
        try:
            conn = context
        except:
            print('Error: no context')
    return conn


# ===========================================================
#               OFFICE
# ===========================================================

class Office:
    """Frequently used methods in office context
    """
    def __init__(self, context=None):
        if context:
            self.ctx = context
        else:
            self.ctx = ConnectOffice()

    def getContext(self):
        return self.ctx

    def getDesktop(self):
        """Access to the desktop environment"""
        desktop = self.ctx.getValueByName('/singletons/com.sun.star.frame.theDesktop')
        return desktop

    def getDocument(self):
        """Access to the current document"""
        return self.getDesktop().getCurrentComponent()

    def getSelection(self):
        """Access to the the current selection"""
        return self.getDocument().getSelection()

    def createUnoService(self, service):
        """Create UNO service"""
        return self.ctx.ServiceManager.createInstance(service)

    def createUnoStruct(self, struct):
        """Initialize without to import the class of your target struct with uno.createUnoStruct function

        :param struct: target struct, "com.sun.star.awt.Point"
        """
        return uno.createUnoStruct(struct)

    def filePathToUrl(self, path):
        return uno.systemPathToFileUrl(path)

    def fileUrlToPath(self, url):
        return uno.fileUrlToSystemPath(url)

# -----------------------------------------------------------
#               GUI CLASSES
# -----------------------------------------------------------


class SimpleDialog(unohelper.Base, XActionListener, XJobExecutor):
    """
    Class documentation...
    """
    def __init__(self, nPositionX=None, nPositionY=None, nWidth=None, nHeight=None, sTitle=None):
        self.ctx = ConnectOffice()
        self.ServiceManager = self.ctx.ServiceManager
        self.Toolkit = self.ServiceManager.createInstanceWithContext("com.sun.star.awt.ExtToolkit", self.ctx)
          #
        # --------------create dialog container and set model and properties
        self.DialogContainer = self.ServiceManager.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", self.ctx)
        self.DialogModel = self.ServiceManager.createInstance("com.sun.star.awt.UnoControlDialogModel")
        self.DialogContainer.setModel(self.DialogModel)
        self.DialogModel.PositionX = nPositionX
        self.DialogModel.PositionY = nPositionY
        self.DialogModel.Height = nHeight
        self.DialogModel.Width = nWidth
        self.DialogModel.Name = "Default"
        self.DialogModel.Closeable = True
        self.DialogModel.Moveable = True

    def addControl(self, sAwtName, sControlName, dProps):
        oControlModel = self.DialogModel.createInstance("com.sun.star.awt.UnoControl" + sAwtName + "Model")
        while dProps:
            prp = dProps.popitem()
            uno.invoke(oControlModel, "setPropertyValue", (prp[0], prp[1]))
            oControlModel.Name = sControlName
        self.DialogModel.insertByName(sControlName, oControlModel)
        if sAwtName == "Button":
            self.DialogContainer.getControl(sControlName).addActionListener(self)
            self.DialogContainer.getControl(sControlName).setActionCommand(sControlName + '_OnClick')
        return oControlModel

    def showDialog(self):
        self.DialogContainer.setVisible(True)
        self.DialogContainer.createPeer(self.Toolkit, None)
        self.DialogContainer.execute()


class SelectBoxClass(SimpleDialog):
    """
    Class documentation...
    """
    def __init__(self, message="Select one item", title="SelectBox", choices=None):
        SimpleDialog.__init__(self, nPositionX=60, nPositionY=60, nWidth=100, nHeight=55, sTitle=None)

        if choices is None:
            choices = ['a', 'b', 'c']

        self.DialogModel.Title = title

        dMessage = {"PositionY": 5, "PositionX": 5, "Height": 15, "Width": 90, "Label": message,}
        self.lbMessage = self.addControl("FixedText", "lbMessage", dMessage)

        dChoices = {"PositionY": 15, "PositionX": 5, "Height": 15, "Width": 90,"Dropdown": True,}
        self.cbChoices = self.addControl("ComboBox", "cbChoices", dChoices)
        self.cbChoices.StringItemList = tuple(choices)

        dOK = {"PositionY": 35, "PositionX": 30, "Height": 15, "Width": 30, "Label": "OK",}
        self.btnOK = self.addControl("Button", "btnOK", dOK)

        dCancel = {"PositionY": 35, "PositionX": 65, "Height": 15, "Width": 30, "Label": "Cancel",}
        self.btnCancel = self.addControl("Button", "btnCancel", dCancel)

        self.returnValue = None

    def actionPerformed(self, oActionEvent):
        if oActionEvent.ActionCommand == 'btnOK_OnClick':
            self.returnValue = self.cbChoices.Text
            self.DialogContainer.endExecute()

        if oActionEvent.ActionCommand == 'btnCancel_OnClick':
            self.DialogContainer.endExecute()

    def returnValue(self):
        pass


class OptionBoxClass(SimpleDialog):
    """
    Class documentation...
    """
    def __init__(self, message="Select multiple items", title="OptionBox", choices=['a', 'b', 'c']):
        SimpleDialog.__init__(self, nPositionX=60, nPositionY=60, nWidth=135, nHeight=120, sTitle=None)
        self.DialogModel.Title = title

        dMessage = {"PositionY": 5, "PositionX": 5, "Height": 15, "Width": 110, "Label": message,}
        self.lbMessage = self.addControl("FixedText", "lbMessage", dMessage)

        dChoices = {"PositionY": 15, "PositionX": 5, "Height": 80, "Width": 125, "MultiSelection": True}
        self.lbChoices = self.addControl("ListBox", "lbChoices", dChoices)
        self.lbChoices.StringItemList = tuple(choices)

        dSelectAll = {"PositionY": 100, "PositionX": 5, "Height": 15, "Width": 30, "Label": "Select All",}
        self.btnSelectAll = self.addControl("Button", "btnSelectAll", dSelectAll)

        dClearAll = {"PositionY": 100, "PositionX": 35, "Height": 15, "Width": 30, "Label": "Clear All",}
        self.btnClearAll = self.addControl("Button", "btnClearAll", dClearAll)

        dOK = {"PositionY": 100, "PositionX": 70, "Height": 15, "Width": 30, "Label": "OK",}
        self.btnOK = self.addControl("Button", "btnOK", dOK)

        dCancel = {"PositionY": 100, "PositionX": 100, "Height": 15, "Width": 30, "Label": "Cancel",}
        self.btnCancel = self.addControl("Button", "btnCancel", dCancel)

        self.returnValue = ()

    def actionPerformed(self, oActionEvent):
        if oActionEvent.ActionCommand == 'btnOK_OnClick':

            n = len(self.DialogContainer.getControl('lbChoices').getSelectedItems())
            if n == 0:
                self.returnValue = ()
            elif n == 1:
                item = self.DialogContainer.getControl('lbChoices').getSelectedItem()
                self.returnValue = (item,)
            else:
                self.returnValue = self.DialogContainer.getControl('lbChoices').getSelectedItems()

            self.DialogContainer.endExecute()

        if oActionEvent.ActionCommand == 'btnCancel_OnClick':
            self.DialogContainer.endExecute()

        if oActionEvent.ActionCommand == 'btnSelectAll_OnClick':
            for item in self.lbChoices.StringItemList:
                self.DialogContainer.getControl('lbChoices').selectItem(item, True)

        if oActionEvent.ActionCommand == 'btnClearAll_OnClick':
            for item in self.lbChoices.StringItemList:
                self.DialogContainer.getControl('lbChoices').selectItem(item, False)

    def returnValue(self):
        pass


class TextBoxClass(SimpleDialog):
    """
    Class documentation...
    """

    def __init__(self, message="Enter a text", title="TextBox", text=""):
        SimpleDialog.__init__(self, nPositionX=60, nPositionY=60, nWidth=100, nHeight=55, sTitle=None)
        self.DialogModel.Title = title

        dMessage = {"PositionY": 5, "PositionX": 5, "Height": 15, "Width": 90, "Label": message,}
        self.lbMessage = self.addControl("FixedText", "lbMessage", dMessage)

        dText = {"PositionY": 15, "PositionX": 5, "Height": 15, "Width": 90, "Text":text}
        self.txtText = self.addControl("Edit", "txtText", dText)

        dOK = {"PositionY": 35, "PositionX": 30, "Height": 15, "Width": 30, "Label": "OK",}
        self.btnOK = self.addControl("Button", "btnOK", dOK)

        dCancel = {"PositionY": 35, "PositionX": 65, "Height": 15, "Width": 30, "Label": "Cancel",}
        self.btnCancel = self.addControl("Button", "btnCancel", dCancel)

        self.returnValue = None

    def actionPerformed(self, oActionEvent):
        if oActionEvent.ActionCommand == 'btnOK_OnClick':
            self.returnValue = self.txtText.Text
            self.DialogContainer.endExecute()

        if oActionEvent.ActionCommand == 'btnCancel_OnClick':
            self.DialogContainer.endExecute()

    def returnValue(self):
        pass


class NumberBoxClass(SimpleDialog):
    """
    Class documentation...
    """

    def __init__(self, message="Enter a number", title="NumberBox", default_value=0, min_=-10000, max_=10000, decimals=0):
        SimpleDialog.__init__(self, nPositionX=60, nPositionY=60, nWidth=100, nHeight=55, sTitle=None)
        self.DialogModel.Title = title

        self.default_value = default_value
        self.min_ = min_
        self.max_ = max_
        self.decimals = decimals

        dMessage = {"PositionY": 5, "PositionX": 5, "Height": 15, "Width": 90, "Label": message,}
        self.lbMessage = self.addControl("FixedText", "lbMessage", dMessage)

        dNumber = {"PositionY": 15, "PositionX": 5, "Height": 15, "Width": 90,}
        self.nfNumber = self.addControl("NumericField", "nfNumber", dNumber)
        self.nfNumber.setPropertyValue("DecimalAccuracy", self.decimals)
        self.nfNumber.setPropertyValue("StrictFormat", True)
        self.nfNumber.setPropertyValue("Value", self.default_value)
        self.nfNumber.setPropertyValue("ValueMin", self.min_)
        self.nfNumber.setPropertyValue("ValueMax", self.max_)

        dOK = {"PositionY": 35, "PositionX": 30, "Height": 15, "Width": 30, "Label": "OK",}
        self.btnOK = self.addControl("Button", "btnOK", dOK)

        dCancel = {"PositionY": 35, "PositionX": 65, "Height": 15, "Width": 30, "Label": "Cancel",}
        self.btnCancel = self.addControl("Button", "btnCancel", dCancel)

        self.returnValue = None

    def actionPerformed(self, oActionEvent):
        if oActionEvent.ActionCommand == 'btnOK_OnClick':
            if self.decimals == 0:
                self.returnValue = int(self.nfNumber.Value)
            else:
                self.returnValue = self.nfNumber.Value

            self.DialogContainer.endExecute()

        if oActionEvent.ActionCommand == 'btnCancel_OnClick':
            self.DialogContainer.endExecute()

    def returnValue(self):
        pass


class DateBoxClass(SimpleDialog):
    """
    Class documentation...
    """

    def __init__(self, message="Choose a date", title='DateBox'):
        """
        the format of the displayed date 9: short YYYYMMDD

        """
        SimpleDialog.__init__(self, nPositionX=60, nPositionY=60, nWidth=100, nHeight=55, sTitle=None)
        self.DialogModel.Title = title

        dMessage = {"PositionY": 5, "PositionX": 5, "Height": 15, "Width": 90, "Label": message,}
        self.lbMessage = self.addControl("FixedText", "lbMessage", dMessage)

        dDate = {"PositionY": 15, "PositionX": 5, "Height": 15, "Width": 90, "Dropdown": True,
                 "StrictFormat": True, "DateFormat": 9}
        self.dbDate = self.addControl("DateField", "dbDate", dDate)

        dOK = {"PositionY": 35, "PositionX": 30, "Height": 15, "Width": 30, "Label": "OK",}
        self.btnOK = self.addControl("Button", "btnOK", dOK)

        dCancel = {"PositionY": 35, "PositionX": 65, "Height": 15, "Width": 30, "Label": "Cancel",}
        self.btnCancel = self.addControl("Button", "btnCancel", dCancel)

        self.returnValue = ""

    def actionPerformed(self, oActionEvent):
        if oActionEvent.ActionCommand == 'btnOK_OnClick':
            self.returnValue = self.dbDate.Text
            self.DialogContainer.endExecute()

        if oActionEvent.ActionCommand == 'btnCancel_OnClick':
            self.DialogContainer.endExecute()

    def returnValue(self):
        pass


class MessageBoxWizardClass(SimpleDialog):
    """
    Class documentation...
    """
    def __init__(self):
        """
        Message Box Wizard

        """
        SimpleDialog.__init__(self, nPositionX=60, nPositionY=60, nWidth=155, nHeight=180, sTitle='Message Box Wizard')
        self.DialogModel.Title = ' MessageBox Wizard'
        # title
        dLabelTitle = {"PositionY": 5, "PositionX": 5, "Height": 15, "Width": 30, "Label": 'Title'}
        self.lbTitle = self.addControl("FixedText", "lbTitle", dLabelTitle)
        dTitle = {"PositionY": 5, "PositionX": 35, "Height": 15, "Width": 115, "Text": 'Title'}
        self.txtTitle = self.addControl("Edit", "txtTitle", dTitle)
        # message
        dLabelMessage = {"PositionY": 20, "PositionX": 5, "Height": 15, "Width": 30, "Label": 'Message'}
        self.lbMessage = self.addControl("FixedText", "lbMessage", dLabelMessage)
        dMessage = {"PositionY": 20, "PositionX": 35, "Height": 15, "Width": 115, "Text": 'Message'}
        self.txtMessage = self.addControl("Edit", "txtMessage", dMessage)
        # type
        dLabelType = {"PositionY": 35, "PositionX": 5, "Height": 15, "Width": 30, "Label": 'Type'}
        self.lbMsgType = self.addControl("FixedText", "lbMsgType", dLabelType)
        mtype = [MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX]
        dMessageType = {"PositionY": 35, "PositionX": 35, "Height": 15, "Width": 115, "Dropdown": True}
        self.cbMsgType = self.addControl("ComboBox", "cbMsgType", dMessageType)
        self.cbMsgType.StringItemList = tuple(mtype)
        # buttons
        dLabelButtons = {"PositionY": 50, "PositionX": 5, "Height": 15, "Width": 30, "Label": 'Buttons'}
        self.lbMsgButtons = self.addControl("FixedText", "lbMsgButtons", dLabelButtons)

        self.mbtn = {'BUTTONS_OK':  1, 'BUTTONS_OK_CANCEL': 2, 'BUTTONS_YES_NO': 3, 'BUTTONS_YES_NO_CANCEL': 4, 'BUTTONS_RETRY_CANCEL': 5, 'BUTTONS_ABORT_IGNORE_RETRY': 6}
        dMessageButtons = {"PositionY": 50, "PositionX": 35, "Height": 15, "Width": 115, "Dropdown": True}
        self.cbMsgButtons = self.addControl("ComboBox", "cbMsgButtons", dMessageButtons)
        self.cbMsgButtons.StringItemList = tuple(self.mbtn.keys())
        # default buttons
        dLabelDefaultButtons = {"PositionY": 65, "PositionX": 5, "Height": 15, "Width": 30, "Label": 'Default',}
        self.lbMsgDefaultButtons = self.addControl("FixedText", "lbMsgDefaultButtons", dLabelDefaultButtons)

        self.mdefbtn = {'DEFAULT_BUTTON_OK': 65536, 'DEFAULT_BUTTON_CANCEL': 131072, 'DEFAULT_BUTTON_RETRY': 196608, 'DEFAULT_BUTTON_YES': 262144, 'DEFAULT_BUTTON_NO': 327680, 'DEFAULT_BUTTON_IGNORE': 393216}
        dMessageDefaultButtons = {"PositionY": 65, "PositionX": 35, "Height": 15, "Width": 115, "Dropdown": True}
        self.cbMsgDefaultButtons = self.addControl("ComboBox", "cbMsgDefaultButtons", dMessageDefaultButtons)
        self.cbMsgDefaultButtons.StringItemList = tuple(self.mdefbtn.keys())

        # code
        text = ''
        dText = {"PositionY": 101, "PositionX": 5, "Height": 60, "Width": 145, "Text": text, 'MultiLine': True, 'HScroll': True, 'VScroll': True}
        self.txtText = self.addControl("Edit", "txtText", dText)

        # imports
        dLabelImports = {"PositionY": 83, "PositionX": 5, "Height": 15, "Width": 30, "Label": 'Imports'}
        self.lbdImports = self.addControl("FixedText", "lbdImports", dLabelImports)
        dImports = {"PositionY": 83, "PositionX": 35, "Height": 15, "Width": 55, "Dropdown": True}
        self.cbImports = self.addControl("ComboBox", "cbImports", dImports)
        self.cbImports.StringItemList = tuple(['Minimal', 'All'])
        # dialog buttons
        dShow = {"PositionY": 83, "PositionX": 90, "Height": 15, "Width": 30, "Label": "Show"}
        self.btnShow = self.addControl("Button", "btnShow", dShow)
        dClear = {"PositionY": 83, "PositionX": 120, "Height": 15, "Width": 30, "Label": "Clear"}
        self.btnClear = self.addControl("Button", "btnClear", dClear)
        dClose = {"PositionY": 163, "PositionX": 120, "Height": 15, "Width": 30, "Label": "Close"}
        self.btnClose = self.addControl("Button", "btnClose", dClose)

        self.returnValue = None

    def actionPerformed(self, oActionEvent):
        if oActionEvent.ActionCommand == 'btnShow_OnClick':
            if self.cbImports.Text == 'Minimal':
                imports = "from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX\n\n"
                buttons = str(self.mbtn[self.cbMsgButtons.Text] + self.mdefbtn[self.cbMsgDefaultButtons.Text])
                t = imports + 'Gui.MessageBox(message="' + self.txtMessage.Text + '", title="' + self.txtTitle.Text + '", messageType=' + self.cbMsgType.Text + ', messageButtons=' + buttons + ')'

            elif self.cbImports.Text == 'All':

                imports = """from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX\nfrom com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_OK_CANCEL, BUTTONS_YES_NO, BUTTONS_YES_NO_CANCEL, BUTTONS_RETRY_CANCEL, BUTTONS_ABORT_IGNORE_RETRY\nfrom com.sun.star.awt.MessageBoxButtons import DEFAULT_BUTTON_OK, DEFAULT_BUTTON_CANCEL, DEFAULT_BUTTON_RETRY,    DEFAULT_BUTTON_YES, DEFAULT_BUTTON_NO, DEFAULT_BUTTON_IGNORE\n\n"""
                t = imports + 'Gui.MessageBox(message="' + self.txtMessage.Text + '", title="' + self.txtTitle.Text + '", messageType=' + self.cbMsgType.Text + ', messageButtons=' + self.cbMsgButtons.Text + ' + ' + self.cbMsgDefaultButtons.Text + ')'

            self.txtText.Text = t
            self.returnValue = 0

        if oActionEvent.ActionCommand == 'btnClear_OnClick':
             self.txtText.Text = ''

        if oActionEvent.ActionCommand == 'btnClose_OnClick':
            self.DialogContainer.endExecute()

    def returnValue(self):
        pass


# -----------------------------------------------------------
#               GUI FUNCTIONS
# -----------------------------------------------------------

class Gui:
    """
    Provides a simple dialog boxes for interaction with a user:

    make choices (SelectBox, OptionBox)
    enter new data (TextBox, NumberBox, DateBox)
    get paths (FolderPathBox, FilePathBox)
    show information (MessageBox)

    In script interactions are invoked by simple function calls.
    """

    def SelectBox(message="Select one item", title="SelectBox", choices=['a', 'b', 'c']):
        """
        Usage: SelectBox(message="Select one item", title="SelectBox", choices=['a','b','c'])

        Return: a string, or None
        """
        app = SelectBoxClass(message, title, choices)
        app.showDialog()
        return app.returnValue

    def OptionBox(message="Select multiple items", title="OptionBox", choices=['a', 'b', 'c']):
        """
        Usage: OptionBox(message="Select multiple items", title="OptionBox", choices=['a','b','c'])

        Return: a tuple of selected items, or empty tuple
        """
        app = OptionBoxClass(message, title, choices)
        app.showDialog()
        return app.returnValue

    def TextBox(message="Enter your input", title="TextBox", text=""):
        """
        Usage: TextBox(message="Enter your input", title="TextBox", text="")

        Return: a string, or None
        """
        app = TextBoxClass(message, title, text)
        app.showDialog()
        return app.returnValue

    def NumberBox(message="Enter a number", title="NumberBox", default_value=0, min_=-10000, max_=10000, decimals=0):
        """
        Usage: NumberBox(message="Enter a number", title="NumberBox", default_value=0, min_=-10000, max_=10000, decimals=0)

        Return: an integer/float or None
        """
        app = NumberBoxClass(message, title, default_value, min_, max_, decimals)
        app.showDialog()
        return app.returnValue

    def DateBox(message="Choose a date", title='DateBox'):
        """
        Usage: DateBox(message="Date of birth", title="BirthDay")

        Return: the selected date in format YYYYMMDD
        """
        app = DateBoxClass(message, title)
        app.showDialog()
        return app.returnValue

    def FolderPathBox(title='Get directory path'):
        """
        Usage: FolderPathBox(title='Get directory path')

        Return: the path of a directory or an empty string
        """
        ctx = ConnectOffice()
        smgr = ctx.getServiceManager()
        folder_picker = smgr.createInstanceWithContext("com.sun.star.ui.dialogs.FolderPicker", ctx)
        folder_picker.setTitle(title)
        folder_picker.execute()
        return folder_picker.getDirectory()

    def FilePathBox(title='Get file path', context=None):
        """
        Usage: FilePathBox(title='Get file path')

        Return: the path of a file or an empty string
        """
        ctx = ConnectOffice()
        smgr = ctx.getServiceManager()
        open_file_picker = smgr.createInstanceWithContext("com.sun.star.ui.dialogs.FilePicker", ctx)
        open_file_picker.setMultiSelectionMode(False)
        open_file_picker.setTitle(title)
        open_file_picker.appendFilter("All files (*.*)", "*.*")
        open_file_picker.execute()
        return open_file_picker.getSelectedFiles()[0]

    def MessageBox(message="Message", title="MessageBox", messageType=INFOBOX, messageButtons=BUTTONS_OK):
        """
        
        :param title: 
        :param messageType: 
        :param messageButtons: 
        :return:CANCEL = 0, OK = 1, YES = 2, NO = 3, RETRY = 4, IGNORE = 5 
 
        """
        ctx = ConnectOffice()
        sm = ctx.ServiceManager
        toolkit = sm.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
        parent_win = sm.createInstanceWithContext("com.sun.star.awt.ExtToolkit", ctx)
        messageBox = toolkit.createMessageBox(parent_win, messageType, messageButtons, title, message)
        rval = messageBox.execute()
        return rval

    def MBWizard():
        """
        Allows developers to quickly generate code for message boxes.
        """
        app = MessageBoxWizardClass()
        app.showDialog()
        return None


# -----------------------------------------------------------
#               INSPECTION
# -----------------------------------------------------------

class Inspector:
    """Frequently used methods in development context

    """
    def __init__(self, context=None):
        if context:
            self.ctx = context
        else:
            self.ctx = ConnectOffice()
        self.smgr = self.ctx.ServiceManager
        self.desktop = self.ctx.getValueByName('/singletons/com.sun.star.frame.theDesktop')
        self.introspection = self.ctx.getValueByName("/singletons/com.sun.star.beans.theIntrospection")
        self.documenter = self.ctx.getValueByName('/singletons/com.sun.star.util.theServiceDocumenter')

    def _inspectProperties(self, object, items=None):
        """Inspect properties

        :param object: object to inspect
        "param items: list method items
        """

        if items is None:
            items = []

        all = ""
        part = ""

        MP = {}
        try:
            inspector = self.introspection.inspect(object)
            # properties
            properties = inspector.getProperties(PROPERTY_CONCEPT_ALL)
            for ii in properties:
                try:
                    v = object.getPropertyValue(str(ii.Name))
                    t = str(v)
                    if t.startswith("pyuno object"):
                        v = "()"
                    if t.startswith("("):
                        v = "()"

                    MP['{:<35}'.format(str(ii.Name))] = str(v)
                except:
                    MP['{:<35}'.format(str(ii.Name))] = "()"

            for key, value in sorted(MP.items()):
                if key.strip() in items:
                    part = part + key + "  " + value + "\n"
                else:
                    all = all + key + "  " + value + "\n"
        except:
            pass

        if len(items) == 0:
            return all
        else:
            return part

    def _inspectMethods(self, object, items=None):
        """Inspect Methods

        :param object:
        :param items; list
        """

        if items is None:
            items = []

        all = ""
        part = ""

        MP = {}
        try:
            inspector = self.introspection.inspect(object)
            # methods
            methods = inspector.getMethods(METHOD_CONCEPT_ALL)
            for m in methods:
                args = m.ParameterTypes
                infos = m.ParameterInfos
                params = "("
                for i in range(0, len(args)):
                    params = params + _mode_to_str(infos[i].aMode) + " " + str(args[i].Name) + " " + str(infos[i].aName) + ", "
                params = params + ")"

                if params == "()":
                    params = "()"

                MP['{:<35}'.format(str(m.Name))] = str(params)

            for key, value in sorted(MP.items()):
                if key.strip() in items:
                    part = part + key + "  " + value + "\n"
                else:
                    all = all + key + "  " + value + "\n"

        except:
            pass

        if len(items) == 0:
            return all
        else:
            return part

    def callMRI(self, obj=None):
        """Create an instance of MRI inspector and inspect the given object"""
        try:
            if not obj:
                obj = self.desktop.getCurrentComponent().getSelection()
            mri = self.ctx.ServiceManager.createInstance("mytools.Mri")
            mri.inspect(obj)
        except:
            raise RuntimeException("\n MRI is not installed", self.ctx)

    def inspect(self, object, items=None):
        """Inspect object

        :param object: object to inspect
        "param items: list method items

        Return properties and methods
        """
        if items is None:
            items = []
        result = ''
        p = self._inspectProperties(object, items)
        m = self._inspectMethods(object, items)
        result = result + p + m
        return result

    def showServiceDocs(self, object):
        """Open browser to show service documentation
        :param object:
        """
        return self.documenter.showServiceDocs(object)

    def showInterfaceDoc(self, object):
        """Open browser to show interface documentation
        :param object:
        """
        return self.documenter.showInterfaceDoc(object)

if __name__ == "__main__":
    Gui.MBWizard()





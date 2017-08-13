# unostarter

Unostarter is a PyUNO project template and module which integrates frequently used methods in LibreOffice macro development. 

## Features

The unostarter provides:
* project template for easy start for beginners and newcomers
* example macro to start with 
* python module which integrates:
  * frequently used methods in macro development and
  * simple dialog boxes for interaction with a user
* execute macro:
   * inside the office `LibreOffice-Tools-Macro` menu
   * connect to remote procces 
   

## Usage

Unzip and copy my_project in `.../Script/python` directory. 
   
    my_project/                         > project root dir
            src/                        > source dir
                pythonpath/
                    unostarter.py          
            Test_my_project.py          > write your code here

Read the instructions in file `Test_my_project.py` and adapt to your needs.

## Office context

The class `Office` provides frequently used methods in office context
    
    getContext()  
    
    getDesktop()
    
    getDocument() 
    
    getSelection()       
    
    createUnoService(service)  
    
    createUnoStruct(struct)  
    
    filePathToUrl(path)  
    
    fileUrlToPath(url)
       
    
## Object inspection

The class `Inspector` provides frequently used methods in development context


    callMRI(obj=None)
    
    inspect(object, items=None)
    
    showServiceDocs(object)
    
    showInterfaceDoc(object)
    
## Basic GUI

The class `Gui` provides basic GUI boxes for interaction with a user

 
    SelectBox(message="Select one item", title="SelectBox", choices=['a', 'b', 'c'])
    
    OptionBox(message="Select multiple items", title="OptionBox", choices=['a', 'b', 'c'])  
    
    TextBox(message="Enter your input", title="TextBox", text="")   
    
    NumberBox(message="Enter a number", title="NumberBox", default_value=0, min_=-10000, max_=10000, decimals=0)   
    
    DateBox(message="Choose a date", title='DateBox')   
    
    FolderPathBox(title='Get directory path')   
    
    FilePathBox(title='Get file path', context=None)   
    
    MessageBox(message="Message", title="MessageBox", messageType=INFOBOX, messageButtons=BUTTONS_OK)   
    
    MBWizard() # MessageBox Wizard

    
    
    



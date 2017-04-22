# unostarter

PyUNO module which integrates frequently used methods in LibreOffice macro development.


## Usage
Unzip and copy my_project in `.../Script/python` directory. 
   
    my_project/                         > project root dir
            src/                        > source dir
                pythonpath/
                    unostarter.py          
            Test_my_project.py          > write your code here

Read the instructions in file `Test_my_project.py` and adapt to your needs.

## Frequently used methods in office context

    getContext()  
    
    getDesktop()
    
    getDocument() 
    
    getSelection()       
    
    createUnoService(service)  
    
    createUnoStruct(struct)  
    
    filePathToUrl(path)  
    
    fileUrlToPath(url)
       
    
## Frequently used methods in development context

    callMRI(obj=None)
    
    inspect(object, items=None)
    
    showServiceDocs(object)
    
    showInterfaceDoc(object)
    
## Basic GUI boxes for interaction with a user
 
    SelectBox(message="Select one item", title="SelectBox", choices=['a', 'b', 'c'])
    
    OptionBox(message="Select multiple items", title="OptionBox", choices=['a', 'b', 'c'])  
    
    TextBox(message="Enter your input", title="TextBox", text="")   
    
    NumberBox(message="Enter a number", title="NumberBox", default_value=0, min_=-10000, max_=10000, decimals=0)   
    
    DateBox(message="Choose a date", title='DateBox')   
    
    FolderPathBox(title='Get directory path')   
    
    FilePathBox(title='Get file path', context=None)   
    
    MessageBox(message="Message", title="MessageBox", messageType=INFOBOX, messageButtons=BUTTONS_OK)   
    
    MBWizard()

    
    
    



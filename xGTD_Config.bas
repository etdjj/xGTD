Attribute VB_Name = "xGTD_Config"
' xGTD Version 7 config

Function LoadSettings()

    'The local folder to store the EMAIL.
    strGTDFolderBase = "e:\03_DelphiTech\GTD-Reference\"
    
    'The Email address you want to send the task.
    strGTDMail = "etdjj.SQcAC@doitim.in"
    
    'The Note Email address.
    strNoteMail = "zhuce_cgf_56@mywiz.cn"
    
    'The folder in Outlook to store the archive Email.
    strGTDAchriveFoler = "Archive"
    
    'Control if add the subject to the name of local Email.- true or false
    AddSubjectInEMAILName = "true"
    
    'Config the GTD Tool - "doit" , "ZenDone", "RTM" supported
    GTDTOOL = "doit"
    
    'When no EMAIL selected, create the task without EMAIL. - true or false
    NewActWhenNoEmailSelect = "false"

End Function

Attribute VB_Name = "xGTD_Config"
Function LoadSettings()

    'The local folder to store the EMAIL.
    strGTDFolderBase = "e:\03_DelphiTech\GTD-Reference\"
    
    'The EMAIL address you want to send the task.
    strGTDMail = "etdjj.SQcAC@doitim.in"
    
    'The folder in Outlook to store the archive EMAIL. NOTE: it must be a subfolder of INBOX.
    strGTDAchriveFolerInOL = "Archive"
    
    'Control if add the subject to the name of local EMAIL.
    AddSubjectInEMAILName = "true"
    
    'Config the GTD Tool - "doit" , "ZenDone" supported
    GTDTOOL = "doit"

End Function

Attribute VB_Name = "xGTD"
' xGTD
' a outlook GTD plugin, work together with (EverNote,ZonDone);(Doit.im)
' Log
' Version 1: XuHui:first version support create action
' Version 2: XuHui:archive processed mail to specified folder
' Version 3: Guanfeng:support create action at Explore View
' Version 4: XuHui: fix ZenDoen creating action bug, add "-"

Public strGTDFolderBase As String
Public strGTDMail As String
Public strGTDAchriveFolerInOL As String
Public AddSubjectInEMAILName As String
Public GTDTOOL As String

Public myInbox As Outlook.Folder
Public myDestFolder As Outlook.Folder



Private Sub SendEmail(strSubject As String, strBody As String)

    Dim objMsg As MailItem
    Set objMsg = Application.CreateItem(olMailItem)

    With objMsg
        .To = strGTDMail
        .subject = strSubject
        .BodyFormat = olFormatHTML
        .HTMLBody = strBody
        .DeleteAfterSubmit = True
        .Send
    End With
     
    Set objMsg = Nothing
End Sub

Sub Initialize()

    LoadSettings

    If Dir(strGTDFolderBase, vbDirectory) = "" Then
        MkDir strGTDFolderBase
        MsgBox "Create GTD folder " & strGTDFolderBase
    Else
        MsgBox "Aleady have GTD folder " & strGTDFolderBase
    End If
End Sub


Sub CreateActionFromMail()
    Dim myNameSpace As Outlook.NameSpace
    
    LoadSettings

    Set myNameSpace = Application.GetNamespace("MAPI")
    Set myInbox = myNameSpace.GetDefaultFolder(olFolderInbox)
    Set myDestFolder = myInbox.Folders(strGTDAchriveFolerInOL)
   
    If TypeName(Application.ActiveWindow) = "Inspector" Then
        CreateFromInspector
    ElseIf TypeName(Application.ActiveWindow) = "Explorer" Then
        CreateFromExplore
    Else
        MsgBox "You are in the wrong active window." & TypeName(Application.ActiveWindow)
        Exit Sub
    End If
End Sub

Sub AchriveItem()
    Dim myNameSpace As Outlook.NameSpace
    
    LoadSettings

    Set myNameSpace = Application.GetNamespace("MAPI")
    Set myInbox = myNameSpace.GetDefaultFolder(olFolderInbox)
    Set myDestFolder = myInbox.Folders(strGTDAchriveFolerInOL)
   
    If TypeName(Application.ActiveWindow) = "Inspector" Then
        AchriveFromInspector
    ElseIf TypeName(Application.ActiveWindow) = "Explorer" Then
        AchriveFromExplore
    Else
        MsgBox "You are in the wrong active window." & TypeName(Application.ActiveWindow)
        Exit Sub
    End If
    
End Sub

Private Sub AchriveFromInspector()
    Set myInspector = Application.ActiveInspector
    If TypeName(myInspector.CurrentItem) = "MailItem" Then
        Set myItem = myInspector.CurrentItem
        
        On Error Resume Next
            myItem.Move myDestFolder

    End If
End Sub

Private Sub AchriveFromExplore()
    Dim MailSelection As Selection
    Dim SelectNum As Byte
    Dim MailObject As Object
    

    Set MailSelection = Application.ActiveExplorer.Selection
    SelectNum = MailSelection.Count

    For i = 1 To SelectNum
       Set MailObject = MailSelection.Item(i)
       
       If TypeName(MailObject) = "MailItem" Then
          MailObject.Move myDestFolder
          MailExist = "true"
       Else
          ExceptExist = "true"
       End If
    Next i
    
    If MailExist = "true" Then
        If ExceptExist = "true" Then
            MsgBox "Item which is not EMIAL Selected."
        End If
    Else
        MsgBox "Not Any EMAIL Selected."
    End If
End Sub

Private Sub CreateFromInspector()
    Dim myInspector As Outlook.Inspector
    Set myInspector = Application.ActiveInspector
    Dim myItem As Outlook.MailItem
    Dim strActionName As String
    Dim strGTDFolder As String
    Dim SendMailContent As String

    If TypeName(myInspector.CurrentItem) = "MailItem" Then
        Set myItem = myInspector.CurrentItem
        
        strGTDFolder = strGTDFolderBase & Format(DateValue(myItem.ReceivedTime), "yyyymmdd")
        If Dir(strGTDFolder, vbDirectory) = "" Then
            MkDir strGTDFolder
        End If
    
        strActionName = GetActionName()
        mailPath = FomatMailPath(strActionName, myItem.subject, strGTDFolder)
        
        myItem.SaveAs mailPath, olMSG
        
        SendMailContent = mailPath
        
        If GTDTOOL = "ZenDone" Then
            strActionName = "- " & strActionName
        End If
        SendEmail strActionName, SendMailContent
        
        On Error Resume Next
            myItem.Move myDestFolder
    Else
        MsgBox "The item is of the wrong type."
        Exit Sub
    End If

End Sub

Private Sub CreateFromExplore()
    Dim MailSelection As Selection
    Dim SelectNum As Byte
    Dim strActionName As String
    Dim MailObject As Object
    Dim SendMailContent As String
    Dim mailname As String
    Dim strGTDFolder As String
    
    strActionName = GetActionName()

    Set MailSelection = Application.ActiveExplorer.Selection
    SelectNum = MailSelection.Count

    For i = 1 To SelectNum
       Set MailObject = MailSelection.Item(i)
       
       If TypeName(MailObject) = "MailItem" Then
       
           strGTDFolder = strGTDFolderBase & Format(DateValue(MailObject.ReceivedTime), "yyyymmdd")
           If Dir(strGTDFolder, vbDirectory) = "" Then
               MkDir strGTDFolder
           End If
    
           mailPath = FomatMailPath(strActionName, MailObject.subject, strGTDFolder)
           MailObject.SaveAs mailPath, olMSG
           
           If i = 1 Then
              SendMailContent = mailPath
           Else
              SendMailContent = SendMailContent & "<br>" & mailPath
           End If
           
           MailObject.Move myDestFolder
           
           MailExist = "true"
       Else
           ExceptExist = "true"
       End If
    Next i
    
    If MailExist = "true" Then
        If GTDTOOL = "ZenDone" Then
            strActionName = "- " & strActionName
        End If
        SendEmail strActionName, SendMailContent
        If ExceptExist = "true" Then
            MsgBox "Item which is not EMIAL Selected."
        End If
    Else
        MsgBox "Not Any EMAIL Selected."
    End If
    

End Sub

Private Function FomatEMAILName(name As String) As String
    name = Replace(name, ".", " ")
    name = Replace(name, "/", " ")
    name = Replace(name, "\", " ")
    name = Replace(name, ":", " ")
    name = Replace(name, "~", " ")
    name = Replace(name, "#", " ")
    name = Replace(name, "$", " ")
    name = Replace(name, "%", " ")
    name = Replace(name, "^", " ")
    name = Replace(name, "|", " ")
    name = Replace(name, "&", " ")
    name = Replace(name, ";", " ")
    FomatEMAILName = name
End Function

Private Function GetActionName() As String
    strActionHelp = "Action with a due date tomorrow and contained in the project invitations " & vbNewLine
    strActionHelp = strActionHelp & "   - some action. tomorrow. invitations" & vbNewLine
    strActionHelp = strActionHelp & "Action contained in a new project named improve documentation that belongs to your home area of responsibility" & vbNewLine
    strActionHelp = strActionHelp & "   - some action. tomorrow. p: improve documentation. home " & vbNewLine
    strActionHelp = strActionHelp & "Action delegated to Mike" & vbNewLine
    strActionHelp = strActionHelp & "   - some action. mike" & vbNewLine
    strActionHelp = strActionHelp & "Next action with 2 contexts: errands and a new one named shopping" & vbNewLine
    strActionHelp = strActionHelp & "   - some action. errands. t: shopping. focus"
    
    If GTDTOOL = "doit" Then
        strActionHelp = "Input the task name"
    End If
    GetActionName = InputBox(strActionHelp, "Action Name")
    
    If GetActionName = "" Then
        MsgBox "Please type an action name"
        End
    End If
End Function

Private Function FomatMailPath(ActName As String, SubName As String, GTDFolder As String) As String
    Dim mailname As String
    If AddSubjectInEMAILName = "true" Then
        mailname = ActName & "-" & SubName
    Else
        mailname = ActName
    End If
    mailname = FomatEMAILName(mailname)
    FomatMailPath = GTDFolder & "\" & mailname & ".msg"
End Function

Sub GetCurrent_xGTDVersion()
    MsgBox "Version 4"
End Sub

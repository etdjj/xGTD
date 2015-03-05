Attribute VB_Name = "xGTD"
' xGTD
' a outlook GTD plugin, work together with (EverNote,ZonDone);(Doit.im)
' Log
' Version 1: XuHui:first version support create action
' Version 2: XuHui:archive processed mail to specified folder
' Version 3: Guanfeng:support create action at Explore View
' Version 4: XuHui: fix ZenDoen creating action bug, add "-"
' Version 5: Guanfeng:suppport create action without email
'                     support send email to note
'                     make email read after achrive
'                     move achrive folder out of Inbox
'                     optimize input box
'                     fix the issue when add subject to action name
'                     fix the issue when config AddSubjectInEMAILName = false
' Version 6  Guanfeng fix Achrive folder fix
'                     fix some bug for FomatEMAILName
' Version 7  XuHui: Support new GTD tool RTM

Public strGTDFolderBase As String
Public strGTDMail As String
Public strGTDAchriveFoler As String
Public AddSubjectInEMAILName As String
Public GTDTOOL As String
Public NewActWhenNoEmailSelect As String
Public strNoteMail As String

Sub GetCurrent_xGTDVersion()
    MsgBox "Version 7"
End Sub

Sub Initialize()

    LoadSettings

    If Dir(strGTDFolderBase, vbDirectory) = "" Then
        MkDir strGTDFolderBase
        MsgBox "Create GTD folder " & strGTDFolderBase
    End If

    On Error GoTo ErrorHandler
    Dim myNameSpace As Outlook.NameSpace
    Set myNameSpace = Application.GetNamespace("MAPI")
    Set myInbox = myNameSpace.GetDefaultFolder(olFolderInbox)
    myInbox.Parent.Folders.Add (strGTDAchriveFoler)

ErrorHandler:
     Set myAchrFolder = myInbox.Parent.Folders.Item(strGTDAchriveFoler)
     MsgBox "GTD Folder       =  " & strGTDFolderBase & vbCrLf & "Archive Folder  =  " & myAchrFolder.FolderPath
End Sub


Sub CreateActionFromMail()
    If TypeName(Application.ActiveWindow) = "Inspector" Then
        CreateFromInspector
    ElseIf TypeName(Application.ActiveWindow) = "Explorer" Then
        CreateFromExplore
    Else
        MsgBox "You are in the wrong active window." & TypeName(Application.ActiveWindow)
        Exit Sub
    End If
End Sub

Sub CreateActionFree()
    Dim strActionName As String
    LoadSettings

    strActionName = GetActionName()
    
    SendEmail strActionName, ""
End Sub

Sub AchriveItem()
    Dim myNameSpace As Outlook.NameSpace
    
    LoadSettings
   
    If TypeName(Application.ActiveWindow) = "Inspector" Then
        AchriveFromInspector
    ElseIf TypeName(Application.ActiveWindow) = "Explorer" Then
        AchriveFromExplore
    Else
        MsgBox "You are in the wrong active window." & TypeName(Application.ActiveWindow)
        Exit Sub
    End If
    
End Sub

Sub StoreToNote()
    If TypeName(Application.ActiveWindow) = "Inspector" Then
        StoreFromInspector
    ElseIf TypeName(Application.ActiveWindow) = "Explorer" Then
        StoreFromExplore
    Else
        MsgBox "You are in the wrong active window." & TypeName(Application.ActiveWindow)
        Exit Sub
    End If

End Sub



Private Sub StoreFromInspector()
    Dim myInspector As Outlook.Inspector
    Set myInspector = Application.ActiveInspector
    Dim myItem As Outlook.MailItem
    Dim strNoteName As String

    If TypeName(myInspector.CurrentItem) <> "" Then
        Set myItem = myInspector.CurrentItem
        strNoteName = GetNoteName(myItem.Subject)
        ForwardMail myItem, strNoteName, strNoteMail
        AchriveMailItem myItem
    Else
        MsgBox "The item is of the wrong type."
        Exit Sub
    End If

End Sub

Private Sub StoreFromExplore()
    Dim MailSelection As Selection
    Dim SelectNum As Byte
    Dim MailObject As Object
    Dim NoteName As String
    
    LoadSettings
    
    Set MailSelection = Application.ActiveExplorer.Selection
    SelectNum = MailSelection.Count

    If SelectNum = 0 Then
        MsgBox "Nothing Selected"
        Exit Sub
    End If

    If SelectNum = 1 Then
        Set MailObject = MailSelection.Item(1)
        NoteName = GetNoteName(MailObject.Subject)
        ForwardMail MailObject, NoteName, strNoteMail
        AchriveMailItem MailObject
    Else
        strNoteName = GetNoteName("null")
        For i = 1 To SelectNum
           Set MailObject = MailSelection.Item(i)
           If strNoteName = "null" Then
              NoteName = MailObject.Subject
           Else
              NoteName = strNoteName & "-" & MailObject.Subject
           End If
           ForwardMail MailObject, NoteName, strNoteMail
           AchriveMailItem MailObject
        Next i
    End If
End Sub


Private Sub AchriveFromInspector()
    Set myInspector = Application.ActiveInspector
    If TypeName(myInspector.CurrentItem) = "MailItem" Then
        Set myItem = myInspector.CurrentItem
        AchriveMailItem myItem
    End If
End Sub

Private Sub AchriveFromExplore()
    Dim MailSelection As Selection
    Dim SelectNum As Byte
    Dim MailObject As MailItem
    

    Set MailSelection = Application.ActiveExplorer.Selection
    SelectNum = MailSelection.Count

    For i = 1 To SelectNum
       Set MailObject = MailSelection.Item(i)
       
       If TypeName(MailObject) = "MailItem" Then
          AchriveMailItem MailObject
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
        mailPath = FomatMailPath(strActionName, myItem.Subject, strGTDFolder, 1)
        
        myItem.SaveAs mailPath, olMSG
        
        SendMailContent = "Reference:" & vbNewLine & mailPath
        
        If GTDTOOL = "ZenDone" Then
            strActionName = "- " & strActionName
        End If
        SendEmail strActionName, SendMailContent
        
        AchriveMailItem myItem
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
    Dim i As Byte
    
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
    
           mailPath = FomatMailPath(strActionName, MailObject.Subject, strGTDFolder, i)
           
           MailObject.SaveAs mailPath, olMSG
           
           If i = 1 Then
              SendMailContent = SendMailContent & mailPath
           Else
              SendMailContent = SendMailContent & "<br>" & mailPath
           End If
           
           AchriveMailItem MailObject
           
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
        If NewActWhenNoEmailSelect = "true" Then
            If GTDTOOL = "ZenDone" Then
                strActionName = "- " & strActionName
            End If
            SendEmail strActionName, ""
        Else
            MsgBox "No EMAIL is selected."
        End If
    End If
    

End Sub

Private Function FomatEMAILName(name As String) As String
    name = Replace(name, ".", "_")
    name = Replace(name, "/", "_")
    name = Replace(name, "\", "_")
    name = Replace(name, ":", "_")
    name = Replace(name, "*", "_")
    name = Replace(name, "?", "_")
    name = Replace(name, "<", "_")
    name = Replace(name, ">", "_")
    name = Replace(name, "|", "_")
    name = Replace(name, """", "_")
    name = Replace(name, "_ ", "_")
    name = Replace(name, " _", "_")
    name = Replace(name, "__", "_")
    name = Replace(name, "  ", " ")
    name = Replace(name, "  ", " ")
    name = Replace(name, "  ", " ")
    FomatEMAILName = name
End Function

Private Function GetActionName() As String

    Dim strActionHelp As String
    LoadSettings
    If GTDTOOL = "ZenDone" Then
        strActionHelp = "Action with a due date tomorrow and contained in the project invitations " & vbNewLine
        strActionHelp = strActionHelp & "   - some action. tomorrow. invitations" & vbNewLine
        strActionHelp = strActionHelp & "Action contained in a new project named improve documentation that belongs to your home area of responsibility" & vbNewLine
        strActionHelp = strActionHelp & "   - some action. tomorrow. p: improve documentation. home " & vbNewLine
        strActionHelp = strActionHelp & "Action delegated to Mike" & vbNewLine
        strActionHelp = strActionHelp & "   - some action. mike" & vbNewLine
        strActionHelp = strActionHelp & "Next action with 2 contexts: errands and a new one named shopping" & vbNewLine
        strActionHelp = strActionHelp & "   - some action. errands. t: shopping. focus"
    ElseIf GTDTOOL = "doit" Then
        strActionHelp = "Doit.im-GTD"
    ElseIf GTDTOOL = "RTM" Then
        strActionHelp = "Remeber The Milk" & vbNewLine
        strActionHelp = strActionHelp & "Example:" & vbNewLine
        strActionHelp = strActionHelp & "Take out the trash Monday at 8pm !1 *weekly =15min #Personal #errand" & vbNewLine
        strActionHelp = strActionHelp & "Result:A task named Take out the trash will be added to your Personal List (with the properties due Monday at 8pm, high priority, repeat weekly, time estimate 15 minutes, tagged errand)."
    Else
        strActionHelp = ""
    End If
    
    InputRet = GetInput(strActionHelp, "Action Name", "To Do")
    
    If InputRet = "cancel" Then
        End
    ElseIf InputRet = "null" Then
        MsgBox "Please type the action name."
        End
    Else
        GetActionName = InputRet
    End If
End Function

Private Function GetNoteName(Subject As String) As String
    Dim strHelp As String

    LoadSettings
    
    strHelp = "Forward Email to " & strNoteMail & vbCrLf & vbCrLf
    strHelp = strHelp & "As same as Email subject if keeping default name"
    
    InputRet = GetInput(strHelp, "Note Name", "plz input note name")
    
    If InputRet = "cancel" Then
        End
    ElseIf InputRet = "null" Then
        GetNoteName = Subject
    Else
        GetNoteName = InputRet
    End If
End Function

Private Function GetInput(Prompt As String, Title As String, default As String) As String

    InputStr = InputBox(Prompt, Title, default)
    
    If InputStr = "" Then
        GetInput = "cancel"
    ElseIf InputStr = default Then
        GetInput = "null"
    Else
        GetInput = InputStr
    End If
End Function


Private Function FomatMailPath(ActName As String, SubName As String, GTDFolder As String, index As Byte) As String
    Dim mailname As String
    If AddSubjectInEMAILName = "true" Then
        mailname = ActName & "-" & SubName
    Else
        If index = 1 Then
            mailname = ActName
        Else
            mailname = ActName & "-" & (index - 1)
        End If
    End If
    mailname = FomatEMAILName(mailname)
    FomatMailPath = GTDFolder & "\" & mailname & ".msg"
End Function


Private Function GetDestFolder() As Outlook.folder
    LoadSettings
    
    Dim myNameSpace As Outlook.NameSpace
    Set myNameSpace = Application.GetNamespace("MAPI")
    Set myInbox = myNameSpace.GetDefaultFolder(olFolderInbox)
    
    On Error GoTo CreateFolder
        Set GetDestFolder = myInbox.Parent.Folders.Item(strGTDAchriveFoler)
    Exit Function

CreateFolder:
    myInbox.Parent.Folders.Add (strGTDAchriveFoler)
    Set GetDestFolder = myInbox.Parent.Folders.Item(strGTDAchriveFoler)
    MsgBox "Archive Folder  =  " & GetDestFolder.FolderPath
End Function

Private Sub AchriveMailItem(ByVal MyMail As MailItem)
    MyMail.UnRead = False

    On Error Resume Next
    MyMail.Move GetDestFolder()
End Sub

Private Sub SendEmail(strSubject As String, strBody As String)

    Dim objMsg As MailItem
    Set objMsg = Application.CreateItem(olMailItem)

    With objMsg
        .To = strGTDMail
        .Subject = strSubject
        .BodyFormat = olFormatHTML
        .HTMLBody = strBody
        .DeleteAfterSubmit = True
        .Send
    End With
     
    Set objMsg = Nothing
End Sub

Private Sub ForwardMail(ByVal MailObject As MailItem, Subject As String, Receiver As String)
    Set objMsg = MailObject.Forward
    
    With objMsg
        .To = Receiver
        .Subject = Subject
        .DeleteAfterSubmit = True
        .Send
    End With
End Sub

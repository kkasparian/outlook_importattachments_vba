# outlook_importattachments_vba
Exports attachments (excel files) from emails and stores them in a folder in C: drive. Then filters through the attachments and exports a specific row onto another master excel on C: drive
Sub GetAttachments()

Dim ns As Namespace

Dim inbox As MAPIFolder

Dim Item As Object

Dim Atmt As Attachment

Dim SubFolder As MAPIFolder

Dim I As Integer

Dim NewFileName As String

' file in C drive that stores the CR Email attachments called "CR ATTACHMENTS"

Const AttachmentPath As String = "C:\CR ATTACHMENTS\"
      
Set ns = GetNamespace("MAPI")

Set inbox = ns.GetDefaultFolder(olFolderInbox)

' Folder in inbox where CR Emails automatically get forwarded to called "CR Emails"

Set SubFolder = inbox.Folders("CR Emails")
 
NewFileName = AttachmentPath & "CR ATTACHMENTS"



 Dim y As Workbook
 
'Opens up folder stored in C drive called "MASTER" that contains the "McLeod Summary" Excel

Set y = Workbooks.Open("C:\MASTER\McLeod Summary")

    If SubFolder.Items.Restrict("[UnRead]=True").Count > 0 Then
    
        For Each Item In SubFolder.Items.Restrict("[UnRead]=True")
        
        For Each Atmt In Item.Attachments
        
 ' Extracts all attachments from the unread emails and stores them in the "ATTACHMENTS" folder as "Latest CR Attachment" (.csv file)
 
        Filename = "C:\CR ATTACHMENTS\Latest CR Attachment.csv"
        
        Atmt.SaveAsFile Filename
         
        Dim x As Workbook
        
        'Opens up the "Latest CR Attachment" file
        
        Set x = Workbooks.Open("C:\CR ATTACHMENTS\Latest CR Attachment")
    
        I = I + 1
    
        Const strTest = "MDL_DEVTN EOD"
        
        Dim wsSource As Worksheet
        
        Dim wsDest As Worksheet
        
        Dim NoRows As Long
        
        Dim J As Long
        
        Dim rngCells As Range
        
        Dim rngFind As Range
        
      
        Dim copy_from As Range
        
        Dim copy_to As Range
  
  ' Setting the "Latest CR Attachment" as the worksheet source.
  
        Set wsSource = x.Sheets("Latest CR Attachment")
 
        NoRows = wsSource.Range("A65536").End(xlUp).Row
     
  ' Setting the "CR Data" worksheet in the "McLeod Summary" excel as the destination.
  
        Set wsDest = y.Sheets("CR Data")
        
        For J = 1 To NoRows
    
        Set rngCells = wsSource.Range("A" & J)
        
   ' If statement that looks for "MDL_DEVTN EOD" and pastes only that row to the next available row in master sheet
        
            If Not (rngCells.Find(strTest) Is Nothing) Then
        
                Set copy_from = wsSource.Range("A" & J)
                
                Set copy_to = wsDest.Range("A" & Rows.Count).End(xlUp).Offset(1, 0)

                copy_from.EntireRow.Copy Destination:=copy_to
                
                Application.CutCopyMode = False
               
            End If
        
        
    Next J
    
    x.Close
    
    Next Atmt
    
    Next Item
   
    End If
    
    
   

' Runs the "PFimport" Macro from the Master excel

Application.Run "'McLeod Summary.xlsm'!PFimport"

'Runs the "PAimport" Macro from the Master excel

Application.Run "'McLeod Summary.xlsm'!PAimport"

'Runs the "FCimport" Macro from the Master excel

Application.Run "'McLeod Summary.xlsm'!FCimport"

Application.Run "MarkUnreadCR()"

Application.Run "MarkUnreadPF()"

Application.Run "MarkUnreadPA()"

Application.Run "MarkUnreadFC()"

GetAttachments_exit:

Set Atmt = Nothing

Set Item = Nothing

Set ns = Nothing

Exit Sub

End Sub


Sub MarkUnreadCR()

Application.ScreenUpdating = False

Dim objInbox As Outlook.MAPIFolder

Dim objOutlook As Object, objnSpace As Object, objMessage As Object

Dim objSubfolder As Outlook.MAPIFolder

Set objOutlook = CreateObject("Outlook.Application")

Set objnSpace = objOutlook.GetNamespace("MAPI")

Set objInbox = objnSpace.GetDefaultFolder(olFolderInbox)

Set objSubfolder = objInbox.Folders.Item("CR Emails")

For Each objMessage In objSubfolder.Items

objMessage.UnRead = False

Next

Set objOutlook = Nothing

Set objnSpace = Nothing

Set objInbox = Nothing

Set objSubfolder = Nothing

Application.ScreenUpdating = True


End Sub


Sub MarkUnreadPF()

Application.ScreenUpdating = False

Dim objInbox As Outlook.MAPIFolder

Dim objOutlook As Object, objnSpace As Object, objMessage As Object

Dim objSubfolder As Outlook.MAPIFolder

Set objOutlook = CreateObject("Outlook.Application")

Set objnSpace = objOutlook.GetNamespace("MAPI")

Set objInbox = objnSpace.GetDefaultFolder(olFolderInbox)

Set objSubfolder = objInbox.Folders.Item("PF Emails")

For Each objMessage In objSubfolder.Items

objMessage.UnRead = False

Next

Set objOutlook = Nothing

Set objnSpace = Nothing

Set objInbox = Nothing

Set objSubfolder = Nothing

Application.ScreenUpdating = True


End Sub

Sub MarkUnreadPA()

Application.ScreenUpdating = False

Dim objInbox As Outlook.MAPIFolder

Dim objOutlook As Object, objnSpace As Object, objMessage As Object

Dim objSubfolder As Outlook.MAPIFolder

Set objOutlook = CreateObject("Outlook.Application")

Set objnSpace = objOutlook.GetNamespace("MAPI")

Set objInbox = objnSpace.GetDefaultFolder(olFolderInbox)

Set objSubfolder = objInbox.Folders.Item("PA Emails")

For Each objMessage In objSubfolder.Items

objMessage.UnRead = False

Next

Set objOutlook = Nothing

Set objnSpace = Nothing

Set objInbox = Nothing

Set objSubfolder = Nothing

Application.ScreenUpdating = True


End Sub

Sub MarkUnreadFC()

Application.ScreenUpdating = False

Dim objInbox As Outlook.MAPIFolder

Dim objOutlook As Object, objnSpace As Object, objMessage As Object

Dim objSubfolder As Outlook.MAPIFolder

Set objOutlook = CreateObject("Outlook.Application")

Set objnSpace = objOutlook.GetNamespace("MAPI")

Set objInbox = objnSpace.GetDefaultFolder(olFolderInbox)

Set objSubfolder = objInbox.Folders.Item("FC Emails")

For Each objMessage In objSubfolder.Items

objMessage.UnRead = False

Next

Set objOutlook = Nothing

Set objnSpace = Nothing

Set objInbox = Nothing

Set objSubfolder = Nothing

Application.ScreenUpdating = True


End Sub

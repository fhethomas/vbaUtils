Sub pickUpAttachment()

' *** Define your save location below ***
Dim folderDestStr As String
folderDestStr = "\\My Documents\Test"

' *** Identify your file type below ***
Dim fileTypeStr As String
fileTypeStr = ".xlsx"
Dim fileLengthInt As Integer 'length is defined by file type
fileLengthInt = Len(fileTypeStr)

' Define Outlook Variables
Dim objNS As Outlook.Namespace: Set objNS = GetNamespace("MAPI")
Dim olFolder As Outlook.MAPIFolder
Set olFolder = objNS.GetDefaultFolder(olFolderInbox)
Dim Item As Object
Dim attachmentObjs As Object
Dim attachment As Object 'Outlook.Attachment

' iterate over mailbox
For Each Item In olFolder.Items
    If TypeOf Item Is Outlook.MailItem Then
        Dim oMail As Outlook.MailItem: Set oMail = Item
        
            ' find attachments
            Set attachmentObjs = oMail.Attachments
            If attachmentObjs.Count > 0 Then
                ' Only look at xlsx files
                For Each attachment In attachmentObjs
                    If Right(attachment.Filename, fileLengthInt) = fileTypeStr Then
                        ' Print to console
                        Debug.Print oMail.SenderEmailAddress
                        Debug.Print oMail.Subject
                        Debug.Print "Attachments"
                        Debug.Print attachment.Filename
                        ' Save file to location
                        attachment.SaveAsFile folderDestStr & "\" & attachment.Filename
                    End If
                Next attachment

            End If

    End If
Next

End Sub

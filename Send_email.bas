Sub Send_email()

Dim olApp As Outlook.Application
Dim olEmail As Outlook.MailItem

Set olApp = New Outlook.Application
Set olEmail = olApp.CreateItem(olMailItem)

olEmail.Display

With olEmail
    .BodyFormat = olFormatRichText
    .Display
    .Body = "I am sending sheet_name file" & vbNewLine & .Body
    .Attachments.Add Environ("UserProfile") & "\Desktop\excel\sheet_name.docx"
    .To = "person1@gmail.com; person2@gmail.com"
    .Subject = "sheet_name"
    .Send
End With

End Sub

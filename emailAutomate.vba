Sub testEmail()
	Dim OutApp, OutMail As Object
	Dim strbody As String
	Set OutApp = CreateObject(“Outlook.Application”)
	Set OutMail=OutApp.CreateItem(0)
	Strbody = “Hi there “ & vbNewLine & vbNewLine
	On Error Resume Next
	With OutMail
		.To=testemail@gmail.com
		.CC=””
		.BCC=””
		.Subject=”This is the subject line”
		.Body=strbody
		‘you can add files like:
		.Attachments.Add(“C:\test.txt”)
		.Send ‘or use .Display
	End With
	On Error GoTo 0
	Set OutMail=Nothing
	Set OutApp=Nothing
End Sub

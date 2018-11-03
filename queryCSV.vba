Sub qryCSV()
	Dim sqlStr, strcon, fileLoc As String
	Dim sqlFile As String
	Dim outputSht as Worksheet
	Set outputSht  = worksheets(“Output”)
	Dim cn as Object
	Dim rs As Object
	‘Define our files and queries
	sqlFile="datafile.csv"
	sqlStr="SELECT * FROM [datafile.csv];"
	fileLoc="C:\Users\"
	‘creates connection us MS Access
	Strcon= “Provider=Microsoft.ACE.OLEDB.12.0; Data Source=” & fileLoc & “; Extended Properties = ‘text;HDR=YES;FMT=Delimited’;”
	‘set up objects
	Set cn = CreateObject(“ADODB.Connection”)
	Set rs = CreateObject(“ADODB.RECORDSET”)
	‘open connection and run SQL – output record set
	cn.Open strcon
	rs.activeconnection=cn
	rs.Open sqlStr
	OutputSht.Range(“A2”).CopyFromRecordset rs
	rs.close
	cn.close
	Set cn=Nothing
	Set rs=Nothing
	Application.CutCopyMode = False
End Sub

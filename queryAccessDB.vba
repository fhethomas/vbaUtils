Sub sqlQuery(outputSht, strSQL, strFile)
        'outputSht accepts output
	'strSQL is the SQL to run
	'strFile is the MSAccess DB 
	Dim cn As Object
	Dim rs As Object
	Dim strConnection As String

	‘Link to MS Access & uses our file location as data source
	strConnection=”Provider=Microsoft.ACE.OLEDB.12.0; Data Source=” & strFile
	Set cn = CreateObject(“ADODB.Connection”)
	cn.Open strConnection
	Set rs = CreateObject(“ADODB.RECORDSET”)
	rs.activeconnection=cn
	rs.Open strlSQL
	outputSht.Range(“A2”).CopyFromRecordset rs
	rs.Close
	cn.Close
	Set cn=Nothing
	Set rs=Nothing
End Sub

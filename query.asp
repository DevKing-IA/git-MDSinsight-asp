<%
'declare the variables
Dim Connection

Dim Recordset
Dim SQL


'declare the SQL statement that will query the database
SQL = "SELECT * FROM tblUsers"
'Response.write SQL

'create an instance of the ADO connection and recordset objects
Set Connection = Server.CreateObject("ADODB.Connection")
Set Recordset = Server.CreateObject("ADODB.Recordset")

'Open the connection to the database
Connection.Open InsightCnnString

'Open the recordset object executing the SQL statement and return records
Recordset.Open SQL,Connection,3,3


'first of all determine whether there are any records
If Recordset.EOF Then
Response.Write("No records returned.")
Else
'if there are records then loop through the fields
Do While NOT Recordset.Eof   
'Response.write Recordset("id")
Response.write Recordset("userEmail")
Response.write " - "   
Response.write Recordset("userpassword")
Response.write "<br>"   
Recordset.MoveNext    
Loop
End If

'close the connection and recordset objects to free up resources
Recordset.Close
Set Recordset=nothing
Connection.Close
Set Connection=nothing
%>
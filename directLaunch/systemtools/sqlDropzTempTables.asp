<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<script type="text/javascript">
	function closeme() {
		window.open('', '_parent', '');
		window.close();  }
</script>
<%
Server.ScriptTimeout = 7000
'SQL Table Creation and Modification Script
'Designed to be launched via a scheduled process (Win Task Scheduler)
'Self contained page will check the alerts db and take the appropriate actions
'Usage = "http://{xxx}.{domain}.com/directLaunch/systemtools/sqlDropzTempTables.asp?runlevel=run_now

'The runlevel parameter is inconsequential to the operation 
'of the page. It is only used so that the page will not run
'if it is loaded via an unexpected method (spiders, etc)
If Request.QueryString("runlevel") <> "run_now" then response.end

'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)

'This single page loops through and handles autocompletes for all databases
SQL = "SELECT * FROM tblServerInfo WHERE Active = 1"

Set TopConnection = Server.CreateObject("ADODB.Connection")
Set TopRecordset = Server.CreateObject("ADODB.Recordset")
TopConnection.Open InsightCnnString
'Open the recordset object executing the SQL statement and return records
TopRecordset.Open SQL,TopConnection,3,3


'First lookup the ClientKey in tblServerInfo
'If there is no record with the entered client key, close connection
'and exit
If Not TopRecordset.Eof Then

	Do While Not TopRecordset.EOF

		PassPhrase = TopRecordset.Fields("directLaunchPassphrase")
		ClientKey = TopRecordset.Fields("clientkey")

		Response.Write("******** START DELETING SQL Z TEMP TABLES For " & ClientKey  & "************<br>")
		
		Call SetClientCnnString
		
		Session("ClientCnnString") = MUV_READ("ClientCnnString") ' Until session vars are gone, then delete this
		
		If MUV_READ("cnnStatus") = "OK" Then ' else it loops and excludes dev client keys
			
			'****************************************
			'Begin Modify, Create SQL Tables
			'****************************************
			 Response.Write("Begin DELETING SQL Z TEMP TABLES<br>")
			'******************************************

			Server.ScriptTimeout = 500
			
			Set cnnTableDropzTempTables = Server.CreateObject("ADODB.Connection")
			cnnTableDropzTempTables.open (Session("ClientCnnString"))
			Set rsTableDropzTempTables = Server.CreateObject("ADODB.Recordset")
			rsTableDropzTempTables.CursorLocation = 3 
		
			
			Set cnnTableDrop = Server.CreateObject("ADODB.Connection")
			cnnTableDrop.open (Session("ClientCnnString"))
			Set rsTableDrop = Server.CreateObject("ADODB.Recordset")
			rsTableDrop.CursorLocation = 3 
		
		
		
			SQLTableDropzTempTables = "SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME LIKE 'z%'"
			Response.Write("SQLTableDropzTempTables: " & SQLTableDropzTempTables & "<br>")
			
			Set rsTableDropzTempTables = cnnTableDropzTempTables.Execute(SQLTableDropzTempTables)
			
			If Not rsTableDropzTempTables.EOF Then
			
				Do While Not rsTableDropzTempTables.EOF
				
					tableToDrop = rsTableDropzTempTables("TABLE_NAME")
					
					Response.Write("About to drop table: " & tableToDrop & "<br>")
					
					SQLTableDrop = "DROP TABLE " & tableToDrop & ";"
					
					Response.Write("SQLTableDrop : " & SQLTableDrop & "<br>")
					
					Set rsTableDrop = cnnTableDrop.Execute(SQLTableDrop)
					
					rsTableDropzTempTables.MoveNext
				
				Loop
				
			End If
			
	
			set rsTableDrop = nothing
			cnnTableDrop.close
			set cnnTableDrop = nothing
					
			set rsTableDropzTempTables = nothing
			cnnTableDropzTempTables.close
			set cnnTableDropzTempTables = nothing
			
			Response.Write("End DELETING SQL Z TEMP TABLES<br>")
			'******************************************	
			
		End If
		
		Response.Write("******** DONE DELETING SQL Z TEMP TABLES For " & ClientKey  & "************<br>")
			
					
	TopRecordset.movenext
	
	Loop

	TopRecordset.Close
	Set TopRecordset = Nothing
	TopConnection.Close
	Set TopConnection = Nothing
	
End If

Response.write("<script type=""text/javascript"">closeme();</script>")	
'Response.End
'*************************
'*************************
'Subs and funcs begin here


Sub SetClientCnnString

	dummy=MUV_WRITE("cnnStatus","")

	SQL = "SELECT * FROM tblServerInfo where clientKey='"& ClientKey &"'"

	Set Connection = Server.CreateObject("ADODB.Connection")
	Set Recordset = Server.CreateObject("ADODB.Recordset")
	Connection.Open InsightCnnString
	
	'Open the recordset object executing the SQL statement and return records
	Recordset.Open SQL,Connection,3,3

	
	'First lookup the ClientKey in tblServerInfo
	'If there is no record with the entered client key, close connection
	'and exit
	If Recordset.recordcount <= 0 then
		Recordset.close
		Connection.close
		set Recordset=nothing
		set Connection=nothing
	Else
		ClientCnnString = "Driver={SQL Server};Server=" & Recordset.Fields("dbServer")
		ClientCnnString = ClientCnnString & ";Database=" & Recordset.Fields("dbCatalog")
		ClientCnnString = ClientCnnString & ";Uid=" & Recordset.Fields("dbLogin")
		ClientCnnString = ClientCnnString & ";Pwd=" & Recordset.Fields("dbPassword") & ";"
		dummy = MUV_Write("ClientCnnString",ClientCnnString)
		dummy = MUV_Write("SQL_Owner",Recordset.Fields("dbLogin"))
		dummy = MUV_Write("ClientID",Recordset.Fields("clientkey"))
		Recordset.close
		Connection.close
		dummy=MUV_WRITE("cnnStatus","OK")
	End If
End Sub



%>
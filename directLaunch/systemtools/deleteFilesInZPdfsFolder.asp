<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<script type="text/javascript">
	function closeme() {
		window.open('', '_parent', '');
		window.close();  }
</script>
<%
	Server.ScriptTimeout = 7000
	
	'File Folder Contents Deletion Script
	'Designed to be launched via a scheduled process (Win Task Scheduler)
	'Self contained page will check the clientfiles/zPdfs directory and delete all files within it
	
	'Usage = "http://{xxx}.{domain}.com/directLaunch/systemtools/deleteFilesInZPdfsFolder.asp?runlevel=run_now
	
	'The runlevel parameter is inconsequential to the operation 
	'of the page. It is only used so that the page will not run
	'if it is loaded via an unexpected method (spiders, etc)
	If Request.QueryString("runlevel") <> "run_now" then response.end

	
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
	
			Response.Write("******** START DELETING CLIENTFILES/zPDFS FOLDER CONTENTS FOR " & ClientKey & "************<br>")
			
			Call SetClientCnnString
			
			Session("ClientCnnString") = MUV_READ("ClientCnnString") ' Until session vars are gone, then delete this
			
			If MUV_READ("cnnStatus") = "OK" Then ' else it loops and excludes dev client keys
				
				'****************************************
				'Begin Modify, Create SQL Tables
				'****************************************
				 Response.Write("***BEGIN DELETING CLIENTFILES/zPDFS FOLDER CONTENTS*********<br>")
				'******************************************
	
				Server.ScriptTimeout = 500
				
				Set objFS = CreateObject("Scripting.FileSystemObject")
				Set objFolder = objFS.GetFolder(Server.MapPath("..\..\") & "\clientfiles\"  & ClientKey & "\z_pdfs") 
				Set objFiles = objFolder.Files
				dim curFile
				
				For each curFile in objFiles
					Response.Write("Currently deleting file clientfiles\" & ClientKey & "\zPdfs\" & curFile.Name & "<br>")
					objFS.DeleteFile(curFile)
				Next
				
				Response.Write("***END DELETING CLIENTFILES/zPDFS FOLDER CONTENTS*********<br>")
				'******************************************	
				
			End If
			
			Response.Write("******** DONE DELETING CLIENTFILES/zPDFS FOLDER CONTENTS FOR " & ClientKey  & "************<br>")
				
						
		TopRecordset.movenext
		
		Loop
	
		TopRecordset.Close
		Set TopRecordset = Nothing
		TopConnection.Close
		Set TopConnection = Nothing
		
	End If
	
	Response.write("<script type=""text/javascript"">closeme();</script>")	


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
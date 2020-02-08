<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<script type="text/javascript">
	function closeme() {
		window.open('', '_parent', '');
		window.close();  }
</script>
<%
Server.ScriptTimeout = 7000
'Client Files Virtual Directory File Folder Modification Script
'Designed to be launched via a scheduled process (Win Task Scheduler)
'Self contained page will check the alerts db and take the appropriate actions
'Usage = "http://{xxx}.{domain}.com/directLaunch/systemtools/createClientFileVirtualDirectories.asp?runlevel=run_now

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

dim filesys, pathToCurrentFolderToFind, newFolderToCreate
set filesys = CreateObject("Scripting.FileSystemObject")

If Not TopRecordset.Eof Then

	Do While Not TopRecordset.EOF

		PassPhrase = TopRecordset.Fields("directLaunchPassphrase")
		ClientKey = TopRecordset.Fields("clientkey")
		
		Response.Write("<br><br><br>******** START Processing Client File Folder Synchronizing For " & ClientKey  & "************<br><br>")
		
		Call SetClientCnnString
		
		Session("ClientCnnString") = MUV_READ("ClientCnnString") ' Until session vars are gone, then delete this
		
		If MUV_READ("cnnStatus") = "OK" Then ' else it loops and excludes dev client keys
			
			'*******************************************************************
			'Begin Client File Folder Synchronizing
			'*******************************************************************
			
			 Response.Write("Begin Client File Folder Synchronizing<br><br>")
			 'Response.write(Server.MapPath(".") & "<br><br>")
			'******************************************

			Server.ScriptTimeout = 500
			
			'***********************************************	
			'CURRENT TOP LEVEL FOLDERS AND SUBFOLDERS
			'***********************************************	
			
			' 1. emails
			' 2. logos
			' 3. ftp/outbound
			' 4. z_pdfs
			' 5. accountsreceivable/consolidated
			' 6. Copy logo.png from individual clientfiles folder to clientfilesV
			' 7. prospecting/emailAttachments
			
			'******************************************	
			
			serverName = Request.ServerVariables("SERVER_NAME")
			
			If serverName = "www.mdsinsight.com" Then serverName = "mdsinsight.com"
			
			Response.Write("serverName: " & serverName & "<br>")
			
			
			If (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") = 0 AND Instr(ucase(ClientKey),"D") = 0)_
			or (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 AND Instr(ucase(ClientKey),"D") <> 0) Then 

				'**************************************************
				'Let's start with the {clientKey} under ClientFiles
				'**************************************************
				
				pathToCurrentFolderToFind = "C:\home\clientfilesV\" & ClientKey & "\"
				
				Response.Write("path to check: " & pathToCurrentFolderToFind & "<br>")
				
				If Not filesys.FolderExists(pathToCurrentFolderToFind) Then 
					Response.write("creating: " & pathToCurrentFolderToFind & "<br>")
					newFolderToCreate = filesys.CreateFolder (pathToCurrentFolderToFind) 
					Response.Write ("A new folder for Client " & ClientKey & " <strong>(" & pathToCurrentFolderToFind & ")</strong> has been created.<br>") 
				Else
					Response.Write ("The folder <strong>(" & pathToCurrentFolderToFind & ")</strong> already exists for Client " & ClientKey & ".<br>")
				End If

				
				'***********************
				'emails
				'***********************
				
				pathToCurrentFolderToFind = "C:\home\clientfilesV\" & ClientKey & "\emails\"
				
				Response.Write("path to check: " & pathToCurrentFolderToFind & "<br>")
				
				If Not filesys.FolderExists(pathToCurrentFolderToFind) Then 
					newFolderToCreate = filesys.CreateFolder (pathToCurrentFolderToFind) 
					Response.Write ("A new folder for Client " & ClientKey & " <strong>(" & pathToCurrentFolderToFind & ")</strong> has been created.<br>") 
				Else
					Response.Write ("The folder <strong>(" & pathToCurrentFolderToFind & ")</strong> already exists for Client " & ClientKey & ".<br>")
				End If
				

				'***********************
				'logos
				'***********************
				
				pathToCurrentFolderToFind = "C:\home\clientfilesV\" & ClientKey & "\logos\"
				
				Response.Write("path to check: " & pathToCurrentFolderToFind & "<br>")
				
				If Not filesys.FolderExists(pathToCurrentFolderToFind) Then 
					newFolderToCreate = filesys.CreateFolder (pathToCurrentFolderToFind) 
					Response.Write ("A new folder for Client " & ClientKey & " <strong>(" & pathToCurrentFolderToFind & ")</strong> has been created.<br>") 
				Else
					Response.Write ("The folder <strong>(" & pathToCurrentFolderToFind & ")</strong> already exists for Client " & ClientKey & ".<br>")
				End If


				'***********************
				'z_pdfs
				'***********************
				
				pathToCurrentFolderToFind = "C:\home\clientfilesV\" & ClientKey & "\z_pdfs\"
				
				Response.Write("path to check: " & pathToCurrentFolderToFind & "<br>")
				
				If Not filesys.FolderExists(pathToCurrentFolderToFind) Then 
					newFolderToCreate = filesys.CreateFolder (pathToCurrentFolderToFind) 
					Response.Write ("A new folder for Client " & ClientKey & " <strong>(" & pathToCurrentFolderToFind & ")</strong> has been created.<br>") 
				Else
					Response.Write ("The folder <strong>(" & pathToCurrentFolderToFind & ")</strong> already exists for Client " & ClientKey & ".<br>")
				End If

				'***********************
				'uploaded_data
				'***********************
				
				pathToCurrentFolderToFind = "C:\home\clientfilesV\" & ClientKey & "\uploaded_data\"
				
				Response.Write("path to check: " & pathToCurrentFolderToFind & "<br>")
				
				If Not filesys.FolderExists(pathToCurrentFolderToFind) Then 
					newFolderToCreate = filesys.CreateFolder (pathToCurrentFolderToFind) 
					Response.Write ("A new folder for Client " & ClientKey & " <strong>(" & pathToCurrentFolderToFind & ")</strong> has been created.<br>") 
				Else
					Response.Write ("The folder <strong>(" & pathToCurrentFolderToFind & ")</strong> already exists for Client " & ClientKey & ".<br>")
				End If


				'***********************
				'ftp/outbound
				'***********************
				
				pathToCurrentFolderToFind = "C:\home\clientfilesV\" & ClientKey & "\ftp\"
				
				Response.Write("path to check: " & pathToCurrentFolderToFind & "<br>")
				
				If Not filesys.FolderExists(pathToCurrentFolderToFind) Then 
					newFolderToCreate = filesys.CreateFolder (pathToCurrentFolderToFind) 
					Response.Write ("A new folder for Client " & ClientKey & " <strong>(" & pathToCurrentFolderToFind & ")</strong> has been created.<br>") 
				Else
					Response.Write ("The folder <strong>(" & pathToCurrentFolderToFind & ")</strong> already exists for Client " & ClientKey & ".<br>")
				End If

				pathToCurrentFolderToFind = "C:\home\clientfilesV\" & ClientKey & "\ftp\outbound\"
				
				Response.Write("path to check: " & pathToCurrentFolderToFind & "<br>")
				
				If Not filesys.FolderExists(pathToCurrentFolderToFind) Then 
					newFolderToCreate = filesys.CreateFolder (pathToCurrentFolderToFind) 
					Response.Write ("A new folder for Client " & ClientKey & " <strong>(" & pathToCurrentFolderToFind & ")</strong> has been created.<br>") 
				Else
					Response.Write ("The folder <strong>(" & pathToCurrentFolderToFind & ")</strong> already exists for Client " & ClientKey & ".<br>")
				End If



				'***********************
				'accountsreceivable/consolidated
				'***********************
				
				pathToCurrentFolderToFind = "C:\home\clientfilesV\" & ClientKey & "\accountsreceivable\"
				
				Response.Write("path to check: " & pathToCurrentFolderToFind & "<br>")
				
				If Not filesys.FolderExists(pathToCurrentFolderToFind) Then 
					newFolderToCreate = filesys.CreateFolder (pathToCurrentFolderToFind) 
					Response.Write ("A new folder for Client " & ClientKey & " <strong>(" & pathToCurrentFolderToFind & ")</strong> has been created.<br>") 
				Else
					Response.Write ("The folder <strong>(" & pathToCurrentFolderToFind & ")</strong> already exists for Client " & ClientKey & ".<br>")
				End If

				pathToCurrentFolderToFind = "C:\home\clientfilesV\" & ClientKey & "\accountsreceivable\consolidated\"
				
				Response.Write("path to check: " & pathToCurrentFolderToFind & "<br>")
				
				If Not filesys.FolderExists(pathToCurrentFolderToFind) Then 
					newFolderToCreate = filesys.CreateFolder (pathToCurrentFolderToFind) 
					Response.Write ("A new folder for Client " & ClientKey & " <strong>(" & pathToCurrentFolderToFind & ")</strong> has been created.<br>") 
				Else
					Response.Write ("The folder <strong>(" & pathToCurrentFolderToFind & ")</strong> already exists for Client " & ClientKey & ".<br>")
				End If
				

				'****************************************************************
				'Copy logo.png from individual clientfiles folder to clientfilesV
				'****************************************************************
				
				pathToLogoToFind = "C:\home\clientfilesV\" & ClientKey & "\logos\logo.png"
				
				Response.Write("path & file to check: " & pathToLogoToFind & "<br>")
				
				If Not filesys.FileExists(pathToLogoToFind) Then 
					Response.Write ("Client logo not found in clientFilesV, checking clientfiles (not virtual)<br>") 
					
					strippedbaseURL = Replace(baseURL,"https://","")
					strippedbaseURL = Replace(strippedbaseURL,"http://","")

					pathToLogoToFind = "C:\home\" & strippedbaseURL & "wwwroot\clientfiles\" & ClientKey & "\logos\logo.png"
				
					Response.Write("path & file to check: " & pathToLogoToFind & "<br>")
					
					If filesys.FileExists(pathToLogoToFind) Then 
					
						Response.Write("found logo file at: " & pathToLogoToFind & "<br>")
						Response.Write("copying: " & pathToLogoToFind & " to: " & "C:\home\clientfilesV\" & ClientKey & "\logos\logo.png" & "<br>")
						
						filesys.CopyFile pathToLogoToFind , "C:\home\clientfilesV\" & ClientKey & "\logos\logo.png"
						
					Else
					
						Response.Write("NO logo file found at: " & pathToLogoToFind & "<br>")
					
					End If


				End If

				
				'****************************************
				'prospecting/emailAttachments
				'****************************************
				

				pathToCurrentFolderToFind = "C:\home\clientfilesV\" & ClientKey & "\prospecting\"
				
				Response.Write("path to check: " & pathToCurrentFolderToFind & "<br>")
				
				If Not filesys.FolderExists(pathToCurrentFolderToFind) Then 
					newFolderToCreate = filesys.CreateFolder (pathToCurrentFolderToFind) 
					Response.Write ("A new folder for Client " & ClientKey & " <strong>(" & pathToCurrentFolderToFind & ")</strong> has been created.<br>") 
				Else
					Response.Write ("The folder <strong>(" & pathToCurrentFolderToFind & ")</strong> already exists for Client " & ClientKey & ".<br>")
				End If

				pathToCurrentFolderToFind = "C:\home\clientfilesV\" & ClientKey & "\prospecting\emailAttachments\"
				
				Response.Write("path to check: " & pathToCurrentFolderToFind & "<br>")
				
				If Not filesys.FolderExists(pathToCurrentFolderToFind) Then 
					newFolderToCreate = filesys.CreateFolder (pathToCurrentFolderToFind) 
					Response.Write ("A new folder for Client " & ClientKey & " <strong>(" & pathToCurrentFolderToFind & ")</strong> has been created.<br>") 
				Else
					Response.Write ("The folder <strong>(" & pathToCurrentFolderToFind & ")</strong> already exists for Client " & ClientKey & ".<br>")
				End If

		
				Response.Write("<br>End Client File Folder Synchronizing<br>")
				'*******************************************************************	

				
			End If
			
		End If
		
		Response.Write("<br>******** DONE Processing Client File Folder Synchronizing For " & ClientKey & "************<br>")
			
					
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
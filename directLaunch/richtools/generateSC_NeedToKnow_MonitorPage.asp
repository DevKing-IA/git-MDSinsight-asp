<!DOCTYPE html>
<html>
  <head>
    <meta charset="UTF-8">
    <title>NeedToKnow Last Run Monitoring Page</title>
	
	<!--#include file="../../inc/SubsAndFuncs.asp"-->
	<!--#include file="../../inc/InsightFuncs.asp"-->

	<!-- Latest compiled and minified CSS -->
	<link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.1.3/css/bootstrap.min.css" integrity="sha384-MCw98/SFnGE8fJT3GXwEOngsV7Zt27NXFoaoApmYm81iuXoPkFOJwJ8ERdknLPMO" crossorigin="anonymous">

	<!-- Latest compiled and minified JavaScript -->
	<script src="https://code.jquery.com/jquery-3.3.1.slim.min.js" integrity="sha384-q8i/X+965DzO0rT7abK41JStQIAqVgRVzpbzo5smXKp4YfRvH+8abtTE1Pi6jizo" crossorigin="anonymous"></script>
	<script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.14.3/umd/popper.min.js" integrity="sha384-ZMP7rVo3mIykV+2+9J3UJ46jBk0WLaUAdn689aCwoqbBJiSnjAK/l8WvCWPIPm49" crossorigin="anonymous"></script>
	<script src="https://stackpath.bootstrapcdn.com/bootstrap/4.1.3/js/bootstrap.min.js" integrity="sha384-ChfqqxuZUCnJSK3+MXmPNIyE6ZbWh2IMqE241rYiqJxyMiZ6OW/JmZQ5stwEULTy" crossorigin="anonymous"></script>

	<style>
				
		.table-xs {
		    width:544px;
		}
		
		.table-sm {
		    width: 576px;
		}
		
		.table-md {
		    width: 768px;
		}
		
		.table-lg {
		    width: 992px;
		}
		
		.table-xl {
		    width: 1200px;
		}
		
		/* Small devices (landscape phones, 544px and up) */
		@media (min-width: 576px) {  
		    .table-sm {
		        width: 100%;
		    }
		}
		
		/* Medium devices (tablets, 768px and up) The navbar toggle appears at this breakpoint */
		@media (min-width: 768px) {
		    .table-sm {
		        width: 100%;
		    }
		
		    .table-md {
		        width: 100%;
		    }
		}
		
		/* Large devices (desktops, 992px and up) */
		@media (min-width: 992px) {
		    .table-sm {
		        width: 100%;
		    }
		
		    .table-md {
		        width: 100%;
		    }
		
		    .table-lg {
		        width: 100%;
		    }
		}
		
		/* Extra large devices (large desktops, 1200px and up) */
		@media (min-width: 1200px) {
		    .table-sm {
		        width: 100%;
		    }
		
		    .table-md {
		        width: 100%;
		    }
		
		    .table-lg {
		        width: 100%;
		    }
		
		    .table-xl {
		        width: 100%;
		    }
		}		
	</style>
	
  </head>

<%
Server.ScriptTimeout = 7000
'Monitor the last udpated runtime for the SQL Table SC_NeedToKnow for Each Client
'Designed to be launched via a scheduled process (Win Task Scheduler)
'Self contained page will check the alerts db and take the appropriate actions
'Usage = "http://{xxx}.{domain}.com/directLaunch/richtools/generateSC_NeedToKnow_MonitorPage.asp?runlevel=run_now"

'The runlevel parameter is inconsequential to the operation 
'of the page. It is only used so that the page will not run
'if it is loaded via an unexpected method (spiders, etc)
If Request.QueryString("runlevel") <> "run_now" then response.end

If Request.QueryString("q") = 1 Then QuietMode = True Else QuietMode = False

'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)

'This single page loops through and handles autocompletes for all databases
SQL = "SELECT * FROM tblServerInfo WHERE Active = 1 ORDER BY ClientKey"

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

%>
  <body>

<%
If Not TopRecordset.Eof Then

	Do While Not TopRecordset.EOF

		PassPhrase = TopRecordset.Fields("directLaunchPassphrase")
		ClientKey = TopRecordset.Fields("clientkey")
		CompanyName = TopRecordset.Fields("CompanyName")
		
		'Response.Write("<br><br><br>******** START Processing Last Updated SC_NeedToKnow Modules For " & ClientKey  & "************<br><br>")
		
		Call SetClientCnnString
		
		Session("ClientCnnString") = MUV_READ("ClientCnnString") ' Until session vars are gone, then delete this
		If (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") = 0 AND Instr(ucase(ClientKey),"D") = 0)_
		or (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 AND Instr(ucase(ClientKey),"D") <> 0)_
		or (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"QB.") <> 0 AND Instr(ucase(ClientKey),"D") <> 0)_
		or (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"FL") <> 0 AND Instr(ucase(ClientKey),"D") <> 0)_
		or 1 = 1 Then 

		'Response.Write("Request.ServerVariables(SERVER_NAME):" & Request.ServerVariables("SERVER_NAME") & "<br><br>")
		'Response.Write("ClientKey:" & ClientKey & "<br><br>")
		
		If MUV_READ("cnnStatus") = "OK" Then ' else it loops and excludes dev client keys
			
			'*******************************************************************
			'Begin Processing Each Client
			'*******************************************************************
			 'Response.Write("Begin Processing Last Updated SC_NeedToKnow Modules<br><br>")
			'******************************************
			
			Server.ScriptTimeout = 500
			
	
			Set cnnSCNeedToKnow = Server.CreateObject("ADODB.Connection")
			cnnSCNeedToKnow.open (Session("ClientCnnString"))
			Set rsSCNeedToKnow = Server.CreateObject("ADODB.Recordset")
			rsSCNeedToKnow.CursorLocation = 3 
			
			
			Set cnnSCNeedToKnowMonitor = Server.CreateObject("ADODB.Connection")
			cnnSCNeedToKnowMonitor.open (Session("ClientCnnString"))
			Set rsSCNeedToKnowMonitor = Server.CreateObject("ADODB.Recordset")
			rsSCNeedToKnowMonitor.CursorLocation = 3 
			
			
			SQL_SCNeedToKnowMonitor = "SELECT Module, Submodule, MAX(RecordCreationDateTime) AS Expr1 "
			SQL_SCNeedToKnowMonitor = SQL_SCNeedToKnowMonitor & " FROM  SC_NeedToKnow "
			SQL_SCNeedToKnowMonitor = SQL_SCNeedToKnowMonitor & " GROUP BY Module, Submodule "
			SQL_SCNeedToKnowMonitor = SQL_SCNeedToKnowMonitor & " ORDER BY Module, Submodule "


			If QuietMode = False Then
				'Response.Write("<br><br><br>" & SQL_SCNeedToKnowMonitor & "<br>")
			End If
		
			Set rsSCNeedToKnowMonitor = cnnSCNeedToKnowMonitor.Execute(SQL_SCNeedToKnowMonitor)
			
			
			%>
			<div class="table-responsive-lg">
			<table class="table table-bordered table-lg"><%
			
			%>
			
			<thead class="thead-dark">	
				<tr class="d-flex">
			      <th scope="col" class="col-3">Client ID</th>
			      <th scope="col" class="col-3">Module</th>
			      <th scope="col" class="col-3">Submodule</th>
			      <th scope="col" class="col-3">Last Updated</th>
			    </tr>
		   </thead>
		  
		   <tbody>		
			<%
			
			If NOT rsSCNeedToKnowMonitor.EOF Then
			
				Do While NOT rsSCNeedToKnowMonitor.EOF
			
					N2K_Module = rsSCNeedToKnowMonitor("Module")
					N2K_Submodule = rsSCNeedToKnowMonitor("Submodule")
					N2K_LastUpdated = rsSCNeedToKnowMonitor("Expr1")
					
					%>
					
					<% If Abs(DateDiff("d",N2K_LastUpdated,Now())) > 2 Then %>
				    	<tr class="d-flex table-danger">
			      	<% Else %>
			      		<tr class="d-flex">
			      	<% End If %>
				    
				      <td class="col-3"><%= ClientKey %></td>
				      <td class="col-3"><%= N2K_Module %></td>
				      <td class="col-3"><%= N2K_Submodule %></td>
				      <td class="col-3"><%= N2K_LastUpdated %></td>
				      
				    </tr>					
					
					<%					
					rsSCNeedToKnowMonitor.MoveNext
				Loop
					
			End If
			
			%>
			
				</tbody>
			</table>
			</div>
			
			<br><hr><br>
			<%

			'Response.Write("<br>End Processing Last Updated SC_NeedToKnow Modules<br>")
			'*******************************************************************	
			
		End If
		
		'Response.Write("<br>******** DONE Processing Last Updated SC_NeedToKnow Modules For " & ClientKey & "************<br>")
		
	End If	
					
	TopRecordset.movenext
	
	Loop

	TopRecordset.Close
	Set TopRecordset = Nothing
	TopConnection.Close
	Set TopConnection = Nothing
	
End If




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

  
  </body>
</html>

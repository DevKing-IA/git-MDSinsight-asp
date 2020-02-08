<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs_AR_AP.asp"-->
<script type="text/javascript">
	function closeme() {
		window.open('', '_parent', '');
		window.close();  }
</script>
<%
Server.ScriptTimeout = 7000
'Populate the SQL Table AR_CustomerCounts for Each Client Script
'Designed to be launched via a scheduled process (Win Task Scheduler)
'Self contained page will check the alerts db and take the appropriate actions
'Usage = "http://{xxx}.{domain}.com/directLaunch/systemtools/populateAR_CustomerCounts_activeClientLoop.asp?runlevel=run_now"

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


If Not TopRecordset.Eof Then

	Do While Not TopRecordset.EOF

		PassPhrase = TopRecordset.Fields("directLaunchPassphrase")
		ClientKey = TopRecordset.Fields("clientkey")
		CompanyName = TopRecordset.Fields("CompanyName")
		
		Response.Write("<br><br><br>******** START Processing Populate the SQL Table AR_CustomerCounts For " & ClientKey  & "************<br><br>")
		
		Call SetClientCnnString
		
		Session("ClientCnnString") = MUV_READ("ClientCnnString") ' Until session vars are gone, then delete this
		
		If MUV_READ("cnnStatus") = "OK" Then ' else it loops and excludes dev client keys
			
			'*******************************************************************
			'Begin Client AR_CustomerCounts
			'*******************************************************************
			
			 Response.Write("Begin Populating the SQL Table AR_CustomerCounts<br><br>")
			'***********************************************************************
						
			Set cnnARCustomer = Server.CreateObject("ADODB.Connection")
			cnnARCustomer.open (Session("ClientCnnString"))
			Set rsARCustomer = Server.CreateObject("ADODB.Recordset")
			rsARCustomer.CursorLocation = 3 		
			
			numTotalAccounts = NumberOfARCustAccounts()
			numActiveAccounts = NumberOfActiveARCustAccounts()
			numInactiveAccounts = NumberOfInactiveARCustAccounts()
			
			SQL_ARCustomer = "INSERT INTO AR_CustomerCounts (numTotalAccounts, numActiveAccounts, numInactiveAccounts) "
			SQL_ARCustomer = SQL_ARCustomer & " VALUES (" & numTotalAccounts & "," & numActiveAccounts & "," & numInactiveAccounts & ")"
			Set rsARCustomer = cnnARCustomer.Execute(SQL_ARCustomer)	
			
			Response.Write("Client " & ClientKey & " has " & numTotalAccounts & " total accounts. " & numActiveAccounts & " are active and " & numInactiveAccounts & " are inactive.<br>")

			Set rsARCustomer = Nothing
			cnnARCustomer.Close
			Set cnnARCustomer = Nothing
					
			'***********************************************************************
			Response.Write("<br>End Populating the SQL Table AR_CustomerCounts<br>")
			'***********************************************************************	
			
		End If
		
		Response.Write("<br>******** DONE Processing Populate the SQL Table AR_CustomerCounts For " & ClientKey & "************<br>")
			
					
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
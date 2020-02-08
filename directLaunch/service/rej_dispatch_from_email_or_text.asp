<% @ Language = VBScript %>
<%
Response.Buffer = True
Response.Expires = 0
Response.Clear
%>
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/settings.asp"-->

<%
sURL = Request.ServerVariables("SERVER_NAME")
'baseURL should alwats have a trailing /slash, just in case, handle either way

ServiceTicketNumber = Request.QueryString("t")
UserNumber = Request.QueryString("u")
CustNum = Request.QueryString("c")
ClientKey =  Request.QueryString("cl")

SQL = "SELECT * FROM tblServerInfo where clientKey='"& ClientKey &"'"

Set Connection = Server.CreateObject("ADODB.Connection")
Set Recordset = Server.CreateObject("ADODB.Recordset")
Connection.Open InsightCnnString
'Response.Write("InsightCnnString:" & InsightCnnString & "<br>")

'Open the recordset object executing the SQL statement and return records
Recordset.Open SQL,Connection,3,3
'Response.Write("SQL:" & SQL& "<br>")

'First lookup the ClientKey in tblServerInfo
'If there is no record with the entered client key, close connection
'and go back to login with QueryString
If Recordset.recordcount <= 0 then
	Recordset.close
	Connection.close
	set Recordset=nothing
	set Connection=nothing
	%>MDS Insight: Unable to connect to SQL database. The server is not available or the credentials specified are incorrect. - Please contact your administrator.<%
	Response.End
Else
	Session("ClientCnnString") = "Driver={SQL Server};Server=" & Recordset.Fields("dbServer")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Database=" & Recordset.Fields("dbCatalog")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Uid=" & Recordset.Fields("dbLogin")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Pwd=" & Recordset.Fields("dbPassword") & ";"
	dummy = MUV_Write("SQL_Owner",Recordset.Fields("dbLogin"))
	dummy = MUV_Write("ClientID",Recordset.Fields("clientkey"))
	Recordset.close
	Connection.close	
End If	


'Response.Write("ServiceTicketNumber :" & ServiceTicketNumber & "<br>")
'Response.Write("UserNumber :" & UserNumber & "<br>")
'Response.Write("CustNum :" & CustNum & "<br>")
'Response.End

If ServiceTicketNumber = "" Then
	%>MDS Insight is unable to decline this dispatch due to a blank service ticket id. Please contact your administrator.<% 
	Response.End
End If
If UserNumber= "" Then
	%>MDS Insight is unable to decline this dispatch due to a blank user number. Please contact your administrator.<% 
	Response.End
End If
If CustNum = "" Then
	%>MDS Insight is unable to decline this dispatch due to a blank customer id. Please contact your administrator.<% 
	Response.End
End If

'Only insert the acknowledgement record if it isn't there already
Set cnnDispatch = Server.CreateObject("ADODB.Connection")
cnnDispatch.open (Session("ClientCnnString"))
Set rsDispatch = Server.CreateObject("ADODB.Recordset")

SQLDispatch = "SELECT * FROM FS_ServiceMemosDetail WHERE "
SQLDispatch = SQLDispatch & "MemoNumber = '"  & ServiceTicketNumber & "' AND "
SQLDispatch = SQLDispatch & " MemoStage ='Dispatch Declined' AND "
SQLDispatch = SQLDispatch & "UserNoOfServiceTech = "  & UserNumber 

Set rsDispatch = cnnDispatch.Execute(SQLDispatch)
'response.write(SQLDispatch)
If rsDispatch.Eof Then

	SQLDispatch = "INSERT INTO FS_ServiceMemosDetail (MemoNumber, CustNum, MemoStage, "
	SQLDispatch = SQLDispatch & " SubmissionDateTime, USerNoSubmittingRecord,UserNoOfServiceTech,OriginalDispatchDateTime)"
	SQLDispatch = SQLDispatch &  " VALUES (" 
	SQLDispatch = SQLDispatch & "'"  & ServiceTicketNumber & "'"
	SQLDispatch = SQLDispatch & ",'"  & CustNum & "'"
	SQLDispatch = SQLDispatch & ",'Dispatch Declined'"
	SQLDispatch = SQLDispatch & ",getdate() "
	SQLDispatch = SQLDispatch & ","  & UserNumber 
	SQLDispatch = SQLDispatch & ","  & GetServiceTicketDispatchedTech(ServiceTicketNumber )
	SQLDispatch = SQLDispatch & ", '" & TicketOriginalDispatchDateTime(ServiceTicketNumber) & "')"
	Set rsDispatch = cnnDispatch.Execute(SQLDispatch)
	
End If

Set rsDispatch = Nothing
cnnDispatch.Close
Set cnnDispatch = Nothing

'Write audit trail for dispatch
'*******************************
Description = GetUserDisplayNameByUserNo(UserNumber) & " declined the dispatch for service ticket number " & ServiceTicketNumber & " at " & NOW()
CreateAuditLogEntry "Service Ticket System","Dispatched","Minor",0,Description 

dummy = Redispatch(ServiceTicketNumber)
Description = "Service ticket #" & ServiceTicketNumber & " was set for redispatch due to being set to dispatch declined at " & NOW()
CreateAuditLogEntry "Service Ticket System","Redispatch","Minor",0,Description 


Response.Write("<font color='blue' size='6'><center><br><br><br><br>Your decline has been recorded.<br><br><br><br> Thank You.</font></center>")
Response.End
%>
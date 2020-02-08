<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<% If Session("UserNo") = "" Then Response.End() %>
<%
Response.ContentType = "application/json"

ServiceTicketID = Request.QueryString("serviceTicketID") 
InternalRecordIdentifier = Request.QueryString("i")

If ServiceTicketID = "" Then Response.End()

Set cnnServiceTicketNote = Server.CreateObject("ADODB.Connection")
cnnServiceTicketNote.open (Session("ClientCnnString"))
Set rsServiceTicketNote = Server.CreateObject("ADODB.Recordset")
rsServiceTicketNote.CursorLocation = 3 


If Request.Form("updateAction")="save" Then

	'***************************************************************************************
	'Lookup the record as it exists now so we can fillin the audit trail
	'***************************************************************************************
	SQL = "SELECT * FROM FS_ServiceMemosNotes WHERE InternalRecordIdentifier='" & Request.Form("updateActionId") & "'"		
	Set cnnServiceTicketNote8 = Server.CreateObject("ADODB.Connection")
	cnnServiceTicketNote8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnnServiceTicketNote8.Execute(SQL)
		
	If not rs.EOF Then
		Orig_Note = rs("Note")
	End If

	'***************************************************************************************
	'After SQL update, record entries in audit trail
	'***************************************************************************************
	
	Note = Request.Form("LogNote")
	ServiceTicketID = Request.Form("updateServiceTicketID")
	CustID = GetServiceTicketCust(ServiceTicketID)
	CustomerName = GetCustNameByCustNum(CustID)
	UserNo = Session("UserNo")
	UserName = GetUserDisplayNameByUserNo(UserNo)

	If Orig_Note  <> Note Then
	
		Description = "The service ticket note changed from " & Orig_Note  & " to " & Note  & ", by " & UserName & " on " & FormatDateTime(Now(),2) & " for ticket #"
		Description = Description & ServiceTicketID & ", for customer " & CustomerName & "(" & CustID & ")."
		
		CreateAuditLogEntry "Service ticket note edited ",GetTerm("Service"),"Minor",0,Description

	End If
	
	Query = "UPDATE FS_ServiceMemosNotes SET Note='"&EscapeSingleQuotes(Request.Form("LogNote"))&"' WHERE InternalRecordIdentifier='"&Request.Form("updateActionId")&"'"
	
	Set rsServiceTicketNote = cnnServiceTicketNote.Execute(Query)

End If





If Request.Form("updateAction")="insert" Then

	Note = Request.Form("LogNote")
	ServiceTicketID = Request.Form("updateServiceTicketID")
	CustID = GetServiceTicketCust(ServiceTicketID)
	CustomerName = GetCustNameByCustNum(CustID)
	UserNo = Session("UserNo")
	UserName = GetUserDisplayNameByUserNo(UserNo)
		
	Description = "The service ticket note " & Note  & ", was created by " & UserName & " on " & FormatDateTime(Now(),2) & " for ticket #"
	Description = Description & ServiceTicketID & ", for customer " & CustomerName & "(" & CustID & ")."

	CreateAuditLogEntry "Service ticket note created ",GetTerm("Service"),"Minor",0,Description

	Query = "INSERT INTO FS_ServiceMemosNotes (ServiceTicketID, EnteredByUserNo, Note) "
	Query = Query & "VALUES ('" & ServiceTicketID & "'," & Session("Userno") & ",'" & EscapeSingleQuotes(Request.Form("LogNote")) & "')"
	
	'Response.Write(Query)
	
	Set rsServiceTicketNote = cnnServiceTicketNote.Execute(Query)
	
End If




If Request.Form("updateAction")="delete" Then

	SQL = "SELECT * FROM FS_ServiceMemosNotes WHERE InternalRecordIdentifier='"&Request.Form("updateActionId")&"'"		
	Set cnnServiceTicketNote8 = Server.CreateObject("ADODB.Connection")
	cnnServiceTicketNote8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnnServiceTicketNote8.Execute(SQL)
		
	If not rs.EOF Then
		InternalRecordIdentifier = rs("InternalRecordIdentifier")
		RecordCreationDateTime = rs("RecordCreationDateTime")
		ServiceTicketID = rs("ServiceTicketID")
		CustID = GetServiceTicketCust(ServiceTicketID)
		CustomerName = GetCustNameByCustNum(CustID)
		EnteredByUserNo	= rs("EnteredByUserNo")
		Note = rs("Note")
		UserName = GetUserDisplayNameByUserNo(rs("EnteredByUserNo"))
	End If
	
	set rs = Nothing
	cnnServiceTicketNote8.close
	set cnnServiceTicketNote8 = Nothing
	
	Description = "The service ticket note " & Note  & ", originally entered by " & UserName & " on " & FormatDateTime(RecordCreationDateTime,2) & " for ticket #"
	Description = Description & ServiceTicketID & ", for customer " & CustomerName & "(" & CustID & ") was deleted."
	
	CreateAuditLogEntry "Service ticket note deleted ",GetTerm("Service"),"Minor",0,Description

	Query = "DELETE FROM FS_ServiceMemosNotes WHERE InternalRecordIdentifier='"&Request.Form("updateActionId")&"'"
	
	Set rsServiceTicketNote = cnnServiceTicketNote.Execute(Query)
	
End If


Query = "SELECT InternalRecordIdentifier, RecordCreationDateTime, ServiceTicketID, EnteredByUserNo, Note "
Query = Query & "FROM FS_ServiceMemosNotes WHERE ServiceTicketID = '"& Request.Form("updateServiceTicketID") & "' "
Query = Query & " ORDER BY RecordCreationDateTime DESC"

Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnnServiceTicketNote.Execute(Query)

Response.Write("[")
If not rs.EOF Then
	sep = ""
	Do While Not rs.EOF			
			Response.Write(sep)
			sep = ","
			Response.Write("{")
			Response.Write("""id"":""" & EscapeQuotes(rs("InternalRecordIdentifier")) & """")
			Response.Write(",""Date"":""" & EscapeQuotes(FormatDateTime(rs("RecordCreationDateTime"),2)) & """")
			Response.Write(",""Time"":""" & EscapeQuotes(FormatDateTime(rs("RecordCreationDateTime"),3)) & """")
			Response.Write(",""User"":""" & EscapeQuotes(GetUserDisplayNameByUserNo(rs("EnteredByUserNo"))) & """")
		
			If rs("Note") = "" OR IsNull(rs("Note")) OR IsEmpty(rs("Note")) Then
				Response.Write(",""LogNote"":""" & EscapeNewLine(EscapeQuotes(rs("Note"))) & """")
			Else
				Response.Write(",""LogNote"":""" & EscapeNewLine(EscapeQuotes(rs("Note"))) & """")
			End If
			Response.Write("}")

		rs.MoveNext						
	Loop
End If
Response.Write("]")

cnnServiceTicketNote.Close
Set cnnServiceTicketNote = Nothing

Function EscapeNewLine(val)
	EscapeNewLine = Replace(val, vbcr, "\r")
	EscapeNewLine = Replace(EscapeNewLine, vblf, "\n")
End Function
Function EscapeQuotes(val)
	EscapeQuotes = Replace(val, """", "\""")
End Function
Function EscapeSingleQuotes(val)
	EscapeSingleQuotes = Replace(val, "'", "''")
End Function

%> 
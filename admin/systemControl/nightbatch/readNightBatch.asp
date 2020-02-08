<!--#include file="../../../inc/header.asp"-->
<!--#include file="../../../inc/mail.asp"-->
<%
'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)
MUV_REMOVE("CRONQUERY") ' To be used below
'***********************************
'Post the night batch query to UNIX
'***********************************
data = "<DATASTREAM>"
data = data & "<IDENTITY>Pm8316wyc011</IDENTITY>"
data = data & "<MODE>" & GetPOSTParams("Mode") & "</MODE>"
data = data & "<RECORD_TYPE>NIGHTBATCH</RECORD_TYPE>"
data = data & "<RECORD_SUBTYPE>CRONQUERY</RECORD_SUBTYPE>"
data = data & "<SERNO>" & GetPOSTParams("Serno") & "</SERNO>"
data = data & "</DATASTREAM>"

Description = "Post to " & GetPOSTParams("CUSTOMERURL1")
CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("Mode")

Description = "data:" & data 
CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("Mode")

Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")

httpRequest.Open "POST", GetPOSTParams("CUSTOMERURL1"), False
httpRequest.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
httpRequest.Send data
	
IF httpRequest.status = 200 THEN 

	If Instr(httpRequest.responseText,"success") <> 0 Then ' Success
	
		Description ="success! httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>NIGHTBATCH and <RECORD_SUBTYPE>CRONQUERY "& "<br>"
		Description = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
		Description = Description & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
		Description = Description & "Posted to " & GetPOSTParams("CUSTOMERURL1") & "<br>"
		Description = Description & "POSTED DATA:" & data & "<br>"
		Description = Description & "SERNO:" & MUV_READ("ClientID") & "<br>"
		
		CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("Mode")
		
		'*****************************************
		'Kind of a different way to do things
		'Set the results in MUV CRONQUERY and the
		'next page will grab the values from there
		'******************************************
		tmpvar = httpRequest.responseText
		tmpvar=ucase(tmpvar)
		tmpvar = Replace(tmpvar,"SUCCESS","")
		tmpvar = trim(tmpvar)
		dummy = MUV_WRITE("CRONQUERY",tmpvar) 		
		Response.Redirect("main.asp")
	Else
		'FAILURE
		emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>NIGHTBATCH and <RECORD_SUBTYPE>CRONQUERY "& "<br>"
		emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
		emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
		emailBody = emailBody & "Posted to " & GetPOSTParams("CUSTOMERURL1") & "<br>"
		emailBody = emailBody & "POSTED DATA:" & data & "<br>"
		emailBody = emailBody & "SERNO:" & MUV_READ("ClientID") & "<br>"
		SendMail "mailsender@" & maildomain ,"rich@ocsaccess.com",MUV_READ("ClientID") & " NIGHT BATCH CRONQUERY POST ERROR",emailBody, "Night Batch", "Post Failure"
	
		Description = emailBody 
		CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("Mode")
		
		'Go to main with a Querystring, indicationg a read error
		'Doesn't matter what the value is. Any querystring being present
		'indicates a read error
		Response.Redirect("main.asp?s=0")
		
		'This code doesn't execute anymore
		Response.Write("<br><br><br><br><br><br><strong><center>Unable to read night batch settings from " & GetTerm("Backend") & "</center></strong>")
		Response.Write("<br><br><br><strong><center>Please try again and contact techincal support if the problem continues.</center></strong>")
	End If
	
Else
	
	emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>NIGHTBATCH and <RECORD_SUBTYPE>CRONQUERY "& "<br>"
	emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
	emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
	emailBody = emailBody & "Posted to " & GetPOSTParams("CUSTOMERURL1") & "<br>"
	emailBody = emailBody & "POSTED DATA:" & data & "<br>"
	emailBody = emailBody & "SERNO:" & MUV_READ("ClientID") & "<br>"
	SendMail "mailsender@" & maildomain ,"projects@metroplexdata.com",MUV_READ("ClientID") & " NIGHT BATCH CRONQUERY POST ERROR",emailBody, "Night Batch", "Post Failure"

	Description = emailBody 
	CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("Mode")

	'Go to main with a Querystring, indicationg a read error
	'Doesn't matter what the value is. Any querystring being present
	'indicates a read error
	Response.Redirect("main.asp?s=0")

	Response.Write("<br><br><br><br><br><br><strong><center>Unable to read night batch settings from " & GetTerm("Backend") & "</center></strong>")
	Response.Write("<br><br><br><strong><center>Please try again and contact techincal support if the problem continues.</center></strong>")

End If


%><!--#include file="../../../inc/footer-main.asp"-->
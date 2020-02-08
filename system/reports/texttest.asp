<!--#include file="../../inc/header.asp"-->


<% 

Dim sUrl
Dim sAPI_ID, sPassword, sUsername, sMobileNo, sText
Dim oXMLHTTP, sPostData, sResult
sUrl = "http://api.clickatell.com/http/sendmsg"
sAPI_ID = "3543646"
sPassword = "JDBUaRfGODKKgc"
sUsername = "OCSAccess"
sMobileNo = "16099294430"
sText = "This is an example message"
sPostData = "api_id=" & sAPI_ID
sPostData = sPostData & "&user=" & sUsername
sPostData = sPostData & "&password=" & sPassword
sPostData = sPostData & "&to=" & sMobileNo
sPostData = sPostData & "&text=" & sText
Set oXMLHTTP = Server.CreateObject("Microsoft.XMLHTTP")
oXMLHTTP.Open "POST", sUrl, false
oXMLHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
oXMLHTTP.Send sPostData
sResult = oXMLHTTP.responseText
Set oXMLHTTP = nothing
Response.Write sResult

	    
	    
%><!-- eof row !--><!--#include file="../../inc/footer-main.asp"-->
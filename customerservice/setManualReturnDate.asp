<!--#include file="../inc/subsandfuncs.asp"-->
<!--#include file="../inc/InsightFuncs.asp"-->

<%


CstNum = Request.Form("Cstnum")
rtdte = Request.Form("rtdte")
Reason  = Request.Form("selReturnDateReason")
If Reason = "" Then Reason = 0

Response.Write("CstNum :" & CstNum & "<br>")
Response.Write("rtdte :" & rtdte & "<br>")
Response.Write("Reason  :" & Reason  & "<br>")
'Response.Write("ServiceTicketNumber :" & ServiceTicketNumber & "<br>")
'Response.Write("BaseURL :" & BaseURL & "<br>")
Response.End

If  Month(rtdte) < 10 Then
	dteHold = "0" & Month(rtdte) & "/"
Else
	dteHold = Month(rtdte) & "/"
End IF

If  Day(rtdte) < 10 Then
	dteHold = dteHold  & "0" & Day(rtdte) & "/"
Else
	dteHold = dteHold  & Day(rtdte) & "/"
End IF
dteHold = dteHold & Right(rtdte,2)
rtdte = dteHold

Set cnnrteDte = Server.CreateObject("ADODB.Connection")
cnnrteDte.open (Session("ClientCnnString"))

Set rsrteDte = Server.CreateObject("ADODB.Recordset")
rsrteDte.CursorLocation = 3 

SQL = "Select ReturnDate from AR_Customer WHERE CustNum= " & CstNum
Set rsrteDte = cnnrteDte.Execute(SQL)
If not rsrteDte.Eof Then ORDte = rsrteDte("ReturnDate")

SQL = "UPDATE AR_Customer Set ReturnDate = '" & rtdte & "' WHERE CustNum= " & CstNum
Set rsrteDte = cnnrteDte.Execute(SQL)

set rsrteDte = Nothing
cnnrteDte.Close
Set cnnrteDte = Nothing

Description = ""
Description = Description & "The return date was manually changed for account # "  & CstNum 
Description = Description & "     The new return date is: "  & rtdte
Description = Description & "     The original return date was: "  & ORDte
 
CreateAuditLogEntry "Return Date Changed","Return Date Changed","Minor",0,Description

'Post to MDS goes here
data = "<DATASTREAM>"
data = data & "<IDENTITY>Pm8316wyc011</IDENTITY>"
data = data & "<MODE>" & GetPOSTParams("Mode") & "</MODE>"
data = data & "<RECORD_TYPE>UPDATE_CUSTOMER</RECORD_TYPE>"
data = data & "<RECORD_SUBTYPE>RETURN_DATE</RECORD_SUBTYPE>"
data = data & "<CLIENT_ID>CCS</CLIENT_ID>"
data = data & "<SERNO>" & GetPOSTParams("SERNO") & "</SERNO>"
data = data & "<SUBMISSION_SOURCE>MDS Insight</SUBMISSION_SOURCE>"
data = data & "<ACCOUNT_NUM>" & CstNum & "</ACCOUNT_NUM>"
data = data & "<FIELD_DATA>" & rtdte & "</FIELD_DATA>"
data = data & "<FIELD_DATA1>" & Reason & "</FIELD_DATA1>"
data = data & "</DATASTREAM>"

Description = "Post to " & GetPOSTParams("CustomerURL1")
CreateINSIGHTAuditLogEntry sURL,Description,GetPOSTParams("Mode")
Description = "data:" & data 
CreateINSIGHTAuditLogEntry sURL,Description,GetPOSTParams("Mode")

Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
httpRequest.Open "POST", GetPOSTParams("CustomerURL1"), False
httpRequest.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
httpRequest.Send data
	
IF httpRequest.status = 200 THEN 
	Description = "httpRequest.responseText:" & httpRequest.responseText
	CreateINSIGHTAuditLogEntry sURL,Description,GetPOSTParams("Mode")
	postResponse = httpRequest.responseText
ELSE
	'In here it must email us if there are problems
	Description = "httpRequest.responseText:" & httpRequest.responseText
	CreateINSIGHTAuditLogEntry sURL,Description,GetPOSTParams("Mode")
	postResponse= "Could not get data to metroplex."
END IF
	
If postResponse <> "success" then 
	postResponse = httpRequest.responseText
	'In here it must email us if there are problems
End If
	

%>

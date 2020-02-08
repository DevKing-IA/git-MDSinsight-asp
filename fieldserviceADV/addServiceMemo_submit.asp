<!--#include file="../inc/header-field-service-mobile.asp"-->

<%

AccountNumber = Request.Form("txtAccount")
Company = Request.Form("txtCompany")
ContactName = Request.Form("txtContactName")
ContactPhone = Request.Form("txtContactPhone")
ContactEmail = Request.Form("txtContactEmail")
ProblemLocation = Request.Form("txtLocation")
ProblemDescription = Request.Form("txtDescription")
MemoType = Request.Form("selMemoType")
SumissionSource = "Insight Field Service WebApp"
sURL = Request.ServerVariables("SERVER_NAME")


' Write Audit trail first, then post
Set cnn8 = Server.CreateObject("ADODB.Connection")

cnn8.open (Session("ClientCnnString"))
Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
'Set rs8 = cnn8.Execute(SQL)

set rs8 = Nothing

Description = ucase(MemoType) & ",    "
Description = Description & "Account: "  & AccountNumber & " - " & Company 
Description = Description & ",     Location: "  & ProblemLocation 
Description = Description & ",    Description: "  & ProblemDescription 

CreateAuditLogEntry "Service Memo Added","Service Memo Added","Major",0,Description


'Post to MDS goes here
postmessage=Replace(ProblemDescription," ","%20")
postmessage= Replace(postmessage,"&","%26") 
data = "create_service_request&account="
data = data & AccountNumber
data = data & "&prob=" &  postmessage & "%0A%0D" 
data = data & "Opened by: " & 	MUV_Read("DisplayName") & " - " & Session("UserEmail") & "%0A%0D"
data = data & "&probloc=" & ProblemLocation
data = data & "&sbn=" & ContactName 
data = data & "&sbe=" & ContactEmail 
data = data & "&sbp=" & ContactPhone 
data = data & "&contnm=" & ContactName
data = data & "&st=" & MemoType
data = data & "&md=" & GetPOSTParams("Mode")
data = data & "&serno=" & GetPOSTParams("Serno")
data = data & "&src=" & SumissionSource
data = data & "&usr=" & Session("userNo")
data = data & "&egcy=0" ' 0 - No 1 - Yes
If GetPOSTParams("NeverPutOnHold") = 0  Then data = data & "&hld=1" Else data = data & "&hld=0" 'I know it's wierd. It's the opposite of how it is stored in the table


For x = 1 to 2

		DO_Post = 0
		
		If x = 1 Then Do_Post = GetPOSTParams("ServiceMemoURL1ONOFF") 
		If x = 2 Then Do_Post = GetPOSTParams("ServiceMemoURL2ONOFF") 
		
		If IsNull(Do_Post) or Do_Post = "" Then Do_Post = 0
	
		If cint(Do_Post) = 1 Then
		
				CreateINSIGHTAuditLogEntry sURL,"Post Loop "& x,GetPOSTParams("Mode")

				If x = 1 Then
					Description = "Post to " & GetPOSTParams("ServiceMemoURL1")
				Else
					Description = "Post to " & GetPOSTParams("ServiceMemoURL2")
				End If

				CreateINSIGHTAuditLogEntry sURL,Description,GetPOSTParams("Mode")
				Description = "data:" & data 
				CreateINSIGHTAuditLogEntry sURL,Description,GetPOSTParams("Mode")
				
				Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
				
				If x = 1 Then
					httpRequest.Open "POST", GetPOSTParams("ServiceMemoURL1"), False
				Else
					httpRequest.Open "POST", GetPOSTParams("ServiceMemoURL2"), False				
				End If

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
		End If
Next 
Response.Redirect("addServiceMemo_ThankYou.asp")
	
%><!--#include file="../inc/footer-field-service-noTimeout.asp"-->





<!--#include file="../inc/header-field-service-mobile.asp"-->
<!--#include file="../inc/mail.asp"-->

<%
Set Upload = Server.CreateObject("Persits.Upload.1")
Upload.OverwriteFiles = False
Upload.Save
SelectedMemoNumber = Upload.Form("txtTicketNumber")

' Use AspJpeg to resize image
Set Jpeg = Server.CreateObject("Persits.Jpeg")
jpeg.PreserveAspectRatio = True

'Rename the files
' Construct the save path
Pth ="../clientfiles/" & trim(GetPOSTParams("Serno")) & "/SvcMemoPics/"

x =1

For Each File in Upload.Files

   File.SaveAsVirtual  Pth & SelectedMemoNumber & "-" & x & File.Ext
   
   Jpeg.Open File.Path
   
   ' New width for thumbnails
	L = 150
	
	If Jpeg.OriginalWidth > Jpeg.OriginalHeight Then
	   Jpeg.Width = L
	Else
	   Jpeg.Height = L
	End If

   Jpeg.Save Server.MapPath(Pth & SelectedMemoNumber & "-" & x & "-thumb" & File.Ext)
   
   x=x+1
   
Next

Account = GetServiceTicketCust(SelectedMemoNumber)
ServiceNotes = Upload.Form("ServiceNotes")

'Might come from a dropdown or typed in
If Upload.Form("txtAssetTagNumber")<> "" Then AssetTagNumber = Upload.Form("txtAssetTagNumber")
If Upload.Form("selAssetID")<> "" Then AssetTagNumber = Upload.Form("selAssetID")

AssetLocation = Upload.Form("txtAssetLocation")
sURL = Request.ServerVariables("SERVER_NAME")
PrintedName =  Upload.Form("txtPrintedName")



'******************************************************************
'Lookup Service Ticket & See If It Has Completed Filters
'******************************************************************

'For Each Item in Upload.Form
	'Response.Write Item.Name & "= " & Item.Value & "<BR>"
'Next

Set cnnFilter = Server.CreateObject("ADODB.Connection")
cnnFilter.open (Session("ClientCnnString"))
Set rsFilter = Server.CreateObject("ADODB.Recordset")
rsFilter.CursorLocation = 3 
	
SQL = "SELECT * FROM FS_ServiceMemosFilterInfo WHERE (ServiceTicketID = '" & SelectedMemoNumber & "')"
'Response.Write(SQL & "<br>")

Set rsFilter = cnnFilter.Execute(SQL)

If NOT rsFilter.EOF Then

	DO WHILE NOT rsFilter.EOF
	
		 InternalRecordIdentifier = rsFilter("InternalRecordIdentifier") 
		 ServiceTicketID = rsFilter("ServiceTicketID")
		 CustFilterIntRecID = rsFilter("CustFilterIntRecID") 
		 ICFilterIntRecID = rsFilter("ICFilterIntRecID")
		
		Filter_Rec_IDs = Filter_Rec_IDs & InternalRecordIdentifier & ","
		
		'*************************************************************
		'GET VALUE OF COMPLETE CHECKBOX FOR CURRENT FILTER FROM FORM
		
		 chkCurrentFilterComplete = Upload.Form("chkComplete" & InternalRecordIdentifier)
		 
		 'Response.Write("checkbox name : chkComplete" & InternalRecordIdentifier & "<br>")
		 'Response.Write("checkbox value : " & chkCurrentFilterComplete & "<br>")
		 '*************************************************************
		 
		Set cnnFilterUpdate = Server.CreateObject("ADODB.Connection")
		cnnFilterUpdate.open (Session("ClientCnnString"))
		Set rsFilterUpdate = Server.CreateObject("ADODB.Recordset")
		rsFilterUpdate.CursorLocation = 3 
		
		'******************************************************************************
		'IF THE FILTER IS CHECKED AS COMPLETE, UPDATE FS_SERVICEMEMOSFILTERINFO
		
		If chkCurrentFilterComplete = "1" OR chkCurrentFilterComplete = "on" Then

			SQLFilterUpdate = "UPDATE FS_ServiceMemosFilterInfo SET Completed=1, CompletedDate = GetDate(), CompletedByUserNo = " & Session("UserNo") & " WHERE (InternalRecordIdentifier = " & InternalRecordIdentifier & ")"
			Set rsFilterUpdate = cnnFilterUpdate.Execute(SQLFilterUpdate)
			'Response.Write("SQLFilterUpdate: " & SQLFilterUpdate& "<br>")
		
		End If
		'******************************************************************************
		 
	rsFilter.MoveNext
	Loop

End If 
'******************************************************************



If Right(Filter_Rec_IDs,1) = "," Then Filter_Rec_IDs= Left(Filter_Rec_IDs,Len(Filter_Rec_IDs)-1)

'**************************************
'This is the post of the service ticket
'**************************************
	
If GetPOSTParams("ServiceMemoURL1ONOFF")  = 1 Then

		CreateINSIGHTAuditLogEntry sURL,"Post Loop "& x,GetPOSTParams("Mode")

		If GetPOSTParams("SERVICEMEMOURL1MPLEXFORMAT") = 1 then ' see if using metroplex format

			CreateINSIGHTAuditLogEntry sURL,"Metroplex post format",GetPOSTParams("Mode")

			ServiceNotes = Replace(ServiceNotes,"&","%26") 
			ServiceNotes = Replace(ServiceNotes," ","%20") & "%0A%0D"
			
			If AssetTagNumber <> "" And AssetLocation <> "" Then
				ServiceNotes = ServiceNotes & "Asset Location updated. Asset Tag#: " & AssetTagNumber & "  New Location:  " & AssetLocation 
				ServiceNotes = Replace(ServiceNotes,"&","%26") 
				ServiceNotes = Replace(ServiceNotes," ","%20") & "%0A%0D"
			End If
			
			If PrintedName <> "" Then
				ServiceNotes = ServiceNotes & "Printed Name: " & PrintedName 
				ServiceNotes = Replace(ServiceNotes,"&","%26") 
				ServiceNotes = Replace(ServiceNotes," ","%20") & "%0A%0D"
			End If
			
			data = "create_service_request&account="
			data = data & Account
			data = data & "&st=CLOSE"
			data = data & "&usr=" & Session("UserNo")
			data = data & "&tnum=" & SelectedMemoNumber 
			data = data & "&prob=" 
			data = data & CloseOrCancelNotes & "%0A%0D"
			data = data & "Closed by: "  & 	GetUserDisplayNameByUserNo(Session("UserNo")) & " - " & GetUserEmailByUserNo(Session("UserNo")) & "%0A%0D"
			data = data & "Service Notes: "  &  ServiceNotes & "%0A%0D"
			data = data & "&md=" &  GetPOSTParams("Mode")
			data = data & "&serno="  & GetPOSTParams("Serno")
			data = data & "&src=Insight Field Service WebApp"
			data = data & "&rids=" & Filter_Rec_IDs 

			


		Else ' Regular XML format

			'Post to APIs goes here
			
			data = ""
			
			data = data & "<POST_DATA>"
			data = data & "<IDENTITY>Pm8316wyc011</IDENTITY>"
			data = data & "<ACCOUNT_NUM>" & Account & "</ACCOUNT_NUM>"
			data = data & "<PROBLEM_DESCRIPTION>" & CloseOrCancelNotes & "</PROBLEM_DESCRIPTION>"
			data = data & "<RECORD_TYPE>SENDSERVICEMSG</RECORD_TYPE>"
			data = data & "<RECORD_SUBTYPE>CLOSE</RECORD_SUBTYPE>"
			data = data & "<SERVICE_TICKET_NUMBER>" & SelectedMemoNumber & "</SERVICE_TICKET_NUMBER>"
			data = data & "<SERVICE_NOTES>" & ServiceNotes & "</SERVICE_NOTES>"
			data = data & "<SUBMISSION_SOURCE>Insight Field Service WebApp</SUBMISSION_SOURCE>"
			data = data & "<SERNO>" & GetPOSTParams("Serno") & "</SERNO>"
			data = data & "<CLIENT_ID>" & MUV_READ("ClientID") & "</CLIENT_ID>"
			data = data & "<USER_NO>" & Session("UserNo") & "</USER_NO>"
			data = data & "<MODE>" & GetPOSTParams("Mode") & "</MODE>"
			data = data & "</POST_DATA>"
				
		End IF

		Description = "Post to " & GetPOSTParams("ServiceMemoURL1")
		
		CreateINSIGHTAuditLogEntry sURL,Description,GetPOSTParams("Mode")
		Description = "data:" & data 
		CreateINSIGHTAuditLogEntry sURL,Description,GetPOSTParams("Mode")
		
		Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
		
		httpRequest.Open "POST", GetPOSTParams("ServiceMemoURL1"), False
		
		'httpRequest.SetRequestHeader "Content-Type", "text/xml"
		httpRequest.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"

		httpRequest.Send data
			
		If httpRequest.status = 200 THEN 
		
			If Instr(httpRequest.responseText,"success") <> 0 Then ' Success
			
				Description ="success! httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>SENDSERVICEMSG and <RECORD_SUBTYPE>CANCEL"& "<br>"
				Description = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
				Description = Description & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
				If x = 1 Then
					Description = Description & "Posted to " & GetPOSTParams("ServiceMemoURL1") & "<br>"
				Else
					Description = Description & "Posted to " & GetPOSTParams("ServiceMemoURL2") & "<br>"
				End IF
				Description = Description & "POSTED DATA:" & data & "<br>"
				Description = Description & "SERNO:" & GetPOSTParams("Serno") & "<br>"
				Description = Description & "MODE:" & GetPOSTParams("Mode") & "<br>"
		
				CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("Mode")

			Else
				'FAILURE
				emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>SENDSERVICEMSG and <RECORD_SUBTYPE>CANCEL"& "<br>"
				emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
				emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
				If x = 1 Then
					emailBody = emailBody & "Posted to " & GetPOSTParams("ServiceMemoURL1") & "<br>"
				Else
					emailBody = emailBody  & "Posted to " & GetPOSTParams("ServiceMemoURL2") & "<br>"
				End IF
				emailBody = emailBody & "POSTED DATA:" & data & "<br>"
				emailBody = emailBody & "SERNO:" & GetPOSTParams("Serno") & "<br>"
				emailBody = emailBody & "MODE:" & GetPOSTParams("Mode") & "<br>"
				
			If Len(GetPOSTParams("EMAILFORNON200RESPONSES")) > 1 Then
				SendErrorToArray = Split(GetPOSTParams("EMAILFORNON200RESPONSES"),";")
				For z = 0 to Ubound(SendErrorToArray)
					If isEmailValid(SendErrorToArray(z)) = 0 Then 
						SendMail "mailsender@" & maildomain ,SendErrorToArray(z),MUV_READ("ClientID") & " ADD SERVICE MEMO POST ERROR",emailBody,GetTerm("Service"),"Post Failure"
					Else
						SendMail "mailsender@" & maildomain ,"rich@ocsaccess.com",MUV_READ("ClientID") & " ADD SERVICE MEMO POST ERROR",emailBody,GetTerm("Service"),"Post Failure"
					End If
				Next 
			Else
				SendMail "mailsender@" & maildomain ,"rich@ocsaccess.com",MUV_READ("ClientID") & " ADD SERVICE MEMO POST ERROR",emailBody,GetTerm("Service"),"Post Failure"
			End If
			
				Description = emailBody 
				CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("Mode")
			
			End If
			
		Else
			
			emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>SENDSERVICEMSG and <RECORD_SUBTYPE>CANCEL"& "<br>"
			emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
			emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
			If x = 1 Then
				emailBody = emailBody & "Posted to " & GetPOSTParams("ServiceMemoURL1") & "<br>"
			Else
				emailBody = emailBody  & "Posted to " & GetPOSTParams("ServiceMemoURL2") & "<br>"
			End IF
			emailBody = emailBody & "POSTED DATA:" & data & "<br>"
			emailBody = emailBody & "SERNO:" & GetPOSTParams("Serno") & "<br>"
			emailBody = emailBody & "MODE:" & GetPOSTParams("Mode") & "<br>"

			If Len(GetPOSTParams("EMAILFORNON200RESPONSES")) > 1 Then
				SendErrorToArray = Split(GetPOSTParams("EMAILFORNON200RESPONSES"),";")
				For z = 0 to Ubound(SendErrorToArray)
					If isEmailValid(SendErrorToArray(z)) = 0 Then 
						SendMail "mailsender@" & maildomain ,SendErrorToArray(z),MUV_READ("ClientID") & " ADD SERVICE MEMO POST ERROR",emailBody,GetTerm("Service"),"Post Failure"
					Else
						SendMail "mailsender@" & maildomain ,"rich@ocsaccess.com",MUV_READ("ClientID") & " ADD SERVICE MEMO POST ERROR",emailBody,GetTerm("Service"),"Post Failure"
					End If
				Next 
			Else
				SendMail "mailsender@" & maildomain ,"rich@ocsaccess.com",MUV_READ("ClientID") & " ADD SERVICE MEMO POST ERROR",emailBody,GetTerm("Service"),"Post Failure"
			End If
		
			Description = emailBody 
			CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("Mode")
							
		End If
		
		Set httpRequest = Nothing
End IF


Response.Redirect("main_OpenTickets.asp")
	

%>
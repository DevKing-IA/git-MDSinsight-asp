<!--#include file="inc/header-tech-and-driver.asp"-->
<!--#include file="../../inc/mail.asp"-->
<!--#include file="../../inc/settings.asp"-->

<%
sURL = Request.ServerVariables("SERVER_NAME")
'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)

Set Upload = Server.CreateObject("Persits.Upload.1")
Upload.OverwriteFiles = False
Upload.Save
SelectedMemoNumber = Upload.Form("txtTicketNumber")
ServiceNotes = Upload.Form("ServiceNotes")

'Rename the files
' Construct the save path
Pth ="../../clientfiles/" & trim(GetPOSTParams("Serno")) & "/SvcMemoPics/"

x =1
For Each File in Upload.Files
   File.SaveAsVirtual  Pth & SelectedMemoNumber & "-" & x & File.Ext
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


For x = 1 to 2
		
		DO_Post = 0
		
		If x = 1 Then Do_Post = GetPOSTParams("AssetLocationURL1ONOFF") 
		If x = 2 Then Do_Post = GetPOSTParams("AssetLocationURL2ONOFF") 
		
		If IsNull(Do_Post) or Do_Post = "" Then Do_Post = 0

		If cint(Do_Post) = 1 Then

			'If we have asset information, post that first
			If AssetTagNumber <> "" And AssetLocation <> "" Then
			
				Description = Description & ",    Account: "  & Account 
				Description = Description & ",    Asset Tag#: "  & AssetTagNumber
				Description = Description & ",    New Asset Location: "  & AssetLocation 
				Description = Description & ",   Submitted via Insight Field Service WebApp"
				CreateAuditLogEntry "Asset Location Updated","Asset Location Updated","Major",0,Description
			
				data = "asset_id=" & AssetTagNumber 
				data = data & "&asset_loc=" & AssetLocation
				data = data & "&md=" & GetPOSTParams("Mode")
				data = data & "&serno=" & GetPOSTParams("Serno")
				data = data & "&src=Insight Field Service WebApp"
				
				Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
				
				If x = 1 Then
					httpRequest.Open "POST", GetPOSTParams("AssetLocationURL1"), False
				Else
					httpRequest.Open "POST", GetPOSTParams("AssetLocationURL2"), False				
				End If

				httpRequest.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
				httpRequest.Send data
			
				Set httpRequest = Nothing
			
			End If
		End If
		
Next


'If AssetTagNumber <> "" And AssetLocation <> "" Then
'	ServiceNotes = ServiceNotes & "Asset Location updated. Asset Tag#: " & AssetTagNumber & "  New Location:  " & AssetLocation 
'	ServiceNotes = Replace(ServiceNotes,"&","%26") 
'	ServiceNotes = Replace(ServiceNotes," ","%20") & "%0A%0D"
'End If



'*******************************
'Now do the wait for parts piece
'*******************************

SQLwait = "INSERT INTO FS_ServiceMemosDetail (MemoNumber, CustNum, MemoStage, "
SQLwait = SQLwait & " SubmissionDateTime, USerNoSubmittingRecord,UserNoOfServiceTech,Urgent,OriginalDispatchDateTime,Remarks)"
SQLwait = SQLwait &  " VALUES (" 
SQLwait = SQLwait & "'"  & SelectedMemoNumber & "'"
SQLwait = SQLwait & ",'"  & GetServiceTicketCust(SelectedMemoNumber)& "'"
SQLwait = SQLwait & ",'Wait for parts'"
SQLwait = SQLwait & ",getdate() "
SQLwait = SQLwait & ","  & Session("UserNo") 
SQLwait = SQLwait & ","  & GetServiceTicketDispatchedTech(SelectedMemoNumber)
If TicketIsUrgent(SelectedMemoNumber) Then
	SQLwait = SQLwait & ",1" 'Urgent
Else
	SQLwait = SQLwait & ",0" 'Not Urgent
End If
SQLwait = SQLwait & ", '" & TicketOriginalDispatchDateTime(SelectedMemoNumber) & "' "
SQLwait = SQLwait & ",'"  & ServiceNotes & "')"

Set cnnwait = Server.CreateObject("ADODB.Connection")
cnnwait.open (Session("ClientCnnString"))
Set rswait = Server.CreateObject("ADODB.Recordset")
Set rswait = cnnwait.Execute(SQLwait)

	
'Write audit trail for dispatch
'*******************************
Description = GetUserDisplayNameByUserNo(Session("UserNo")) & " set their status to wait for parts for service ticket number " & SelectedMemoNumber & " at " & NOW()
CreateAuditLogEntry "Service Ticket System","Wait for parts","Minor",0,Description 
	
Set rswait = Nothing
cnnwait.Close
Set cnnwait = Nothing

dummy = Redispatch(SelectedMemoNumber)
Description = "Service ticket #" & SelectedMemoNumber  & " was set for redispatch due to waitng for parts at " & NOW()
CreateAuditLogEntry "Service Ticket System","Redispatch","Minor",0,Description 

Call Send_Parts_Request_Email

Response.Redirect("main_menu.asp")


Sub Send_Parts_Request_Email

		Send_To=""
		'Get all the service manager email addresses
		Set cnn_CheckAlerts = Server.CreateObject("ADODB.Connection")
		cnn_CheckAlerts.open (Session("ClientCnnString"))
		Set rs_CheckAlerts = Server.CreateObject("ADODB.Recordset")
		rs_CheckAlerts.CursorLocation = 3 
		SQL_CheckAlerts = "SELECT userEmail FROM tblUsers WHERE userType = 'Service Manager' and userArchived <> 1" 
		Set rs_CheckAlerts = cnn_CheckAlerts.Execute(SQL_CheckAlerts)
		If not rs_CheckAlerts.EOF Then
			Do
				If rs_CheckAlerts("userEmail") <> "" AND Not IsNull(rs_CheckAlerts("userEmail")) Then Send_To = Send_To & rs_CheckAlerts("userEmail") & ";"
				rs_CheckAlerts.MoveNext
			Loop Until rs_CheckAlerts.Eof
		End If
		Set rs_CheckAlerts = Nothing
		cnn_CheckAlerts.Close
		Set cnn_CheckAlerts = Nothing

		'Got all the addresses so now break them up
		Send_To_Array = Split(Send_To,";")

		For x = 0 to Ubound(Send_To_Array) -1
			Send_To = Send_To_Array(x)
			%>
			<!--#include file="../../emails/ADVdispatch_parts_request.asp"-->			
			<%	
			'Failsafe for dev
			If Instr(ucase(sURL),"DEV") <> 0 Then Send_To = "rich@ocsaccess.com"
			SendMail "mailsender@" & maildomain,Send_To,emailSubject,emailBody,GetTerm("Field Service"),"Part Request"
			Description = "A part request was sent to " & Send_To & " for ticket #: " & SelectedMemoNumber & " at " & Now() 
			Description = Description & " The text of the email was: " & ServiceNotes 
			CreateAuditLogEntry "Service Part Request","Service Part Request","Minor",0,Description
		Next
		
End Sub

%>
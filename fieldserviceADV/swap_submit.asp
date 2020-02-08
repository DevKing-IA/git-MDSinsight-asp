<!--#include file="../inc/header-field-service-mobile.asp"-->
<!--#include file="../inc/mail.asp"-->

<%
sURL = Request.ServerVariables("SERVER_NAME")
'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)

Set Upload = Server.CreateObject("Persits.Upload.1")
Upload.OverwriteFiles = False
Upload.Save
SelectedMemoNumber = Upload.Form("txtTicketNumber")

AuthorizedBy = Upload.Form("selPreAuth")
If AuthorizedBy ="" Then AuthorizedBy = "noneselected"



If AuthorizedBy <> "noneselected" Then
	AuthorizedBy = GetUserDisplayNameByUserNo(AuthorizedBy) 
	ServiceNotes = "Authorized by: " & AuthorizedBy & " - " & Upload.Form("ServiceNotes")
Else
	ServiceNotes = "Not Preauthorized - " & Upload.Form("ServiceNotes")
End if



'Rename the files
' Construct the save path
Pth ="../clientfiles/" & trim(GetPOSTParams("Serno")) & "/SvcMemoPics/"

x =1
For Each File in Upload.Files
   File.SaveAsVirtual  Pth & SelectedMemoNumber & "-" & x & File.Ext
   x=x+1
Next

Account = GetServiceTicketCust(SelectedMemoNumber)

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



For x = 1 to 2

		DO_Post = 0
		
		If x = 1 Then Do_Post = GetPOSTParams("AssetLocationURL1ONOFF") 
		If x = 2 Then Do_Post = GetPOSTParams("AssetLocationURL2ONOFF") 
		
				
		If IsNull(Do_Post) or Do_Post = "" Then Do_Post = 0

		If cint(Do_Post) = 1 Then
		
				CreateINSIGHTAuditLogEntry sURL,"Post Loop "& x,GetPOSTParams("Mode")

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



'*******************************
'Now do the wait for parts piece
'*******************************

SQLwait = "INSERT INTO FS_ServiceMemosDetail (MemoNumber, CustNum, MemoStage, "
SQLwait = SQLwait & " SubmissionDateTime, USerNoSubmittingRecord,UserNoOfServiceTech,OriginalDispatchDateTime,Remarks)"
SQLwait = SQLwait &  " VALUES (" 
SQLwait = SQLwait & "'"  & SelectedMemoNumber & "'"
SQLwait = SQLwait & ",'"  & GetServiceTicketCust(SelectedMemoNumber)& "'"
SQLwait = SQLwait & ",'Swap'"
SQLwait = SQLwait & ",getdate() "
SQLwait = SQLwait & ","  & Session("UserNo") 
SQLwait = SQLwait & ","  & GetServiceTicketDispatchedTech(SelectedMemoNumber)
SQLwait = SQLwait & ", '" & TicketOriginalDispatchDateTime(SelectedMemoNumber) & "' "
SQLwait = SQLwait & ",'"  & Replace(ServiceNotes,"'","''") & "')"

'response.Write("<br>" & SQLwait  & "<br>")

Set cnnwait = Server.CreateObject("ADODB.Connection")
cnnwait.open (Session("ClientCnnString"))
Set rswait = Server.CreateObject("ADODB.Recordset")
Set rswait = cnnwait.Execute(SQLwait)

	
'Write audit trail for dispatch
'*******************************
Description = GetUserDisplayNameByUserNo(Session("UserNo")) & " set service ticket number " & SelectedMemoNumber & " to swap at " & NOW()
CreateAuditLogEntry "Service Ticket System","Swap","Minor",0,Description 
	
Set rswait = Nothing
cnnwait.Close
Set cnnwait = Nothing

dummy = Redispatch(SelectedMemoNumber)
Description = "Service ticket #" & SelectedMemoNumber  & " was set for redispatch due to waitng for parts at " & NOW()
CreateAuditLogEntry "Service Ticket System","Redispatch","Minor",0,Description 

If AuthorizedBy="noneselected" Then Call Send_Unauth_Swap_Email

'Now post the appropriate memo to MDS
If AuthorizedBy="noneselected" Then
	MType=15
Else
	MType=18
End If

For x = 1 to 1
		
		DO_Post = 0
		
		If x = 1 Then Do_Post = GetPOSTParams("CustomerURL1ONOFF") 
		If x = 2 Then Do_Post = GetPOSTParams("CustomerURL2ONOFF") 
		
		If IsNull(Do_Post) or Do_Post = "" Then Do_Post = 0
		
		If Not IsNumeric(Do_Post) Then Do_Post = 0
		

		If cint(Do_Post) = 1 Then

			'MDS post which creates a memo
					data = "<DATASTREAM>"
					data = data & "<IDENTITY>Pm8316wyc011</IDENTITY>"
					data = data & "<MODE>" & GetPOSTParams("Mode") & "</MODE>"
					data = data & "<RECORD_TYPE>MEMO</RECORD_TYPE>"
					data = data & "<RECORD_SUBTYPE>CREATE_MEMO</RECORD_SUBTYPE>"
					data = data & "<CLIENT_ID>" & MUV_Read("ClientID") & "</CLIENT_ID>"
					data = data & "<SERNO>" & GetPOSTParams("SERNO") & "</SERNO>"
					data = data & "<SUBMISSION_SOURCE>Insight Field Service WebApp</SUBMISSION_SOURCE>"
					data = data & "<ACCOUNT_NUM>" & Account & "</ACCOUNT_NUM>"
					data = data & "<FIELD_DATA>" & MType & "</FIELD_DATA>"
					data = data & "<FIELD_DATA1>" & ServiceNotes & "</FIELD_DATA1>"
					data = data & "</DATASTREAM>"
		
					If x = 1 Then
						Description = "Post to " & GetPOSTParams("CustomerURL1")
					Else
						Description = "Post to " & GetPOSTParams("CustomerURL2")
					End If

					CreateINSIGHTAuditLogEntry sURL,Description,GetPOSTParams("Mode")
					Description = "data:" & data 
					CreateINSIGHTAuditLogEntry sURL,Description,GetPOSTParams("Mode")
		
					Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
					
					If x = 1 Then
						httpRequest.Open "POST", GetPOSTParams("CustomerURL1"), False
					Else
						httpRequest.Open "POST", GetPOSTParams("CustomerURL2"), False				
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
			
					Set httpRequest = Nothing
		End If
		
Next	


Response.Redirect("main_OpenTickets.asp")


Sub Send_Unauth_Swap_Email

		Send_To=""
		'Get all the service manager email addresses
		Set cnn_CheckAlerts = Server.CreateObject("ADODB.Connection")
		cnn_CheckAlerts.open (Session("ClientCnnString"))
		Set rs_CheckAlerts = Server.CreateObject("ADODB.Recordset")
		rs_CheckAlerts.CursorLocation = 3 
		SQL_CheckAlerts = "SELECT userEmail FROM tblUsers WHERE userType = 'Service Manager' and userArchived <> 1 and userEnabled = 1" 
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
			<!--#include file="../emails/ADVdispatch_NonPreAuth_Swap.asp"-->
			<%	
			'Failsafe for dev
			If Instr(ucase(sURL),"DEV") <> 0 Then Send_To = "rich@ocsaccess.com"
			SendMail "mailsender@" & maildomain,Send_To,emailSubject,emailBody,GetTerm("Field Service"),"Swap Not Pre-Authorized"
			Description = "An non-preauthorized swap email was sent to " & Send_To & " for ticket #: " & SelectedMemoNumber & " at " & Now() 
			CreateAuditLogEntry "Service Part Request","Service Part Request","Minor",0,Description
		Next
		
End Sub

%>
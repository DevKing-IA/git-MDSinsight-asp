<!--#include file="../inc/header-field-service-mobile.asp"-->
<!--#include file="../inc/InSightFuncs.asp"-->

<%

Set Upload = Server.CreateObject("Persits.Upload.1")
Upload.OverwriteFiles = False
Upload.Save
SelectedMemoNumber = Upload.Form("txtTicketNumber")
ServiceNotes = Upload.Form("ServiceNotes")

'Rename the files
' Construct the save path
Pth ="../clientfiles/" & trim(GetPOSTParams("Serno")) & "/SvcMemoPics/"

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
SQLwait = SQLwait & ",'Follow Up'"
SQLwait = SQLwait & ",getdate() "
SQLwait = SQLwait & ","  & Session("UserNo") 
SQLwait = SQLwait & ","  & GetServiceTicketDispatchedTech(SelectedMemoNumber)
SQLwait = SQLwait & ", '" & TicketOriginalDispatchDateTime(SelectedMemoNumber) & "' "
SQLwait = SQLwait & ",'"  & ServiceNotes & "')"

Set cnnwait = Server.CreateObject("ADODB.Connection")
cnnwait.open (Session("ClientCnnString"))
Set rswait = Server.CreateObject("ADODB.Recordset")
Set rswait = cnnwait.Execute(SQLwait)

	
'Write audit trail for dispatch
'*******************************
Description = GetUserDisplayNameByUserNo(Session("UserNo")) & " set the status of service ticket number " & SelectedMemoNumber & "to follow up at " & NOW()
CreateAuditLogEntry "Service Ticket System","Wait for parts","Minor",0,Description 
	
Set rswait = Nothing
cnnwait.Close
Set cnnwait = Nothing

dummy = Redispatch(SelectedMemoNumber)
Description = "Service ticket #" & SelectedMemoNumber  & " was set for redispatch due to being set to follow up at " & NOW()
CreateAuditLogEntry "Service Ticket System","Redispatch","Minor",0,Description 

Response.Redirect("main_OpenTickets.asp")

%>
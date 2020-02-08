<!--#include file="../../../inc/subsandfuncs.asp"-->
<!--#include file="../../../inc/InsightFuncs.asp"-->
<!--#include file="../../../inc/mail.asp"-->



<%
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then 
	
	'baseURL should always have a trailing /slash, just in case, handle either way
	If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)
	
	MondayStartTime = Request.Form("txtMonday")
	MondayOn = Request.Form("chkMonday")
	TuesdayStartTime = Request.Form("txtTuesday")
	TuesdayOn = Request.Form("chkTuesday")
	WednesdayStartTime = Request.Form("txtWednesday")
	WednesdayOn = Request.Form("chkWednesday")
	ThursdayStartTime = Request.Form("txtThursday")
	ThursdayOn = Request.Form("chkThursday")
	FridayStartTime = Request.Form("txtFriday")
	FridayOn = Request.Form("chkFriday")
	SaturdayStartTime = Request.Form("txtSaturday")
	SaturdayOn = Request.Form("chkSaturday")
	SundayStartTime = Request.Form("txtSunday")
	SundayOn = Request.Form("chkSunday")
	NightBatchRunReportEmail = Request.Form("txtRunOrDont")
	NightBatchRunReportTime = Request.Form("txtemailtime")
	NightBatchRunReportOn = Request.Form("chkRunorDont")
	
	If MondayOn = "on" then MondayOn = vbTrue else MondayOn = vbFalse
	If TuesdayOn = "on" then TuesdayOn = vbTrue else TuesdayOn = vbFalse
	If WednesdayOn = "on" then WednesdayOn = vbTrue else WednesdayOn = vbFalse
	If ThursdayOn = "on" then ThursdayOn = vbTrue else ThursdayOn = vbFalse
	If FridayOn = "on" then FridayOn = vbTrue else FridayOn = vbFalse
	If SaturdayOn = "on" then SaturdayOn = vbTrue else SaturdayOn = vbFalse
	If SundayOn = "on" then SundayOn = vbTrue else SundayOn = vbFalse	
	If NightBatchRunReportOn = "on" then NightBatchRunReportOn = vbTrue else NightBatchRunReportOn = vbFalse
	
	
	'*************************************************
	'Now read in the original values for we know 
	'what the differences are
	'*************************************************	
	Orig_SundayOn= MUV_ReadAndRemove("Orig_SundayOn")
	Orig_SundayStartTime= MUV_ReadAndRemove("Orig_SundayStartTime")
	Orig_MondayOn= MUV_ReadAndRemove("Orig_MondayOn")
	Orig_MondayStartTime= MUV_ReadAndRemove("Orig_MondayStartTime")
	Orig_TuesdayOn= MUV_ReadAndRemove("Orig_TuesdayOn")
	Orig_TuesdayStartTime= MUV_ReadAndRemove("Orig_TuesdayStartTime")
	Orig_WednesdayOn= MUV_ReadAndRemove("Orig_WednesdayOn")
	Orig_WednesdayStartTime = MUV_ReadAndRemove("Orig_WednesdayStartTime")
	Orig_ThursdayOn= MUV_ReadAndRemove("Orig_ThursdayOn")
	Orig_ThursdayStartTime= MUV_ReadAndRemove("Orig_ThursdayStartTime")
	Orig_FridayOn = MUV_ReadAndRemove("Orig_FridayOn")
	Orig_FridayStartTime= MUV_ReadAndRemove("Orig_FridayStartTime")
	Orig_SaturdayOn= MUV_ReadAndRemove("Orig_SaturdayOn")
	Orig_SaturdayStartTime= MUV_ReadAndRemove("Orig_SaturdayStartTime")
	'*******************************************************
	' Run/Don't run report settings are in a different table
	'*******************************************************	

	SQL = "SELECT NightBatchRunReportTime, NightBatchRunReportEmail,NightBatchRunReportOn  FROM Settings_Global"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	If not rs.EOF Then
		Orig_NightBatchRunReportTime = rs("NightBatchRunReportTime")
		Orig_NightBatchRunReportEmail = rs("NightBatchRunReportEmail")
		Orig_NightBatchRunReportOn = rs("NightBatchRunReportOn")
	End If
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	'*************************************************
	'End read original values
	'*************************************************	
	
	'***********************************
	'Update the run / don't run settings
	'***********************************
	SQL = "Update Settings_Global SET "
	SQL = SQL & "NightBatchRunReportTime = '" & NightBatchRunReportTime & "', "
	SQL = SQL & "NightBatchRunReportEmail = '" & NightBatchRunReportEmail & "', "
	SQL = SQL & "NightBatchRunReportOn  = " & NightBatchRunReportOn
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	'***************************************
	'eof Update the run / don't run settings
	'***************************************
	
	'Post first, we don't wrute the audit trail unless
	'the post comes back with 200
	
	For x = 1 to 2
	
			DO_Post = 0
			
			If x = 1 Then Do_Post = GetPOSTParams("CustomerURL1ONOFF") 
			If x = 2 Then Do_Post = GetPOSTParams("CustomerURL2ONOFF") 
			
			If IsNull(Do_Post) or Do_Post = "" Then Do_Post = 0
		
			If cint(Do_Post) = 1 Then
			
					CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),"Post Loop "& x,GetPOSTParams("Mode")
			
					'Post to APIs goes here
					data = "<DATASTREAM>"
					data = data & "<IDENTITY>Pm8316wyc011</IDENTITY>"
					data = data & "<MODE>" & GetPOSTParams("Mode") & "</MODE>"
					data = data & "<RECORD_TYPE>NIGHTBATCH</RECORD_TYPE>"
					data = data & "<RECORD_SUBTYPE>CRONSET</RECORD_SUBTYPE>"
					data = data & "<SERNO>" & GetPOSTParams("Serno") & "</SERNO>"
					
					'Now construct the field data
					FieldData =""
					
					If SundayOn = vbTrue Then
						FieldData = FieldData & "0"
						FieldData = FieldData & Replace(FormatDateTime(SundayStartTime,4),":","")
					End If

					If MondayOn = vbTrue Then
						FieldData = FieldData & "1"
						FieldData = FieldData & Replace(FormatDateTime(MondayStartTime,4),":","")
					End If					
					
					If TuesdayOn = vbTrue Then
						FieldData = FieldData & "2"
						FieldData = FieldData & Replace(FormatDateTime(TuesdayStartTime,4),":","")
					End If							

					If WednesdayOn = vbTrue Then
						FieldData = FieldData & "3"
						FieldData = FieldData & Replace(FormatDateTime(WednesdayStartTime,4),":","")
					End If							
					
					If ThursdayOn = vbTrue Then
						FieldData = FieldData & "4"
						FieldData = FieldData & Replace(FormatDateTime(ThursdayStartTime,4),":","")
					End If							

					If FridayOn = vbTrue Then
						FieldData = FieldData & "5"
						FieldData = FieldData & Replace(FormatDateTime(FridayStartTime,4),":","")
					End If							

					If SaturdayOn = vbTrue Then
						FieldData = FieldData & "6"
						FieldData = FieldData & Replace(FormatDateTime(SaturdayStartTime,4),":","")
					End If							
	
					data = data & "<FIELD_DATA>" & FieldData  & "</FIELD_DATA>"
					data = data & "</DATASTREAM>"

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
					
						If Instr(httpRequest.responseText,"success") <> 0 Then ' Success
						
							Description ="success! httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>NIGHTBATCH and <RECORD_SUBTYPE>CRONSET "& "<br>"
							Description = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
							Description = Description & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
							Description = Description & "Posted to " & GetPOSTParams("CUSTOMERURL1") & "<br>"
							Description = Description & "POSTED DATA:" & data & "<br>"
							Description = Description & "SERNO:" & MUV_READ("ClientID") & "<br>"
					
							CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("Mode")

							dummy=MUV_WRITE("NIGHTBATCHOK",1)

						Else
							'FAILURE
							emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>NIGHTBATCH and <RECORD_SUBTYPE>CRONSET "& "<br>"
							emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
							emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
							emailBody = emailBody & "Posted to " & GetPOSTParams("CUSTOMERURL1") & "<br>"
							emailBody = emailBody & "POSTED DATA:" & data & "<br>"
							emailBody = emailBody & "SERNO:" & MUV_READ("ClientID") & "<br>"
							SendMail "mailsender@" & maildomain ,"rich@ocsaccess.com",MUV_READ("ClientID") & " NIGHT BATCH CRONSET POST ERROR",emailBody, "Night Batch", "Post Failure"
						
							Description = emailBody 
							CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("Mode")
						
							dummy=MUV_WRITE("NIGHTBATCHOK",0)
						End If
						
					Else
						
						emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>NIGHTBATCH and <RECORD_SUBTYPE>CRONSET "& "<br>"
						emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
						emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
						emailBody = emailBody & "Posted to " & GetPOSTParams("CUSTOMERURL1") & "<br>"
						emailBody = emailBody & "POSTED DATA:" & data & "<br>"
						emailBody = emailBody & "SERNO:" & MUV_READ("ClientID") & "<br>"
						SendMail "mailsender@" & maildomain ,"projects@metroplexdata.com",MUV_READ("ClientID") & " NIGHT BATCH CRONSET POST ERROR",emailBody, "Night Batch", "Post Failure"
					
						Description = emailBody 
						CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("Mode")
					
						dummy=MUV_WRITE("NIGHTBATCHOK",0)
											
					End If
					
					Set httpRequest = Nothing
			End IF
			
			If MUV_READ("NIGHTBATCHOK") = 0 Then Exit For ' Dont try another if the first failed
	Next
	
	If MUV_READ("NIGHTBATCHOK") <> 0 Then Call HandleAuditTrail
	
	Response.Redirect("readNightBatchloading.asp")	
		
	
End If ' If it came from a post

'Just put this here to keep the code above a little neater
Sub HandleAuditTrail()

	If Orig_MondayStartTime <> MondayStartTime Then CreateAuditLogEntry "Night batch schedule change", "Night batch schedule change", "Major", 1, MUV_Read("DisplayName") & " changed Monday start time from " & VerbiageForReport & " to " & MondayStartTime
	If Orig_MondayOn <> MondayOn Then
		' Just make it say On/Off instead of True/False
		If Orig_MondayOn = vbTrue then Orig_MondayOnForReport = "On" else Orig_MondayOnForReport = "Off" 
		If MondayOn = vbTrue then MondayOnForReport = "On" else MondayOnForReport = "Off" 
		CreateAuditLogEntry "Night batch schedule change", "Night batch schedule change", "Major", 1, MUV_Read("DisplayName") & " changed Monday on/off from " & Orig_MondayOnForReport & " to " & MondayOnForReport 
	End If
	If Orig_TuesdayStartTime <> TuesdayStartTime Then CreateAuditLogEntry "Night batch schedule change", "Night batch schedule change", "Major", 1, MUV_Read("DisplayName") & " changed Tuesday start time from " & VerbiageForReport & " to " & TuesdayStartTime
	If Orig_TuesdayOn <> TuesdayOn Then
		' Just make it say On/Off instead of True/False
		If Orig_TuesdayOn = vbTrue then Orig_TuesdayOnForReport = "On" else Orig_TuesdayOnForReport = "Off" 
		If TuesdayOn = vbTrue then TuesdayOnForReport = "On" else TuesdayOnForReport = "Off" 
		CreateAuditLogEntry "Night batch schedule change", "Night batch schedule change", "Major", 1, MUV_Read("DisplayName") & " changed Tuesday on/off from " & Orig_TuesdayOnForReport & " to " & TuesdayOnForReport 
	End If
	If Orig_WednesdayStartTime <> WednesdayStartTime Then CreateAuditLogEntry "Night batch schedule change", "Night batch schedule change", "Major", 1, MUV_Read("DisplayName") & " changed Wednesday start time from " & VerbiageForReport & " to " & WednesdayStartTime
	If Orig_WednesdayOn <> WednesdayOn Then
		' Just make it say On/Off instead of True/False
		If Orig_WednesdayOn = vbTrue then Orig_WednesdayOnForReport = "On" else Orig_WednesdayOnForReport = "Off" 
		If WednesdayOn = vbTrue then WednesdayOnForReport = "On" else WednesdayOnForReport = "Off" 
		CreateAuditLogEntry "Night batch schedule change", "Night batch schedule change", "Major", 1, MUV_Read("DisplayName") & " changed Wednesday on/off from " & Orig_WednesdayOnForReport & " to " & WednesdayOnForReport 
	End If
	If Orig_ThursdayStartTime <> ThursdayStartTime Then CreateAuditLogEntry "Night batch schedule change", "Night batch schedule change", "Major", 1, MUV_Read("DisplayName") & " changed Thursday start time from " & VerbiageForReport & " to " & ThursdayStartTime
	If Orig_ThursdayOn <> ThursdayOn Then
		' Just make it say On/Off instead of True/False
		If Orig_ThursdayOn = vbTrue then Orig_ThursdayOnForReport = "On" else Orig_ThursdayOnForReport = "Off" 
		If ThursdayOn = vbTrue then ThursdayOnForReport = "On" else ThursdayOnForReport = "Off" 
		CreateAuditLogEntry "Night batch schedule change", "Night batch schedule change", "Major", 1, MUV_Read("DisplayName") & " changed Thursday on/off from " & Orig_ThursdayOnForReport & " to " & ThursdayOnForReport 
	End If
	If Orig_FridayStartTime <> FridayStartTime Then CreateAuditLogEntry "Night batch schedule change", "Night batch schedule change", "Major", 1, MUV_Read("DisplayName") & " changed Friday start time from " & VerbiageForReport & " to " & FridayStartTime
	If Orig_FridayOn <> FridayOn Then
		' Just make it say On/Off instead of True/False
		If Orig_FridayOn = vbTrue then Orig_FridayOnForReport = "On" else Orig_FridayOnForReport = "Off" 
		If FridayOn = vbTrue then FridayOnForReport = "On" else FridayOnForReport = "Off" 
		CreateAuditLogEntry "Night batch schedule change", "Night batch schedule change", "Major", 1, MUV_Read("DisplayName") & " changed Friday on/off from " & Orig_FridayOnForReport & " to " & FridayOnForReport 
	End If
	If Orig_SaturdayStartTime <> SaturdayStartTime Then CreateAuditLogEntry "Night batch schedule change", "Night batch schedule change", "Major", 1, MUV_Read("DisplayName") & " changed Saturday start time from " & VerbiageForReport & " to " & SaturdayStartTime
	If Orig_SaturdayOn <> SaturdayOn Then
		' Just make it say On/Off instead of True/False
		If Orig_SaturdayOn = vbTrue then Orig_SaturdayOnForReport = "On" else Orig_SaturdayOnForReport = "Off" 
		If SaturdayOn = vbTrue then SaturdayOnForReport = "On" else SaturdayOnForReport = "Off" 
		CreateAuditLogEntry "Night batch schedule change", "Night batch schedule change", "Major", 1, MUV_Read("DisplayName") & " changed Saturday on/off from " & Orig_SaturdayOnForReport & " to " & SaturdayOnForReport 
	End If
	If Orig_SundayStartTime <> SundayStartTime Then CreateAuditLogEntry "Night batch schedule change", "Night batch schedule change", "Major", 1, MUV_Read("DisplayName") & " changed Sunday start time from " & VerbiageForReport & " to " & SundayStartTime
	If Orig_SundayOn <> SundayOn Then
		' Just make it say On/Off instead of True/False
		If Orig_SundayOn = vbTrue then Orig_SundayOnForReport = "On" else Orig_SundayOnForReport = "Off" 
		If SundayOn = vbTrue then SundayOnForReport = "On" else SundayOnForReport = "Off" 
		CreateAuditLogEntry "Night batch schedule change", "Night batch schedule change", "Major", 1, MUV_Read("DisplayName") & " changed Sunday on/off from " & Orig_SundayOnForReport & " to " & SundayOnForReport 
	End If
	
	If Orig_NightBatchRunReportEmail <> NightBatchRunReportEmail Then CreateAuditLogEntry "Night batch schedule change", "Night batch schedule change", "Major", 1, MUV_Read("DisplayName") & " changed Send a daily run/don't run status email to the following addresses  " & Orig_NightBatchRunReportEmail & " to " & NightBatchRunReportEmail
		
	If Orig_NightBatchRunReportTime <> NightBatchRunReportTime Then CreateAuditLogEntry "Night batch schedule change", "Night batch schedule change", "Major", 1, MUV_Read("DisplayName") & " changed Send a daily run/don't run status email at the following time  " & Orig_NightBatchRunReportTime & " to " & NightBatchRunReportTime		

	If Orig_NightBatchRunReportOn <> NightBatchRunReportOn Then
		' Just make it say On/Off instead of True/False
		If Orig_NightBatchRunReportOn = vbTrue then Orig_NightBatchRunReportOn = "On" else Orig_NightBatchRunReportOn = "Off" 
		If NightBatchRunReportOn = vbTrue then NightBatchRunReportOn = "On" else NightBatchRunReportOn = "Off" 
		CreateAuditLogEntry "Night batch schedule change", "Night batch schedule change", "Major", 1, MUV_Read("DisplayName") & " changed Turn on daily run / don't run status email from " & Orig_NightBatchRunReportOn  & " to " & NightBatchRunReportOn 
	End If

End Sub
%>
















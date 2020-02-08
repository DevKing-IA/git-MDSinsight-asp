<!--#include file="../inc/InSightFuncs.asp"-->
<!--#include file="../inc/InSightFuncs_Users.asp"-->
<!--#include file="../inc/settings.asp"-->
<!--#include file="../inc/mail.asp"-->
<!--#include file="../inc/InsightFuncs_Service.asp"-->
<%
sURL = Request.ServerVariables("SERVER_NAME")
'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)

Set cnnDispatch = Server.CreateObject("ADODB.Connection")
cnnDispatch.open (Session("ClientCnnString"))


ListOfPossibleTickets = Request.Form("txtListOfPossibleTickets")
If ListOfPossibleTickets <> "" Then ListOfPossibleTicketsArray = Split(ListOfPossibleTickets,",")


SendText = "on"

For Each TickNum in ListOfPossibleTicketsArray 

	If Left(TickNum,1) <> "F" OR (Left(TickNum,1)="F" AND TicketInServiceMemosFilterInfo(TickNum) = True) Then 'Not the PENDING filters, we will do them next

		ServiceTicketNumber = TickNum 
		
		selectBoxToGet = "selFieldTech" & TickNum
		
		If Request.Form(selectBoxToGet) <> 0 Then ' Zero means do not dospatch
		
			UserToDispatch = Request.Form(selectBoxToGet)
			
			CustNum = GetServiceTicketCust(TickNum)
			
			'Now do the dispatching
			SQLDispatch = "INSERT INTO FS_ServiceMemosDetail (MemoNumber, CustNum, MemoStage, "
			SQLDispatch = SQLDispatch & "UserNoOfServiceTech, SubmissionDateTime, USerNoSubmittingRecord,EmailAddressSentTo,TextNumberSentTo,OriginalDispatchDateTime,Remarks)"
			SQLDispatch = SQLDispatch &  " VALUES (" 
			SQLDispatch = SQLDispatch & "'"  & ServiceTicketNumber & "'"
			SQLDispatch = SQLDispatch & ",'"  & CustNum & "'"
			SQLDispatch = SQLDispatch & ",'Dispatched'"
			SQLDispatch = SQLDispatch & ","  & UserToDispatch 
			SQLDispatch = SQLDispatch & ",getdate() "
			SQLDispatch = SQLDispatch & ","  & Session("UserNo")
			SQLDispatch = SQLDispatch & ",'"  & getUserEmailAddress(UserToDispatch) & "'"
			SQLDispatch = SQLDispatch & ",'" & getUserCellNumber(UserToDispatch) & "' "
			SQLDispatch = SQLDispatch & ", getDate() "
			SQLDispatch = SQLDispatch & ",'" &  GetUserDisplayNameByUserNo(UserToDispatch) & " has been dispatched.')"	
			
			Set rsDispatch = Server.CreateObject("ADODB.Recordset")
			Set rsDispatch = cnnDispatch.Execute(SQLDispatch)
			
			
			'Write audit trail for dispatch
			'*******************************
			Description = GetUserDisplayNameByUserNo(UserToDispatch) & " was dispatched to service ticket number " & ServiceTicketNumber & " by " & GetUserDisplayNameByUserNo(Session("UserNo")) & " at " & NOW()
			CreateAuditLogEntry "Service Ticket System","Dispatched","Minor",0,Description 
			
			'Also set dispatched flag for simple dispatch model
			SQLDispatch = "Update FS_ServiceMemos Set Dispatched = CASE WHEN Dispatched = 0 THEN -1 ELSE 0 END Where MemoNumber = '"  & ServiceTicketNumber & "'"
			Set rsDispatch = cnnDispatch.Execute(SQLDispatch)
			
			
			Set rsDispatch = Nothing
			
			
			
			'Send text if necessary
			'**********************
			If SendText="on" then
			
				'See if ACK link should be included
				DLinkInText = False
				SQLtxt = "SELECT IncludeACKInDispatchText FROM Settings_EmailService"
				Set rstxt = Server.CreateObject("ADODB.Recordset")
				rstxt.CursorLocation = 3 
				Set rstxt = cnnDispatch.Execute(SQLtxt)
			
				If not rstxt.EOF Then DLinkInText = rstxt("IncludeACKInDispatchText")
				set rstxt = Nothing
				
				If getUserCellNumber(UserToDispatch) <> "" Then
				
					Send_To = getUserCellNumber(UserToDispatch)
	
					'Only do this if there are actually texts to send
					If Send_To <> "" Then
												
						Send_To = Replace(Send_To,"-","") ' EZ Texting doesn't like dashes
				
						'*****Text numbers don't get split into an array, the php takes multiple #'s seprated by commas	
							
						If Right(Send_To,1) = "," Then Send_To = Left(Send_To,Len(Send_To)-1)
					
						TEXT_TO = Send_To
						
						'Split Text_To for recording the alerts sent
						TextNumberArray = Split(TEXT_TO&",",",")
						
			
						str_data="txtSubject=" & "DISPATCH"
	
	
						If GetCustNameByCustNum(CustNum) <> "" Then
							txtMessage = GetTerm("Account") & ": " & EZTexting_Filter1(Replace(GetCustNameByCustNum(CustNum),"&"," "))
						Else
							txtMessage = GetTerm("Account") & ": " & CustNum 
							txtMessage = txtMessage &  "   Ticket:" & ServiceTicketNumber
						End If
			
						If DLinkInText = 1 Then
							txtMessage = txtMessage & "    Tap the link for more info "
							txtMessage = txtMessage & Server.URLEncode(baseURL & "directlaunch/service/moreinfo_dispatch_from_email_or_text.asp?t=" & ServiceTicketNumber & "&u=" & UserToDispatch & "&c=" & CustNum & "&cl=" & MUV_READ("SERNO"))
						End If
			
						txtMessage = Replace(txtMessage ," ", "%20")
	
						
	 					str_data = str_data & "&txtMessage=" & txtMessage 
	 					
	 					str_data = str_data &  "&txtTEXT_TO=" & TEXT_TO
				
						str_data = str_data & "&txtu1=" & EzTextingUserID() & "&txtu2=" & EzTextingPassword()
						
						str_data = str_data & "&txtCountry=" & GetCompanyCountry()
						
						Set obj_post=Server.CreateObject("Msxml2.SERVERXMLHTTP")
						obj_post.Open "POST", BaseURL & "inc/sendtext_post.php",False
						obj_post.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
						obj_post.Send str_data
	
						Description = "A dispatch text message was sent to " & GetUserDisplayNameByUserNo(UserToDispatch) & " (" & getUserCellNumber(UserToDispatch) & ") at " & NOW()
						CreateAuditLogEntry "Service Ticket System","Dispatch email sent","Minor",0,Description
						
						Description = str_data
						CreateAuditLogEntry "Texting data",str_data,"Minor",0,Description
						
					End If
			End If
			
			dummy=RemoveFromRedispatch(ServiceTicketNumber)
	
			End If
		End If
	
	End If	
Next

'*******************************
' Now do all the PENDING filters
'*******************************
Filter_Rec_IDs = ""

For Each TickNum in ListOfPossibleTicketsArray 

	If Left(TickNum,1) = "F" AND TicketInServiceMemosFilterInfo(TickNum) = False Then 'ONLY the PENDING filters

		ServiceTicketNumber = TickNum 
		
		selectBoxToGet = "selFieldTech" & TickNum
		
		If Request.Form(selectBoxToGet) <> 0 Then ' Zero means do not dospatch
		
			If Request.Form(selectBoxToGet) <> 9999 Then UserToDispatch = Request.Form(selectBoxToGet)
	
			' First create the service ticket for this filter change
			'Lookup everything else we need
			SQL = "SELECT * FROM FS_CustomerFilters "
			SQL = SQL & "WHERE InternalRecordIdentifier = " & Right(TickNum,Len(TickNum)-1)
			


			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.CursorLocation = 3 
			Set rs = cnnDispatch.Execute(SQL)
			If not rs.eof then


				AccountNumber = rs("CustID")
				Company = GetCustNameByCustNum(rs("CustID"))
				ProblemDescription = "FILTER CHANGE"
				SumissionSource = "MDS Insight"
				FilterInRecID = rs("FilterIntRecID")
				IntRecID =  rs("InternalRecordIdentifier")
			Else

				SendMail "mailsender@" & maildomain ,"rich@ocsaccess.com","ERROR ERROR ERROR ERROR ",SQL,GetTerm("Service"),"Post Failure"
			end if
			set rs = nothing
				
			filters = filters & vbcrlf & "Filter: " & GetFilterIDByIntRecID(FilterInRecID) & " - " & GetFilterDescByIntRecID(FilterInRecID)
	
			Filter_Rec_IDs = Filter_Rec_IDs & IntRecID  & ","
				
		End If
	
	'End If
	
'Next		

If Right(Filter_Rec_IDs,1) = "," Then Filter_Rec_IDs = Left(Filter_Rec_IDs,len(Filter_Rec_IDs)-1)



	
			Description = "OPEN,    "
			Description = Description & "Account: "  & AccountNumber & " - " & Company 
			Description = Description & ",    Description: "  & ProblemDescription 
			CreateAuditLogEntry "Service Memo Added","Service Memo Added","Major",0,Description

			DO_Post = 0
			
			Do_Post = GetPOSTParams("ServiceMemoURL1ONOFF") 
					
			If IsNull(Do_Post) or Do_Post = "" Then Do_Post = 0
		
			If cint(Do_Post) = 1 Then

				If GetPOSTParams("SERVICEMEMOURL1MPLEXFORMAT") = 1 then 'see if using metroplex format
					
					CreateINSIGHTAuditLogEntry sURL,"Metroplex post format",GetPOSTParams("Mode")
					
					'Post to APIs goes here
					data = "create_service_request&account="
					data = data & AccountNumber 
					data = data & "&prob=FILTER CHANGE - " & filters & "%0A%0D"
					data = data & "Opened by: " & 	MUV_Read("DisplayName") & " - " & Session("UserEmail")
					data = data & "&probloc=" & ProblemLocation 
					data = data & "&sbn=" & ContactName 
					data = data & "&sbe=" & ContactEmail
					data = data & "&sbp=" & ContactPhone 
					data = data & "&contnm=" & ContactName
					data = data & "&st=" & "OPEN"
					data = data & "&md=" & GetPOSTParams("Mode")
					data = data & "&serno=" & GetPOSTParams("Serno")
					data = data & "&src=" & SumissionSource
					data = data & "&usr=" & Session("userNo")
					data = data & "&pm=" & Server.URLencode("Filter Change")
					data = data & "&rids=" & Filter_Rec_IDs


'If GetPOSTParams("NeverPutOnHold") = 0  Then data = data & "&hld=1" Else data = data & "&hld=0" 'I know it's wierd. It's the opposite of how it is stored in the table
data = data & "&hld=0"

							
				Else ' Regular XML format
			
					'Post to APIs goes here
							
					data = ""
							
					data = data & "<POST_DATA>"
					data = data & "<IDENTITY>Pm8316wyc011</IDENTITY>"
					data = data & "<ACCOUNT_NUM>" & AccountNumber & "</ACCOUNT_NUM>"
					data = data & "<PROBLEM_LOCATION>" & ProblemLocation & "</PROBLEM_LOCATION>"
					data = data & "<PROBLEM_DESCRIPTION>FILTER CHANGE - " & filters & "</PROBLEM_DESCRIPTION>"
					data = data & "<RECORD_TYPE>SENDSERVICEMSG</RECORD_TYPE>"
					data = data & "<RECORD_SUBTYPE>OPEN</RECORD_SUBTYPE>"
					data = data & "<SERVICE_TICKET_NUMBER>AUTO</SERVICE_TICKET_NUMBER>"
					data = data & "<COMPANY_NAME>" & Replace(GetCustNameByCustNum(AccountNumber),"&","&amp;") & "</COMPANY_NAME>"
					data = data & "<SUBMISSION_SOURCE>MDS Insight</SUBMISSION_SOURCE>"
					data = data & "<SERNO>" & GetPOSTParams("Serno") & "</SERNO>"
					data = data & "<CLIENT_ID>" & MUV_READ("ClientID") & "</CLIENT_ID>"
					data = data & "<USER_NO>" & Session("UserNo") & "</USER_NO>"
					data = data & "<MODE>" & GetPOSTParams("Mode") & "</MODE>"
					data = data & "<SUBMITTED_BY_PHONE>" & ContactPhone & "</SUBMITTED_BY_PHONE>"
					data = data & "<SUBMITTED_BY_EMAIL>" & ContactEmail & "</SUBMITTED_BY_EMAIL>"
					data = data & "<SUBMITTED_BY_NAME>" & ContactEmail & "</SUBMITTED_BY_NAME>"
					data = data & "<FILTER_REC_IDS>" & Filter_Rec_IDs & "</FILTER_REC_IDS>"					
					data = data & "</POST_DATA>"
					data = Replace(data ,"&","&amp;")
					data = Replace(data ,chr(34),"")
				End IF
			
				Description = "Post to " & GetPOSTParams("ServiceMemoURL1")
	
				CreateINSIGHTAuditLogEntry sURL,Description,GetPOSTParams("Mode")
				Description = "data:" & data 
				CreateINSIGHTAuditLogEntry sURL,Description,GetPOSTParams("Mode")
	
			
				CreateINSIGHTAuditLogEntry sURL,"Post Loop "& x,GetPOSTParams("Mode")
	
				Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")
					
				httpRequest.Open "POST", GetPOSTParams("ServiceMemoURL1"), False
					
				httpRequest.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
				httpRequest.Send data
		
				If httpRequest.status = 200 THEN 
								
					If Instr(httpRequest.responseText,"success") <> 0 Then ' Success
								
						Description ="success! httpRequest.status returned " & httpRequest.status & " when posting a filter change ticket"& "<br>"
						Description = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
						Description = Description & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
						Description = Description & "Posted to " & GetPOSTParams("ServiceMemoURL1") & "<br>"
						Description = Description & "POSTED DATA:" & data & "<br>"
						Description = Description & "SERNO:" & GetPOSTParams("Serno") & "<br>"
						Description = Description & "MODE:" & GetPOSTParams("Mode") & "<br>"
				
						CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("Mode")
			
					Else
						'FAILURE
						emailbody="httpRequest.status returned " & httpRequest.status & " when posting a filter change ticket"& "<br>"
						emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
						emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
						emailBody = emailBody & "Posted to " & GetPOSTParams("ServiceMemoURL1") & "<br>"
						emailBody = emailBody & "POSTED DATA:" & data & "<br>"
						emailBody = emailBody & "SERNO:" & GetPOSTParams("Serno") & "<br>"
						emailBody = emailBody & "MODE:" & GetPOSTParams("Mode") & "<br>"
									
						SendMail "mailsender@" & maildomain ,"rich@ocsaccess.com",MUV_READ("ClientID") & " ADD SERVICE MEMO POST ERROR",emailBody,GetTerm("Service"),"Post Failure"
		
						Description = emailBody 
						CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("Mode")
						
					End If
								
				Else
									
					emailbody="httpRequest.status returned " & httpRequest.status & " when posting a filter change ticket"& "<br>"
					emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br>"
					emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
					emailBody = emailBody & "Posted to " & GetPOSTParams("ServiceMemoURL1") & "<br>"
					emailBody = emailBody & "POSTED DATA:" & data & "<br>"
					emailBody = emailBody & "SERNO:" & GetPOSTParams("Serno") & "<br>"
					emailBody = emailBody & "MODE:" & GetPOSTParams("Mode") & "<br>"
			
					SendMail "mailsender@" & maildomain ,"rich@ocsaccess.com",MUV_READ("ClientID") & " ADD SERVICE MEMO POST ERROR",emailBody,GetTerm("Service"),"Post Failure"
							
					Description = emailBody 
					CreateINSIGHTAuditLogEntry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("Mode")
												
				End If
			End IF
		
	


		'***********************************************************************
		'***********************************************************************		
		'***********************************************************************
		' Now we need to lookup the service ticket we just created & dispatch it
		'***********************************************************************
		'***********************************************************************
		'***********************************************************************
		ServiceTicketNumber = ""
		
		SQL = "SELECT ServiceTicketID AS MemoNumber FROM FS_ServiceMemosFilterInfo WHERE CustID = '" & AccountNumber & "' AND "
		SQL = SQL & "ServiceTicketID IN (SELECT MemoNumber FROM FS_ServiceMemos WHERE CurrentStatus='OPEN')"

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = 3 
Response.Write(SQL)	

		Set rs = cnnDispatch.Execute(SQL)
		IF NOT rs.EOF Then ServiceTicketNumber = rs("MemoNumber")
		set rs = nothing
		
		' Write Audit trail first, then post
		
		Set rs8 = Server.CreateObject("ADODB.Recordset")
		rs8.CursorLocation = 3 
		'Set rs8 = cnnDispatch.Execute(SQL)
		set rs8 = Nothing

		selectBoxToGet = "selFieldTech" & TickNum
		
		If ServiceTicketNumber <> "" Then ' Zero means do not dospatch
			
			
			CustNum = AccountNumber
			
			'Now do the dispatching
			SQLDispatch = "INSERT INTO FS_ServiceMemosDetail (MemoNumber, CustNum, MemoStage, "
			SQLDispatch = SQLDispatch & "UserNoOfServiceTech, SubmissionDateTime, USerNoSubmittingRecord,EmailAddressSentTo,TextNumberSentTo,OriginalDispatchDateTime,Remarks)"
			SQLDispatch = SQLDispatch &  " VALUES (" 
			SQLDispatch = SQLDispatch & "'"  & ServiceTicketNumber & "'"
			SQLDispatch = SQLDispatch & ",'"  & CustNum & "'"
			SQLDispatch = SQLDispatch & ",'Dispatched'"
			SQLDispatch = SQLDispatch & ","  & UserToDispatch 
			SQLDispatch = SQLDispatch & ",getdate() "
			SQLDispatch = SQLDispatch & ","  & Session("UserNo")
			SQLDispatch = SQLDispatch & ",'"  & getUserEmailAddress(UserToDispatch) & "'"
			SQLDispatch = SQLDispatch & ",'" & getUserCellNumber(UserToDispatch) & "' "
			SQLDispatch = SQLDispatch & ", getDate() "
			SQLDispatch = SQLDispatch & ",'" &  GetUserDisplayNameByUserNo(UserToDispatch) & " has been dispatched.')"	
			
			
			Response.Write(SQLDispatch )
			
			Set cnnDispatch = Server.CreateObject("ADODB.Connection")
			cnnDispatch.open (Session("ClientCnnString"))
			Set rsDispatch = Server.CreateObject("ADODB.Recordset")
			Set rsDispatch = cnnDispatch.Execute(SQLDispatch)
			
			
			'Write audit trail for dispatch
			'*******************************
			Description = GetUserDisplayNameByUserNo(UserToDispatch) & " was dispatched to service ticket number " & ServiceTicketNumber & " by " & GetUserDisplayNameByUserNo(Session("UserNo")) & " at " & NOW()
			CreateAuditLogEntry "Service Ticket System","Dispatched","Minor",0,Description 			
			
			
			'Send text if necessary
			'**********************
			If SendText="on" then
			
				'See if ACK link should be included
				DLinkInText = False
				SQLtxt = "SELECT IncludeACKInDispatchText FROM Settings_EmailService"
				Set rstxt = Server.CreateObject("ADODB.Recordset")
				rstxt.CursorLocation = 3 
				Set rstxt = cnnDispatch.Execute(SQLtxt)
			
				If not rstxt.EOF Then DLinkInText = rstxt("IncludeACKInDispatchText")
				set rstxt = Nothing
				
				If getUserCellNumber(UserToDispatch) <> "" Then
				
					Send_To = getUserCellNumber(UserToDispatch)
	
					'Only do this if there are actually texts to send
					If Send_To <> "" Then
												
						Send_To = Replace(Send_To,"-","") ' EZ Texting doesn't like dashes
				
						'*****Text numbers don't get split into an array, the php takes multiple #'s seprated by commas	
							
						If Right(Send_To,1) = "," Then Send_To = Left(Send_To,Len(Send_To)-1)
					
						TEXT_TO = Send_To
						
						'Split Text_To for recording the alerts sent
						TextNumberArray = Split(TEXT_TO&",",",")
						
			
						str_data="txtSubject=" & "DISPATCH"
	
	
						If GetCustNameByCustNum(CustNum) <> "" Then
							txtMessage = GetTerm("Account") & ": " & EZTexting_Filter1(Replace(GetCustNameByCustNum(CustNum),"&"," "))
						Else
							txtMessage = GetTerm("Account") & ": " & CustNum 
							txtMessage = txtMessage &  "   Ticket:" & ServiceTicketNumber
						End If
			
						If DLinkInText = 1 Then
							txtMessage = txtMessage & "    Tap the link for more info "
							txtMessage = txtMessage & Server.URLEncode(baseURL & "directlaunch/service/moreinfo_dispatch_from_email_or_text.asp?t=" & ServiceTicketNumber & "&u=" & UserToDispatch & "&c=" & CustNum & "&cl=" & MUV_READ("SERNO"))
						End If
			
						txtMessage = Replace(txtMessage ," ", "%20")
	
						
	 					str_data = str_data & "&txtMessage=" & txtMessage 
	 					
	 					str_data = str_data &  "&txtTEXT_TO=" & TEXT_TO
				
						str_data = str_data & "&txtu1=" & EzTextingUserID() & "&txtu2=" & EzTextingPassword()
						
						str_data = str_data & "&txtCountry=" & GetCompanyCountry()
				
						Set obj_post=Server.CreateObject("Msxml2.SERVERXMLHTTP")
						obj_post.Open "POST", BaseURL & "inc/sendtext_post.php",False
						obj_post.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
						obj_post.Send str_data
	
						Description = "A dispatch text message was sent to " & GetUserDisplayNameByUserNo(UserToDispatch) & " (" & getUserCellNumber(UserToDispatch) & ") at " & NOW()
						CreateAuditLogEntry "Service Ticket System","Dispatch email sent","Minor",0,Description
						
						Description = str_data
						CreateAuditLogEntry "Texting data",str_data,"Minor",0,Description
						
					End If
			End If
			
			dummy=RemoveFromRedispatch(ServiceTicketNumber)
	
			End If
		End If
		'***************************************************************************
		'***************************************************************************		
		'***************************************************************************
		' END Now we need to lookup the service ticket we just created & dispatch it
		'***************************************************************************
		'***************************************************************************
	'***************************************************************************
	end if
Next	

'Response.Write("TickNum :" & TickNum & "<br>")
'Response.Write("Right :" & cstr(Right(TickNum,Len(TickNum)-1)) & "<br>")
'Response.Write("filters :" & filters & "<br>")
		



'Response.Write("SendEmail:" & SendEmail & "<br>")
'Response.Write("SendText:" & SendText& "<br>")
'Response.Write("UserToDispatch :" & UserToDispatch & "<br>")
'Response.Write("ServiceTicketNumber :" & ServiceTicketNumber & "<br>")

'Response.End

cnnDispatch.Close
Set cnnDispatch = Nothing



Response.Redirect(BaseURL & "service/main.asp")
%>

 

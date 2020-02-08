<!--#include file="../../../inc/subsandfuncs.asp"-->
<!--#include file="../../../inc/InsightFuncs.asp"-->
<!--#include file="../../../inc/InsightFuncs_service.asp"-->
<!--#include file="../../../inc/mail.asp"-->

<%
sURL = Request.ServerVariables("SERVER_NAME")
'baseURL should always have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)

FilterIntRecID = "" : CustID = "" : Filter_Rec_IDs = ""

'**********************************************************
'**********************************************************
'**********************************************************
FilterIntRecID = Request.Form("FilterIntRecID")
CustID = Request.Form("CustID")
'**********************************************************
'**********************************************************
'**********************************************************

Set cnnCreateTicket = Server.CreateObject("ADODB.Connection")
cnnCreateTicket.open (Session("ClientCnnString"))
Set rsCreateTicket = Server.CreateObject("ADODB.Recordset")

ReDim FilterList(1)


If CustID <> "" and FilterIntRecID = "" Then   ' All filters for a given customer

	SQLCreateTicket = "SELECT * FROM FS_CustomerFilters WHERE CustID = '" & CustID & "' AND InternalRecordIdentifier NOT IN (SELECT InternalRecordIdentifier FROM FS_ServiceMemosFilterInfo WHERE CustID = '" & CustID & "' AND Completed=0)"
	'SQLCreateTicket = "SELECT * FROM FS_CustomerFilters WHERE CustID = '" & CustID & "'"

End If

If FilterIntRecID <> "" Then   ' One particular filter

	SQLCreateTicket = "SELECT * FROM FS_CustomerFilters WHERE InternalRecordIdentifier = " & FilterIntRecID
	
End If

Set rsCreateTicket = cnnCreateTicket.Execute(SQLCreateTicket)

If NOT rsCreateTicket.EOF Then

	Do While Not rsCreateTicket.EOF
	
		CustID = rsCreateTicket("CustID") 
		
		
		FilterList(Ubound(FilterList)) = rsCreateTicket("FilterIntRecID")
		
		Filter_Rec_IDs = Filter_Rec_IDs & rsCreateTicket("InternalRecordIdentifier") & ","

		ReDim Preserve FilterList(Ubound(FilterList) + 1 )
		'addTicket rsCreateTicket("CustID") , rsCreateTicket("FilterIntRecID") 
		rsCreateTicket.Movenext
	Loop
	
End If

rsCreateTicket.Close

DIM varSQL

If right(Filter_Rec_IDs,1) = "," Then Filter_Rec_IDs = Left(Filter_Rec_IDs,LEN(Filter_Rec_IDs)-1)


' If there is already an open filter change ticket for this customer, then we dont post to the API
' All we do is add the remaining filters to FS_ServiceMemosFilterInfo
ExistingTicketNumber = ""

SQLCreateTicket = "SELECT * FROM FS_ServiceMemosFilterInfo WHERE CustID = '" & CustID & "' AND Completed=0"
Set rsCreateTicket = cnnCreateTicket.Execute(SQLCreateTicket)
If Not rsCreateTicket.EOF Then
	ExistingTicketNumber = rsCreateTicket("ServiceTicketID")
End If
rsCreateTicket.Close

Dim ResultResponse
ResultResponse=""

DIM resultID
resultID=0

If ExistingTicketNumber = "" Then

	AccountNumber = CustID
	Company = GetCustNameByCustNum(CustID)
	ProblemDescription = "FILTER CHANGE"
	SumissionSource = "MDS Insight"
	filters = ""
	
	For x = 0 to Ubound(FilterList) -1
		If FilterList(x) <> "" Then
			filters = filters & vbcrlf & "Filter: " & GetFilterIDByIntRecID(FilterList(x)) & " - " & GetFilterDescByIntRecID(FilterList(x))
		End If
	Next

	Description = "OPEN,    "
	Description = Description & "Account: "  & AccountNumber & " - " & Company 
	Description = Description & ",    Description: "  & ProblemDescription 
	CreateAuditLogEntry "Service Memo Added","Service Memo Added","Major",0,Description

	If GetPOSTParams("SERVICEMEMOURL1MPLEXFORMAT") = 1 then 'see if using metroplex format
	
		'Post to MDS goes here
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
	
		If GetPOSTParams("NeverPutOnHold") = 0  Then data = data & "&hld=1" Else data = data & "&hld=0" 'I know it's wierd. It's the opposite of how it is stored in the table
	
	Else
	
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
		'data = data & "<CLIENT_ID>" & MUV_READ("ClientID") & "</CLIENT_ID>"
		data = data & "<USER_NO>" & Session("UserNo") & "</USER_NO>"
		data = data & "<MODE>" & GetPOSTParams("Mode") & "</MODE>"
		data = data & "<SUBMITTED_BY_PHONE>" & ContactPhone & "</SUBMITTED_BY_PHONE>"
		data = data & "<SUBMITTED_BY_EMAIL>" & ContactEmail & "</SUBMITTED_BY_EMAIL>"
		data = data & "<SUBMITTED_BY_NAME>" & ContactEmail & "</SUBMITTED_BY_NAME>"
		data = data & "<FILTER_REC_IDS>" & Filter_Rec_IDs & "</FILTER_REC_IDS>"		
		data = data & "<URGENT>0</URGENT>"			
		data = Replace(data ,"&","&amp;")
		data = Replace(data ,chr(34),"")
		data = data & "</POST_DATA>"


	

	
	End If

	'x=5/0


		DO_Post = 0
		
		Do_Post = GetPOSTParams("ServiceMemoURL1ONOFF") 
				
		If IsNull(Do_Post) or Do_Post = "" Then Do_Post = 0
	
		If cint(Do_Post) = 1 Then
		
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

	If Instr(httpRequest.responseText,"success") <> 0 Then
		ResultResponse="""result"":""0"",""message"":""success"""
		resultID=0
	Else
		ResultResponse="""result"":""1"",""message"":"""+httpRequest.responseText+""""
		resultID=1
	End If
		
	'	Response.Write("<br>XX" &  data & "XX<br>")
	'	Response.Write("<br>XX" &  postResponse & "XX<br>")
	'	Response.end
	
End If

'If there is an exisitng ticket, just add records to FS_ServiceMemoFilterInfo for that ticket
If ExistingTicketNumber <> "" Then

	Set rsForInsert = Server.CreateObject("ADODB.Recordset")
	
	SQLCreateTicket = "SELECT * FROM FS_CustomerFilters WHERE CustID = '" & CustID & "' AND InternalRecordIdentifier NOT IN (SELECT InternalRecordIdentifier FROM FS_ServiceMemosFilterInfo WHERE CustID = '" & CustID & "' AND Completed=0)"
	Set rsCreateTicket = cnnCreateTicket.Execute(SQLCreateTicket)
	
	If Not rsCreateTicket.EOF Then
	
		Do While NOT rsCreateTicket.EOF
		
			
			SQLForInsert = "INSERT INTO FS_ServiceMemosFilterInfo (CustID, ServiceTicketID, CustFilterIntRecID, ICFilterIntRecID, Completed) "
			SQLForInsert = SQLForInsert & " VALUES ( "


			SQLForInsert = SQLForInsert & "'" & CustID & "' "
			SQLForInsert = SQLForInsert & ",'" & ExistingTicketNumber & "' "

			SQLForInsert = SQLForInsert & "," & rsCreateTicket("InternalRecordIdentifier")
			SQLForInsert = SQLForInsert & "," & rsCreateTicket("FilterIntRecID")
			
			SQLForInsert = SQLForInsert & ",0"
			
			SQLForInsert = SQLForInsert & " ) "
			
			
			Set rsForInsert = cnnCreateTicket.Execute(SQLForInsert)

			Description = "Filter added to service ticket # " & ServiceTicketID & " for customer " & CustID  & "    Filter: " & GetFilterIDByIntRecID(rsCreateTicket("FilterIntRecID")) & " - " & GetFilterDescByIntRecID(rsCreateTicket("FilterIntRecID"))
			CreateAuditLogEntry "Service Memo Added","Service Memo Added","Major",0,Description
			
			rsCreateTicket.movenext
		Loop

	End If
	ResultResponse="""result"":""0"",""message"":""success"""
	resultID=0
End If
if resultID=0 THEN

	ActiveFilterTicketNumber = ""
	Set rsActiveTicket = Server.CreateObject("ADODB.Recordset")
	rsActiveTicket.CursorLocation = 3
			
	SQActiveTicket = "SELECT TOP 1 ServiceTicketID FROM FS_ServiceMemosFilterInfo WHERE CustID = '" &  CustID & "' AND "
	SQActiveTicket = SQActiveTicket & " ServiceTicketID IN "
	SQActiveTicket = SQActiveTicket & "(SELECT MemoNumber FROM FS_ServiceMemos WHERE AccountNumber = '" & CustID & "' AND CurrentStatus='OPEN')"
	IF LEN(FilterIntRecID)>0 THEN
		SQActiveTicket = SQActiveTicket & " AND CustFilterIntRecID = '" & FilterIntRecID & "' "
	END IF
	Set rsActiveTicket = cnnCreateTicket.Execute(SQActiveTicket )
	
	If NOT rsActiveTicket.EOF Then 
		ResultResponse="{""resultID"":""" & resultID & """,""ticketNumber"":""" & rsActiveTicket("ServiceTicketID") & """," & ResultResponse & "}"
		ELSE
			ResultResponse="{""resultID"":""" & resultID & """,""ticketNumber"":""""," & ResultResponse & "}"
	END IF
	rsActiveTicket.Close
	
	ELSE 
		ResultResponse="{""resultID"":""" & resultID & """," & ResultResponse & "}"
END IF

Response.Write ResultResponse 

SUB addTicket(custID, filterID) 
	Set cnnAddTicket = Server.CreateObject("ADODB.Connection")
	cnnAddTicket.open (Session("ClientCnnString"))
	Set rsAddTicket = Server.CreateObject("ADODB.Recordset")
	
	DIM varSQL
	
	varSQL="SELECT * FROM FS_ServiceMemosFilterInfo WHERE CustFilterIntRecID=" & filterID & " AND CustID='" & custID & "' AND completed=0"
	rsAddTicket=cnnAddTicket.Execute(varSQL)
	IF rsAddTicket.EOF Then
		rsAddTicket.Close
		varSQL="INSERT INTO FS_ServiceMemosFilterInfo (CustFilterIntRecID, CustID, completed) VALUES (" & filterID & ",'" & custID & "',0)"
		cnnAddTicket.Execute(varSQL)
		ELSE
			rsAddTicket.Close
	END IF
	
	
	cnnAddTicket.Close

END SUB


%>
<!--#include file="../inc/subsandfuncs.asp"-->
<!--#include file="../inc/InsightFuncs.asp"-->

<%

' This is the auto return date advance function
'Looks up the return date info for the customer
'an automatically advances it accordingly
CstNum = Request.Form("Cstnum")
Reason  = Request.Form("rsn")
If Reason = "" Then Reason = 0

'Text for the auto generation of the customer note
Select Case Reason
	Case "119"
		ReasonText = "No order needed"
	Case "120"
		ReasonText = "Left message"
	Case "121"
		ReasonText = "Will call back</option>"
	Case "122"
		ReasonText = "Email sent - last option"
	Case "123"
		ReasonText = "Charge rent - no order"
	Case "124"
		ReasonText = "Received order"
	Case "125"
		ReasonText = "Updated return date"
	Case "126"
		ReasonText = "Change return date to 7/7/77"
	Case "127"
		ReasonText = "Referred to sales department"
	Case Else
		ReasonText = "No reason selected"
End Select



Set cnnrteDte = Server.CreateObject("ADODB.Connection")
cnnrteDte.open (Session("ClientCnnString"))
Set rsrteDte = Server.CreateObject("ADODB.Recordset")
rsrteDte.CursorLocation = 3 

SQL = "Select ReturnDate,ReturnType,ReturnTime From AR_Customer WHERE CustNum= " & CstNum
Set rsrteDte = cnnrteDte.Execute(SQL)

If not rsrteDte.Eof Then
	RDte = rsrteDte("ReturnDate")
	ORDte = rsrteDte("ReturnDate")
	RTyp = Trim(Ucase(rsrteDte("ReturnType")))
	RTim = rsrteDte("ReturnTime")		
End If

set rsrteDte = Nothing
cnnrteDte.Close
Set cnnrteDte = Nothing

NewDte = RDte 
'Now we will see if we can auto advance the date
If RTyp = "D" Then
	NewDte = DateAdd("d",RTim,RDte)
End If
If RTyp = "M" Then
	NewDte = DateAdd("m",RTim,RDte)
End If

If NewDte = RDte Then ' Still the same, could not set

		Description = ""
		Description = Description & "Unable to auto advance return date for account # "  & CstNum 
		Description = Description & "     The original return date is: "  & RDte
		Description = Description & "     The return type is: "  & RTyp 
		Description = Description & "     The return time is: "  & RTim
		 
		CreateAuditLogEntry "Return Date Changed","Return Date Changed","Minor",0,Description

Else ' OK, set it

		rtdte = NewDte 
		
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
		
		SQL = "UPDATE AR_Customer Set ReturnDate = '" & rtdte & "' WHERE CustNum= " & CstNum
		Set rsrteDte = cnnrteDte.Execute(SQL)
		
		set rsrteDte = Nothing
		cnnrteDte.Close
		Set cnnrteDte = Nothing
		
		Description = ""
		Description = Description & "The return date was changed via auto advance for account # "  & CstNum 
		Description = Description & "     The new return date is: "  & rtdte
		Description = Description & "     The original return date was: "  & ORDte 
		 
		CreateAuditLogEntry "Return Date Changed","Return Date Changed","Minor",0,Description
		
		'**************************************************************
		'Before we run the post, make a new entry in the customer notes
		'**************************************************************
		SQL = "INSERT INTO tblCustomerNotes (CustNum,Note,UserNo,Sequence,Sticky,ExpirationDate) "
		SQL = SQL &  " VALUES (" 
		SQL = SQL & "'"  & CstNum & "'"
		SQL = SQL & ",'Changed return date from " & FormatDateTime(ORDte,2) & " to " & FormatDateTime(NewDte,2) & ". Reason: "  & ReasonText & "'"
		SQL = SQL & ","  & Session("UserNo") & ",0," & 0 & ", '"
		SQL = SQL & DateAdd("d",60,Now()) & "')"
			
		Set cnn8 = Server.CreateObject("ADODB.Connection")
		cnn8.open (Session("ClientCnnString"))
		
		Set rs8 = Server.CreateObject("ADODB.Recordset")
		rs8.CursorLocation = 3 
		Set rs8 = cnn8.Execute(SQL)
		set rs8 = Nothing
		
		Description = ""
		Description = Description & "A new note was added to account # "  & CstNum 
		Description = Description & "     The text of the note is as follows: Changed return date from " & FormatDateTime(ORDte,2) & " to " & FormatDateTime(NewDte,2) & ". Reason: "  & ReasonText 

		 
		CreateAuditLogEntry "Account Note Added","Account Note Added","Minor",0,Description

		'******************************************************************
		'END Before we run the post, make a new entry in the customer notes
		'******************************************************************

		
'Post to MDS goes here
'First post which changes the rerun date
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

		Set httpRequest = Nothing

'Second post which creates a memo
		data = "<DATASTREAM>"
		data = data & "<IDENTITY>Pm8316wyc011</IDENTITY>"
		data = data & "<MODE>" & GetPOSTParams("Mode") & "</MODE>"
		data = data & "<RECORD_TYPE>MEMO</RECORD_TYPE>"
		data = data & "<RECORD_SUBTYPE>CREATE_MEMO</RECORD_SUBTYPE>"
		data = data & "<CLIENT_ID>CCS</CLIENT_ID>"
		data = data & "<SERNO>" & GetPOSTParams("SERNO") & "</SERNO>"
		data = data & "<SUBMISSION_SOURCE>MDS Insight</SUBMISSION_SOURCE>"
		data = data & "<ACCOUNT_NUM>" & CstNum & "</ACCOUNT_NUM>"
		data = data & "<FIELD_DATA>" & Reason & "</FIELD_DATA>"
		data = data & "<FIELD_DATA1>" & ReasonText & "</FIELD_DATA1>"
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

		Set httpRequest = Nothing
			
End If
%>

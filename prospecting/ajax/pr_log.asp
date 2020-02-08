<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<!--#include file="../../inc/InSightFuncs_Prospecting.asp"-->
<!--#include file="../../inc/SubsAndFuncs.asp"-->
<% If Session("Userno") = "" Then Response.End() %>
<%
Response.ContentType = "application/json"
ProspectIntRecID = Request.QueryString("i") 
If ProspectIntRecID = "" Then Response.End()

Set cnn = Server.CreateObject("ADODB.Connection")
cnn.open (Session("ClientCnnString"))



If Request.Form("updateAction")="save" Then

	'***************************************************************************************
	'Lookup the record as it exists now so we can fillin the audit trail
	'***************************************************************************************
	SQL = "SELECT * FROM PR_ProspectNotes WHERE InternalRecordIdentifier='"&Request.Form("updateActionId")&"'"		
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
		
	If not rs.EOF Then
	
		Orig_Note = rs("Note")

	End If

	'***************************************************************************************
	'After SQL update, record entries in audit trail
	'***************************************************************************************
	
	Note 		= Request.Form("LogNote")

	If Orig_Note  <> Note Then
	
		Description =  "Prospect note changed from " & Orig_Note  & " to " & Note  & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
		CreateAuditLogEntry GetTerm("Prospecting") & " note change ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description =  "Note changed from <em><strong>" & Orig_Note  & "</em></strong> to <em><strong>" & Note & "</em></strong>"
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")

	End If


	Query = "UPDATE PR_ProspectNotes SET Note='"&EscapeSingleQuotes(Request.Form("LogNote"))&"', "
	Query = Query & " Sticky='"&Request.Form("Sticky")&"' WHERE InternalRecordIdentifier='"&Request.Form("updateActionId")&"'"
	cnn.Execute(Query)
End If





If Request.Form("updateAction")="insert" Then

	Note = Request.Form("LogNote")
	
	Description = "A note " & Note & " was created for prospect " & GetProspectNameByNumber(ProspectIntRecID) 
	CreateAuditLogEntry GetTerm("Prospecting") & " prospect note added ",GetTerm("Prospecting"),"Minor",0,Description
	
	Description = "The note <em><strong> " & Note  & "</em></strong> was added."
	Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")


	Query = "INSERT INTO PR_ProspectNotes (ProspectIntRecID, DateAndTime, EnteredByUserNo, Note, Sticky) "
	Query = Query & "VALUES (" & ProspectIntRecID & ", getdate(), " &Session("Userno")& ", "
	Query = Query & "'"&EscapeSingleQuotes(Request.Form("LogNote"))&"', '"&Request.Form("Sticky")&"')"
	cnn.Execute(Query)
	
End If




If Request.Form("updateAction")="delete" Then

	SQL = "SELECT * FROM PR_ProspectNotes WHERE InternalRecordIdentifier='"&Request.Form("updateActionId")&"'"		
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
		
	If not rs.EOF Then
		ProspectIntRecID = rs("ProspectIntRecID")
		DateAndTime		= rs("DateAndTime")
		EnteredByUserNo	= rs("EnteredByUserNo")
		Note = rs("Note")
		UserName = GetUserDisplayNameByUserNo(rs("EnteredByUserNo"))
	End If
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	Description = "The note " & Note  & ", originally entered by " & UserName & " on " & FormatDateTime(DateAndTime,2) & " for prospect " & GetProspectNameByNumber(ProspectIntRecID) & " was deleted."
	CreateAuditLogEntry GetTerm("Prospecting") & " prospect note deleted ",GetTerm("Prospecting"),"Minor",0,Description
	
	Description = "The note <em><strong> " & Note  & "</em></strong>, originally entered by " & UserName & " on " & FormatDateTime(DateAndTime,2) & " for this prospect was deleted."
	Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")


	Query = "DELETE FROM PR_ProspectNotes WHERE InternalRecordIdentifier='"&Request.Form("updateActionId")&"'"
	cnn.Execute(Query)
	
End If




If Request.Form("updateAction")="Sticky-1" Then

	If Request.Form("updateLogNoteType") = "Email" Then
		Query = "UPDATE PR_ProspectEmailLog SET Sticky=1 WHERE InternalRecordIdentifier='"&Request.Form("updateActionId")&"'"
	Else
		Query = "UPDATE PR_ProspectNotes SET Sticky=1 WHERE InternalRecordIdentifier='"&Request.Form("updateActionId")&"'"
	End If
	cnn.Execute(Query)
	
End If


If Request.Form("updateAction")="Sticky-0" Then

	If Request.Form("updateLogNoteType") = "Email" Then
		Query = "UPDATE PR_ProspectEmailLog SET Sticky=0 WHERE InternalRecordIdentifier='"&Request.Form("updateActionId")&"'"
	Else
		Query = "UPDATE PR_ProspectNotes SET Sticky=0 WHERE InternalRecordIdentifier='"&Request.Form("updateActionId")&"'"
	End If
	cnn.Execute(Query)
	
End If


Query = "SELECT DetailType, id, DateAndTime, EnteredByUserNo, Note, Sticky, StageNumber, ActivityStatus, ActivityNumber FROM "
Query = Query & "(SELECT 'Note' AS DetailType, InternalRecordIdentifier as id, DateAndTime, EnteredByUserNo, Note, Sticky, 0 as StageNumber, '' AS ActivityStatus, 0 as ActivityNumber "
Query = Query & "FROM PR_ProspectNotes WHERE ProspectIntRecID = " & ProspectIntRecID & " "
Query = Query & "UNION "
Query = Query & "SELECT 'Activity' AS DetailType, InternalRecordIdentifier as id, RecordCreationDateTime, StatusChangedByUserNo, Notes AS Expr2, 0 AS Expr3, 0 as StageNumber, Status, ActivityRecID "
Query = Query & "FROM PR_ProspectActivities WHERE ProspectRecID = " & ProspectIntRecID & " "
Query = Query & "UNION "
Query = Query & "SELECT 'Stage Change' AS DetailType,InternalRecordIdentifier as id, RecordCreationDateTime, StageChangedByUserNo As UserNo, Notes AS Expr2, 0 AS Expr3, StageRecID AS StageNumber, '' AS ActivityStatus, 0 as ActivityNumber   "
Query = Query & "FROM PR_ProspectStages AS PR_ProspectStages_1 WHERE ProspectRecID = " & ProspectIntRecID & " "
Query = Query & "UNION "
Query = Query & "SELECT 'Email' AS DetailType,InternalRecordIdentifier as id, RecordCreationDateTime, 0 As UserNo, '' AS Expr2, Sticky, 0 AS StageNumber, '' AS ActivityStatus, 0 as ActivityNumber   "
Query = Query & "FROM PR_ProspectEmailLog AS PR_ProspectEmailLog1_1 WHERE ProspectRecID = " & ProspectIntRecID & ") AS t1 "
Query = Query & "ORDER BY Sticky DESC, DateAndTime DESC"

'response.write(Query)
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn.Execute(Query)

Response.Write("[")
If not rs.EOF Then
	sep = ""
	Do While Not rs.EOF 
			Response.Write(sep)
			sep = ","
			Response.Write("{")
			Response.Write("""id"":""" & EscapeQuotes(rs("id")) & """")
			Response.Write(",""LogDetailType"":""" & EscapeQuotes(rs("DetailType")) & """")
			Response.Write(",""Date"":""" & EscapeQuotes(FormatDateTime(rs("DateAndTime"),2)) & """")
			Response.Write(",""Time"":""" & EscapeQuotes(FormatDateTime(rs("DateAndTime"),3)) & """")
						
			If rs("DetailType") = "Stage Change" Then
			
				Response.Write(",""User"":""" & EscapeQuotes(GetUserDisplayNameByUserNo(rs("EnteredByUserNo"))) & """")
			
				If rs("Note") = "" OR IsNull(rs("Note")) OR IsEmpty(rs("Note")) Then
					LogNoteVar = "Stage changed to " & GetStageByNum(rs("StageNumber")) & ". " & rs("Note")
					Response.Write(",""LogNote"":""" & LogNoteVar & """")
				Else
					LogNoteVar = "Stage changed to " & GetStageByNum(rs("StageNumber")) & ". " & rs("Note")
					Response.Write(",""LogNote"":""" & EscapeQuotes(LogNoteVar) & """")
				End If
				
				
			ElseIf rs("DetailType") = "Email" Then
			
				FromAddress = GetFromAddressByRecID(rs("id"))
				EnteredByUserNo = GetUserNoByEmailAddress(FromAddress)
				
				If EnteredByUserNo <> "" Then
					Response.Write(",""User"":""" & EscapeQuotes(GetUserDisplayNameByUserNo(EnteredByUserNo)) & """")
				Else
					Response.Write(",""User"":""" & EscapeQuotes(FromAddress) & """")
				End If
				
				
				SQLEmail = "SELECT * FROM PR_ProspectEmailLog WHERE InternalRecordIdentifier='"& rs("id") &"'"		
				Set cnnEmail = Server.CreateObject("ADODB.Connection")
				cnnEmail.open (Session("ClientCnnString"))
				Set rsEmail = Server.CreateObject("ADODB.Recordset")
				rsEmail.CursorLocation = 3 
				Set rsEmail = cnnEmail.Execute(SQLEmail)
					
				If not rsEmail.EOF Then
					EmailTo = rsEmail("to_addr")
					EmailCC	= rsEmail("cc_addr")
					EmailBcc = rsEmail("bcc_addr")
					EmailDateTime = rsEmail("EmailDateTime")
					EmailSubject = rsEmail("sub")
					EmailBodyText  = rsEmail("body_text")
				End If
				
				set rsEmail = Nothing
				cnnEmail.close
				set cnnEmail = Nothing
				
				EmailDateTime = cDate(EmailDateTime)
				
				EmailSentDate = EscapeQuotes(FormatDateTime(EmailDateTime,2))
				EmailSentTime = EscapeQuotes(FormatDateTime(EmailDateTime,3))


				EmailBodyText = Replace(EmailBodyText,chr(13),"<br>")  
				EmailBodyText = Hacker_Filter2(EmailBodyText)
				'EmailBodyText = "test"
				EmailBodyText = Server.HTMLEncode(EmailBodyText)
				EmailBodyText = StripSpecialChar(EmailBodyText)
				
				Note = "Sent to <a href='mailto:" & EmailTo & "' class='address'>" & EmailTo & "</a> at " & EmailSentDate & " " & EmailSentTime & " . <strong>Subject</strong>: " & EmailSubject & " <strong>Message</strong><em>: " & EmailBodyText & "</em>"


				Response.Write(",""LogNote"":""" & EscapeQuotes(Note) & """")
				
				
			ElseIf rs("DetailType") = "Activity" Then
			
				If rs("EnteredByUserNo") <> "" Then
					Response.Write(",""User"":""" & EscapeQuotes(GetUserDisplayNameByUserNo(rs("EnteredByUserNo"))) & """")
				Else
					Response.Write(",""User"":""" & EscapeQuotes(GetUserDisplayNameByUserNo(GetActivityCreatedByUserNo(rs("id")))) & """")
				End If
				
				If rs("ActivityNumber") = "" OR IsNull(rs("ActivityNumber")) OR IsEmpty(rs("ActivityNumber")) Then
					ActivityNumber = 0
				Else
					ActivityNumber = rs("ActivityNumber")
				End If

			
				If Not IsNull(rs("ActivityStatus")) Then 
					ActivityVar = "The activity, " & GetActivityByNum(ActivityNumber) & ", has been marked as " & rs("ActivityStatus") & " "
					If rs("Note") <> "" Then
						ActivityVar = ActivityVar & " with the following notes: " & rs("Note")
					End If
					Response.Write(",""LogNote"":""" & ActivityVar & """")
				Else
					If ActivityNumber = 0 Then
						ActivityVar = GetActivityByNum(ActivityNumber)
					Else
						ActivityVar =  GetActivityByNum(ActivityNumber)
					End If
					
					Response.Write(",""LogNote"":""" & ActivityVar & """")
				End If
					
			Else
			
				Response.Write(",""User"":""" & EscapeQuotes(GetUserDisplayNameByUserNo(rs("EnteredByUserNo"))) & """")
			
				If rs("Note") = "" OR IsNull(rs("Note")) OR IsEmpty(rs("Note")) Then
					Response.Write(",""LogNote"":""" & rs("Note") & """")
				Else
					Response.Write(",""LogNote"":""" & EscapeQuotes(rs("Note")) & """")
				End If
				
			End If
			Response.Write(",""LogNoteType"":""" & EscapeQuotes(NoteType) & """")
			If rs("Sticky") = 1  Then
				Response.Write(",""Sticky"":1")
			Else
				Response.Write(",""Sticky"":0")
			End If
			Response.Write("}")

		rs.MoveNext						
	Loop
End If
Response.Write("]")
Set rs = Nothing
cnn.Close
Set cnn = Nothing

Function EscapeQuotes(val)
	EscapeQuotes = Replace(val, """", "\""")
End Function
Function EscapeSingleQuotes(val)
	EscapeSingleQuotes = Replace(val, "'", "''")
End Function

Function StripSpecialChar(strInput)
	Dim objRegExp
	Set objRegExp = New Regexp
	objRegExp.IgnoreCase = True
	objRegExp.Global = True
	objRegExp.Pattern = "[\+\#\@\&\%\$\?\*]"
	'//Replace all character matches with an Empty String
	StripSpecialChar = objRegExp.Replace(strInput, "")  
	Set objRegExp = Nothing
End Function
%> 

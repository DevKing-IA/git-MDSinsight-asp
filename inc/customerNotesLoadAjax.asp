<!--#include file="InSightFuncs.asp"-->
<!--#include file="InSightFuncs_Users.asp"-->
<!--#include file="InsightFuncs_AR_AP.asp"-->

<% If Session("Userno") = "" Then Response.End() %>
<%
Response.ContentType = "application/json"

CustID = Request.QueryString("custID") 
notesall = Request.QueryString("notesall")
noteTypeIntRecIDPassed = Request.QueryString("notetypeintrecid")

'************************************************************************************************************************
'The code below is meant to fix conflicts between the note ID on the ALL tab and the note ID on its tab by note type
'If a note has 5 zeros appended to its right side, then it was being updated from the ALL NOTES tab, which
'means we have to strip off the 5 right zeros to get the correct note internal record identifier for AR_CustomerNotes
'************************************************************************************************************************

UpdateActionID = Request.Form("updateActionId")

If InStr(UpdateActionID, "000000000") Then
	UpdateActionID = Left(Request.Form("updateActionId"), Len(Request.Form("updateActionId"))-9)
End If	


If CustID = "" Then Response.End()

Set cnnCatNote = Server.CreateObject("ADODB.Connection")
cnnCatNote.open (Session("ClientCnnString"))
Set rsCatNote = Server.CreateObject("ADODB.Recordset")
rsCatNote.CursorLocation = 3 




If Request.Form("updateAction")="save" Then

	'***************************************************************************************
	'Lookup the record as it exists now so we can fillin the audit trail
	'***************************************************************************************
	SQL = "SELECT * FROM AR_CustomerNotes WHERE InternalRecordIdentifier='" & UpdateActionID & "'"		
	Set cnnCatNote8 = Server.CreateObject("ADODB.Connection")
	cnnCatNote8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnnCatNote8.Execute(SQL)
		
	If not rs.EOF Then
		Orig_Note = rs("Note")
		Orig_NoteTypeIntRecID = rs("NoteTypeIntRecID")
		NoteInternalRecordIdentifier = rs("InternalRecordIdentifier")
	End If

	'***************************************************************************************
	'After SQL update, record entries in audit trail
	'***************************************************************************************
	
	Note = Request.Form("CustomerNote")
	NoteTypeIntRecID = Request.Form("NoteType")
		
	CustomerName = GetCustNameByCustNum(CustID)
	
	UserNo = Session("UserNo")
	UserName = GetUserDisplayNameByUserNo(UserNo)

	If Orig_Note <> Note Then
	
		Description = "The customer note with record ID " & NoteInternalRecordIdentifier & " changed from " & Orig_Note  & " to " & Note & ", by " & UserName & " on " & FormatDateTime(Now(),2) & ", "
		Description = Description & "for customer " & CustomerName & "(" & CustID & ")."
		CreateAuditLogEntry "Customer Note Edited",GetTerm("Accounts Receivable"),"Minor",0,Description

	End If
	
	If Orig_NoteTypeIntRecID <> NoteTypeIntRecID Then
	
		Description = "The customer note with record ID " & NoteInternalRecordIdentifier & " changed note type from " & GetCustNoteTypeByNoteIntRecID(Orig_NoteTypeIntRecID) & " to " & GetCustNoteTypeByNoteIntRecID(NoteTypeIntRecID) & ", by " & UserName & " on " & FormatDateTime(Now(),2) & ", "
		Description = Description & "for customer " & CustomerName & "(" & CustID & ")."
		CreateAuditLogEntry "Customer Note Edited",GetTerm("Accounts Receivable"),"Minor",0,Description

	End If
	
	
	Query = "UPDATE AR_CustomerNotes SET Note='"&EscapeSingleQuotes(Note)&"', NoteTypeIntRecID="&EscapeSingleQuotes(NoteTypeIntRecID)&" WHERE InternalRecordIdentifier='"&UpdateActionID&"'"
	
	'Response.Write(Query)
	
	Set rsCatNote = cnnCatNote.Execute(Query)

End If





If Request.Form("updateAction")="insert" Then

	Note = Request.Form("CustomerNote")
	NoteTypeIntRecID = Request.Form("NoteType")
	
	CustomerName = GetCustNameByCustNum(CustID)
	
	UserNo = Session("UserNo")
	UserName = GetUserDisplayNameByUserNo(UserNo)

	Description = "The a new note, " & Note  & ", was created by " & UserName & " on " & FormatDateTime(Now(),2) & ", for customer " & CustomerName & "(" & CustID & ")."
	CreateAuditLogEntry "Customer Note Created",GetTerm("Accounts Receivable"),"Minor",0,Description

	Query = "INSERT INTO AR_CustomerNotes (CustID, Category, EnteredByUserNo, Note, NoteTypeIntRecID) "
	Query = Query & "VALUES ('" & CustID & "',-2," & Session("Userno") & ",'" & EscapeSingleQuotes(Note) & "'," & EscapeSingleQuotes(NoteTypeIntRecID) & ")"
	
	'Response.Write(Query)
	
	Set rsCatNote = cnnCatNote.Execute(Query)
	
	Call MarkNewNoteNoteTypeForUserAsRead(NoteTypeIntRecID, CustID)
	
End If




If Request.Form("updateAction")="delete" Then

	SQL = "SELECT * FROM AR_CustomerNotes WHERE InternalRecordIdentifier='"&UpdateActionID&"'"		
	Set cnnCatNote8 = Server.CreateObject("ADODB.Connection")
	cnnCatNote8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnnCatNote8.Execute(SQL)
		
	If not rs.EOF Then
		InternalRecordIdentifier = rs("InternalRecordIdentifier")
		RecordCreationDateTime = rs("RecordCreationDateTime")
		CustID = rs("CustID")
		CustomerName = GetCustNameByCustNum(CustID)
		EnteredByUserNo	= rs("EnteredByUserNo")
		Note = rs("Note")
		NoteTypeIntRecID = rs("NoteTypeIntRecID")
		UserName = GetUserDisplayNameByUserNo(rs("EnteredByUserNo"))
	End If
	
	set rs = Nothing
	cnnCatNote8.close
	set cnnCatNote8 = Nothing
	
	Description = "The " & GetCustNoteTypeByNoteIntRecID(NoteTypeIntRecID) & " note, " & Note  & ", originally entered by " & UserName & " on " & FormatDateTime(RecordCreationDateTime,2) & " for customer " & CustomerName & "(" & CustID & ") was deleted."
	CreateAuditLogEntry "Customer Note Deleted ",GetTerm("Accounts Receivable"),"Minor",0,Description

	Query = "DELETE FROM AR_CustomerNotes WHERE InternalRecordIdentifier='"&UpdateActionID&"'"
	
	Set rsCatNote = cnnCatNote.Execute(Query)
	
	'**************************************************************************************************************
	'If the note being deleted, is the last note of that note type, then delete corresponding records from
	'AR_CustomerNotesUserViewed to ensure read/unread notifications are correct
	'**************************************************************************************************************
	
	Query = "SELECT COUNT(*) AS NoteCount FROM AR_CustomerNotes WHERE CustID = '" & CustID  & "' AND NoteTypeIntRecID = " & NoteTypeIntRecID
	
	Set rsCatNote = cnnCatNote.Execute(Query)

	If NOT rsCatNote.EOF Then
		If rsCatNote("NoteCount") < 1 Then
			Query = "DELETE FROM AR_CustomerNotesUserViewed WHERE CustID = '" & CustID  & "' AND NoteTypeIntRecID = " & NoteTypeIntRecID & " AND UserNo = " & Session("Userno")
			Set rsCatNote = cnnCatNote.Execute(Query)
		End If
	Else
		Query = "DELETE FROM AR_CustomerNotesUserViewed WHERE CustID = '" & CustID  & "' AND NoteTypeIntRecID = " & NoteTypeIntRecID & " AND UserNo = " & Session("Userno")
		Set rsCatNote = cnnCatNote.Execute(Query)
	End If
	
	
End If



'************************************************* 
'NoteTypes By IntRecID - to show only the notes
' for each tab that match that tab note type
'************************************************* 
'(1) General
'(2) Backend
'(3) System
'(4) MCS
'(5) Service
'(6) A/R
'(7) CRM
'************************************************* 

If noteTypeIntRecIDPassed <> "" Then

	If noteTypeIntRecIDPassed = 0 Then 'SHOW ALL NOTES
	
		If notesall = "true" Then
			WhereClause = " AND EnteredByUserNo = " & Session("UserNo") & " AND NoteTypeIntRecID <> '' "
		Else 
			WhereClause = " AND NoteTypeIntRecID <> '' "
		End If
		
	Else
	
		If notesall = "true" Then
			WhereClause = " AND EnteredByUserNo = " & Session("UserNo") & " AND NoteTypeIntRecID = " & noteTypeIntRecIDPassed & " "
		Else 
			WhereClause = " AND NoteTypeIntRecID = " & noteTypeIntRecIDPassed & " "
		End If
	
	End If

Else
	If notesall = "true" Then
		WhereClause = " AND EnteredByUserNo = " & Session("UserNo") & " "
	Else 
		WhereClause = ""
	End If

End IF

hasReason = 0

Query = "SELECT InternalRecordIdentifier, RecordCreationDateTime, CustID, Category, "
Query = Query & " EnteredByUserNo, Note, NoteType, NoteTypeIntRecID, MCSReasonIntRecID "
Query = Query & " FROM AR_CustomerNotes  "
Query = Query & " WHERE CustID = '"& CustID & "' " & WhereClause				
Query = Query & " ORDER BY RecordCreationDateTime DESC"

'Response.Write(Query)

Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnnCatNote.Execute(Query)

Response.Write("[")
If not rs.EOF Then
	sep = ""
	Do While Not rs.EOF		
			Response.Write(sep)
			sep = ","
			Response.Write("{")
			Response.Write("""id"":""" & EscapeQuotes(rs("InternalRecordIdentifier")) & """")
			Response.Write(",""query"":""" & EscapeQuotes(Query) & """")
			Response.Write(",""NoteTypeGetTerm"":""" & EscapeQuotes(GetTerm(GetCustNoteTypeByNoteIntRecID(rs("NoteTypeIntRecID")))) & """")
			Response.Write(",""NoteTypeIntRecID"":""" & EscapeQuotes(rs("NoteTypeIntRecID")) & """")
			Response.Write(",""NoteTypeCanBeCreatedByUser"":""" & EscapeQuotes(GetCustNoteTypeCanBeEdited(rs("NoteTypeIntRecID"))) & """")
			Response.Write(",""Date"":""" & EscapeQuotes(FormatDateTime(rs("RecordCreationDateTime"),2)) & """")
			Response.Write(",""Time"":""" & EscapeQuotes(FormatDateTime(rs("RecordCreationDateTime"),3)) & """")
			Response.Write(",""User"":""" & EscapeQuotes(GetUserDisplayNameByUserNo(rs("EnteredByUserNo"))) & """")
			
			If rs("Note") = "" OR IsNull(rs("Note")) OR IsEmpty(rs("Note")) Then
				Response.Write(",""CustomerNote"":""" & EscapeCustomerNote(rs("Note")) & """")
			Else
				Response.Write(",""CustomerNote"":""" & EscapeCustomerNote(rs("Note")) & """")
			End If
			
			if rs("MCSReasonIntRecID") > 0 Then
				Query2 = "SELECT Reason FROM BI_MCSReasons WHERE InternalRecordIdentifier=" & rs("MCSReasonIntRecID") 
				Set rs1 = cnnCatNote.Execute(Query2)
				If not rs1.EOF Then
					Response.Write(",""Reason"":""" & EscapeQuotes(rs1("Reason")) & """")
					hasReason = 1
				Else
					hasReason = 0
				End If
				rs1.close()				
			End If
			Response.Write(",""hasReason"":" & hasReason)
			Response.Write("}")

		rs.MoveNext						
	Loop
End If
Response.Write("]")

cnnCatNote.Close
Set cnnCatNote = Nothing

Function EscapeNewLine(val)
	EscapeNewLine = Replace(val, vbcr, "\r")
	EscapeNewLine = Replace(EscapeNewLine, vblf, "\n")
End Function
Function EscapeQuotes(val)
	EscapeQuotes = Replace(val, """", "\""")
End Function
Function EscapeSingleQuotes(val)
	EscapeSingleQuotes = Replace(val, "'", " ")
End Function

Function EscapeCustomerNote(val)
	'New Line
	EscapedNote = val
	
	EscapedNote = Replace(EscapedNote, vbcr, "\r")
	EscapedNote = Replace(EscapedNote, vblf, "\n")
	'Double Quotes
	EscapedNote = Replace(EscapedNote, """", "\""")
	'Single Quotes
	EscapedNote = Replace(EscapedNote, "'", " ")
	'Ampersand
	EscapedNote = Replace(EscapedNote, "&", "and")
	
	EscapeCustomerNote = EscapedNote
	
End Function


%> 
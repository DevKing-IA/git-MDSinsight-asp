<!--#include file="../inc/InsightFuncs.asp"-->
<!--#include file="../inc/InSightFuncs_Users.asp"-->

<%
InternalNoteNumber = Request.QueryString("nt")

If InternalNoteNumber  <> "" Then
	
	SQLArchiveNote = "Select * FROM tblCustomerNotes WHERE InternalNoteNumber = "& InternalNoteNumber 
	
	Set cnnArchiveNote = Server.CreateObject("ADODB.Connection")
	cnnArchiveNote.open (Session("ClientCnnString"))
	Set rsArchiveNote = Server.CreateObject("ADODB.Recordset")
	rsArchiveNote.CursorLocation = 3 
	Set rsArchiveNote = cnnArchiveNote.Execute(SQLArchiveNote)
	
	If not rsArchiveNote.eof then
		heldCustNum = rsArchiveNote ("CustNum")
		heldNote = rsArchiveNote ("Note")
		heldUserNo = rsArchiveNote ("UserNo")
	End If

	
	Description = ""
	Description = Description & "An account note was archives on account # "  & heldCustNum 
	Description = Description & "     The text of the note was as follows: "  & heldNote 
	Description = Description & "     The note was originally created by "  & GetUserDisplayNameByUserNo(heldUserNo) 
 
	CreateAuditLogEntry "Account Note Archived","Account Note Archived","Minor",0,Description


	SQLArchiveNote = "UPDATE tblCustomerNotes Set Archived = 1, Sticky = 0, ArchivedByUserNo = " & Session("UserNo") & " WHERE InternalNoteNumber = "& InternalNoteNumber 'Archive cant be sticky
	Set rsArchiveNote = cnnArchiveNote.Execute(SQLArchiveNote)
	
	set rsArchiveNote = Nothing
	cnnArchiveNote.Close
	set cnnArchiveNote = Nothing
	
End If

Response.Redirect ("main.asp#Archived")
%>
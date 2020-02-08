<!--#include file="../inc/subsandfuncs.asp"-->
<!--#include file="../inc/InSightFuncs.asp"-->
<!--#include file="../inc/InSightFuncs_Prospecting.asp"-->
<%
' add note to multiple prospects
ids = Request.Form("addnotesmultipleids")

If ids<>"" Then
	temparr = Split(ids,",")	
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	
	Note = Request.Form("txtProspectingNote")
	
	For t = 0 To Ubound(temparr)
	ProspectIntRecID = temparr(t)
	
	If IsNumeric(ProspectIntRecID) AND ProspectIntRecID<>"" AND Note<>"" Then
		
	NoteType = "Note"
	sticky = 0
	LogNoteTypeNumber = ""
	
	Description = "A note " & Note & " of type " & NoteType & ", was created for prospect " & GetProspectNameByNumber(ProspectIntRecID) 
	CreateAuditLogEntry GetTerm("Prospecting") & " prospect note added ",GetTerm("Prospecting"),"Minor",0,Description
	
	Description = "The note <em><strong> " & Note  & "</em></strong> of type <em><strong>" & NoteType & "</em></strong>, was added."
	Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")


	Query = "INSERT INTO PR_ProspectNotes (ProspectIntRecID, DateAndTime, EnteredByUserNo, NoteTypeNumber, Note, Sticky) "
	Query = Query & "VALUES (" & ProspectIntRecID & ", getdate(), " &Session("Userno")& ", '"&LogNoteTypeNumber&"', "
	Query = Query & "'"&EscapeSingleQuotes(Note)&"', '"&Sticky&"')"
	cnn8.Execute(Query)
	
	End If ' check if numeric id
	
	Next
End If	

Function EscapeQuotes(val)
	EscapeQuotes = Replace(val, """", "\""")
End Function
Function EscapeSingleQuotes(val)
	EscapeSingleQuotes = Replace(val, "'", "''")
End Function
%>

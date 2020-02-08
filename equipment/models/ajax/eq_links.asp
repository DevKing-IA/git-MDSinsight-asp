<!--#include file="../../../inc/InSightFuncs.asp"-->
<!--#include file="../../../inc/InSightFuncs_Users.asp"-->
<!--#include file="../../../inc/InSightFuncs_Equipment.asp"-->
<% If Session("Userno") = "" Then Response.End() %>
<%

Response.ContentType = "application/json"

ModelIntRecID = Request.QueryString("i") 
If ModelIntRecID = "" Then Response.End()

Set cnn = Server.CreateObject("ADODB.Connection")
cnn.open (Session("ClientCnnString"))



If Request.Form("updateActionLinks")="save" Then

	'***************************************************************************************
	'Lookup the record as it exists now so we can fill in the audit trail
	'***************************************************************************************
	SQL = "SELECT * FROM EQ_ModelLinks WHERE InternalRecordIdentifier='"&Request.Form("updateActionIdLinks")&"'"		
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
		
	If not rs.EOF Then
	
		Orig_LinkNote = rs("Note")
		Orig_LinkURL = rs("LinkURL")

	End If

	'***************************************************************************************
	'After SQL update, record entries in audit trail
	'***************************************************************************************
	
	LinkNote = Request.Form("LinkNote")
	LinkNote = Replace(LinkNote,"'","''")
	LinkURL = Request.Form("LinkURL")
	
    Sticky = Request.Form("Sticky")
    If Sticky = "on" Then 
        Sticky = 1 
    Else 
        Sticky = 0 
   End If
		

	If Orig_LinkNote <> LinkNote Then
		Description =  GetTerm("Equipment") & " Model Link Note changed from " & Orig_LinkNote  & " to " & LinkNote  & " for the model, " & GetModelNameByIntRecID(ModelIntRecID)
		CreateAuditLogEntry GetTerm("Equipment") & " Model Link Note change ", GetTerm("Equipment"),"Minor",0,Description
	End If

	If Orig_LinkURL <> LinkURL Then
		Description =  GetTerm("Equipment") & " Model Link URL changed from " & Orig_LinkURL & " to " & LinkURL & " for the model, " & GetModelNameByIntRecID(ModelIntRecID)
		CreateAuditLogEntry GetTerm("Equipment") & " Model Link URL change ", GetTerm("Equipment"),"Minor",0,Description
	End If

	Query = "UPDATE EQ_ModelLinks SET Note ='"& Request.Form("LinkNote") &"', LinkURL='"&EscapeSingleQuotes(LinkURL)&"', "
	Query = Query & " Sticky='"& Sticky &"' WHERE InternalRecordIdentifier='"&Request.Form("updateActionIdLinks")&"'"
	cnn.Execute(Query)

End If





If Request.Form("updateActionLinks")="insert" Then

	LinkNote = Request.Form("LinkNote")
	LinkNote = Replace(LinkNote,"'","''")
	LinkURL = Request.Form("LinkURL")
	
	Sticky = Request.Form("Sticky")
    If Sticky = "on" Then 
        Sticky = 1 
    Else 
        Sticky = 0 
   End If
	
	If LinkNote = "" Then
		Description = "The link " & LinkURL & " for model, " & GetModelNameByIntRecID(ModelIntRecID) & " was uploaded via the documents tab." 
		CreateAuditLogEntry GetTerm("Equipment") & " model link uploaded ", GetTerm("Equipment"),"Minor",0,Description
	Else
		Description = "The link " & LinkURL & ", with the note <em><strong>" & LinkNote & "</em></strong> for model, " & GetModelNameByIntRecID(ModelIntRecID)  & " was uploaded via the documents tab."
		CreateAuditLogEntry GetTerm("Equipment") & " model link uploaded ", GetTerm("Equipment"),"Minor",0,Description	
	End If
		
	Query = "INSERT INTO EQ_ModelLinks (ModelIntRecID, CreatedByUserNo, Note, LinkURL ,Sticky) "
	Query = Query & "VALUES (" & ModelIntRecID & ", " & Session("Userno") & ", '"& EscapeSingleQuotes(LinkNote) &"','" & LinkURL &"', '"& Sticky &"')"
	cnn.Execute(Query)
	
	
End If




If Request.Form("updateActionLinks")="delete" Then

	SQL = "SELECT * FROM EQ_ModelLinks WHERE InternalRecordIdentifier='"&Request.Form("updateActionIdLinksImages")&"'"		
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
		
	If not rs.EOF Then
		ModelIntRecID = rs("ModelIntRecID")
		RecordCreationDateTime	= rs("RecordCreationDateTime")
		CreatedByUserNo	= rs("CreatedByUserNo")
		LinkNote = rs("Note")
		LinkURL = rs("LinkURL")
		UserName = GetUserDisplayNameByUserNo(rs("CreatedByUserNo"))
	End If
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing

	If LinkNote = "" Then
		Description = "The link " & LinkURL & " was deleted for model, " & GetModelNameByIntRecID(ModelIntRecID) 
		CreateAuditLogEntry GetTerm("Equipment") & " model link deleted ", GetTerm("Equipment"),"Minor",0,Description
	Else
		Description = "The link " & LinkURL & ", with the note <em><strong>" & LinkNote & "</em></strong> was deleted from the images for model, " & GetModelNameByIntRecID(ModelIntRecID) 
		CreateAuditLogEntry GetTerm("Equipment") & " model link deleted ", GetTerm("Equipment"),"Minor",0,Description
	End If

	Query = "DELETE FROM EQ_ModelLinks WHERE InternalRecordIdentifier='"&Request.Form("updateActionIdLinks")&"'"
	cnn.Execute(Query)
	

End If


If Request.Form("updateActionLinks")="Sticky-1" Then
	Query = "UPDATE EQ_ModelLinks SET Sticky=1 WHERE InternalRecordIdentifier='"&Request.Form("updateActionIdLinks")&"'"
	cnn.Execute(Query)
End If

If Request.Form("updateActionLinks")="Sticky-0" Then
	Query = "UPDATE EQ_ModelLinks SET Sticky=0 WHERE InternalRecordIdentifier='"&Request.Form("updateActionIdLinks")&"'"
	cnn.Execute(Query)
End If


Query = "SELECT *, InternalRecordIdentifier as ID FROM EQ_ModelLinks WHERE ModelIntRecID = " & ModelIntRecID & " ORDER BY Sticky, RecordCreationDateTime desc"
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
			Response.Write("""id"":""" & EscapeQuotes(rs("ID")) & """")
			Response.Write(",""Date"":""" & EscapeQuotes(FormatDateTime(rs("RecordCreationDateTime"),2)) & """")
			Response.Write(",""Time"":""" & EscapeQuotes(FormatDateTime(rs("RecordCreationDateTime"),3)) & """")		
			Response.Write(",""User"":""" & EscapeQuotes(GetUserDisplayNameByUserNo(rs("CreatedByUserNo"))) & """")
			If rs("Note") = "" OR IsNull(rs("Note")) OR IsEmpty(rs("Note")) Then
				Response.Write(",""LinkNote"":""" & rs("Note") & """")
			Else
				Response.Write(",""LinkNote"":""" & EscapeQuotes(rs("Note")) & """")
			End If
			Response.Write(",""LinkURL"":""" & rs("LinkURL") & """")
			If rs("Sticky") = 1 OR rs("Sticky") = True Then
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

%> 

<!--#include file="../../../inc/InSightFuncs.asp"-->
<!--#include file="../../../inc/InSightFuncs_Users.asp"-->
<!--#include file="../../../inc/InSightFuncs_Equipment.asp"-->
<% If Session("Userno") = "" Then Response.End() %>

<%

ModelIntRecID = Request.QueryString("i") 
If ModelIntRecID = "" Then Response.End()

Set cnn = Server.CreateObject("ADODB.Connection")
cnn.open (Session("ClientCnnString"))


If Request.Form("updateAction")="delete" Then

	SQL = "SELECT * FROM EQ_ModelDocuments WHERE InternalRecordIdentifier='"&Request.Form("updateActionId")&"'"		
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
		
	If not rs.EOF Then
		ModelIntRecID = rs("ModelIntRecID")
		RecordCreationDateTime	= rs("RecordCreationDateTime")
		CreatedByUserNo	= rs("CreatedByUserNo")
		AttachmentNote = rs("Note")
		AttachmentFileName = rs("AttachmentFileName")
		UserName = GetUserDisplayNameByUserNo(rs("CreatedByUserNo"))
			
		If AttachmentNote = "" Then
			Description = "The file " & AttachmentFilename & " was deleted from documents for model, " & GetModelNameByIntRecID(ModelIntRecID) 
			CreateAuditLogEntry GetTerm("Equipment") & " model document deleted ", GetTerm("Equipment"),"Minor",0,Description
		Else
			Description = "The file " & AttachmentFilename & ", with the note <em><strong>" & AttachmentNote & "</em></strong> was deleted from the documents for model, " & GetModelNameByIntRecID(ModelIntRecID) 
			CreateAuditLogEntry GetTerm("Equipment") & " model document deleted ", GetTerm("Equipment"),"Minor",0,Description
		End If
	
	
		Query = "DELETE FROM EQ_ModelDocuments WHERE InternalRecordIdentifier='"&Request.Form("updateActionId")&"'"
		cnn.Execute(Query)
		
	
	    Pth ="../../../clientfiles/" & MUV_READ("CLIENTID") & "/attachments/equipment/models/documents/"
	    dim fs
	    set fs=Server.CreateObject("Scripting.FileSystemObject")
	    
	    'Response.Write("Path: " & Pth & AttachmentFileName)
	    
	    fs.DeleteFile Server.MapPath(Pth & AttachmentFileName)
	    set fs=nothing
		
	End If
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing

	
End If




If Request.Form("updateAction")="Sticky-1" Then
	Query = "UPDATE EQ_ModelDocuments SET Sticky=1 WHERE InternalRecordIdentifier='"&Request.Form("updateActionId")&"'"
	cnn.Execute(Query)
End If
If Request.Form("updateAction")="Sticky-0" Then
	Query = "UPDATE EQ_ModelDocuments SET Sticky=0 WHERE InternalRecordIdentifier='"&Request.Form("updateActionId")&"'"
	cnn.Execute(Query)
End If



Query = "SELECT *, InternalRecordIdentifier as id FROM EQ_ModelDocuments WHERE ModelIntRecID = " & ModelIntRecID & " ORDER BY Sticky, RecordCreationDateTime desc"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn.Execute(Query)

Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn.Execute(Query)

Response.Write("[")
If not rs.EOF Then
	sep = ""
	Do While Not rs.EOF
	
			blankPath = ""
			blankExt = ""
			fileExtention = ""
			
			
			Response.Write(sep)
			sep = ","
			Response.Write("{")
			Response.Write("""id"":""" & EscapeQuotes(rs("id")) & """")
			Response.Write(",""Date"":""" & EscapeQuotes(FormatDateTime(rs("RecordCreationDateTime"),2)) & """")
			Response.Write(",""Time"":""" & EscapeQuotes(FormatDateTime(rs("RecordCreationDateTime"),3)) & """")		
			Response.Write(",""User"":""" & EscapeQuotes(GetUserDisplayNameByUserNo(rs("CreatedByUserNo"))) & """")
			If rs("Note") = "" OR IsNull(rs("Note")) OR IsEmpty(rs("Note")) Then
				Response.Write(",""DocumentNotes"":""" & rs("Note") & """")
			Else
				Response.Write(",""DocumentNotes"":""" & EscapeQuotes(rs("Note")) & """")
			End If
			If rs("AttachmentFileName") = "" OR IsNull(rs("AttachmentFileName")) OR IsEmpty(rs("AttachmentFileName")) Then
				Response.Write(",""DocumentAttachment"":""" & rs("AttachmentFileName") & """")
				Response.Write(",""DocumentPath"":""" & blankPath & """")
				Response.Write(",""DocumentExt"":""" & blankExt & """")
			Else

				fileExtention = Right(rs("AttachmentFileName"),len(rs("AttachmentFileName"))-InstrRev(rs("AttachmentFileName"),"."))
				Select Case fileExtention 
					Case "csv","doc","docx","gif","jpg","jpeg","pdf","png","ppt","pptx","txt","xls","xlsx","zip"
						' Do nothing
					Case Else
						fileExtention = "file"
				End Select

				Response.Write(",""DocumentAttachment"":""" & EscapeQuotes(rs("AttachmentFileName")) & """")
				Response.Write(",""DocumentPath"":""" & BaseURL & "/clientfiles/" & MUV_READ("CLIENTID") & "/attachments/equipment/models/documents/" & rs("AttachmentFileName") & """")
				Response.Write(",""DocumentExt"":""" & fileExtention & """")
			End If
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

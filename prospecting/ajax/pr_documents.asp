<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<!--#include file="../../inc/InSightFuncs_Prospecting.asp"-->
<% If Session("Userno") = "" Then Response.End() %>
<%


Response.ContentType = "application/json"
'Response.ContentType = "multipart/form-data"


Set Upload = Server.CreateObject("Persits.Upload.1")
Upload.OverwriteFiles = False
Upload.IgnoreNoPost = True
Upload.Save 




ProspectIntRecID = Request.QueryString("i") 
If ProspectIntRecID = "" Then Response.End()

Set cnn = Server.CreateObject("ADODB.Connection")
cnn.open (Session("ClientCnnString"))



If Upload.Form("updateAction")="save" Then

	'***************************************************************************************
	'Lookup the record as it exists now so we can fillin the audit trail
	'***************************************************************************************
	SQL = "SELECT * FROM PR_ProspectDocuments WHERE InternalRecordIdentifier='"&Upload.Form("updateActionId")&"'"		
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
		
	If not rs.EOF Then
	
		Orig_AttachmentNote = rs("Note")
		Orig_AttachmentFileName = rs("AttachmentFileName")

	End If

	'***************************************************************************************
	'After SQL update, record entries in audit trail
	'***************************************************************************************
	
	AttachmentNote = Upload.Form("DocumentNotes")
	'AttachmentFileName	= Upload.Form("DocumentAttachment")
	
	'Pth ="/upload/"
	Pth ="../../clientfiles/" & MUV_READ("CLIENTID") & "/attachments/prospecting/"
	
	AttachmentFileName = ""
	NewCompleteName = ""
	
	Set File = Upload.Files("DocumentAttachment")
	If Not File Is Nothing Then
		AttachmentFileName = File.Filename
		fn = File.Filename
		File.SaveAs  Server.MapPath(Pth) & "\" & Upload.Form("updateActionId") & "-" & fn
		NewCompleteName = Upload.Form("updateActionId") & "-" & fn
	End If
	
	

	If Orig_AttachmentNote <> AttachmentNote Then
	
		Description =  "Prospect Attachment Note changed from " & Orig_AttachmentNote  & " to " & AttachmentNote  & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
		CreateAuditLogEntry GetTerm("Prospecting") & " AttachmentNote change ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description =  "Attachment Note changed from <em><strong>" & Orig_AttachmentNote  & "</em></strong> to <em><strong>" & AttachmentNote & "</em></strong>"
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")

	End If

	If Orig_AttachmentFileName <> AttachmentFileName AND NewCompleteName<>"" Then
	
		Description =  "Prospect document file attachment name changed from " & Orig_AttachmentFileName & " to " & NewCompleteName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
		CreateAuditLogEntry GetTerm("Prospecting") & " file attachment name change ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description =  "Prospect document file attachment name change from <em><strong>" & Orig_AttachmentFileName & "</em></strong> to <em><strong>" & NewCompleteName & "</em></strong>"
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")

	End If
	
	If NewCompleteName<>"" Then
	Query = "UPDATE PR_ProspectDocuments SET Note ='"& Upload.Form("DocumentNotes") &"', AttachmentFileName='"&EscapeSingleQuotes(NewCompleteName)&"', "
	Query = Query & " Sticky="&Upload.Form("Sticky")&" WHERE InternalRecordIdentifier='"&Upload.Form("updateActionId")&"' AND ProspectRecID = " & ProspectIntRecID 
	Else
	Query = "UPDATE PR_ProspectDocuments SET Note ='"& Upload.Form("DocumentNotes") &"', "
	Query = Query & " Sticky="&Upload.Form("Sticky")&" WHERE InternalRecordIdentifier='"&Upload.Form("updateActionId")&"' AND ProspectRecID = " & ProspectIntRecID 	
	End If
	cnn.Execute(Query)
End If



If Upload.Form("updateAction")="insert" Then
	
	'***************************************************************************************************************
	'ASP UPLOAD OF FILE TO CLIENTFILES/CLIENTID/ATTACHEMENTS/PROSPECTING/PROSPECTID-FILENAME.EXT
	'***************************************************************************************************************
	
	'Rename the files
	' Construct the save path
	Pth ="../../clientfiles/" & MUV_READ("CLIENTID") & "/attachments/prospecting/"
	
	'Pth ="/upload/"
	
	Set File = Upload.Files("DocumentAttachment")
	If Not File Is Nothing Then
		AttachmentFileName = File.Filename
		fn = File.Filename
		File.SaveAs  Server.MapPath(Pth) & "\" & Upload.Form("updateActionId") & "-" & fn
	End If


	AttachmentNote  = Upload.Form("DocumentNotes")
	'AttachmentFileName = Upload.Form("DocumentAttachment")
	Sticky = Upload.Form("Sticky")

	Description = "The file <em><strong> " & AttachmentFilename & "</em></strong> was uploaded via the documents tab."
	Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
	
	'For Each File in Upload.Files
		'fn=File.FileName
	   	'File.SaveAsVirtual  Pth & Upload.Form("updateActionId") & "-" & fn
	'Next
	
	NewCompleteName = Upload.Form("updateActionId") & "-" & fn 
	
		
	'***************************************************************************************************************
	
	If AttachmentNote = "" Then
	
		Description = "The file " & AttachmentFilename & " was uploaded to documents for prospect " & GetProspectNameByNumber(ProspectIntRecID) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect document uploaded ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description = "The file <em><strong> " & AttachmentFilename & "</em></strong> was uploaded via the documents tab."
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
		
	Else
	
		Description = "The file " & AttachmentFilename & ", with the note <em><strong>" & AttachmentNote & "</em></strong> was uploaded to documents for prospect " & GetProspectNameByNumber(ProspectIntRecID) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect document uploaded ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description = "The file <em><strong> " & AttachmentFilename & "</em></strong>" & ", with the note <em><strong>" & AttachmentNote & "</em></strong> was uploaded via the documents tab."
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
		
	End If


	Query = "INSERT INTO PR_ProspectDocuments (ProspectRecID, CreatedByUserNo, Note, AttachmentFileName ,Sticky) "
	Query = Query & "VALUES (" & ProspectIntRecID & ", " & Session("Userno") & ", '"& EscapeSingleQuotes(AttachmentNote) &"','" & NewCompleteName &"', "& Sticky &")"
	cnn.Execute(Query)
	
End If




If Upload.Form("updateAction")="delete" Then

	SQL = "SELECT * FROM PR_ProspectDocuments WHERE InternalRecordIdentifier='"&Upload.Form("updateActionId")&"'"		
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
		
	If not rs.EOF Then
		ProspectIntRecID = rs("ProspectRecID")
		RecordCreationDateTime	= rs("RecordCreationDateTime")
		CreatedByUserNo	= rs("CreatedByUserNo")
		AttachmentNote = rs("Note")
		AttachmentFileName = rs("AttachmentFileName")
		UserName = GetUserDisplayNameByUserNo(rs("CreatedByUserNo"))
	End If
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	If AttachmentNote = "" Then
	
		Description = "The file " & AttachmentFilename & " was deleted from documents for prospect " & GetProspectNameByNumber(ProspectIntRecID) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect document deleted ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description = "The file <em><strong> " & AttachmentFilename & "</em></strong> was uploaded via the documents tab."
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
		
	Else
	
		Description = "The file " & AttachmentFilename & ", with the note <em><strong>" & AttachmentNote & "</em></strong> was deleted from the documents for prospect " & GetProspectNameByNumber(ProspectIntRecID) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect document deleted ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description = "The file <em><strong> " & AttachmentFilename & "</em></strong>" & ", with the note <em><strong>" & AttachmentNote & "</em></strong> was uploaded via the documents tab."
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
		
	End If


	Query = "DELETE FROM PR_ProspectDocuments WHERE InternalRecordIdentifier='"&Upload.Form("updateActionId")&"'"
	cnn.Execute(Query)
	
End If




If Upload.Form("updateAction")="Sticky-1" Then
	Query = "UPDATE PR_ProspectDocuments SET Sticky=1 WHERE InternalRecordIdentifier='"&Upload.Form("updateActionId") & "'"
	cnn.Execute(Query)
End If
If Upload.Form("updateAction")="Sticky-0" Then
	Query = "UPDATE PR_ProspectDocuments SET Sticky=0 WHERE InternalRecordIdentifier='"&Upload.Form("updateActionId")&"'"
	cnn.Execute(Query)
End If





Query = "SELECT *, InternalRecordIdentifier as id FROM PR_ProspectDocuments WHERE ProspectRecID="&ProspectIntRecID&" ORDER BY Sticky, RecordCreationDateTime desc"
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
			Response.Write("""id"":""" & EscapeQuotes(rs("InternalRecordIdentifier")) & """")
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
				Response.Write(",""DocumentPath"":""" & BaseURL & "/clientfiles/" & MUV_READ("CLIENTID") & "/attachments/prospecting/" & rs("AttachmentFileName") & """")
				Response.Write(",""DocumentExt"":""" & fileExtention & """")
			End If
			If rs("Sticky") = True  Then
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

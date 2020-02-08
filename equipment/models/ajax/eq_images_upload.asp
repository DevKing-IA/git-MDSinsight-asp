<!--#include file="../../../inc/InSightFuncs.asp"-->
<!--#include file="../../../inc/InSightFuncs_Users.asp"-->
<!--#include file="../../../inc/InSightFuncs_Equipment.asp"-->
<% If Session("Userno") = "" Then Response.End() %>
<%

Set cnn = Server.CreateObject("ADODB.Connection")
cnn.open (Session("ClientCnnString"))

Set Upload = Server.CreateObject("Persits.Upload.1")
Upload.OverwriteFiles = True
If (Request.ServerVariables("REQUEST_METHOD") = "POST") Then
	Upload.Save Server.MapPath("../../../clientfiles/" & MUV_READ("CLIENTID") & "/attachments/equipment/models/images/")
	ModelIntRecID = Upload.Form("i")
End If

If ModelIntRecID = "" Then Response.End()

Set cnn = Server.CreateObject("ADODB.Connection")
cnn.open (Session("ClientCnnString"))

dim fs

If Upload.Form("updateActionImages")="save" Then

    ' Construct the save path
	Pth ="../../clientfiles/" & MUV_READ("CLIENTID") & "/attachments/equipment/models/images/"
    For Each File in Upload.Files
		fn=File.FileName
	   	File.SaveAsVirtual  Pth & fn        
	Next
    NewCompleteName = Upload.Form("updateActionIdImages") & "-" & fn

	'***************************************************************************************
	'Lookup the record as it exists now so we can fill in the audit trail
	'***************************************************************************************
	SQL = "SELECT * FROM EQ_ModelImages WHERE InternalRecordIdentifier='"&Upload.Form("updateActionIdImages")&"'"		
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
		
	If not rs.EOF Then
	
		Orig_AttachmentNote = rs("Note")
		Orig_ImageFileName = rs("ImageFileName")

	End If

	'***************************************************************************************
	'After SQL update, record entries in audit trail
	'***************************************************************************************
	
	AttachmentNote = Upload.Form("ImageNotes")
	ImageFileName	= NewCompleteName

	If Orig_AttachmentNote <> AttachmentNote Then
	
		Description =  GetTerm("Equipment") & " Image Attachment Note changed from " & Orig_AttachmentNote  & " to " & AttachmentNote  & " for the model, " & GetModelNameByIntRecID(ModelIntRecID)
		CreateAuditLogEntry GetTerm("Equipment") & " Image Attachment Note change ", GetTerm("Equipment"),"Minor",0,Description
		
	End If
	
	'*************************************************************************************************************
	'NO EDITING OF FILES IS ALLOWED, YOU CAN DELETE AN IMAGE OR CREATE A NEW ONE, BUT YOU CANNOT UPDATE THE FILE
	'*************************************************************************************************************
	
	'If Orig_ImageFileName <> ImageFileName Then
	
		'Description =  GetTerm("Equipment") & " image file attachment name changed from " & Orig_ImageFileName & " to " & ImageFileName & " for the model, " & GetModelNameByIntRecID(ModelIntRecID)
		'CreateAuditLogEntry GetTerm("Equipment") & " image file attachment name change ", GetTerm("Equipment"),"Minor",0,Description

        'set fs=Server.CreateObject("Scripting.FileSystemObject")
        'fs.DeleteFile Server.MapPath(Pth & Orig_ImageFileName)
        'set fs=nothing
	'End If

    Sticky = Upload.Form("Sticky")
    If Sticky = "on" Then 
        Sticky = 1 
    Else 
        Sticky = 0 
   End If
	
	Query = "UPDATE EQ_ModelImages SET Note ='"& Upload.Form("ImageNotes") &"', ImageFileName='"&EscapeSingleQuotes(ImageFileName)&"', "
	Query = Query & " Sticky='"& Sticky &"' WHERE InternalRecordIdentifier='"&Upload.Form("updateActionIdImages")&"' AND ModelIntRecID = '" & ModelIntRecID & "'"
	cnn.Execute(Query)
End If



If Upload.Form("updateActionImages")="insert" Then

	'***************************************************************************************************************
	'ASP UPLOAD OF FILE TO CLIENTFILES/CLIENTID/ATTACHEMENTS/EQUIPMENT/models/images/MODELRECID-FILENAME.EXT
	'***************************************************************************************************************
	
	'Rename the files
	' Construct the save path
	Pth ="../../../clientfiles/" & MUV_READ("CLIENTID") & "/attachments/equipment/models/images/"


	AttachmentNote  = Upload.Form("ImageNotes")
	ImageFileName = Upload.Form("ImageAttachment")
	Sticky = Upload.Form("Sticky")
    If Sticky = "on" Then 
        Sticky = 1 
    Else 
        Sticky = 0 
   End If

	Description = "The " & GetTerm("Equipment") & " model image file <em><strong> " & ImageFileName & "</em></strong> was uploaded via the documents tab."
	CreateAuditLogEntry GetTerm("Equipment") & " image uploaded ", GetTerm("Equipment"),"Minor",0,Description
	
	For Each File in Upload.Files
		fn=File.FileName 
	   	File.SaveAsVirtual  Pth & Upload.Form("updateActionIdImages") & "-" & fn
        NewPartialName = fn       
	Next	
	'***************************************************************************************************************
	
	If AttachmentNote = "" Then
		Description = "The file " & ImageFileName & " was uploaded to images for model, " & GetModelNameByIntRecID(ModelIntRecID) 
		CreateAuditLogEntry GetTerm("Equipment") & " model document uploaded ", GetTerm("Equipment"),"Minor",0,Description
	Else
		Description = "The file " & ImageFileName & ", with the note <em><strong>" & AttachmentNote & "</em></strong> was uploaded to images for model, " & GetModelNameByIntRecID(ModelIntRecID) 
		CreateAuditLogEntry GetTerm("Equipment") & " model image uploaded ", GetTerm("Equipment"),"Minor",0,Description	
	End If


	Query = "INSERT INTO EQ_ModelImages (ModelIntRecID, CreatedByUserNo, Note, ImageFileName ,Sticky) "
	Query = Query & "VALUES (" & ModelIntRecID & ", " & Session("Userno") & ", '"& EscapeSingleQuotes(AttachmentNote) &"','" & Upload.Form("updateActionIdImages") & "-" & NewPartialName &"', '"& Sticky &"')"
	cnn.Execute(Query)

    NewRecordID = 0
    Set rs = cnn.Execute("SELECT @@IDENTITY AS 'Identity'") 
    If Not rs.Eof Then
        NewRecordID = rs("Identity")
    End If

    Query = "UPDATE EQ_ModelImages SET ImageFileName = '" & NewRecordID & "-" & NewPartialName & "' WHERE InternalRecordIdentifier='"& NewRecordID &"' AND ModelIntRecID = '" & ModelIntRecID & "'"
	cnn.Execute(Query)

    set fs=Server.CreateObject("Scripting.FileSystemObject")
    
    fs.CopyFile Server.MapPath(Pth & Upload.Form("updateActionIdImages") & "-" & NewPartialName), Server.MapPath(Pth & NewRecordID & "-" & NewPartialName)
    fs.DeleteFile Server.MapPath(Pth & Upload.Form("updateActionIdImages") & "-" & NewPartialName)
    fs.DeleteFile Server.MapPath(Pth & NewPartialName)
    set fs=nothing

End If



Query = "SELECT *, InternalRecordIdentifier as id FROM EQ_ModelImages ORDER BY Sticky, RecordCreationDateTime desc"
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
				Response.Write(",""ImageNotes"":""" & rs("Note") & """")
			Else
				Response.Write(",""ImageNotes"":""" & EscapeQuotes(rs("Note")) & """")
			End If
			If rs("ImageFileName") = "" OR IsNull(rs("ImageFileName")) OR IsEmpty(rs("ImageFileName")) Then
				Response.Write(",""ImageAttachment"":""" & rs("ImageFileName") & """")
				Response.Write(",""ImagePath"":""" & blankPath & """")
				Response.Write(",""ImageExt"":""" & blankExt & """")
			Else

				fileExtention = Right(rs("ImageFileName"),len(rs("ImageFileName"))-InstrRev(rs("ImageFileName"),"."))
				Select Case fileExtention 
					Case "csv","doc","docx","gif","jpg","jpeg","pdf","png","ppt","pptx","txt","xls","xlsx","zip"
						' Do nothing
					Case Else
						fileExtention = "file"
				End Select

				Response.Write(",""ImageAttachment"":""" & EscapeQuotes(rs("ImageFileName")) & """")
				Response.Write(",""ImagePath"":""" & BaseURL & "/clientfiles/" & MUV_READ("CLIENTID") & "/attachments/equipment/models/images/" & rs("ImageFileName") & """")
				Response.Write(",""ImageExt"":""" & fileExtention & """")
			End If
			If rs("Sticky") = "1" OR rs("Sticky")=True  Then
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

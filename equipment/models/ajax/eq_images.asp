<!--#include file="../../../inc/InSightFuncs.asp"-->
<!--#include file="../../../inc/InSightFuncs_Users.asp"-->
<!--#include file="../../../inc/InSightFuncs_Equipment.asp"-->
<% If Session("Userno") = "" Then Response.End() %>

<%

ModelIntRecID = Request.QueryString("i") 
If ModelIntRecID = "" Then Response.End()

Set cnn = Server.CreateObject("ADODB.Connection")
cnn.open (Session("ClientCnnString"))


If Request.Form("updateActionImages")="delete" Then

	SQL = "SELECT * FROM EQ_ModelImages WHERE InternalRecordIdentifier='"&Request.Form("updateActionIdImages")&"'"
	
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
		ImageFileName = rs("ImageFileName")
		UserName = GetUserDisplayNameByUserNo(rs("CreatedByUserNo"))
	
		If AttachmentNote = "" Then
			Description = "The image file " & ImageFileName & " was deleted for model, " & GetModelNameByIntRecID(ModelIntRecID) 
			CreateAuditLogEntry GetTerm("Equipment") & " model image deleted ", GetTerm("Equipment"),"Minor",0,Description
		Else
			Description = "The image file " & ImageFileName & ", with the note <em><strong>" & AttachmentNote & "</em></strong> was deleted from the images for model, " & GetModelNameByIntRecID(ModelIntRecID) 
			CreateAuditLogEntry GetTerm("Equipment") & " model image deleted ", GetTerm("Equipment"),"Minor",0,Description
		End If
	
	
		Query = "DELETE FROM EQ_ModelImages WHERE InternalRecordIdentifier='"&Request.Form("updateActionIdImages")&"'"
		cnn.Execute(Query)
		
	    Pth ="../../../clientfiles/" & MUV_READ("CLIENTID") & "/attachments/equipment/models/images/"
	    dim fs
	    set fs=Server.CreateObject("Scripting.FileSystemObject")
	    fs.DeleteFile Server.MapPath(Pth & ImageFileName)
	    set fs=nothing
		
	End If
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	
	
End If




If Request.Form("updateActionImages")="Sticky-1" Then
	Query = "UPDATE EQ_ModelImages SET Sticky=1 WHERE InternalRecordIdentifier='"&Request.Form("updateActionIdImages")&"'"
	cnn.Execute(Query)
End If
If Request.Form("updateActionImages")="Sticky-0" Then
	Query = "UPDATE EQ_ModelImages SET Sticky=0 WHERE InternalRecordIdentifier='"&Request.Form("updateActionIdImages")&"'"
	cnn.Execute(Query)
End If



Query = "SELECT *, InternalRecordIdentifier as ID FROM EQ_ModelImages WHERE ModelIntRecID = " & ModelIntRecID & " ORDER BY Sticky, RecordCreationDateTime desc"
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
			Response.Write("""id"":""" & EscapeQuotes(rs("ID")) & """")
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

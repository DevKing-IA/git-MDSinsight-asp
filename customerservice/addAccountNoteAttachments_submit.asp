<!--#include file="../inc/subsandfuncs.asp"-->
<!--#include file="../inc/InsightFuncs.asp"-->

<%

'We do some funky stuff in here temporarily setting the CustID to a random number
'because the form field values don't actually become available until
'AFTER the upload save method is called, so we dont initially have them


Set Upload = Server.CreateObject("Persits.Upload.1")
Upload.OverwriteFiles = False

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 

SQL = "Select FLOOR(RAND()*(25000-10000)+10000) as Expr1"
Set rs8 = cnn8.Execute(SQL)
rdnum = rs8("Expr1")

SQL = "INSERT INTO tblCustomerNotesAttachments (CustNum,UserNo,Sequence) "
SQL = SQL &  " VALUES (" 
SQL = SQL & "'"  & rdnum & "'"
SQL = SQL & ","  & Session("UserNo") & ",0)"

Set rs8 = cnn8.Execute(SQL)

'Now we have ot get it back so that we know what internal note id to tie it to
SQL = "Select Top 1 * from tblCustomerNotesAttachments where CustNum = '" & rdnum & "'"
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
InternalNoteeNumber = rs8("InternalNoteNumber")

Upload.Save

'Rename the files
' Construct the save path
Pth ="../clientfiles/" & trim(GetPOSTParams("Serno")) & "/attachments/"

For Each File in Upload.Files
	fn=File.FileName
   File.SaveAsVirtual  Pth & InternalNoteeNumber & "-" & fn
Next


AccountNumber = Upload.Form("txtAccount")
AccountNote = Upload.Form("txtAccountNote")
'Replace all vbCrLf with <BR>s
AccountNote = Replace(AccountNote , vbCrLf, "<BR>")


NewCompleteName = InternalNoteeNumber & "-" & fn 

'OK, now update the record with what the filename is
SQL = "UPDATE tblCustomerNotesAttachments Set AttachmentFilename = '" & NewCompleteName & "', "
SQL = SQL & "CustNum = '" & AccountNumber & "', "
SQL = SQL & "Note = '" & AccountNote & "' "
SQL = SQL & " where InternalNoteNumber = " & InternalNoteeNumber 
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs9 = Server.CreateObject("ADODB.Recordset")
Set rs9 = cnn8.Execute(SQL)
cnn8.close
set cnn8 = Nothing
set rs8 = Nothing

Description = ""
Description = Description & "A new note attachments was added to account # "  & AccountNumber
Description = Description & "     The text of the note is as follows: "  & AccountNote
 
CreateAuditLogEntry "Account Note With Attachment Added","Account Note With Attachment Added","Minor",0,Description


set rs8 = Nothing
cnn8.Close
Set cbb8=Nothing


Response.Redirect("main.asp#Attachments")

%>
















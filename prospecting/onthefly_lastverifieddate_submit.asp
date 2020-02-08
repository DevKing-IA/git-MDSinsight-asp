<!--#include file="../inc/subsandfuncs.asp"-->
<!--#include file="../inc/InSightFuncs.asp"-->
<%


InternalRecordIdentifier = Request.Form("dateInternalRecordIdentifier")

SQL = "SELECT LastVerifiedDate FROM PR_Prospects where InternalRecordIdentifier = " & InternalRecordIdentifier 
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	Orig_Date = rs("LastVerifiedDate")
End If

set rs = Nothing



verifydate 			= CDate(Request.Form("txtProspectEditVerifyDate"))



If Orig_Date<>verifydate Then

SQL = "UPDATE PR_Prospects SET "
SQL = SQL &  " LastVerifiedDate = '" & CDate(verifydate) & "'"
SQL = SQL & " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier


cnn8.Execute(SQL)

End If

cnn8.Close
Set cnn8 = Nothing


Description = ""
If Orig_Date  <> verifydate  Then
	Description = Description & GetTerm("Prospecting") & " last verify date changed from  " & Orig_Date & " to " & verifydate
End If


CreateAuditLogEntry GetTerm("Prospecting") & " last verify date edited ",GetTerm("Prospecting") ,"Minor",0,Description



%>


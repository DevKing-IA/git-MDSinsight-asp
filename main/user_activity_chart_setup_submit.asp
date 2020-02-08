<!--#include file="../inc/subsandfuncs.asp"-->
<!--#include file="../inc/InSightFuncs.asp"-->

<%
NumberOfDays = Request.Form("selNumberOfDays")
If Request.Form("chkFservCloseOnly") <> "" Then
	FservCloseOnly = Request.Form("chkFservCloseOnly")
	'response.write(Request.Form("chkFservCloseOnly"))
	'response.end
	If FservCloseOnly = "on" Then FservCloseOnly = "1" Else FservCloseOnly  = "0"
Else ' It wasn't showing
	FservCloseOnly = 0
End If

'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail
SQL = "SELECT * from Settings_Screens where ScreenNumber = 1000 AND UserNo = " & Session("userNo")
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
If not rs.EOF Then
	Orig_NumberOfDays = rs("ScreenSpecificData1")
	Orig_SelectedUserDisplayNames = rs("ScreenSpecificData2")
	Orig_FservCloseOnly = rs("ScreenSpecificData3")
End If
set rs = Nothing
cnn8.close
set cnn8 = Nothing
'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************


'This is a wierd way to figure out how all the check boxes are named, but it works
SelectedUserDisplayNames = ","
SQL = "SELECT userNo,userFirstName,userLastName,userEmail,userDisplayName FROM tblUsers where userArchived <> 1 order by userFirstName"
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
If not rs.EOF Then
	Do While Not rs.EOF
		If Request.Form("chk" & rs("userNo")) = "on" Then SelectedUserDisplayNames = SelectedUserDisplayNames & GetUserDisplayNameByUserNo(rs("userNo")) & ","
	rs.movenext
	loop
End If


SQL = "UPDATE Settings_Screens SET "
SQL = SQL &  "ScreenSpecificData1 = '" & NumberOfDays & "',"
SQL = SQL &  "ScreenSpecificData2= '" & SelectedUserDisplayNames & "', "
SQL = SQL &  "ScreenSpecificData3= '" & FservCloseOnly & "' "
SQL = SQL &  " WHERE userNo = " & Session("UserNo") & " AND ScreenNumber = 1000"
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing


Description = ""
If cint(Orig_NumberOfDays) <> cint(NumberOfDays) Then
	Description = Description & "Number of days to chart changed from " & Orig_NumberOfDays & " to " & NumberOfDays 
End If
If Orig_SelectedUserDisplayNames <> SelectedUserDisplayNames Then
	Description = Description & "  User names to chart changed from " & Orig_SelectedUserDisplayNames & " to " & SelectedUserDisplayNames 
End If
If Orig_FservCloseOnly  <> FservCloseOnly Then
	Description = Description & "  Only chart close tickets for field techs " & Orig_FservCloseOnly & " to " & FservCloseOnly 
End If


CreateAuditLogEntry "User Activity Chart Settings","User Activity Chart Settings","Minor",0,Description

Response.Redirect(BaseURL & "main/default.asp")

%>
















<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%

'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail
SQL = "SELECT * from Settings_Screens where ScreenNumber = 1100 AND UserNo = " & Session("userNo")
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
If not rs.EOF Then
	Orig_SelectedUserNumbers = rs("ScreenSpecificData2")
Else
	Orig_SelectedUserNumbers = ""
End If
set rs = Nothing
cnn8.close
set cnn8 = Nothing
'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************


'This is a wierd way to figure out how all the check boxes are named, but it works
SelectedUserNumbers = ","
SQL = "SELECT userNo,userFirstName,userLastName,userEmail,userDisplayName FROM tblUsers where userArchived <> 1 order by userFirstName"
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
If not rs.EOF Then
	Do While Not rs.EOF
		If Request.Form("chk" & rs("userNo")) = "on" Then SelectedUserNumbers = SelectedUserNumbers & rs("userNo") & ","
		'Response.write("Current User Numbers : " & rs("userNo") & "---" & Request.Form("chk" & rs("userNo")) & "<br>")
	rs.movenext
	loop
End If




SQL1 = "SELECT * FROM Settings_Screens WHERE userNo = " & Session("UserNo") & " AND ScreenNumber = 1100"

Set rsInsight1 = Server.CreateObject("ADODB.Recordset")
rsInsight1.CursorLocation = 3 
Set rsInsight1 = cnn8.Execute(SQL1)
If NOT rsInsight1.EOF Then
	SQL = "UPDATE Settings_Screens SET ScreenSpecificData1= '" & SelectedUserNumbers & "' WHERE userNo = " & Session("UserNo") & " AND ScreenNumber = 1100"
Else
	SQL = "INSERT INTO Settings_Screens (ScreenNumber,UserNo,ScreenSpecificData1) Values (1100," & Session("UserNo") & ",'" & SelectedUserNumbers & "')"
End If
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing


Description = ""
If Orig_SelectedUserNumbers <> SelectedUserNumbers Then
	Description = Description & "  Prospecting dashboard sales rep chart changed from " & Orig_SelectedUserNumbers & " to " & SelectedUserNumbers 
End If


CreateAuditLogEntry "Prospecting Dashboard Sales Rep Settings","Prospecting Dashboard Sales Rep Settings","Minor",0,Description

Response.Redirect("dashboard_salesrep_display_setup.asp")

%>
















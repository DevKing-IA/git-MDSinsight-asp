<!--#include file="inc/settings.asp"-->
<!--#include file="inc/InSightFuncs_Users.asp"-->
<!--#include file="inc/InSightFuncs.asp"-->

<%

SQL = "SELECT * FROM tblServerInfo where clientKey='"& MUV_READ("ClientID") &"'"
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (InsightCnnString)
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then

	userQuickLoginURL = rs("QuickLoginURL")
	

	If Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV.") <> 0 AND Left(ucase(userQuickLoginURL),11) <> "HTTP://DEV." Then 
	
		'get rid of http://
		userQuickLoginURL = Right(userQuickLoginURL,len(userQuickLoginURL)-7)
	
		'Strip the URL first part
		For x = 1 to len(userQuickLoginURL)
			If Mid(userQuickLoginURL,x,1)="." Then
				userQuickLoginURL = right(userQuickLoginURL,len(userQuickLoginURL)-(x))		
				Exit For
			End If
		Next 
	
		userQuickLoginURL = "http://dev." & userQuickLoginURL 
	
	End If
	
	If Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV2.") <> 0 AND Left(ucase(userQuickLoginURL),12) <> "HTTP://DEV2." Then 
	
		'get rid of http://
		userQuickLoginURL = Right(userQuickLoginURL,len(userQuickLoginURL)-7)
	
		'Strip the URL first part
		For x = 1 to len(userQuickLoginURL)
			If Mid(userQuickLoginURL,x,1)="." Then
				userQuickLoginURL = right(userQuickLoginURL,len(userQuickLoginURL)-(x))		
				Exit For
			End If
		Next 
	
		userQuickLoginURL = "http://dev2." & userQuickLoginURL 
	
	End If
	
End If
	
set rs = Nothing
cnn8.close
set cnn8 = Nothing



If MUV_Read("ClientID") <>  "" Then ' In case it has timed out already
	CreateAuditLogEntry "Logout","Logout","Major",0, Session("userEmail") & " logged out."
End If

'***************************************************************************************************************
'IF CUSTOMER IS CORPORATE COFFEE SYSTEMS, GO TO CUSTOM CCS LOGIN PAGE
'***************************************************************************************************************

If HasCustomLoginPage(MUV_Read("ClientID")) AND MUV_Read("ClientID")<> "" AND (MUV_Read("ClientID")="1071" OR MUV_Read("ClientID")="1071d") Then
	
	If MUV_Read("QuickLoginUsed") = "1" Then
		logoutURL = BaseURL & "ql-CCS.asp?u=" & Session("UserNo") & "&c=" & MUV_Read("ClientID")
	Else
		logoutURL = BaseURL & "default_customLoginCCS.asp"
	End If
	
	
'***************************************************************************************************************	
'IF CUSTOMER IS NOT CCS, BUT HAS A CUSTOM LOGIN PAGE, GO THERE
'***************************************************************************************************************
ElseIf HasCustomLoginPage(MUV_Read("ClientID")) AND MUV_Read("ClientID")<> "" AND MUV_Read("ClientID")<>"1071" AND MUV_Read("ClientID")<>"1071d" Then
	
	If MUV_Read("QuickLoginUsed") = "1" Then
		logoutURL = userQuickLoginURL & "?u=" & Session("UserNo") & "&c=" & MUV_Read("ClientID")
	Else
		logoutURL = BaseURL & "default.asp?clientID=" & MUV_Read("ClientID")
	End If
	
'***************************************************************************************************************
'IF CUSTOMER DOES NOT HAVE A CUSTOM LOGIN, GO TO NON-BRANDED LOGIN PAGE
'***************************************************************************************************************
Else

	If MUV_Read("QuickLoginUsed") = "1" Then
		logoutURL = userQuickLoginURL & "?u=" & Session("UserNo") & "&c=" & MUV_Read("ClientID")
	Else
		logoutURL = BaseURL & "default.asp"
	End If

	
End If


dummy = MUV_Init

Session("ClientCnnString") =""

Session.Abandon
Response.write(logoutURL)
Response.Redirect(logoutURL)
 
%>


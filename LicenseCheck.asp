<!--#include file="inc/InSightFuncs.asp"-->
<!--#include file="inc/InSightFuncs_Users.asp"-->
<!--#include file="inc/settings.asp"-->
<script src="<%=baseURL%>/js/sweetalert/sweet-alert.js"></script>
<link rel="stylesheet" type="text/css" href="<%=baseURL%>/js/sweetalert/sweetalert.css">

<SCRIPT LANGUAGE="JavaScript">
<!--
    function showExpired()
    {
		swal({
		  title: "License Expired",
		  text: "Your license is expired and the grace period has ended. Please contact your Admin or technical support to continue using MDS Insight.",
		  type: "warning",
		  confirmButtonText: "Continue"
		}, function () {
		  window.location = "logout.asp"
		}); 
	}  

    function showAlmostExpired(exdays,lpage)
        {
        
        var edays = exdays
        var lpag = lpage
        
		swal({
			title: "License Expired - Grace Period",
	        text: "Your license is expired! You are currently running under a grace period license which will end in " + edays + " days. After the grace period ends you will no longer be able to log into MDS Insight." ,
		  type: "warning",
		  confirmButtonText: "Continue"
		}, function () {
		  window.location = lpag
		}); 
	}  


// -->
</SCRIPT> 

<%

'PROGRAMMING 

'Response.Write("LICENSE TYPE: " & GetLicenseTypeByUser(Session("UserNo")) & "<br>")
'Response.Write("LOGIN PAGE: " & MUV_Read("LoginPage") & "<br>")

If Trim(Ucase(GetLicenseTypeByUser(Session("UserNo")))) = "PROGRAMMING" Then
	ColorAndTitleAndMsg = "blue~"
	ColorAndTitleAndMsg = ColorAndTitleAndMsg  & "MDS Insight Programmer's Super Ultimate Elite license~"
	ColorAndTitleAndMsg = ColorAndTitleAndMsg  & "You are running under an MDS Insight Programmer's Super Ultimate Elite license.<br>This license never expires."
	dummy = MUV_WRITE("LicenseStatus",ColorAndTitleAndMsg)
	Response.Redirect (MUV_ReadAndRemove("LoginPage"))
End If


'LICENSED

If Ucase(GetLicenseTypeByUser(Session("UserNo"))) = "LICENSED" Then
	'Check expiration
	ExpDate = GetLicenseExpDateByUser(Session("UserNo"))
	If DateDiff("d",Now(),cdate(ExpDate)) > 5 Then ' Valid, no problems
		ColorAndTitleAndMsg = "green~"
		ColorAndTitleAndMsg = ColorAndTitleAndMsg  & "MDS Insight License - Valid~"
		ColorAndTitleAndMsg = ColorAndTitleAndMsg  & "You are running under an valid MDS Insight license.<br>This license is valid through " & FormatDateTime(GetLicenseExpDateByUser(Session("UserNo")),2) & "."
		dummy = MUV_WRITE("LicenseStatus",ColorAndTitleAndMsg)
		Response.Redirect (MUV_ReadAndRemove("LoginPage"))
	End If
	If DateDiff("d",Now(),cdate(ExpDate)) < 6 AND DateDiff("d",Now(),cdate(ExpDate)) > 0  Then ' About to expire
		ColorAndTitleAndMsg = "orange~"
		ColorAndTitleAndMsg = ColorAndTitleAndMsg  & "MDS Insight License - Expires in  " & DateDiff("d",Now(),cdate(ExpDate)) & " day(s)~"
		ColorAndTitleAndMsg = ColorAndTitleAndMsg  & "You are running under an valid MDS Insight license.<br>"
		ColorAndTitleAndMsg = ColorAndTitleAndMsg  & "This license is valid through " & FormatDateTime(GetLicenseExpDateByUser(Session("UserNo")),2) & "."
		dummy = MUV_WRITE("LicenseStatus",ColorAndTitleAndMsg)
		Response.Redirect (MUV_ReadAndRemove("LoginPage"))
	End If
	If DateDiff("d",Now(),cdate(ExpDate)) < 1 AND DateDiff("d",Now(),cdate(ExpDate)) > -5  Then ' Expired, running in grace
		ColorAndTitleAndMsg = "red~"
		ColorAndTitleAndMsg = ColorAndTitleAndMsg  & "MDS Insight License Expired - Running In Grace~"
		ColorAndTitleAndMsg = ColorAndTitleAndMsg  & "You're MDS Insight license is expired.<br>"
		ColorAndTitleAndMsg = ColorAndTitleAndMsg  & "You are currently running under a Grace Period license which will end in " & DateDiff("d",Now(),DateAdd("d",5,cdate(ExpDate))) & " day(s).<br>"
		ColorAndTitleAndMsg = ColorAndTitleAndMsg  & "After the grace period end you will no longer be able to log into MDS Insight.<br>"
		dummy = MUV_WRITE("LicenseStatus",ColorAndTitleAndMsg)
		Response.write("<script type=""text/javascript"">showAlmostExpired(" & DateDiff("d",Now(),DateAdd("d",5,cdate(ExpDate))) & ",'" & MUV_ReadAndRemove("LoginPage") & "');</script>")
	End If
	If DateDiff("d",Now(),cdate(ExpDate)) < 1 AND DateDiff("d",Now(),cdate(ExpDate)) <= -5  Then
		Response.write("<script type=""text/javascript"">showExpired();</script>")		
	End If
End If


'FREE

If Ucase(GetLicenseTypeByUser(Session("UserNo"))) = "FREE" Then
	'Check expiration
	ExpDate = GetLicenseExpDateByUser(Session("UserNo"))
	If DateDiff("d",Now(),cdate(ExpDate)) > 5 Then ' Valid, no problems
		ColorAndTitleAndMsg = "green~"
		ColorAndTitleAndMsg = ColorAndTitleAndMsg  & "MDS Insight FREE License - Valid~"
		ColorAndTitleAndMsg = ColorAndTitleAndMsg  & "You are running under an MDS Insight FREE license.<br>This is a fully functional license free of any restrictions.<br>This license is valid through " & FormatDateTime(GetLicenseExpDateByUser(Session("UserNo")),2) & "."
		dummy = MUV_WRITE("LicenseStatus",ColorAndTitleAndMsg)
		Response.Redirect (MUV_ReadAndRemove("LoginPage"))
	End If
	If DateDiff("d",Now(),cdate(ExpDate)) < 6 AND DateDiff("d",Now(),cdate(ExpDate)) > 0  Then ' About to expire
		ColorAndTitleAndMsg = "orange~"
		ColorAndTitleAndMsg = ColorAndTitleAndMsg  & "MDS Insight FREE License - Expires in  " & DateDiff("d",Now(),cdate(ExpDate)) & " day(s)~"
		ColorAndTitleAndMsg = ColorAndTitleAndMsg  & "You are running under an valid MDS Insight FREE license.<br>"
		ColorAndTitleAndMsg = ColorAndTitleAndMsg  & "This license is valid through " & FormatDateTime(GetLicenseExpDateByUser(Session("UserNo")),2) & "."
		dummy = MUV_WRITE("LicenseStatus",ColorAndTitleAndMsg)
		Response.Redirect (MUV_ReadAndRemove("LoginPage"))
	End If
	If DateDiff("d",Now(),cdate(ExpDate)) < 1 AND DateDiff("d",Now(),cdate(ExpDate)) > -5  Then ' Expired, running in grace
		ColorAndTitleAndMsg = "red~"
		ColorAndTitleAndMsg = ColorAndTitleAndMsg  & "MDS Insight License Expired - Running In Grace~"
		ColorAndTitleAndMsg = ColorAndTitleAndMsg  & "You're MDS Insight license is expired.<br>"
		ColorAndTitleAndMsg = ColorAndTitleAndMsg  & "You are currently running under a Grace Period license which will end in " & DateDiff("d",Now(),DateAdd("d",5,cdate(ExpDate))) & " day(s).<br>"
		ColorAndTitleAndMsg = ColorAndTitleAndMsg  & "After the grace period end you will no longer be able to log into MDS Insight.<br>"
		dummy = MUV_WRITE("LicenseStatus",ColorAndTitleAndMsg)
		Response.write("<script type=""text/javascript"">showAlmostExpired(" & DateDiff("d",Now(),DateAdd("d",5,cdate(ExpDate))) & ",'" & MUV_ReadAndRemove("LoginPage") & "');</script>")
	End If
	If DateDiff("d",Now(),cdate(ExpDate)) < 1 AND DateDiff("d",Now(),cdate(ExpDate)) <= -5  Then
		Response.write("<script type=""text/javascript"">showExpired();</script>")		
	End If
End If



   
'*************************   
'Subs And Funcs Start Here

Function GetLicenseTypeByUser(passedUserNo) 
	
	resultGetLicenseTypeByUser = ""

	Set cnnGetLicenseTypeByUser = Server.CreateObject("ADODB.Connection")
	cnnGetLicenseTypeByUser.open Session("ClientCnnString")
	Set rsGetLicenseTypeByUser = Server.CreateObject("ADODB.Recordset")

	SQLGetLicenseTypeByUser = "Select * from tblUsers where UserNo = " & passedUserNo
	
	rsGetLicenseTypeByUser.CursorLocation = 3 
	Set rsGetLicenseTypeByUser = cnnGetLicenseTypeByUser.Execute(SQLGetLicenseTypeByUser)
	
	If not rsGetLicenseTypeByUser.eof then resultGetLicenseTypeByUser = rsGetLicenseTypeByUser("userLicense") 
	
	Set rsGetLicenseTypeByUser = Nothing
	cnnGetLicenseTypeByUser.Close
	Set cnnGetLicenseTypeByUser = Nothing

	GetLicenseTypeByUser = resultGetLicenseTypeByUser
	
End Function

Function GetLicenseExpDateByUser(passedUserNo) 
	
	resultGetLicenseExpDateByUser = "01/01/1980"

	Set cnnGetLicenseExpDateByUser = Server.CreateObject("ADODB.Connection")
	cnnGetLicenseExpDateByUser.open Session("ClientCnnString")
	Set rsGetLicenseExpDateByUser = Server.CreateObject("ADODB.Recordset")

	SQLGetLicenseExpDateByUser = "Select * from tblUsers where UserNo = " & passedUserNo
	
	rsGetLicenseExpDateByUser.CursorLocation = 3 
	Set rsGetLicenseExpDateByUser = cnnGetLicenseExpDateByUser.Execute(SQLGetLicenseExpDateByUser)
	
	If not rsGetLicenseExpDateByUser.eof then resultGetLicenseExpDateByUser = rsGetLicenseExpDateByUser("userLicenseExpiration") 
	
	Set rsGetLicenseExpDateByUser = Nothing
	cnnGetLicenseExpDateByUser.Close
	Set cnnGetLicenseExpDateByUser = Nothing

	GetLicenseExpDateByUser = resultGetLicenseExpDateByUser 
	
End Function

%>
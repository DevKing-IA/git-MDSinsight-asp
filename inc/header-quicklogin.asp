<% @ Language = VBScript %>

<!--#include file="SubsAndFuncs.asp"-->
<!--#include file="InsightFuncs.asp"-->

<% MUV_Init() %>

<%
	QuickUserNo = Request.QueryString("u")
	QuickClientID = Request.QueryString("c")
	QuickClientDestination = Request.QueryString("d")

	If QuickClientID = "" Then
		QuickClientID = Request.Form("txtClientKeyCustom")
		QuickUserNo = Request.Form("txtUserNo")
		QuickClientDestination = Request.Form("txtDestinationURL")
	End If
	
	IF QuickClientID = "" Then
		QuickClientID = Request.Form("txtClientKey")
		QuickUserNo = Request.Form("txtUserNo")	
		QuickClientDestination = Request.Form("txtDestinationURL")
	End If
	
	If Right(QuickClientID,1) = "d" Then
		ClientKeyForFileNames = LEFT(QuickClientID, (LEN(QuickClientID)-1))
	Else
		ClientKeyForFileNames = QuickClientID
	End If	
	
	dummy = MUV_Write("ClientKeyForFileNames",ClientKeyForFileNames)


	ClientKey = QuickClientID
	UserNo = QuickUserNo
	ClientDestination = QuickClientDestination
	
	customLoginPage = false

	'**************************************************************************
    'Get Company Information
    '**************************************************************************
    
    If QuickClientID <> "" Then
    
		SQLCustomLogin = "SELECT * FROM tblServerInfo where clientKey='"& QuickClientID &"'"
		Set ConnectionCustomLogin = Server.CreateObject("ADODB.Connection")
		Set RecordsetCustomLogin = Server.CreateObject("ADODB.Recordset")
		ConnectionCustomLogin.Open InsightCnnString
	
		'Open the recordset object executing the SQL statement and return records
		RecordsetCustomLogin.Open SQLCustomLogin,ConnectionCustomLogin,3,3
	
		'First lookup the ClientKey in tblServerInfo
		If NOT RecordsetCustomLogin.EOF then
			companyName = RecordsetCustomLogin.Fields("companyName")
			shortCompanyName = RecordsetCustomLogin.Fields("shortCompanyName")
			companyDomainName = RecordsetCustomLogin.Fields("companyDomainName")
			RecordsetCustomLogin.close
			ConnectionCustomLogin.close	
		End If	
		
		If shortCompanyName = "" Then
			shortCompanyName = companyName
		End If
		If companyDomainName = "" Then
			companyDomainName = "mydomain.com"
		End If
		
		If HasCustomLoginPage(QuickClientID) AND (QuickClientID = "1071" OR QuickClientID = "1071d") Then
			Response.Redirect("ql-CCS.asp?u=" & QuickUserNo & "&c=" & QuickClientID & "&d=" & QuickClientDestination)
		ElseIf HasCustomLoginPage(QuickClientID) AND clientID <> "1071" AND QuickClientID <> "1071d" Then
			customLoginPage = true
		Else
			customLoginPage = false
		End If
		

	Else
		companyName = "MDS INSIGHT"
		shortCompanyName = "MDS"
		companyDomainName = "mydomain.com"
		customLoginPage = false
	End If
  	
	'Response.write("QuickUserNo : " &  QuickUserNo & "<br>")
	'Response.write("QuickClientID : " &  QuickClientID & "<br>")
	'Response.write("customLoginPage : " &  customLoginPage & "<br>")
      
    'Use the QuickUserNo & QuickClientID to lookup the users credentials
    
	SQL = "SELECT * FROM tblServerInfo where clientKey='"& QuickClientID &"'"
	Set Connection = Server.CreateObject("ADODB.Connection")
	Set Recordset = Server.CreateObject("ADODB.Recordset")
	Connection.Open InsightCnnString

	'Open the recordset object executing the SQL statement and return records
	Recordset.Open SQL,Connection,3,3

	'First lookup the ClientKey in tblServerInfo
	'If there is no record with the entered client key, close connection
	'and go back to login with QueryString
	If Recordset.recordcount <= 0 then
		info = "Your client key could not be found. Please contact your administrator."
		Connection.close
	Else
		tmpCnnString = "Driver={SQL Server};Server=" & Recordset.Fields("dbServer")
		tmpCnnString = tmpCnnString  & ";Database=" & Recordset.Fields("dbCatalog")
		tmpCnnString  = tmpCnnString  & ";Uid=" & Recordset.Fields("dbLogin")
		tmpCnnString  = tmpCnnString  & ";Pwd=" & Recordset.Fields("dbPassword") & ";"
			
		tmpSQL_Owner = Recordset.Fields("dbLogin")
			
		Recordset.close	
		Connection.close
		
	    Connection.Open tmpCnnString 
	
		SQL = "SELECT * FROM tblUsers where userNo = " & QuickUserNo
				
		'Open the recordset object executing the SQL statement and return records
		Recordset.Open SQL,Connection,3,3
	
		'If there is no record with the entered userNo, close connection
		'and go back to default
		If Recordset.recordcount <= 0 then
			info = "Your user account count not be found. Please contact your administrator."
			Recordset.close
			Connection.close
			set Recordset=nothing
			set Connection=nothing
		Else
			If Recordset.Fields("userEnabled") <> True Then
				Fname = Recordset.Fields("userFirstName")
				Lname = Recordset.Fields("userLastName")
				
				info = "Your login is currently disabled. Please contact your administrator."
				Recordset.close
				Connection.close
				set Recordset=nothing
				set Connection=nothing
	
				Session("ClientCnnString") = tmpCnnString
				dummy = MUV_Write("ClientID",QuickClientID)
				
				Description = "The user " & Fname & " " & Lname & " attempted to login via quick login but their account is disabled."
				CreateAuditLogEntry "Disabled Login Attempt","Disabled Login Attempt","Major",0,Description 
				
				Session.Abandon
				MUV_Init()
				Response.End
			Else
				dummy = MUV_Write("QuickLoginUsed",1)
				'Found a valid & enabled user number
				QuickUserEmail = Recordset.Fields("useremail")

				'Automatic login for techs and drivers
				If Recordset.Fields("userType") = "Driver" or Recordset.Fields("userType") = "Field Service" or Recordset.Fields("userType") = "TechAndDriver" Then 
					QuickUserPassword = Recordset.Fields("userPassword")
				Else
					QuickUserPassword = ""
				End If

	
				Recordset.close
				Connection.close
				set Recordset=nothing
				set Connection=nothing
						
			End If
		End If
		
	End If	

%>   

<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title><%= CompanyName %> | Insight for <%= shortCompanyName %></title>

    <!-- Bootstrap core CSS -->
    <link href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" rel="stylesheet">
    <!-- End Bootstrap core CSS -->
    
    <!-- custom CSS for LOGIN !-->
    <% If customLoginPage = true Then %>
		<!-- Custom Login CSS !-->
		<link href="<%= BaseURL %>clientFiles/<%= MUV_READ("ClientKeyForFileNames") %>/loginPage/css/dashboard-login.css" rel="stylesheet">
		<!-- End Custom Login CSS -->
    <% Else %>
    	<!-- Generic Login OCS !-->
    	<link href="<%= BaseURL %>css/dashboard-login.css" rel="stylesheet">
    	<!-- End Generic Login OCS -->    
    <% End If %>

	<link href="<%= BaseURL %>css/global-insight-styles.css" rel="stylesheet">
	
    <!-- HTML5 shim and Respond.js for IE8 support of HTML5 elements and media queries -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/html5shiv/3.7.2/html5shiv.min.js"></script>
      <script src="https://oss.maxcdn.com/respond/1.4.2/respond.min.js"></script>
    <![endif]-->
    
    <!-- icons and notification styles !-->
     <!--<link href="<%= BaseURL %>css/font-awesome/css/font-awesome.min.css" rel="stylesheet">-->
    <!--<link rel="stylesheet" href="https://pro.fontawesome.com/releases/v5.10.1/css/all.css" integrity="sha384-y++enYq9sdV7msNmXr08kJdkX4zEI1gMjjkw0l9ttOepH7fMdhb7CePwuRQCfwCr" crossorigin="anonymous">-->
    <script src="https://kit.fontawesome.com/43bb408351.js" crossorigin="anonymous"></script>
	<link href="<%= BaseURL %>css/notifications.css" rel="stylesheet">
	<link href="https://fonts.googleapis.com/css?family=Roboto+Condensed:300,400,700" rel="stylesheet">
	<!-- eof icons and notification styles !-->
    
	<!-- sweet alert jquery modal alerts !-->	
	<script src="<%= BaseURL %>js/sweetalert/sweetalert.min.js"></script>
	<link rel="stylesheet" type="text/css" href="<%= BaseURL %>js/sweetalert/sweetalert.css">
	<!-- end sweet alert jquery modal alerts !-->	
    

	<!-- *********************************************************************** -->
	<!-- IMPORTANT - USE OLDER VERSION OF JQUERY FOR SORTABLE PLUGIN             -->
	<!-- *********************************************************************** -->
  	<!--<script src="http://code.jquery.com/jquery-1.11.2.min.js"></script>-->
	<script src="https://code.jquery.com/jquery-3.1.1.js" integrity="sha256-16cdPddA6VdVInumRGo6IbivbERE8p7CQR3HzTBuELA=" crossorigin="anonymous"></script>
	<!-- *********************************************************************** --> 
		 	
	<!-- Bootstrap core JS - must load after jQuery -->
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"></script>
    <!-- End Bootstrap core JS - must load after jQuery -->

	<!-----------------IMPORTANT FILE FOR DELIVERY BOARD HEADER ------------------------------------------->
    <!-- jQuery Cookie Files To Save State of Dismissed Alerts -->
    <script src="<%= BaseURL %>js/jquery.cookie.js"></script>
    <!-- End jQuery Cookie -->
    <!-----------------END IMPORTANT FILE FOR DELIVERY BOARD HEADER ---------------------------------------->

	<!-- validation javascript !-->
	<script type="text/javascript">
	
	
	$(document).ready(function() {
	    $.removeCookie("alert-delboard-hidden-routes");
	});
	
	 function checkFormQuickLogin()
	 {

        if (document.customQuickLoginForm.txtPassword.value == "") {
           swal("Please enter your insight password.");
           return false;
        }

        return true;
		
	}
	
	
	</script>



<!-- eof validation javascript !-->


  </head>
  <!-- site starts here !-->

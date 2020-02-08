<% @ Language = VBScript %>

<!--#include file="inc/SubsAndFuncs.asp"-->
<!--#include file="inc/InsightFuncs.asp"-->

<% MUV_Init() %>

<%
	QuickUserNo = Request.QueryString("u")
	QuickClientID = Request.QueryString("c")
	QuickClientDestination = Request.QueryString("d")
	

	'**************************************************************************
    'Get Company Information
    '**************************************************************************
    
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
		RecordsetCustomLogin.close
		ConnectionCustomLogin.close	
	End If	
	
	If shortCompanyName = "" Then
		shortCompanyName = companyName
	End If


        
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
	
				Recordset.close
				Connection.close
				set Recordset=nothing
				set Connection=nothing
						
			End If
		End If
		
	End If	

    

%>   

<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <title><%= CompanyName %> | Insight by <%= shortCompanyName %></title>

    <!-- Bootstrap core CSS -->
    <link href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" rel="stylesheet">
    <!-- End Bootstrap core CSS -->
	
    <!-- Custom Login CSS !-->
    <link href="<%= BaseURL %>clientFiles/<%= QuickClientID %>/loginPage/css/dashboard-login.css" rel="stylesheet">
    <!-- End Custom Login CSS -->

    
    <!-- HTML5 shim and Respond.js for IE8 support of HTML5 elements and media queries -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/html5shiv/3.7.2/html5shiv.min.js"></script>
      <script src="https://oss.maxcdn.com/respond/1.4.2/respond.min.js"></script>
    <![endif]-->

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
	    
	<!-- validation javascript !-->
	<script type="text/javascript">
	
	 function checkFormQuickLogin()
	 {

        if (document.customQuickLoginForm.txtPassword.value == "") {
           swal("Please enter your insight password.");
           return false;
        }

        return true;
		
	}	
	</script>
    
  </head>
  <body>
	  
	<!-- logo !-->
    <div class="container text-center">
    	<img src="<%= BaseURL %>/clientFiles/<%= QuickClientID %>/loginPage/img/logo.png" class="img-responsive-quick-login" alt="<%= CompanyName %>" title="<%= CompanyName %>">
    </div>
    <!-- eof logo !-->
    
    <!-- main heading !-->
    <div class="container text-center">
    	<h1 class="quicklogin"><span>insight</span> by <%= shortCompanyName %></h1>
    </div>
    <!-- eof main heading !-->


 <!-- login box !-->
    <div class="col-lg-12">
    <div class="quick-login-container">
    	
        <h2 class="quicklogin">Quick Sign In</h2>
        
	    <form action="action_login.asp" method="POST" name="customQuickLoginForm" id="customQuickLoginForm" onSubmit="return checkFormQuickLogin();" class="form-signin">
			
	        <input type="hidden" name="txtUsername" id="txtUsername" class="form-control" value="<%= QuickUserEmail %>">
	        <input type="hidden" name="txtClientKeyCustom" id="txtClientKeyCustom" class="form-control" value="<%= QuickClientID %>">
	        <input type="hidden" name="txtQuickLogin" id="txtQuickLogin" class="form-control" value="true">
	        <input type="hidden" name="txtUserNo" id="txtUserNo" class="form-control" value="<%= QuickUserNo %>">
	        <input type="hidden" name="txtDestinationURL" id="txtDestinationURL" class="form-control" value="<%= QuickClientDestination %>">
			
	        <% If info <> "" then %>
				<!-- line !-->
				<div class="row line">
					<div class="col-lg-12 col-md-12 col-sm-12 col-xs-12 text-center">
						<font color="yellow"><%= info %></font>
					</div>
				</div>
			<% End If %>
			
    		<% QStr = Request.QueryString("login") %>
    		
	        <% If QStr="namefailed" then %>
				<!-- line !-->
				<div class="row line">
					<div class="col-lg-12 col-md-12 col-sm-12 col-xs-12 text-center">
						<font color="yellow">Invalid Login, Please Try Again</font>
					</div>
				</div>
			<% End If %>
			
			
	        <% If QStr="disabled" then %>
				<!-- line !-->
				<div class="row line">
					<div class="col-lg-12 col-md-12 col-sm-12 col-xs-12 text-center">
						<font color="yellow">Your login is currently disabled<br>Please contact your administrator.</font>
					</div>
				</div>
				<%
					Description = "The user " & Request.QueryString("fname") & " " & Request.QueryString("lname") & " attempted to login but their account is disabled."
					CreateAuditLogEntry "Disabled Login Attempt","Disabled Login Attempt","Major",0,Description 
				%>
			<% End If %>
					
				
	        <% If QStr="hoursresctriction" then %>
				<!-- line !-->
				<div class="row line">
					<div class="col-lg-12 col-md-12 col-sm-12 col-xs-12 text-center">
						<font color="yellow"><%= Session("restrictedLoginMessage") %> Please contact your administrator if you need to login.</font>
					</div>
				</div>
				<%
					Description = "The user " & Request.QueryString("fname") & " " & Request.QueryString("lname") & " attempted to login outside of defined access hours."
					CreateAuditLogEntry "Login Attempt Outside Allowed Hours","Login Attempt Outside Allowed Hours","Major",0,Description 
				%>
			<% End If %>
			
	        <% If QStr="holidayresctriction" then %>
				<!-- line !-->
				<div class="row line">
					<div class="col-lg-12 col-md-12 col-sm-12 col-xs-12 text-center">
						<font color="yellow"><%= Session("restrictedLoginMessage") %> Please contact your administrator if you need to login.</font>
					</div>
				</div>
				<%
					Description = "The user " & Request.QueryString("fname") & " " & Request.QueryString("lname") & " attempted to login on a company holiday."
					CreateAuditLogEntry "Login Attempt On Company Holiday","Login Attempt On Company Holiday","Major",0,Description 
				%>
			<% End If %>
			
								
	        
	        <!-- line !-->
	        <div class="row line">
	        	<div class="col-lg-10 col-md-10 col-sm-9 col-xs-9">
	            	<input type="password" name="txtPassword" id="txtPassword" placeholder="password" class="input">
	            </div>
	            <div class="col-lg-2 col-md-2 col-sm-3 col-xs-3">
	            	<span class="icon">
					<img src="<%= BaseURL %>/clientFiles/<%= QuickClientID %>/loginPage/img/password-icon.png" class="img-responsive"></span>
	            </div>
	        </div>
	        <!-- eof line !-->
        
	        <!-- reset pass / login btn !-->
	        <div class="row line">
	        	<div class="col-lg-8">
	            	<a href="<%= BaseURL %>reset-password-ql-CCS.asp?u=<%= QuickUserNo %>&c=<%= QuickClientID %>">Forgot your password? Click Here.</a>
	            </div>
	            
	            <div class="col-lg-4">
	            	<button type="submit">Login</button>
	            </div>
	        </div>
	        <!-- eof reset pass / login btn !-->
    
       </form>
        
    </div>
    </div>
    <!-- eof login box !-->
            
  </body>
</html>
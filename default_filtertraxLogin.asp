<% @ Language = VBScript %>

<!--#include file="inc/subsandfuncs.asp"-->
<!--#include file="inc/InsightFuncs.asp"-->

<% MUV_Init() %>

<% 
dummy = MUV_Write("FilterTrax",1)

clientID = Request.QueryString("clientID")

If Right(clientID ,1) = "d" Then
	ClientKeyForFileNames = LEFT(clientID, (LEN(clientID)-1))
Else
	ClientKeyForFileNames = clientID
End If	

dummy = MUV_Write("ClientKeyForFileNames",ClientKeyForFileNames)


'**************************************************************************
'Get Company Information
'**************************************************************************

If clientID <> "" Then

	SQLCustomLogin = "SELECT * FROM tblServerInfo where clientKey='"& clientID &"'"
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

Else
	companyName = "MDS INSIGHT"
	shortCompanyName = "MDS"
	companyDomainName = "mydomain.com"
End If


%>
  

<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <title>FilterTrax | Insight by MDS</title>

    <!-- Bootstrap core CSS -->
    <link href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" rel="stylesheet">
    <!-- End Bootstrap core CSS -->
	
    <!-- Custom Login CSS !-->
    <link href="<%= BaseURL %>css/filtertraxlogin.css" rel="stylesheet">
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
	<script src="https://code.jquery.com/jquery-3.1.1.js" integrity="sha256-16cdPddA6VdVInumRGo6IbivbERE8p7CQR3HzTBuELA=" crossorigin="anonymous"></script>	
	<!-- *********************************************************************** --> 
		
 	
	<!-- Bootstrap core JS - must load after jQuery -->
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"></script>
    <!-- End Bootstrap core JS - must load after jQuery -->
	    
	<!-- validation javascript !-->
	<script type="text/javascript">
	
	 function checkForm()
	 {
        if (document.customLoginForm.txtUsername.value == "") {
            swal("Please enter your email address.");
            return false;
        }

        if (document.customLoginForm.txtPassword.value == "") {
           swal("Please enter your FilterTrax password.");
           return false;
        }

	    $("<table id='overlay'><tbody><tr><td>Preparing Your FilterTrax Experience...</td></tr><tr><td><img src='img/gears.gif'></td></tr></tbody></table>").css({
	        "position": "fixed",
	        "top": 0,
	        "left": 0,
	        "width": "100%",
	        "height": "100%",
	        "background-color": "rgba(0,0,0,.75)",
	        "z-index": 10000,
	        "vertical-align": "middle",
	        "text-align": "center",
	        "color": "#fff",
	        "font-size": "50px",
	        "font-weight": "bold",
	        "cursor": "wait"
	    }).appendTo("body");

        return true;
		
	}	
	
	
	 function checkFormWithClientKey()
	 {
        if (document.customLoginFormWithKey.txtUsername.value == "") {
            swal("Please enter your email address.");
            return false;
        }

        if (document.customLoginFormWithKey.txtPassword.value == "") {
           swal("Please enter your FiltersTrax password.");
           return false;
        }

        if (document.customLoginFormWithKey.txtClientKey.value == "") {
           swal("Please enter your FilterTrax client key.");
           return false;
        }
				
	    $("<table id='overlay'><tbody><tr><td>Preparing Your FilterTrax Experience...</td></tr><tr><td><img src='img/gears.gif'></td></tr></tbody></table>").css({
	        "position": "fixed",
	        "top": 0,
	        "left": 0,
	        "width": "100%",
	        "height": "100%",
	        "background-color": "rgba(0,0,0,.75)",
	        "z-index": 10000,
	        "vertical-align": "middle",
	        "text-align": "center",
	        "color": "#fff",
	        "font-size": "50px",
	        "font-weight": "bold",
	        "cursor": "wait"
	    }).appendTo("body");
    		      		
        return true;
	}	
	
	</script>
    
  </head>
  <body>
	  
	<!-- logo !-->
    <div class="login-container">
    	<img src="<%= BaseURL %>clientfilesV/filtertrax/logo.png" class="img-responsive" alt="FilterTrax" title="FilterTrax">
    </div>
    <!-- eof logo !-->
    
    <!-- main heading !-->
    <div class="container">
    	<h1><span></span></h1>
    </div>
    <!-- eof main heading !-->


 <!-- login box !-->
    <div class="col-lg-12">
    <div class="login-container">
    	
        <h2>Sign In</h2>
        
		 	<% If clientID <> "" Then %>
				<form action="<%= BaseURL %>action_login.asp" method="POST" name="customLoginForm" id="customLoginForm" onSubmit="return checkForm();" class="form-signin">
			<% Else %>
				<form action="<%= BaseURL %>action_login.asp" method="POST" name="customLoginFormWithKey" id="customLoginFormWithKey" onSubmit="return checkFormWithClientKey();" class="form-signin">
			<% End If %>
	   
    		<% QStr = Request.QueryString("login") %>
    		
	        <% If QStr="namefailed" then %>
				<!-- line !-->
				<div class="row line">
					<div class="col-lg-10 col-md-10 col-sm-9 col-xs-9">
						<font color="yellow">Invalid Login, Please Try Again</font>
					</div>
				</div>
			<% End If %>
			
			
	        <% If QStr="disabled" then %>
				<!-- line !-->
				<div class="row line">
					<div class="col-lg-10 col-md-10 col-sm-9 col-xs-9">
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
					<div class="col-lg-10 col-md-10 col-sm-9 col-xs-9">
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
					<div class="col-lg-10 col-md-10 col-sm-9 col-xs-9">
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
	            	<input type="text" name="txtUsername" id="txtUsername" placeholder="email@<%= CompanyDomainName %>" class="input">
	            </div>
	            <div class="col-lg-2 col-md-2 col-sm-3 col-xs-3">
	            	<span class="icon">
					<img src="<%= BaseURL %>img/loginPage/email-icon.png" class="img-responsive"></span>
	            </div>
	        </div>
	        <!-- eof line !-->
	        
	        <!-- line !-->
	        <div class="row line">
	        	<div class="col-lg-10 col-md-10 col-sm-9 col-xs-9">
	            	<input type="password" name="txtPassword" id="txtPassword" placeholder="password" class="input">
	            </div>
	            <div class="col-lg-2 col-md-2 col-sm-3 col-xs-3">
	            	<span class="icon">
					<img src="<%= BaseURL %>img/loginPage/password-icon.png" class="img-responsive"></span>
	            </div>
	        </div>
	        <!-- eof line !-->
	        
	        <% If clientID <> "" Then %>
	        	<input type="hidden" name="txtClientKeyCustom" id="txtClientKeyCustom" value="<%= clientID %>">
	        <% Else %>
		        <!-- line !-->
		        <div class="row line">
		        	<div class="col-lg-10 col-md-10 col-sm-9 col-xs-9">
		            	<input type="text" name="txtClientKey" id="txtClientKey" placeholder="client key" class="input">
		            </div>
		            <div class="col-lg-2 col-md-2 col-sm-3 col-xs-3">
		            	<span class="icon">
						<img src="<%= BaseURL %>img/loginPage/clientkey-icon.png" class="img-responsive"></span>
		            </div>
		        </div>
		        <!-- eof line !-->
	        <% End If %>
	        
	        
	        <!-- reset pass / login btn !-->
	        <div class="row line">
	        	<div class="col-lg-8">
	            	<a href="<%= BaseURL %>reset-password-FilterTrax.asp?clientID=<%= clientID %>">Forgot your password? Click Here.</a>
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
<% @ Language = VBScript %>

<!--#include file="inc/SubsAndFuncs.asp"-->
<!--#include file="inc/InsightFuncs.asp"-->
<!--#include file="inc/mail.asp"-->

<%
	clientID = "1071"
	
	'**************************************************************************
    'Get Company Information
    '**************************************************************************
    
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
		RecordsetCustomLogin.close
		ConnectionCustomLogin.close	
	End If	
	
	If shortCompanyName = "" Then
		shortCompanyName = companyName
	End If

%>   

<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">

	<% If clientID = "1071" OR clientID = "1071d" Then %>
    	<title><%= CompanyName %> | Insight by <%= shortCompanyName %></title>
    <% Else %>
    	<title><%= CompanyName %> | Insight for <%= shortCompanyName %></title>
    <% End If %>


    <!-- Bootstrap core CSS -->
    <link href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" rel="stylesheet">
    <!-- End Bootstrap core CSS -->
	
    <!-- Custom Login CSS !-->
    <link href="<%= BaseURL %>/clientFiles/<%= clientID %>/loginPage/css/dashboard-login.css" rel="stylesheet">
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
	
	 function checkFormWithCaptcha()
	 {
	    if (document.frmPasswordReset.txtUsername.value == "") {
	        swal("Please enter your email address.");
	        return false;
	    }
	
	    if (document.frmPasswordReset.txtCaptcha.value == "") {
	       swal("Please enter human validation value.");
	       return false;
	    }
	
	    return true;
	}	
	</script>
    
  </head>
  <body>
<%
'Specify your captcha length here, this is the only configuration requirement
captchaLength = 7

Function captcha(captchaLength)
	if captchaLength > 15 then captchaLength = 15
	HighestValue = left(100000000000000,captchaLength)
	lowestValue = left(999999999999999,captchaLength)
	Randomize 
	intHighestNumber = Int((HighestValue - LowestValue + 1) * Rnd) + LowestValue
	session("captcha") = Int(intHighestNumber)
	x = 1
	response.write vbcrlf & "<table width = ''>" & vbcrlf & vbtab & "<tr>" & vbcrlf
	while x <= captchaLength
		response.write vbtab & vbtab & "<td align = 'center'><img src = 'captcha.asp?captchaID=" & x & "' ></td>" & vbcrlf
		x = x + 1
	wend
	response.write vbtab & "</tr>" & vbcrlf & "</table>" & vbcrlf
End Function


function RandomString()

    Randomize()

    dim CharacterSetArray
    CharacterSetArray = Array(_
        Array(7, "abcdefghijklmnopqrstuvwxyz"), _
        Array(1, "0123456789") _
    )

    dim i
    dim j
    dim Count
    dim Chars
    dim Index
    dim Temp

    for i = 0 to UBound(CharacterSetArray)

        Count = CharacterSetArray(i)(0)
        Chars = CharacterSetArray(i)(1)

        for j = 1 to Count

            Index = Int(Rnd() * Len(Chars)) + 1
            Temp = Temp & Mid(Chars, Index, 1)

        next

    next

    dim TempCopy

    do until Len(Temp) = 0

        Index = Int(Rnd() * Len(Temp)) + 1
        TempCopy = TempCopy & Mid(Temp, Index, 1)
        Temp = Mid(Temp, 1, Index - 1) & Mid(Temp, Index + 1)

    loop

    RandomString = TempCopy

end function
%>

<%
if request("action") = "captcha" then

	if cstr(request("txtCaptcha")) = cstr(session("captcha")) then

		Username = Request.Form("txtUsername")	
		ClientKey= Request.Form("txtClientKey")
		password = RandomString()
		
		
		SQLClientID = "SELECT * FROM tblServerInfo where clientKey='"& ClientKey &"'"
		
		
		Set ConnectionClientID  = Server.CreateObject("ADODB.Connection")
		Set RecordsetClientID  = Server.CreateObject("ADODB.Recordset")
		
		ConnectionClientID.Open InsightCnnString
		
		'Open the recordset object executing the SQL statement and return records
		RecordsetClientID.Open SQLClientID,ConnectionClientID,3,3
		
		'First lookup the ClientKey in tblServerInfo
		'If there is no record with the entered client key, close connection
		'and go back to login with QueryString
		If RecordsetClientID.recordcount <= 0 then
			RecordsetClientID.close
			ConnectionClientID.close
			set RecordsetClientID =nothing
			set ConnectionClientID =nothing
			info = "<font color='yellow'>Invaild Client Key. " & SQLClientID & "</font>"
			
		Else
			Session("ClientCnnString") = "Driver={SQL Server};Server=" & RecordsetClientID.Fields("dbServer")
			Session("ClientCnnString") = Session("ClientCnnString") & ";Database=" & RecordsetClientID.Fields("dbCatalog")
			Session("ClientCnnString") = Session("ClientCnnString") & ";Uid=" & RecordsetClientID.Fields("dbLogin")
			Session("ClientCnnString") = Session("ClientCnnString") & ";Pwd=" & RecordsetClientID.Fields("dbPassword") & ";"

			'create an instance of the ADO connection and recordset objects
			Set ConnectionUsers = Server.CreateObject("ADODB.Connection")
			Set rsUsers = Server.CreateObject("ADODB.Recordset")
			
			ConnectionUsers.Open Session("ClientCnnString")
		
			'declare the SQL statement that will query the database
			SQL = "SELECT * FROM tblUsers where userEmail='"& Username &"' and userArchived <> 1"
	
		
			'Open the recordset object executing the SQL statement and return records
			Set rsUsers = ConnectionUsers.Execute(SQL)
	
			'If there is no record with the entered username, close connection
			'and go back to login with QueryString
			If rsUsers.EOF then
				rsUsers.close
				ConnectionUsers.close
				set rsUsers =nothing
				set ConnectionUsers=nothing
				info = "<font color='yellow'>Invaild Email ID.</font>"
			Else		
				
				userEmail = username
				userFirstName = rsUsers("userFirstName")
				userLastName = rsUsers("userLastName")
				userDisplayName = rsUsers("userDisplayName")
				If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)

			
				'declare the SQL statement that will query the database
				SQL = "UPDATE tblUsers SET userpassword='"& Password &"' where userEmail='"& Username &"'"
				Set rsUsers = ConnectionUsers.Execute(SQL)
				
				%><!--#include file="emails/user_password_reset_credentials.asp"--><%
				
				SendMail "mailsender@" & maildomain,userEmail,emailSubject,emailBody, "System", "Password Reset"
	
				Description = "Login password reset email sent to " & userEmail & " for user " & userFirstName & " " & userLastName
			
				CreateAuditLogEntry "Password Reset","Password Reset","Minor",0,Description 
	
				info = "<font color='yellow'>Please check your email for your new password.</font>"
	
				ConnectionUsers.close	
			
			End If		
				
			RecordsetClientID.close
			ConnectionClientID.close	
		End If	
	Else
		info = "<font color='yellow'>Please enter correct captcha/human validation code.</font>"
	End If	
End If	
	
%>
	  
	<!-- logo !-->
    <div class="container">
    	<img src="<%= BaseURL %>/clientFiles/<%= clientID %>/loginPage/img/logo.png" class="img-responsive" alt="<%= CompanyName %>" title="<%= CompanyName %>">
    </div>
    <!-- eof logo !-->
    
    <!-- main heading !-->
    <div class="container">
    	<h1><span>insight</span> by <%= shortCompanyName %></h1>
    </div>
    <!-- eof main heading !-->


 <!-- login box !-->
    <div class="col-lg-12">
    <div class="login-container">
    	
        <h2>Reset Insight Password</h2>
        
		<% MUV_Init() %>
	
    	<form action="<%= BaseURL %>reset-password-CCS.asp?action=captcha" method="POST" name="frmPasswordReset" id="frmPasswordReset" onSubmit="return checkFormWithCaptcha();" class="form-signin">
      

			<input type="hidden" name="txtclientkey" id="txtclientkey" value="<%= clientID %>">
    
	        <% If info <> "" then %>
			<!-- line !-->
			<div class="row line">
				<div class="col-lg-10 col-md-10 col-sm-9 col-xs-9">
					<%= info %>
				</div>
			</div>
			<% End If %>
			
	        <!-- line !-->
	        <div class="row line">
	        	<div class="col-lg-10 col-md-10 col-sm-9 col-xs-9">
	            	<input type="text" name="txtUsername" id="txtUsername" placeholder="email@corpcofe.com" class="input">
	            </div>
	            <div class="col-lg-2 col-md-2 col-sm-3 col-xs-3">
	            	<span class="icon">
					<img src="<%= BaseURL %>/clientFiles/<%= clientID %>/loginPage/img/email-icon.png" class="img-responsive"></span>
	            </div>
	        </div>
	        <!-- eof line !-->
	        
	        <!-- line !-->
	        <div class="row line">
	        	<div class="col-lg-10 col-md-10 col-sm-9 col-xs-9">
					Human Validation Code: <% captcha(captchaLength) %>
	            </div>
	        </div>
	        <!-- eof line !-->
		        
	        
	        <!-- line !-->
	        <div class="row line">
	        	<div class="col-lg-10 col-md-10 col-sm-9 col-xs-9">
	            	<input type="text" name="txtCaptcha" id="txtCaptcha" placeholder="type in human validation code" class="input">
	            </div>
	            <div class="col-lg-2 col-md-2 col-sm-3 col-xs-3">
	            	<span class="icon">
					<img src="<%= BaseURL %>/clientFiles/<%= clientID %>/loginPage/img/password-icon.png" class="img-responsive"></span>
	            </div>
	        </div>
	        <!-- eof line !-->
        
	        <!-- remember / login btn !-->
	        <div class="row line">
	        	<div class="col-lg-4">
	        		<p>&nbsp;</p>
	            </div>
	            
	            <div class="col-lg-8">
	            	<button type="submit">Reset My Password</button>
	            </div>
	        </div>
	        <!-- eof remember / login btn !-->


	        <!-- remember / login btn !-->
	        <div class="row line">
	        	<div class="col-lg-12">
	        		<p>After clicking the RESET PASSWORD button, a new password will be generated & emailed to you.</p>
	            </div>
	            
	        </div>
	        <!-- eof remember / login btn !-->
	        
	        <a href="<%= BaseURL %>default_customLoginCCS.asp"><button type="button">Nevermind, return to login</button></a>
    
       </form>
        
    </div>
    </div>
    <!-- eof login box !-->
            
  </body>
</html>
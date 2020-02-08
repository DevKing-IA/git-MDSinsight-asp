<!--#include file="inc/header-default.asp"--> 
<!--#include file="inc/mail.asp"-->

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
<!-- eof validation javascript !-->

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
			info = "<font color='red'>Invaild Client Key. " & SQLClientID & "</font>"
			
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
				info = "<font color='red'>Invaild Email ID.</font>"
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
	
				info = "<font color='green'>Please check your email for your new password.</font>"
	
				ConnectionUsers.close	
			
			End If		
				
			RecordsetClientID.close
			ConnectionClientID.close	
		End If	
	Else
		info = "<font color='red'>Please enter correct captcha/human validation code.</font>"
	End If	
End If	
	
%>

  	<!-- logo / heading !-->
    <div class="container">
    	
        <div class="col-lg-6">
        <% If customLoginPage = true Then %>
    		<img src="<%= BaseURL %>clientFiles/<%= MUV_READ("ClientKeyForFileNames") %>/loginPage/img/logo.png" class="img-responsive logo" alt="<%= CompanyName %>" title="<%= CompanyName %>">
    	<% Else %>
 			<img src="<%= BaseURL %>img/loginpage/logo.png" class="img-responsive logo" alt="<%= CompanyName %>" title="<%= CompanyName %>">  	
    	<% End If %>
        </div>
        
        <div class="col-lg-6">
        <h1>MDS Insight<br>
			<small>Business Analytics Engineered for Your Success</small></h1>
        </div>
        
    </div>
    <!-- eof logo / heading !-->


    <form action="<%= BaseURL %>reset-password.asp?action=captcha" method="POST" name="frmPasswordReset" id="frmPasswordReset" onSubmit="return checkFormWithCaptcha();" class="form-signin">
      
    
 	<!-- login box !-->
    <div class="container">
    <div class="login-container">
    	
        <!-- left col !-->
        <div class="col-lg-6 equal-height">
        	<p align="center"><img src="<%= BaseURL %>img/loginpage/data-insights-actions.png" class="img-responsive"></p>
        </div>
        <!-- eof left col !-->
        
        <!-- right col !-->
        <div class="col-lg-6 equal-height signin-box">
        
        <h2>Reset Password</h2>

		
        <% If info <> "" then %>
		<!-- line !-->
		<div class="row line">
			<div class="col-lg-10 col-md-10 col-sm-9 col-xs-9">
				<span class="error"><%= info %></span>				
			</div>
		</div>
		<% End If %>
		
					
        
        <!-- line !-->
        <div class="row line">
        	<div class="col-lg-10 col-md-10 col-sm-9 col-xs-9">
            	<input type="email" placeholder="email@<%= CompanyDomainName %>" class="input" name="txtUsername" id="txtUsername">
            </div>
            <div class="col-lg-2 col-md-2 col-sm-3 col-xs-3">
            	<span class="icon"><img src="<%= BaseURL %>img/loginpage/email-icon.png" class="img-responsive"></span>
            </div>
        </div>
        <!-- eof line !-->
 

        <% 
        
        clientID = Request.QueryString("clientID")
        
        If clientID <> "" Then %>
	        <!-- line !-->
	        <div class="row line">
	        	<div class="col-lg-10 col-md-10 col-sm-9 col-xs-9">
	            	<input type="text" placeholder="client key" class="input" name="txtClientKey" id="txtClientKey" value="<%= clientID %>">
	            </div>
	            <div class="col-lg-2 col-md-2 col-sm-3 col-xs-3">
	            	<span class="icon"><img src="<%= BaseURL %>img/loginpage/clientkey-icon.png" class="img-responsive"></span>
	            </div>
	        </div>
	        <!-- eof line !-->
        	
        <% Else %>
	        <!-- line !-->
	        <div class="row line">
	        	<div class="col-lg-10 col-md-10 col-sm-9 col-xs-9">
	            	<input type="text" placeholder="client key" class="input" name="txtClientKey" id="txtClientKey">
	            </div>
	            <div class="col-lg-2 col-md-2 col-sm-3 col-xs-3">
	            	<span class="icon"><img src="<%= BaseURL %>img/loginpage/clientkey-icon.png" class="img-responsive"></span>
	            </div>
	        </div>
	        <!-- eof line !-->
        <% End If %>
   
 
        <!-- line !-->
        <div class="row line">
        	<div class="col-lg-10 col-md-10 col-sm-9 col-xs-9">
				<span class="msg">Human Validation Code: <% captcha(captchaLength) %></span>
            </div>
        </div>
        <!-- eof line !-->
   
        <!-- line !-->
        <div class="row line">
        	<div class="col-lg-10 col-md-10 col-sm-9 col-xs-9">
            	<input type="text" placeholder="type in human validation code" class="input" name="txtCaptcha" id="txtCaptcha">
            </div>
            <div class="col-lg-2 col-md-2 col-sm-3 col-xs-3">
            	<span class="icon"><img src="<%= BaseURL %>img/loginpage/password-icon.png" class="img-responsive"></span>
            </div>
        </div>
        <!-- eof line !-->
        

        
        <!-- remember / login btn !-->
        <div class="row line">
        	<div class="col-lg-6">
        		<span class="msg">After clicking the RESET button, a new password will be generated & emailed to you.</span>
            </div>
            
            <div class="col-lg-6">
            	<button type="submit">Reset Password</button>
            </div>
        </div>
        <!-- eof remember / login btn !-->
        
        <!-- forgot password !-->
        <% If customLoginPage = true Then %>
         	<a href="<%= BaseURL %>default.asp?clientID=<%= clientID %>">Nevermind, return to login</a>
         <% Else %> 
         	<a href="<%= BaseURL %>default.asp">Nevermind, return to login</a>
         <% End If %>
         <!-- eof forgot password !-->
        
       </div>
       <!-- eof right col !-->
        
    </div>
    </div>
    <!-- eof login box !-->
    
    <!-- mplex logos !-->
    <!--
    <div class="container footer-logos">
    	<img src="<%= BaseURL %>img/loginpage/footer.png" class="img-responsive">
    </div>
    -->
    <!-- eof mplex logos !-->

        
</form>
    
<!--#include file="inc/footer-login-noTimeout.asp"-->       
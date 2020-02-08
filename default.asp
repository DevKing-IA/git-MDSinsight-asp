<!--#include file="inc/header-default.asp"--> 

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

	<% If customLoginPage = true Then %>
		<form action="<%= BaseURL %>action_login.asp" method="POST" name="customLoginForm" id="customLoginForm" onSubmit="return checkForm();">
	<% Else %>
		<form action="<%= BaseURL %>action_login.asp" method="POST" name="customLoginFormWithKey" id="customLoginFormWithKey" onSubmit="return checkFormWithClientKey();">
	<% End If %>

    
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
        
        <h2>Sign In</h2>

		<% QStr = Request.QueryString("login") %>
		
        <% If QStr="namefailed" then %>
		<!-- line !-->
		<div class="row line">
			<div class="col-lg-10 col-md-10 col-sm-9 col-xs-9">
				<span class="error">Invalid Login, Please Try Again</span>				
			</div>
		</div>
		<% End If %>
		
		
        <% If QStr="disabled" then %>
			<!-- line !-->
			<div class="row line">
				<div class="col-lg-10 col-md-10 col-sm-9 col-xs-9">
					<span class="error">Your login is currently disabled<br>Please contact your administrator.</span>
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
					<span class="error"><%= Session("restrictedLoginMessage") %> Please contact your administrator if you need to login.</span>
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
					<span class="error"><%= Session("restrictedLoginMessage") %> Please contact your administrator if you need to login.</span>
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
            	<input type="email" placeholder="email@<%= CompanyDomainName %>" class="input" name="txtUsername" id="txtUsername">
            </div>
            <div class="col-lg-2 col-md-2 col-sm-3 col-xs-3">
            	<span class="icon"><img src="<%= BaseURL %>img/loginpage/email-icon.png" class="img-responsive"></span>
            </div>
        </div>
        <!-- eof line !-->
        
        <!-- line !-->
        <div class="row line">
        	<div class="col-lg-10 col-md-10 col-sm-9 col-xs-9">
            	<input type="password" placeholder="password" class="input" name="txtPassword" id="txtPassword">
            </div>
            <div class="col-lg-2 col-md-2 col-sm-3 col-xs-3">
            	<span class="icon"><img src="<%= BaseURL %>img/loginpage/password-icon.png" class="img-responsive"></span>
            </div>
        </div>
        <!-- eof line !-->
        
                
        <% If customLoginPage = true Then %>
        	<input type="hidden" name="txtClientKeyCustom" id="txtClientKeyCustom" value="<%= clientID %>">
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

        
        <!-- remember / login btn !-->
        <div class="row line">
        	<div class="col-lg-8">
        		&nbsp;
            </div>
            
            <div class="col-lg-4">
            	<button type="submit">Login</button>
            </div>
        </div>
        <!-- eof remember / login btn !-->
        
        <!-- forgot password !-->
         	<a href="<%= BaseURL %>reset-password.asp?clientID=<%= clientID %>">Can't access your account?</a> 
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
<!--#include file="inc/header-quicklogin.asp"--> 

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
    <h1 class="quicklogin">MDS Insight<br>
		<small>Business Analytics Engineered for Your Success</small></h1>
    </div>
    
</div>
<!-- eof logo / heading !-->


<form action="<%= BaseURL %>action_login.asp" method="POST" name="customQuickLoginForm" id="customQuickLoginForm" onSubmit="return checkFormQuickLogin();">

		
    <input type="hidden" name="txtUsername" id="txtUsername" class="form-control" value="<%= QuickUserEmail %>">
    <input type="hidden" name="txtQuickLogin" id="txtQuickLogin" class="form-control" value="true">
    <input type="hidden" name="txtUserNo" id="txtUserNo" class="form-control" value="<%= QuickUserNo %>">
    <input type="hidden" name="txtDestinationURL" id="txtDestinationURL" class="form-control" value="<%= QuickClientDestination %>">
    
    <% If customLoginPage = true Then %>
    	<!--<input type="hidden" name="txtClientKeyCustom" id="txtClientKeyCustom" class="form-control" value="<%= MUV_READ("ClientKeyForFileNames") %>">-->
    	<input type="hidden" name="txtClientKeyCustom" id="txtClientKeyCustom" class="form-control" value="<%= QuickClientID %>">
    <% Else %>
    	<input type="hidden" name="txtClientKey" id="txtClientKey" class="form-control" value="<%= QuickClientID %>">
    <% End If %>


 	<!-- login box !-->
    <div class="container">
    <div class="quicklogin-container">
    	
        <!-- right col !-->
        <div class="col-lg-6 equal-height signin-box-quicklogin">
        
        <h2 class="quicklogin">Sign In</h2>

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
        		<% If QuickUserPassword = "" Then %>
	            	<input type="password" placeholder="password" class="input" name="txtPassword" id="txtPassword">
	            <% Else %>
   	            	<input type="password" placeholder="password" class="input" name="txtPassword" id="txtPassword" value="<%=QuickUserPassword%>">
	            <%End If %>
            </div>
            <div class="col-lg-2 col-md-2 col-sm-3 col-xs-3">
            	<span class="icon"><img src="<%= BaseURL %>img/loginpage/password-icon.png" class="img-responsive"></span>
            </div>
        </div>
        <!-- eof line !-->
        

        <!-- remember / login btn !-->
        <div class="row line">            
            <div class="col-lg-11">
            	<button type="submit">Login</button>
            </div>
        </div>
        <!-- eof remember / login btn !-->
  

        <!-- remember / login btn !-->
        <div class="row line">
        	<div class="col-lg-12">
	        <!-- forgot password !-->
	         	<a href="<%= BaseURL %>reset-password-ql.asp?u=<%= QuickUserNo %>&c=<%= QuickClientID  %>">Can't access your account?</a> 
	         <!-- eof forgot password !-->
            </div>
        </div>
        <!-- eof remember / login btn !-->
      
        
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

<% If QuickUserPassword <> "" Then %>
	<script type="text/javascript">
		 // $( document ).ready() block.
		$( document ).ready(function() {
		    $("#customQuickLoginForm").submit();
		});  
	</script>
<% End If %><!--#include file="inc/footer-login-noTimeout.asp"-->
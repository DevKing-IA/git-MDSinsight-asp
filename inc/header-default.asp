<!--#include file="subsandfuncs.asp"-->
<!--#include file="InsightFuncs.asp"-->
<% MUV_Init() %>
<%
clientID = Request.QueryString("clientID")

If Right(clientID ,1) = "d" Then
	ClientKeyForFileNames = LEFT(clientID, (LEN(clientID)-1))
Else
	ClientKeyForFileNames = clientID
End If	

dummy = MUV_Write("ClientKeyForFileNames",ClientKeyForFileNames)

customLoginPage = false

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
		
		If HasCustomLoginPage(clientID) AND (clientID = "1071" OR clientID = "1071d") Then
			Response.Redirect("default_customLoginCCS.asp")
		ElseIf HasCustomLoginPage(clientID) AND clientID <> "1071" AND clientID <> "1071d" Then
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
    <!-- eof icons and notification styles !-->
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

	<!-- Easy Autocomplete Files -->
	<!-- JS file -->
	<script src="<%= BaseURL %>js/easyautocomplete/EasyAutocomplete-1.4.0/jquery.easy-autocomplete.js"></script> 
	<!-- CSS file -->
	<link rel="stylesheet" href="<%= BaseURL %>js/easyautocomplete/EasyAutocomplete-1.4.0/easy-autocomplete.css"> 
	<!-- Additional CSS Themes file - not required-->
	<link rel="stylesheet" href="<%= BaseURL %>js/easyautocomplete/EasyAutocomplete-1.4.0/easy-autocomplete.themes.css"> 

   
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
	
	 function checkForm()
	 {
        if (document.customLoginForm.txtUsername.value == "") {
            swal("Please enter your email address.");
            return false;
        }

        if (document.customLoginForm.txtPassword.value == "") {
           swal("Please enter your insight password.");
           return false;
        }

	    $("<table id='overlay'><tbody><tr><td>Preparing Your MDS Insight Experience...</td></tr><tr><td><img src='img/gears.gif'></td></tr></tbody></table>").css({
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
           swal("Please enter your insight password.");
           return false;
        }

        if (document.customLoginFormWithKey.txtClientKey.value == "") {
           swal("Please enter your insight client key.");
           return false;
        }
				
	    $("<table id='overlay'><tbody><tr><td>Loading Your MDS Insight Experience...</td></tr><tr><td><img src='img/gears.gif'></td></tr></tbody></table>").css({
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



<!-- eof validation javascript !-->


  </head>
  <!-- site starts here !-->
  
  
<% If Session("Userno") = "" Then Response.Redirect("../default.asp") %>

<!DOCTYPE html>
<!--[if lt IE 7 ]> <html class="no-js ie6 oldie" lang="en"> <![endif]-->
<!--[if IE 7 ]>    <html class="no-js ie7 oldie" lang="en"> <![endif]-->
<!--[if IE 8 ]>    <html class="no-js ie8 oldie" lang="en"> <![endif]-->
<!--[if IE 9 ]>    <html class="no-js ie9" lang="en"> <![endif]-->
<!--[if (gte IE 9)|!(IE)]><![endif]--><!-->
<html class="no-js" lang="en">
<!--<![endif]-->
<!--#include file="subsandfuncs.asp"-->
<!--#include file="protect.asp"-->
<!--#include file="InsightFuncs.asp"-->
<!--#include file="InSightFuncs_routing.asp"-->

  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <meta name="description" content="">
    <meta name="author" content="">

    <title>MDS Insight Dashboard</title>

    <!-- Bootstrap core CSS -->
    <link href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" rel="stylesheet">
    <!-- End Bootstrap core CSS -->

    <!-- Custom styles for MDS Insight -->
    <link href="<%= BaseURL %>css/dashboard.css" rel="stylesheet">
    <link href="<%= BaseURL %>css/screensize.css" rel="stylesheet">
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
    
    <!-- fonts !-->
    <link href='http://fonts.googleapis.com/css?family=Coda' rel='stylesheet' type='text/css'>
    <link href='http://fonts.googleapis.com/css?family=Oswald:400,300,700' rel='stylesheet' type='text/css'>
	<link href='http://fonts.googleapis.com/css?family=Indie+Flower' rel='stylesheet' type='text/css'>
    
    <!-- eof fonts !-->
	
	<!-- sort table script !-->
	<script src="<%= BaseURL %>js/sorttable.js"></script>
	<script src="<%= BaseURL %>js/sorttable1.js"></script>
	<!-- eof sort table script !-->

	<!-- *********************************************************************** -->
	<!-- IMPORTANT - USE OLDER VERSION OF JQUERY FOR SORTABLE PLUGIN             -->
	<!-- *********************************************************************** -->
  	<script src="http://code.jquery.com/jquery-1.11.2.min.js"></script>
	<!--<script src="https://code.jquery.com/jquery-3.1.1.js" integrity="sha256-16cdPddA6VdVInumRGo6IbivbERE8p7CQR3HzTBuELA=" crossorigin="anonymous"></script>  -->	
	<!-- *********************************************************************** -->
	
	<!-- Including jQuery UI CSS & jQuery Dialog UI Here-->
	<link href="http://ajax.googleapis.com/ajax/libs/jqueryui/1.9.2/themes/ui-darkness/jquery-ui.css" rel="stylesheet">
	<script src="http://ajax.googleapis.com/ajax/libs/jqueryui/1.9.2/jquery-ui.min.js"></script>
	<!-- End Including jQuery UI CSS & jQuery Dialog UI Here-->
	
	<!-- Bootstrap core JS - must load after jQuery -->
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"></script>
    <!-- End Bootstrap core JS - must load after jQuery -->
       
    <!-- jQuery Cookie Files To Save State In Place of Session Variables -->
    <script src="<%= BaseURL %>js/jquery.cookie.js"></script>
    <!-- End jQuery Cookie -->

</head>

<body class="field-service-body">

<% 
			
showForceNextStopMsg = "False"

If cInt(GetRemainingStopsByUserNo(Session("UserNo"))) > 0 Then

	If AutoForceSelectNextStopON(Session("UserNo")) = True Then

		'Check to see if they have made at least one delivery AND a next stop has not been chosen
	
		If cInt(GetTotalStopsByUserNo(Session("UserNo"))) <> cInt(GetRemainingStopsByUserNo(Session("UserNo"))) Then
		
			Set cnnCheckNextStopSelected = Server.CreateObject("ADODB.Connection")
			cnnCheckNextStopSelected.open Session("ClientCnnString")
		
			resultCheckNextStopSelected = False
				
			SQLCheckNextStopSelected = "SELECT CustNum, MIN(SequenceNumber) AS Expr1, Count(CustNum) AS Expr2, Max(Len(DeliveryStatus)) as Expr3, Max(ManualNextStop) as Expr4 FROM RT_DeliveryBoard "
			SQLCheckNextStopSelected = SQLCheckNextStopSelected & "WHERE (CustNum IN "
			SQLCheckNextStopSelected = SQLCheckNextStopSelected & "(SELECT CustNum FROM RT_DeliveryBoard AS RT_DeliveryBoard_1 "
			SQLCheckNextStopSelected = SQLCheckNextStopSelected & "WHERE (TruckNumber = '" & GetTruckNumberByUser(Session("UserNo")) & "' AND ManualNextStop = 1 AND DeliveryStatus IS NULL) GROUP BY CustNum)) GROUP BY CustNum ORDER BY Expr4 Desc,Expr3, Expr1"
							
							 
			Set rsCheckNextStopSelected = Server.CreateObject("ADODB.Recordset")
			rsCheckNextStopSelected.CursorLocation = 3 
			Set rsCheckNextStopSelected= cnnCheckNextStopSelected.Execute(SQLCheckNextStopSelected)
			
			If rsCheckNextStopSelected.EOF then
			
				showForceNextStopMsg = "True"
				
				protocol = "http" 
				domainName= Request.ServerVariables("SERVER_NAME") 
				fileName= Request.ServerVariables("SCRIPT_NAME") 
				queryString= Request.ServerVariables("QUERY_STRING")
				
				url = protocol & "://" & domainName & fileName

				If InStr(url,"viewStops.asp") = 0 Then
					Response.Redirect("viewStops.asp")
				End If
			End If

		End If
	End If
End If
%>
	
<% If Session("UserNo") = "" Then Response.Redirect("../../../logout.asp") ' They are not logged in %>
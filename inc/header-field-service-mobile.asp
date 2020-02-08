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

  <head>
	<meta charset="utf-8">
	<meta http-equiv="X-UA-Compatible" content="IE=edge">
	<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no"> 
	<meta name="description" content="">
	<meta name="author" content="">

    <title>MDS Insight Dashboard - Field Service Menu</title>   

    <!-- Bootstrap 4 Core CSS -->
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous">

    <!-- Custom styles for this template -->
    <link href="<%= BaseURL %>css/dashboard.css" rel="stylesheet">
    <link href="<%= BaseURL %>css/screensize.css" rel="stylesheet">
    <link href="<%= BaseURL %>css/global-insight-styles.css" rel="stylesheet">
 	<link href="<%= BaseURL %>css/segment.css" rel="stylesheet">

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
	<link href='http://fonts.googleapis.com/css?family=Lato:400' rel='stylesheet' type='text/css'>
    <!-- eof fonts !-->
    
    
	<!-- *********************************************************************** -->
	<!-- IMPORTANT - USE OLDER VERSION OF JQUERY FOR SORTABLE PLUGIN             -->
	<!-- *********************************************************************** -->
  	<script src="http://code.jquery.com/jquery-1.11.2.min.js"></script>
	<!--<script src="https://code.jquery.com/jquery-3.1.1.js" integrity="sha256-16cdPddA6VdVInumRGo6IbivbERE8p7CQR3HzTBuELA=" crossorigin="anonymous"></script> -->	
	<!-- *********************************************************************** -->
       
	<!-- Including jQuery UI CSS & jQuery Dialog UI Here-->
	<link href="http://ajax.googleapis.com/ajax/libs/jqueryui/1.9.2/themes/ui-darkness/jquery-ui.css" rel="stylesheet">
	<script src="http://ajax.googleapis.com/ajax/libs/jqueryui/1.9.2/jquery-ui.min.js"></script>
	<!-- End Including jQuery UI CSS & jQuery Dialog UI Here-->
    
  	<!-- Bootstrap 4 Core JS - must load after jQuery -->
	<!--<script src="https://code.jquery.com/jquery-3.2.1.slim.min.js" integrity="sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN" crossorigin="anonymous"></script>-->
	<script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js" integrity="sha384-ApNbgh9B+Y1QKtv3Rn7W3mgPxhU9K/ScQsAP7hUibX39j7fakFPskvXusvfa0b4Q" crossorigin="anonymous"></script>
	<script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js" integrity="sha384-JZR6Spejh4U02d8jOt6vLEHfe/JQGiRRSQQxSfFWpi1MquVdAyjUar5+76PVCmYl" crossorigin="anonymous"></script>
    <!-- End Bootstrap Core JS - must load after jQuery -->

        
	<!-- sweet alert jquery modal alerts !-->	
	<script src="<%= BaseURL %>js/sweetalert/sweetalert.min.js"></script>
	<link rel="stylesheet" type="text/css" href="<%= BaseURL %>js/sweetalert/sweetalert.css">
	<!-- end sweet alert jquery modal alerts !-->	

	<!-- jquery segmented controls for mobile driver navigation !-->	
	<script src="<%= BaseURL %>js/segment.js"></script>
	<!-- end jquery segmented controls  !-->
    
	<!-- Easy Autocomplete Files MOBILE VERSION -->
	<!-- JS file -->
	<script src="<%= BaseURL %>js/easyautocomplete/EasyAutocomplete-1.4.0/jquery.easy-autocomplete.js"></script> 
	<!-- CSS file -->
	<link rel="stylesheet" href="<%= BaseURL %>js/easyautocomplete/EasyAutocomplete-1.4.0/easy-autocomplete-mobile.css"> 
	<!-- Additional CSS Themes file - not required-->
	<link rel="stylesheet" href="<%= BaseURL %>js/easyautocomplete/EasyAutocomplete-1.4.0/easy-autocomplete.themes.css"> 
	<!-- end Easy Autocomplete MOBILE VERSION -->


    <!-- jQuery Cookie Files To Save State In Place of Session Variables -->
    <script src="<%= BaseURL %>js/jquery.cookie.js"></script>
    <!-- End jQuery Cookie -->

</head>

<body class="field-service-body">
	
<% If Session("UserNo") = "" Then Response.Redirect("../logout.asp") ' They are not logged in %>   
 




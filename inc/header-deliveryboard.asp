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
    
    
    <!-- Custom script for Delivery Board -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/highlight.js/8.3/highlight.min.js"></script>


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
	<!--<script src="<%= BaseURL %>js/sorttable1.js"></script>-->
	<!-- eof sort table script !-->

	<!-- *********************************************************************** -->
	<!-- IMPORTANT - USE OLDER VERSION OF JQUERY FOR SORTABLE PLUGIN             -->
	<!-- *********************************************************************** -->
  	<script src="http://code.jquery.com/jquery-1.11.2.min.js"></script>
  	
	<!--<script src="https://code.jquery.com/jquery-3.1.1.js" integrity="sha256-16cdPddA6VdVInumRGo6IbivbERE8p7CQR3HzTBuELA=" crossorigin="anonymous"></script>  -->	
	<!-- *********************************************************************** -->
	
	<!-- Including jQuery UI CSS & jQuery Dialog UI Here-->
	<!--<link href="http://ajax.googleapis.com/ajax/libs/jqueryui/1.9.2/themes/ui-darkness/jquery-ui.css" rel="stylesheet">
	<script src="http://ajax.googleapis.com/ajax/libs/jqueryui/1.9.2/jquery-ui.min.js"></script>-->
	<link rel="stylesheet" href="http://code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<script src="http://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
	<!-- End Including jQuery UI CSS & jQuery Dialog UI Here-->
	
	<!-- Bootstrap core JS - must load after jQuery -->
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"></script>
    <!-- End Bootstrap core JS - must load after jQuery -->
    
	<!-- sweet alert jquery modal alerts !-->	
	<script src="<%= BaseURL %>js/sweetalert/sweetalert.min.js"></script>
	<link rel="stylesheet" type="text/css" href="<%= BaseURL %>js/sweetalert/sweetalert.css">
	<!-- end sweet alert jquery modal alerts !-->	    

    <!-- dashboard sliding panel navigation !-->
    <link href="<%= BaseURL %>css/panel-menu.css" rel="stylesheet">
    <script src="<%= BaseURL %>js/panel-menu/jquery.navgoco.js"></script>
    
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

    
    <%
           
		Set cnnDelBoardSettings = Server.CreateObject("ADODB.Connection")
		cnnDelBoardSettings.open (Session("ClientCnnString"))
		
		SQLDelBoardSettings = "SELECT * FROM Settings_Global"
		Set rsDelBoardSettings = Server.CreateObject("ADODB.Recordset")
		rsDelBoardSettings.CursorLocation = 3 
		
		Set rsDelBoardSettings = cnnDelBoardSettings.Execute(SQLDelBoardSettings)
		
		If NOT rsDelBoardSettings.EOF Then
			DelBoardRoutesToIgnore = rsDelBoardSettings("DelBoardRoutesToIgnore")
			DelBoardPieTimerColor = rsDelBoardSettings("DelBoardPieTimerColor")
			DelBoardTitleText = rsDelBoardSettings("DelBoardTitleText")
			DelBoardTitleText = Replace(DelBoardTitleText,"'","")
			DelBoardTitleText = Replace(DelBoardTitleText,"~today~",FormatDateTime(Now(),2))
			DelBoardTitleText = Replace(DelBoardTitleText,"~dow~",WeekDayName(Datepart("w",Now())))
			DelBoardTitleTextFontColor = rsDelBoardSettings("DelBoardTitleTextFontColor")
		End If
		 
		Set rsDelBoardSettings = Nothing
		cnnDelBoardSettings.Close
		Set cnnDelBoardSettings = Nothing
		
		If DelBoardRoutesToIgnore = "" OR IsNUll(DelBoardRoutesToIgnore) Then 
			DelBoardRoutesToIgnoreCount = 0
		Else
			DelBoardRoutesToIgnoreArray = Split(DelBoardRoutesToIgnore,",")
			DelBoardRoutesToIgnoreCount = UBound(DelBoardRoutesToIgnoreArray) + 1
		End If
			
		If DelBoardPieTimerColor = "" Then DelBoardPieTimerColor = "000000"
		If IsNull(DelBoardPieTimerColor ) Then DelBoardPieTimerColor = "000000"
		If DelBoardTitleText = "" Then DelBoardTitleText = "Deliveries For " & WeekDayName(Datepart("w",Now())) & "," & FormatDateTime(Now(),2)
		If DelBoardTitleTextFontColor = "" Then DelBoardTitleTextFontColor = "000000"
		Session("DelBoardPieTimerColor") = Replace(DelBoardPieTimerColor,"#","") ' Just this one for Javascript
    
    %>

    <script>
	   $(document).ready(function(){												
	
	       //Navigation Menu Slider
	        $('#nav-expander').on('click',function(e){
	      		e.preventDefault();
	      		$('body').toggleClass('nav-expanded');
	      	});
	      	$('#nav-close').on('click',function(e){
	      		e.preventDefault();
	      		$('body').removeClass('nav-expanded');
	      	});
	
	      	// Initialize navgoco with default options
	        $(".main-menu").navgoco({
	            caret: '<span class="caret"></span>',
	            accordion: false,
	            openClass: 'open',
	            save: true,
	            cookie: {
	                name: 'navgoco',
	                expires: false,
	                path: '/'
	            },
	            slide: {
	                duration: 300,
	                easing: 'swing'
	            }
	        });
	        
	        //prepare tooltip for navbar
	        $("[rel=tooltip]").tooltip({ placement: 'right'});
	        
	        //if tabs are used, show tab name in location bar
		    if(location.hash) {
		        $('a[href=' + location.hash + ']').tab('show');
		    }
		    $(document.body).on("click", "a[data-toggle]", function(event) {
		        location.hash = this.getAttribute("href");
		    });

	
	      });     
	      
      </script>
      <!-- eof dashboard sliding panelnavigation !-->
      
	
<!-- delivery board custom styles !-->
<style type="text/css">


	 .material-switch > input[type=checkbox] {
	    display: none;   
	}
	
	.material-switch > label {
	    cursor: pointer;
	    height: 0px;
	    position: relative; 
	    width: 8px;  
	}

	.material-switch > label::before {
	    background: rgb(0, 0, 0);
	    box-shadow: inset 0px 0px 10px rgba(0, 0, 0, 0.5);
	    border-radius: 8px;
	    content: '';
	    height: 16px;
	    margin-top: -7px;
	    position:absolute;
	    opacity: 0.3;
	    transition: all 0.4s ease-in-out;
	    width: 40px;
	}
	.material-switch > label::after {
	    background: rgb(255, 255, 255);
	    border-radius: 16px;
	    box-shadow: 0px 0px 5px rgba(0, 0, 0, 0.3);
	    content: '';
	    height: 24px;
	    left: -4px;
	    margin-top: -8px;
	    position: absolute;
	    top: -4px;
	    transition: all 0.3s ease-in-out;
	    width: 24px;
	}
	.material-switch > input[type=checkbox]:checked + label::before {
	    background: inherit;
	    opacity: 0.5;
	}
	.material-switch > input[type=checkbox]:checked + label::after {
	    background: inherit;
	    left: 20px;
	}  
	
	.wrapper{
		margin-left:10px;
	}
	
	.heading-legend{
		margin-top:15px;
	}
	
	.heading-legend h4{
		font-weight:bold;
		margin:0px;
		padding:0px;
		text-transform:uppercase;
		text-align:center;
	}
	   

	.navbar-inverse .navbar-header{
		max-height:175px;
		font-size:14px;
	}
	
	.pause{
	   float:right;
	   /*margin:10px 30px 0px 0px;*/
	   margin:10px 30px 20px 10px;
	   color:#337ab7;
	}
	
	 .the-timer{
		float:right;
		margin-left:10px;
		margin-top: 5px;
	 }
	
	
	   
	.legend-complete{
		<% Response.Write("background:" & DelBoardCompletedColor  & ";") %>
	   padding:5px 5px 5px 5px;
	}
	
	.legend-inprogress{
		<% Response.Write("background:" & DelBoardInProgressColor  & ";") %>
	   padding:5px 5px 5px 5px;
	}

	.legend-nodelivery{
	   <% Response.Write("background:" & DelBoardSkippedColor  & ";")%>
	   padding:5px 5px 5px 5px;
	}
	
	.legend-nextstop{
	   <% Response.Write("background:" & DelBoardNextStopColor & ";")%>
	   padding:5px 5px 5px 5px;
	}

	.legend-priority{
		<% Response.Write("-webkit-box-shadow:inset 0px 0px 0px 3px " & DelBoardPriorityColor & ";")%>
		<% Response.Write("-moz-box-shadow:inset 0px 0px 0px 3px " & DelBoardPriorityColor & ";")%>
		<% Response.Write("box-shadow:inset 0px 0px 0px 3px " & DelBoardPriorityColor & ";")%>    
		<% Response.Write("background-color:#FFFFFF;")%> 		
	   padding:5px 5px 5px 5px;
	   color:#000;
	   text-align:center;
	}

	.legend-am{
		<% Response.Write("-webkit-box-shadow:inset 0px 0px 0px 3px " & DelBoardAMColor & ";")%>
		<% Response.Write("-moz-box-shadow:inset 0px 0px 0px 3px " & DelBoardAMColor & ";")%>
		<% Response.Write("box-shadow:inset 0px 0px 0px 3px " & DelBoardAMColor & ";")%>    
		<% Response.Write("background-color:#FFFFFF;")%> 		
	   padding:5px 5px 5px 5px;
	   color:#000;
	   text-align:center;
	}
	
	
</style>
<!-- end delivery board custom styles !-->

<!-- countdown script !-->
<script src="<%= BaseURL %>js/countdown/jquery.pietimer.js"></script>

<script type="text/javascript">

	
	$(document).ready(function() {

	    $('#deliveryBoardInvoiceOptionsModal').on('show.bs.modal', function () {
			$('#switchAutomaticRefresh').prop('checked', true).trigger("change");
	    });
	    
	    $('#myDeliveryBoardCompletedOrSkippedModal').on('show.bs.modal', function () {
			$('#switchAutomaticRefresh').prop('checked', true).trigger("change");
	    });
	    
	   	$('#deliveryBoardAddAlertModal').on('show.bs.modal', function () {
			$('#switchAutomaticRefresh').prop('checked', true).trigger("change");
			$('#deliveryBoardInvoiceOptionsModal').modal('hide');
	    });

	   	$('#deliveryBoardEditAlertModal').on('show.bs.modal', function () {
			$('#switchAutomaticRefresh').prop('checked', true).trigger("change");
			$('#deliveryBoardInvoiceOptionsModal').modal('hide');
	    });
	

	   	$('#deliveryBoardMarkAsPriorityWithTextingModal').on('show.bs.modal', function () {
			$('#switchAutomaticRefresh').prop('checked', true).trigger("change");
			$('#deliveryBoardInvoiceOptionsModal').modal('hide');
	    });
	    
	   	$('#deliveryBoardMarkAsNoPrioritytWithTextingModal').on('show.bs.modal', function () {
			$('#switchAutomaticRefresh').prop('checked', true).trigger("change");
			$('#deliveryBoardInvoiceOptionsModal').modal('hide');
	    });
	    

	   	$('#deliveryBoardMarkAsAMDeliverytWithTextingModal').on('show.bs.modal', function () {
			$('#switchAutomaticRefresh').prop('checked', true).trigger("change");
			$('#deliveryBoardInvoiceOptionsModal').modal('hide');
	    });
    
	    
		$('#deliveryBoardRemoveAMDeliverytWithTextingModal').on('show.bs.modal', function () {
			$('#switchAutomaticRefresh').prop('checked', true).trigger("change");
			$('#deliveryBoardInvoiceOptionsModal').modal('hide');
	    });


 		//**************************************************************************
 		//Special code here******
 		//The service ticket options modal leads to the opening of other modals
 		//So when we hide this modal, we want to keep the pause button paused

		$("#btnCloseFromTop").click(function() {
			$('#switchAutomaticRefresh').prop('checked', false).trigger("change");
		});
		
		$("#btnCloseFromBottom").click(function() {
			$('#switchAutomaticRefresh').prop('checked', false).trigger("change");
		});
 		
	    //$('#deliveryBoardInvoiceOptionsModal').on('hidden.bs.modal', function () {
		     //$('#switchAutomaticRefresh').prop('checked', true).trigger("change");
	    //});
	    //**************************************************************************
 
 	    $('#myDeliveryBoardCompletedOrSkippedModal').on('hidden.bs.modal', function () {
			$('#switchAutomaticRefresh').prop('checked', false).trigger("change");
	    });

 
		//These other three modals appear when a user has clicked a button on the
		//first modal, the options modal. So when these modals are closed, if the user
		//has not said to keep the board paused, then we can start the timer again.
		      	
		$('#deliveryBoardAddAlertModal').on('hidden.bs.modal', function () {
		     $('#switchAutomaticRefresh').prop('checked', false).trigger("change");
    	}); 	
    	    	
		$('#deliveryBoardEditAlertModal').on('hidden.bs.modal', function () {
		     $('#switchAutomaticRefresh').prop('checked', false).trigger("change");		
    	}); 	
    	
	   	$('#deliveryBoardMarkAsPriorityWithTextingModal').on('hidden.bs.modal', function () {
			 $('#switchAutomaticRefresh').prop('checked', false).trigger("change");
	    });
	    
	   	$('#deliveryBoardMarkAsNoPrioritytWithTextingModal').on('hidden.bs.modal', function () {
			 $('#switchAutomaticRefresh').prop('checked', false).trigger("change");
	    });
	    

	   	$('#deliveryBoardMarkAsAMDeliverytWithTextingModal').on('hidden.bs.modal', function () {
			 $('#switchAutomaticRefresh').prop('checked', false).trigger("change");
	    });
    
		$('#deliveryBoardRemoveAMDeliverytWithTextingModal').on('hidden.bs.modal', function () {
			 $('#switchAutomaticRefresh').prop('checked', false).trigger("change");
	    });
   
	});

	function Timer(callback, delay) {
	    var timerId, start, remaining = delay;
	
	    this.pause = function() {
	        window.clearTimeout(timerId);
	        remaining -= new Date() - start;
	    };
	
	    this.resume = function() {
	        start = new Date();
	        window.clearTimeout(timerId);
	        timerId = window.setTimeout(callback, remaining);
	    };
	
	    this.resume();
	}


	function hexToRgb(hex) {
		  var arrBuff = new ArrayBuffer(4);
		  var vw = new DataView(arrBuff);
		  vw.setUint32(0,parseInt(hex, 16),false);
		  var arrByte = new Uint8Array(arrBuff);
		
		  return "rgba(" + arrByte[1] + "," + arrByte[2] + "," + arrByte[3] + ",0.8)";
	}


	$(function(){
    

	    // Check if alert has been closed
	    if( $.cookie('alert-delboard-hidden-routes') === 'closed' ){
	
	        $('.alert').hide();
	
	    }
	
	     // Grab your button (based on your posted html)
	    $('.close').click(function( e ){
	
	        // Do not perform default action when button is clicked
	        e.preventDefault();
	
	        /* If you just want the cookie for a session don't provide an expires
	         Set the path as root, so the cookie will be valid across the whole site */
	        $.cookie('alert-delboard-hidden-routes', 'closed', { path: '/' });
	
	    });

	
		var rgbcolor = '<%=Session("DelBoardPieTimerColor")%>';

		var pagetimer = new Timer(function() {
		    location.reload();
		},  30*1000);
				
		$('#timer').pietimer({
			seconds: 30,
			color: hexToRgb(rgbcolor),
			height: 35,
			width: 35,
			is_reversed: true
		});
		
		$('#timer').pietimer('start');
		
		

		$('#switchAutomaticRefresh').on('change', function(e) {
		   if (this.checked) {
				if (!(e.isTrigger))
				{
					//alert ('human change event fired by clicking');
			        //Cookies.set('delivery-board-pause-autorefresh', 'true');
				}		   
				$('#timer').pietimer('pause');
				pagetimer.pause();
				return false;
		   }
		   else {
				if (!(e.isTrigger))
				{
					//alert ('human change event fired by clicking');
					//Cookies.set('delivery-board-pause-autorefresh', 'false');
				}			   
				$('#timer').pietimer('start');
				pagetimer.resume();
				return false;
		    }

		})	
	});
</script>
<!-- eof countdown script !-->

	
  </head>

<body>
	
	
<!-- license modal starts here !-->
<div class="modal fade" id="LicenseModal" tabindex="-1" role="dialog" aria-labelledby="myModalLabel">
	<div class="modal-dialog" role="document">
		<div class="modal-content">

			<div class="modal-header">
				<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
				<%
				LicArray = Split(MUV_READ("LicenseStatus"),"~")
				Response.Write("<h4 class='modal-title' id='myModalLabel'><span class='" & LicArray(0)  & "'><i class='fa fa-shield fa-lg' aria-hidden='true'></i> " & LicArray(1) & "</span></h4>")
				%>			
			</div>

			<div class="modal-body">
				<% Response.Write(LicArray(2))%>
			</div>

			<div class="modal-footer">
				<button type="button" class="btn btn-default" data-dismiss="modal">Close</button>
			</div>
		</div>
	</div>
</div>
<!-- license modal ends here !-->


 <!-- header !-->
 
 
<div class="navbar navbar-inverse navbar-fixed-top"> 
 <div class="container-fluid">     
    <div class="navbar-header pull-right">
	    <div class="row">
	    
	    		<div class="col-lg-1">
			      <a id="nav-expander" class="nav-expander fixed">MENU &nbsp;<i class="fa fa-bars fa-lg white"></i></a>
		 		</div>
		 		
		 		
		 		<div class="col-lg-4">
			    	<a href="<%= BaseURL %>main/default.asp" class="navbar-brand"><img src="<%= BaseURL %>clientfilesV/<%= MUV_Read("ClientID") %>/logos/logo.png"></a>
			    </div>     
			    
		    
			    <div class="col-lg-7">
					<!-- shield !-->
					<div class="pull-right shield-icons">
						<%
						LicArray = Split(MUV_READ("LicenseStatus"),"~")
						Response.Write("<a data-toggle='modal' href='#' data-target='#LicenseModal'><i class='fa fa-shield fa-lg " & LicArray(0) & "-icon' aria-hidden='true'></i></a> ")
						%>
					</div>
					<!-- eof shield !-->
			
				    <div class="dropdown pull-right logout-box">
						  <button id="dLabel" type="button" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false" class="button">
						    <strong><%= MUV_Read("DisplayName") %></strong>  <i class="fa fa-user fa-lg"></i>
						    <span class="caret"></span>
						  </button>
						   	
							<ul class="dropdown-menu" role="menu" aria-labelledby="dLabel">
								<li><a href="<%= BaseURL %>logout.asp"><i class="fa fa-fw fa-sign-out"></i> Sign Out</a></li>
								<%
								If LicArray(0)="blue" then
									Response.Write("<li>")
									Response.Write(Replace(MUV_ReadALL(),"}{","}<br>{"))
									Response.Write("</li>")
								End If
								%>

							</ul>
					</div>
			
					<!--#include file="topnav.asp"-->
				</div>
				
				
		
			</div>
			<!-- eof row!--> 
			
			<div class="row">	            
				<div class="col-lg-12" style="text-align:center;margin-top:-40px;margin-bottom:0px;">    
					<h2 style="display:inline;"><font color="<%=DelBoardTitleTextFontColor%>"><%=DelBoardTitleText%></font></h2>
				</div>
			</div>
				
				
			<div class="row">	            
				    
				<div class="col-lg-12" style="text-align:center;background-color: #DCE6E9;">    
	             
			  		<div style="padding-top:2px;margin-left:270px">

						<!-- complete !-->
				        <div class="col-lg-1">
				        	<div class="legend-complete">
				            	<h4>Complete</h4>
				            </div>
				        </div>
				        <!-- eof complete !-->
				        
				        <!-- no delivery !-->
				        <div class="col-lg-1">
				        	<div class="legend-nodelivery">
				            	<h4>No Delivery</h4>
				            </div>
				        </div>
				        <!-- eof no delivery !-->
				        
				        <!-- next stop !-->
				        <div class="col-lg-1">
				        	<div class="legend-nextstop">
				            	<h4>Next Stop</h4>
				            </div>
				        </div>
				        <!-- eof next stop !-->
				    	
				        <!-- in progress !-->
				        <div class="col-lg-1">
				        	<div class="legend-inprogress">
				            	<h4>In Progress</h4>
				            </div>
				        </div>
				        <!-- eof in progress stop !-->
				        
				        <!-- priority stop !-->
				        <div class="col-lg-1">
				        	<div class="legend-priority">
				            	<h4>Priority</h4>
				            </div>
				        </div>
				        <!-- eof priority stop !-->
					   				    	
				        <!-- am delivery !-->
				        <div class="col-lg-1">
				        	<div class="legend-am">
				            	<h4>AM Delivery</h4>
				            </div>
				        </div>
				        <!-- eof am delivery !-->
				        
				        <!-- in progress !-->
				        <div class="col-lg-2" style="width:19%">
					  		<div class="pause">
								Pause Automatic Refresh&nbsp;&nbsp;
								<div class="material-switch pull-right">
									<input id="switchAutomaticRefresh" name="chkAutomaticRefresh" type="checkbox"/>
									<label for="switchAutomaticRefresh" class="label-primary"></label>
								</div>
							</div>
							<div id="timer" class="the-timer" style="height:30px;"></div>
				        </div>
			        	<!-- eof in progress stop !-->
			        

				        
					</div>
					
				</div>
				
            

			</div>
		</div>
		<!-- eof navbar-header!--> 
	</div>
	<!-- eof container fluid !-->
</div>
<!-- eof navbar !-->

<!--#include file="leftnav.asp"-->


       
 <!-- eof side bar !-->

        <!-- content area !-->
        <div class="wrapper">

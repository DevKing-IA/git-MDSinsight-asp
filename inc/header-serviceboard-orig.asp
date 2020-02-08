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
<!--#include file="InsightFuncs_Service.asp"-->

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

	<!-----------------IMPORTANT FILE FOR SERVICE BOARD HEADER ------------------------------------------->
    <!-- JavaScript Cookie Files To Save State of Dismissed Alerts -->
    <script src="<%= BaseURL %>js/js.cookie.js"></script>
    <!-- End JavaScript Cookie -->
    <!-----------------END IMPORTANT FILE FOR SERVICE BOARD HEADER ---------------------------------------->

	 
	<link rel="stylesheet" href="<%= BaseURL %>js/bootstrap-multiselect/bootstrap-multiselect.css" type="text/css">
	<script type="text/javascript" src="<%= BaseURL %>js/bootstrap-multiselect/bootstrap-multiselect.js"></script>
	   
    <%
		SQL = "SELECT * FROM Settings_FieldService"
		Set cnn8 = Server.CreateObject("ADODB.Connection")
		cnn8.open (Session("ClientCnnString"))
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = 3 
		Set rs = cnn8.Execute(SQL)
		If not rs.EOF Then
			FSBoardKioskGlobalTitleText = rs("FSBoardKioskGlobalTitleText")
			FSBoardKioskGlobalTitleTextFontColor = rs("FSBoardKioskGlobalTitleTextFontColor")
			FSBoardKioskGlobalTitleGradientColor = rs("FSBoardKioskGlobalTitleGradientColor")
			FSBoardKioskGlobalColorPieTimer = rs("FSBoardKioskGlobalColorPieTimer")	
			FSBoardKioskGlobalColorUrgent = rs("FSBoardKioskGlobalColorUrgent")
			FSBoardKioskGlobalColorAwaitingDispatch = rs("FSBoardKioskGlobalColorAwaitingDispatch")
			FSBoardKioskGlobalColorAwaitingAcknowledgement = rs("FSBoardKioskGlobalColorAwaitingAcknowledgement")
			FSBoardKioskGlobalColorDispatchAcknowledged = rs("FSBoardKioskGlobalColorDispatchAcknowledged")
			FSBoardKioskGlobalColorDispatchDeclined = rs("FSBoardKioskGlobalColorDispatchDeclined")
			FSBoardKioskGlobalColorRedoSwap = rs("FSBoardKioskGlobalColorRedoSwap")
			FSBoardKioskGlobalColorRedoWaitForParts = rs("FSBoardKioskGlobalColorRedoWaitForParts")
			FSBoardKioskGlobalColorRedoFollowUp = rs("FSBoardKioskGlobalColorRedoFollowUp")
			FSBoardKioskGlobalColorRedoUnableToWork = rs("FSBoardKioskGlobalColorRedoUnableToWork")
			FSBoardKioskGlobalColorEnRoute = rs("FSBoardKioskGlobalColorEnRoute")
			FSBoardKioskGlobalColorOnSite = rs("FSBoardKioskGlobalColorOnSite")	
			FSBoardKioskGlobalColorClosed = rs("FSBoardKioskGlobalColorClosed")
			FilterChangeIndicatorAndButtonColor = rs("FilterChangeIndicatorAndButtonColor")	
			ShowSeparateFilterChangesTabOnServiceScreen = rs("ShowSeparateFilterChangesTabOnServiceScreen")				
		End If
		set rs = Nothing
		cnn8.close
		set cnn8 = Nothing
		
		FSBoardKioskGlobalTitleText = Replace(FSBoardKioskGlobalTitleText,"'","")
		FSBoardKioskGlobalTitleText = Replace(FSBoardKioskGlobalTitleText,"~today~",FormatDateTime(Now(),2))
		FSBoardKioskGlobalTitleText = Replace(FSBoardKioskGlobalTitleText,"~dow~",WeekDayName(Datepart("w",Now())))
		
		Session("FSBoardKioskGlobalColorPieTimer") = Replace(FSBoardKioskGlobalColorPieTimer,"#","") ' Just this one for Javascript
		
		If FilterChangeIndicatorAndButtonColor = "" Then FilterChangeIndicatorAndButtonColor = "#dddd53"
		If IsNull(FilterChangeIndicatorAndButtonColor) Then FilterChangeIndicatorAndButtonColor = "#dddd53"	
		
    
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

			
			
			$('#lstExistingRegionList').multiselect({
				   buttonTitle: function(options, select) {
					    var selected = '';
					    options.each(function () {
					      selected += $(this).text() + ', ';
					    });
					    return selected.substr(0, selected.length - 2);
					  },
					buttonClass: 'btn btn-primary',
					buttonWidth: '250px',
					maxHeight: 400,
					dropRight:true,
					enableFiltering:true,
					filterPlaceholder:'Search',
					enableCaseInsensitiveFiltering:true,
					// possible options: 'text', 'value', 'both'
					filterBehavior:'text',
					includeFilterClearBtn:true,
					nonSelectedText:'No Regions Selected For Filtering',
					numberDisplayed: 20,
				    onChange: function() {
				        var selected = this.$select.val();
				        $("#lstSelectedRegionList").val(selected);
				        RegionsToView = $("#lstSelectedRegionList").val();
				        UserNo = '<%= Session("UserNo") %>'
			
				    	$.ajax({
							type:"POST",
							url: "../inc/InSightFuncs_AjaxForServiceModals.asp",
							cache: false,
							data: "action=SetRegionFilterListByUserForServiceBoard&RegionsToView=" + encodeURIComponent(RegionsToView) + "&UserNo=" + encodeURIComponent(UserNo),
							success: function(response)
							 {
				               	 //location.reload();              	 
				             }
						});
	
				    },
				    onDropdownHide: function() {
						location.reload();
				    }
				    
		    			
		    });	
		    
			//*************************************************************************************************
			//Load the bootstrap multiselect box with the current threshhold report users preselected
			//*************************************************************************************************
			var data= $("#lstSelectedRegionList").val();
			//Make an array
			
			if (data) {
				var dataarray=data.split(",");
				// Set the value
				$("#lstExistingRegionList").val(dataarray);
				// Then refresh
				$("#lstExistingRegionList").multiselect("refresh");
			}
			//*************************************************************************************************			
	

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
		font-size:14px;
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
	
	.col-lg-1 {
	    /*width: 6.8% !important;*/
	}	

	.btn-group-lg>.btn, .buttonFilterChangeIndicatorAndButtonColor {
	    padding: 10px 16px;
	    font-size: 16px;
	    line-height: 1.3333333;
	    border-radius: 6px;
	    margin-top: 1px;
	    color:#fff;
	    <% Response.Write("background-color:" & FilterChangeIndicatorAndButtonColor & " !important;")%>
	}		
	.navbar>.container .navbar-brand, .navbar>.container-fluid .navbar-brand {
	    margin-left: 10px !important;
	}
	
	.btn-group-lg>.btn, .btn-lg {
	    padding: 10px 16px;
	    font-size: 16px;
	    line-height: 1.3333333;
	    border-radius: 6px;
	    margin-top: 1px;
	}	
	
</style>
<!-- end delivery board custom styles !-->

<!-- countdown script !-->
<script src="<%= BaseURL %>js/countdown/jquery.pietimer.js"></script>
<script type="text/javascript">

	
	$(document).ready(function() {

	    $('#serviceBoardTicketOptionsModal').on('show.bs.modal', function () {
			//if (Cookies.get('service-board-pause-autorefresh') == 'false' ){
				$('#switchAutomaticRefresh').prop('checked', true).trigger("change");
			//}
	    });
	    
	    
	   	$('#serviceBoardXferModal').on('show.bs.modal', function () {
			//if (Cookies.get('service-board-pause-autorefresh') == 'false' ){
				$('#switchAutomaticRefresh').prop('checked', true).trigger("change");
				$('#serviceBoardTicketOptionsModal').modal('hide');
			//}
	    });

	   	$('#serviceBoardSetAlertModal').on('show.bs.modal', function () {
			//if (Cookies.get('service-board-pause-autorefresh') == 'false' ){
				$('#switchAutomaticRefresh').prop('checked', true).trigger("change");
				$('#serviceBoardTicketOptionsModal').modal('hide');
			//}
	    });
	    
	   	$('#serviceBoardRequestETAModal').on('show.bs.modal', function () {
			//if (Cookies.get('service-board-pause-autorefresh') == 'false' ){
				$('#switchAutomaticRefresh').prop('checked', true).trigger("change");
				$('#serviceBoardTicketOptionsModal').modal('hide');
			//}
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
		
	    //$('#serviceBoardTicketOptionsModal').on('hidden.bs.modal', function () {
		   // if (Cookies.get('service-board-pause-autorefresh') == 'false' ){
		       //$('#switchAutomaticRefresh').prop('checked', true).trigger("change");
		    //}
	    //});
	    //**************************************************************************
 
 
 
		//These other three modals appear when a user has clicked a button on the
		//first modal, the options modal. So when these modals are closed, if the user
		//has not said to keep the board paused, then we can start the timer again.
		      	
		$('#serviceBoardXferModal').on('hidden.bs.modal', function () {
		    //if (Cookies.get('service-board-pause-autorefresh') == 'false' ){
		       $('#switchAutomaticRefresh').prop('checked', false).trigger("change");
		   // }	
    	}); 	
    	
 		$('#serviceBoardSetAlertModal').on('hidden.bs.modal', function () {
		    //if (Cookies.get('service-board-pause-autorefresh') == 'false' ){
		       $('#switchAutomaticRefresh').prop('checked', false).trigger("change");
		    //}		
    	}); 	
    	
		$('#serviceBoardRequestETAModal').on('hidden.bs.modal', function () {
		   // if (Cookies.get('service-board-pause-autorefresh') == 'false' ){
		       $('#switchAutomaticRefresh').prop('checked', false).trigger("change");
		    //} 		
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
    
	
		var rgbcolor = '<%=Session("FSBoardKioskGlobalColorPieTimer")%>';

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
			        Cookies.set('service-board-pause-autorefresh', 'true');
				}		   
				$('#timer').pietimer('pause');
				pagetimer.pause();
				return false;
		   }
		   else {
				if (!(e.isTrigger))
				{
					//alert ('human change event fired by clicking');
					Cookies.set('service-board-pause-autorefresh', 'false');
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
		 		
		 		
		 		<div class="col-lg-3">
			    	<a href="<%= BaseURL %>main/default.asp" class="navbar-brand"><img src="<%= BaseURL %>clientfilesV/<%= MUV_Read("ClientID") %>/logos/logo.png"></a>
			    </div>     
			    
		            
				<div class="col-lg-6" style="margin-top:20px">    
					<h2 style="display:inline;word-wrap: break-word;"><font color="<%=FSBoardKioskGlobalTitleTextFontColor%>"><%=FSBoardKioskGlobalTitleText%></font></h2>
			  		<div class="pause">
						Pause Automatic Refresh&nbsp;&nbsp;
						<div class="material-switch pull-right">
							<input id="switchAutomaticRefresh" name="chkAutomaticRefresh" type="checkbox"/>
							<label for="switchAutomaticRefresh" class="label-primary"></label>
						</div>
					</div>
					<div id="timer" class="the-timer" style="height:30px;"></div>
				</div>
				
		    
			    <div class="col-lg-2">
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
				    
				<div class="col-lg-12" style="text-align:center;background-color: #DCE6E9;">    
	             
			  		<div style="padding-top:2px;margin-left:0px">
			  		
						<div class="col-lg-1" style="margin-left:30px">
							<a href="dispatchcenter/main.asp">
								<button type="button" class="btn btn-warning btn-lg">Dispatch Center</button>
							</a>
						</div>

						<% If FilterChangeModuleOn() Then %>
						
							<div class="col-lg-1" style="margin-left:-10px">
								<a href="main.asp">
									<button type="button" class="btn btn-success btn-lg">Service Tickets</button>
								</a>
							</div>
						
							<div class="col-lg-1" style="margin-right:30px; margin-left:-45px">
								<a href="<%= baseURL %>service/filters/main.asp">
									<button type="button" class="btn btn-lg buttonFilterChangeIndicatorAndButtonColor">Filters</button>
								</a>
							</div>							
										
						<% Else %>
						
							<div class="col-lg-1" style="margin-left:-10px">
								<a href="main.asp" style="margin-right:30px;">
									<button type="button" class="btn btn-success btn-lg">Service Tickets</button>
								</a>
							</div>						
			
						<% End If %>
						  		

				        <!-- in progress !-->
				        <div class="col-lg-1" style=";margin-left:20px">
				        	<div class="legend-awaiting-dispatch">
				            	<h4>Awaiting <%= GetTerm("Dispatch") %></h4>
				            </div>
				        </div>
				        <!-- eof in progress stop !-->
	
				        <!-- in progress !-->
				        <div class="col-lg-1">
				        	<div class="legend-awaiting-acknowledgement">
				            	<h4><%= GetTerm("Dispatched") %></h4>
				            </div>
				        </div>
				        <!-- eof in progress stop !-->
				        
				        <!-- in declined !-->
				        <% If FS_TechCanDecline() Then %>
					        <div class="col-lg-1">
					        	<div class="legend-declined">
					            	<h4><%= GetTerm("Declined") %></h4>
					            </div>
					        </div>
					        <!-- eof in declined stop !-->
				        <% End If %>
	
	
				        <!-- in progress !-->
				        <div class="col-lg-1">
				        	<div class="legend-dispatch-acknowledged">
				            	<h4><%= GetTerm("Acknowledged") %></h4>
				            </div>
				        </div>
				        <!-- eof in progress stop !-->
			  		
				        
				        <!-- no delivery !-->
				        <div class="col-lg-1">
				        	<div class="legend-enroute">
				            	<h4>En Route</h4>
				            </div>
				        </div>
				        <!-- eof no delivery !-->
				        
				        <!-- next stop !-->
				        <div class="col-lg-1">
				        	<div class="legend-onsite">
				            	<h4>On Site</h4>
				            </div>
				        </div>
				        <!-- eof next stop !-->
						        
				        <!-- in progress !-->
				        <div class="col-lg-1">
				        	<div class="legend-redo-swap">
				            	<h4>Swap (<%= GetTerm("Redo") %>)</h4>
				            </div>
				        </div>
				        <!-- eof in progress stop !-->
				       			        
				        <!-- in progress !-->
				        <div class="col-lg-1">
				        	<div class="legend-redo-waitforparts">
				            	<h4>Wait For Parts (<%= GetTerm("Redo") %>)</h4>
				            </div>
				        </div>
				        <!-- eof in progress stop !-->
	
				        <!-- in progress !-->
				        <div class="col-lg-1">
				        	<div class="legend-redo-followup">
				            	<h4>Follow Up (<%= GetTerm("Redo") %>)</h4>
				            </div>
				        </div>
				        <!-- eof in progress stop !-->
	
				        <!-- in progress !-->
				        <div class="col-lg-1">
				        	<div class="legend-redo-unabletowork">
				            	<h4>Unable to Work (<%= GetTerm("Redo") %>)</h4>
				            </div>
				        </div>
				        <!-- eof in progress stop !-->
				        
				        <!-- in progress !-->
				        <div class="col-lg-1">
				        	<div class="legend-closed">
				            	<h4>Closed</h4>
				            </div>
				        </div>
				        <!-- eof in progress stop !-->        
				        
				        <!-- in progress !-->
				        <div class="col-lg-1">
				        	<div class="legend-urgent">
				            	<h4><%= GetTerm("Urgent") %></h4>
				            </div>
				        </div>
				        <!-- eof in progress stop !-->
			        
				
					 	<div class="col-lg-1 pull-right" style="margin-right:100px;">
						<% 
							Set cnnUserRegionsForServiceBoard = Server.CreateObject("ADODB.Connection")
							cnnUserRegionsForServiceBoard.open (Session("ClientCnnString"))
							Set rsUserRegionsForServiceBoard = Server.CreateObject("ADODB.Recordset")
							rsUserRegionsForServiceBoard.CursorLocation = 3 
							
							SQLUserRegionsForServiceBoard = "SELECT UserRegionsToViewService FROM tblUsers WHERE UserNo = " & Session("UserNo")
							Set rsUserRegionsForServiceBoard = cnnUserRegionsForServiceBoard.Execute(SQLUserRegionsForServiceBoard)
						
							RegionList = rsUserRegionsForServiceBoard("UserRegionsToViewService")
							
							set rsUserRegionsForServiceBoard = Nothing
							cnnUserRegionsForServiceBoard.close
							set cnnUserRegionsForServiceBoard = Nothing
						%>	
						<input type="hidden" name="lstSelectedRegionList" id="lstSelectedRegionList" value="<%= RegionList %>">
						<select id="lstExistingRegionList" multiple="multiple" name="lstExistingRegionList">
							<%	
								
							Set cnnRegionList = Server.CreateObject("ADODB.Connection")
							cnnRegionList.open Session("ClientCnnString")
				
							SQLRegionList = "SELECT * FROM AR_Regions ORDER BY InternalRecordIdentifier"
							
							Set rsRegionList = Server.CreateObject("ADODB.Recordset")
							rsRegionList.CursorLocation = 3 
							Set rsRegionList = cnnRegionList.Execute(SQLRegionList)
							
							If Not rsRegionList.EOF Then
								Do While Not rsRegionList.EOF
								
									RegionName = rsRegionList("Region")
									Response.Write("<option value='" & rsRegionList("InternalRecordIdentifier") & "'>" & RegionName & "</option>")
							
									rsRegionList.MoveNext
								Loop
							End If
				
							Set rsRegionList = Nothing
							cnnRegionList.Close
							Set cnnRegionList = Nothing
								
							%>
						</select>				
				 	</div>

			        
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

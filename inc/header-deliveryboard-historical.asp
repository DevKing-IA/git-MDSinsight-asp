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
	<link href="http://ajax.googleapis.com/ajax/libs/jqueryui/1.9.2/themes/ui-darkness/jquery-ui.css" rel="stylesheet">
	<script src="http://ajax.googleapis.com/ajax/libs/jqueryui/1.9.2/jquery-ui.min.js"></script>
	<!-- End Including jQuery UI CSS & jQuery Dialog UI Here-->
	
	<!-- Bootstrap core JS - must load after jQuery -->
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"></script>
    <!-- End Bootstrap core JS - must load after jQuery -->

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

    
	<!-- datepicker for historical delivery board !-->
	<script src="<%= baseURL %>js/moment.min.js" type="text/javascript"></script>
	<link href="<%= baseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.css" rel="stylesheet" type="text/css">
	<script src="<%= baseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.js" type="text/javascript"></script>
	<!-- end datepicker for historical delivery board !-->
	
    <!-- jQuery Cookie Files To Save State In Place of Session Variables -->
    <script src="<%= BaseURL %>js/jquery.cookie.js"></script>
    <!-- End jQuery Cookie -->
	<%
	
	txtDeliveryBoardDate = Request.Form("txtDeliveryBoardDate")

	If txtDeliveryBoardDate = "" then
		DateToRetrieve = DelBoardHistMostRecentDate()
	Else
		DateToRetrieve = txtDeliveryBoardDate
	End If

	%>
    <script>
	   $(document).ready(function(){	
	   
	        $('#datepicker1').datetimepicker({
	        	format: 'MM/DD/YYYY',
	        	useCurrent: false,
	        	//defaultDate: moment(momentString),
	        	maxDate:moment().add(-1, 'days')
	        });
	
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
	    width: 40px;  
	}

	.material-switch > label::before {
	    background: rgb(0, 0, 0);
	    box-shadow: inset 0px 0px 10px rgba(0, 0, 0, 0.5);
	    border-radius: 8px;
	    content: '';
	    height: 16px;
	    margin-top: -8px;
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
		margin-top:5px;
	}
	
	.heading-legend h4{
		font-weight:bold;
		margin:0px;
		padding:0px;
		text-transform:uppercase;
		text-align:center;
	}
	   
	.legend-complete{
		<% Response.Write("background:" & DelBoardCompletedColor  & ";") %>
	   padding:20px 15px 20px 15px;
	}
	
	.legend-nodelivery{
	   <% Response.Write("background:" & DelBoardSkippedColor  & ";")%>
	   padding:20px 15px 20px 15px;
	}
	
	.legend-nextstop{
	   <% Response.Write("background:" & DelBoardNextStopColor & ";")%>
	   padding:20px 15px 20px 15px;
	}

	.legend-priority{
		<% Response.Write("-webkit-box-shadow:inset 0px 0px 0px 3px " & DelBoardPriorityColor & ";")%>
		<% Response.Write("-moz-box-shadow:inset 0px 0px 0px 3px " & DelBoardPriorityColor & ";")%>
		<% Response.Write("box-shadow:inset 0px 0px 0px 3px " & DelBoardPriorityColor & ";")%>    
		<% Response.Write("background-color:#FFFFFF;")%> 		
	   padding:20px 15px 20px 15px;
	   color:#000;
	   text-align:center;
	}

	.legend-am{
		<% Response.Write("-webkit-box-shadow:inset 0px 0px 0px 3px " & DelBoardAMColor & ";")%>
		<% Response.Write("-moz-box-shadow:inset 0px 0px 0px 3px " & DelBoardAMColor & ";")%>
		<% Response.Write("box-shadow:inset 0px 0px 0px 3px " & DelBoardAMColor & ";")%>    
		<% Response.Write("background-color:#FFFFFF;")%> 		
	   padding:20px 15px 20px 15px;
	   color:#000;
	   text-align:center;
	}

	.navbar-inverse .navbar-header{
		max-height:75px;
	}
	
	.pause{
	   float:left;
	   margin:10px 30px 0px 0px;
	   color:#337ab7;
	}
	
	 .the-timer{
		 float:left;
	 }
	
	
	
</style>
<!-- end delivery board custom styles !-->

	
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
		 		
		 		
		 		<div class="col-lg-2">
			    	<a href="<%= BaseURL %>main/default.asp" class="navbar-brand"><img src="<%= BaseURL %>clientfilesV/<%= MUV_Read("ClientID") %>/logos/logo.png"></a>
			    </div>     
		    
	          <!-- legend !-->
          	  <div class="col-lg-6">
		          <!-- legend starts here !-->
					<div class="row heading-legend">
					
				        <!-- complete !-->
				        <div class="col-lg-2">
				        	<div class="legend-complete">
				            	<h4>Complete</h4>
				            </div>
				        </div>
				        <!-- eof complete !-->
				        
				        <!-- no delivery !-->
				        <div class="col-lg-2">
				        	<div class="legend-nodelivery">
				            	<h4>No Delivery</h4>
				            </div>
				        </div>
				        <!-- eof no delivery !-->
				        
				        <!-- next stop !-->
				        <div class="col-lg-2">
				        	<div class="legend-nextstop">
				            	<h4>Next Stop</h4>
				            </div>
				        </div>
				        <!-- eof next stop !-->
				        
				        <!-- priority stop !-->
				        <div class="col-lg-2">
				        	<div class="legend-priority">
				            	<h4>Priority</h4>
				            </div>
				        </div>
				        <!-- eof priority stop !-->
					   				    	
				        <!-- am delivery !-->
				        <div class="col-lg-2">
				        	<div class="legend-am">
				            	<h4>AM Delivery</h4>
				            </div>
				        </div>
				        <!-- eof am delivery !-->
				        
	                     <div class="col-lg-2">
	                     	<!-- datepicker !-->
	                     	<form action="deliveryBoardHistorical.asp" method="POST" name="frmSetDeliveryBoardHistoricalDate" id="frmSetDeliveryBoardHistoricalDate">
			                     <div class="col-lg-11">
									<!-- Bootstrap datepicker for filtering leads by date -->
						            <div class="form-group">
						                <div class="input-group date" id="datepicker1">
						                    <input type="text" class="form-control" name="txtDeliveryBoardDate" id="txtDeliveryBoardDate" placeholder="<%= DateToRetrieve %>">
						                    <span class="input-group-addon">
						                        <span class="glyphicon glyphicon-calendar"></span>
						                    </span>
						                </div>
						            </div>	
	            	             </div>
			                    <div class="col-lg-1">
									<button type="submit" class="btn btn-success">Set Date</button>
			                    </div>
			              	</form>
					      </div>
					     
					    <!-- eof datepicker !-->
    
    			</div>
				<!-- legend ends here !-->
	          </div>
	          <!-- eof legend !-->

	          
					 
		
			
	          
			    
		    <div class="col-lg-3">
            
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
			
					
				</div>
				
				
		
			</div>
			<!-- eof row!--> 
		</div>
		<!-- eof navbar-header!--> 
	</div>
	<!-- eof container fluid !-->
</div>
<!-- eof navbar !-->


<!-- dashboard starts here !-->

<!--#include file="leftnav.asp"-->


       
 <!-- eof side bar !-->

        <!-- content area !-->
        <div class="wrapper">
        


 <!DOCTYPE html>
<!--[if lt IE 7 ]> <html class="no-js ie6 oldie" lang="en"> <![endif]-->
<!--[if IE 7 ]>    <html class="no-js ie7 oldie" lang="en"> <![endif]-->
<!--[if IE 8 ]>    <html class="no-js ie8 oldie" lang="en"> <![endif]-->
<!--[if IE 9 ]>    <html class="no-js ie9" lang="en"> <![endif]-->
<!--[if (gte IE 9)|!(IE)]><![endif]--><!-->
<html class="no-js" lang="en">
<!--<![endif]-->
<!--#include file="../../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../../inc/InsightFuncs.asp"-->
<!--#include file="../../../inc/InSightFuncs_Routing.asp"-->
<%
UsageMessage = "http://www.mdsinsight.com/directLaunch/kiosks/routing/deliveryboardKioskNoPaging.asp?pp={Your Passphrase}&cl={Your Client ID}&ri={Refresh Interval In Seconds}&tn={Truck#,Truck#,Trick#} (Max 5 trucks)"
UsageMessage = UsageMessage & "<br>-OR-<br>"
UsageMessage = UsageMessage & "For last parameter use &tn=auto"

'These must be declared here
Dim DelBoardNextStopColor
Dim DelBoardScheduledColor	
Dim DelBoardCompletedColor	
Dim DelBoardInProgressColor			
Dim DelBoardSkippedColor		
Dim DelBoardProfitDollars		
Dim DelBoardAtOrAboveProfitColor			
Dim DelBoardBelowProfitColor			
Dim DelBoardUserAlertColor
Dim DelBoardAMColor
Dim DelBoardPriorityColor
Dim DelBoardRoutesToIgnore
Dim DelBoardTitleText
Dim DelBoardTitleTextFontColor
Dim DelBoardTitleGradientColor
Dim DelBoardPieTimerColor	

PassPhrase = Request.QueryString("pp")
ClientKey = Request.QueryString("cl")

If Request.QueryString("ri") <> "" Then Session("RefreshInterval") = Request.QueryString("ri") else Session("RefreshInterval") = 60

If PassPhrase = "" or ClientKey = "" Then
	Response.Write(UsageMessage)
	Respnse.End
End If

Session("PassPhrase") = PassPhrase 
Session("ClientKey") = ClientKey 


SQL = "SELECT * FROM tblServerInfo where clientKey='"& ClientKey &"'"

Set Connection = Server.CreateObject("ADODB.Connection")
Set Recordset = Server.CreateObject("ADODB.Recordset")
Connection.Open InsightCnnString
'Response.Write("InsightCnnString:" & InsightCnnString & "<br>")

'Open the recordset object executing the SQL statement and return records
Recordset.Open SQL,Connection,3,3
'Response.Write("SQL:" & SQL& "<br>")

'First lookup the ClientKey in tblServerInfo
'If there is no record with the entered client key, close connection
'and go back to login with QueryString
If Recordset.recordcount <= 0 then
	Recordset.close
	Connection.close
	set Recordset=nothing
	set Connection=nothing
	%>MDS Insight: Unable to connect to SQL database.<%
	Response.End
Else
	Session("ClientCnnString") = "Driver={SQL Server};Server=" & Recordset.Fields("dbServer")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Database=" & Recordset.Fields("dbCatalog")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Uid=" & Recordset.Fields("dbLogin")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Pwd=" & Recordset.Fields("dbPassword") & ";"
	dummy = MUV_Write("SQL_Owner",Recordset.Fields("dbLogin"))
	dummy = MUV_Write("ClientID",Recordset.Fields("clientkey"))
	dummy = MUV_Write("serviceModule",Recordset.Fields("serviceModule"))
	If PassPhrase <>  Recordset.Fields("directLaunchPassphrase") Then
		Response.Write("Access Denied")
		Session.Abandon
		Response.End
	End If
	Recordset.close
	Connection.close
End If

Call Read_Settings_Global
%>
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
    
    <!-- Custom script for Delivery Board -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/highlight.js/8.3/highlight.min.js"></script>


    <!-- HTML5 shim and Respond.js for IE8 support of HTML5 elements and media queries -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/html5shiv/3.7.2/html5shiv.min.js"></script>
      <script src="https://oss.maxcdn.com/respond/1.4.2/respond.min.js"></script>
    <![endif]-->
    
    <!-- icons and notification styles !-->
     <link href="<%= BaseURL %>css/font-awesome/css/font-awesome.min.css" rel="stylesheet">
     <link href="<%= BaseURL %>css/notifications.css" rel="stylesheet">
    <!-- eof icons and notification styles !-->
    
    <!-- fonts !-->
    <link href='http://fonts.googleapis.com/css?family=Coda' rel='stylesheet' type='text/css'>
    <link href='http://fonts.googleapis.com/css?family=Oswald:400,300,700' rel='stylesheet' type='text/css'>
	<link href='http://fonts.googleapis.com/css?family=Indie+Flower' rel='stylesheet' type='text/css'>
    
    <!-- eof fonts !-->
	

	<!-- *********************************************************************** -->
	<!-- IMPORTANT - USE OLDER VERSION OF JQUERY FOR SORTABLE PLUGIN             -->
	<!-- *********************************************************************** -->
  	<script src="http://code.jquery.com/jquery-1.11.2.min.js"></script>
	<!--<script src="https://code.jquery.com/jquery-3.1.1.js" integrity="sha256-16cdPddA6VdVInumRGo6IbivbERE8p7CQR3HzTBuELA=" crossorigin="anonymous"></script>  -->	
	<!-- *********************************************************************** -->
	
	
	<!-- Including jQuery UI CSS & jQuery Dialog UI Here-->
	<link rel="stylesheet" href="http://code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<script src="http://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
	<!-- End Including jQuery UI CSS & jQuery Dialog UI Here-->
	
	
	<!-- Bootstrap core JS - must load after jQuery -->
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"></script>
    <!-- End Bootstrap core JS - must load after jQuery -->
    

	<% 'Read delivery board settings
	
	SQL = "SELECT * FROM Settings_Global"
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	If not rs.EOF Then
			DelBoardNextStopColor = rs("DelBoardNextStopColor")
			DelBoardScheduledColor = rs("DelBoardScheduledColor")	
			DelBoardCompletedColor = rs("DelBoardCompletedColor")	
			DelBoardInProgressColor = rs("DelBoardInProgressColor")			
			DelBoardSkippedColor = rs("DelBoardSkippedColor")		
			DelBoardProfitDollars = rs("DelBoardProfitDollars")			
			DelBoardAtOrAboveProfitColor = rs("DelBoardAtOrAboveProfitColor")			
			DelBoardBelowProfitColor = rs("DelBoardBelowProfitColor")			
			DelBoardUserAlertColor = rs("DelBoardUserAlertColor")
			DelBoardAMColor = rs("DelBoardAMColor")
			DelBoardPriorityColor = rs("DelBoardPriorityColor")
			DelBoardRoutesToIgnore = rs("DelBoardRoutesToIgnore")
			DelBoardTitleText = rs("DelBoardTitleText")
			DelBoardTitleTextFontColor = rs("DelBoardTitleTextFontColor")
			DelBoardTitleGradientColor = rs("DelBoardTitleGradientColor")
			DelBoardPieTimerColor = rs("DelBoardPieTimerColor")	
	End If
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing

	If DelBoardNextStopColor = "" Then DelBoardNextStopColor = "#FFA500"
	If IsNull(DelBoardNextStopColor) Then DelBoardNextStopColor = "#FFA500"
	If DelBoardScheduledColor = "" Then DelBoardScheduledColor = "#F6F6F6"
	If IsNull(DelBoardScheduledColor) Then DelBoardScheduledColor = "#F6F6F6"
	If DelBoardCompletedColor = "" Then DelBoardCompletedColor = "#D8F9D1"
	If IsNull(DelBoardCompletedColor) Then DelBoardCompletedColor = "#D8F9D1"
	If DelBoardSkippedColor = "" Then DelBoardSkippedColor = "#FCB3B3"
	If IsNull(DelBoardSkippedColor) Then DelBoardSkippedColor = "#FCB3B3"
	If DelBoardAtOrAboveProfitColor = "" Then DelBoardAtOrAboveProfitColor = "#D8F9D1"
	If IsNull(DelBoardAtOrAboveProfitColor) Then DelBoardAtOrAboveProfitColor = "#D8F9D1"
	If DelBoardBelowProfitColor = "" Then DelBoardBelowProfitColor = "#FCB3B3"
	If IsNull(DelBoardBelowProfitColor) Then DelBoardBelowProfitColor = "#FCB3B3"
	If DelBoardUserAlertColor = "" Then DelBoardUserAlertColor = "#FFA500"
	If IsNull(DelBoardUserAlertColor) Then DelBoardUserAlertColor = "#FFA500"
	If DelBoardAMColor = "" Then DelBoardAMColor = "#000000"
	If IsNull(DelBoardAMColor) Then DelBoardAMColor = "#000000"
	If DelBoardPriorityColor = "" Then DelBoardPriorityColor = "#000000"
	If IsNull(DelBoardPriorityColor) Then DelBoardPriorityColor = "#000000"
			
	If DelBoardRoutesToIgnore = "" OR IsNUll(DelBoardRoutesToIgnore) Then 
		DelBoardRoutesToIgnoreCount = 0
	Else
		DelBoardRoutesToIgnoreArray = Split(DelBoardRoutesToIgnore,",")
		DelBoardRoutesToIgnoreCount = UBound(DelBoardRoutesToIgnoreArray) + 1
	End If

	If DelBoardTitleTextFontColor = "" Then DelBoardTitleTextFontColor = "#000000"
	If IsNull(DelBoardTitleTextFontColor) Then DelBoardTitleTextFontColor = "#000000"
	If DelBoardTitleText = "" Then DelBoardTitleText = "Delivery Status"
	If IsNull(DelBoardTitleText ) Then DelBoardTitleText = "Delivery Status"
	DelBoardTitleText = Replace(DelBoardTitleText,"'","")
	DelBoardTitleText = Replace(DelBoardTitleText,"~today~",FormatDateTime(Now(),2))
	DelBoardTitleText = Replace(DelBoardTitleText,"~dow~",WeekDayName(Datepart("w",Now())))
	If DelBoardTitleGradientColor = "" Then DelBoardTitleGradientColor = "#80B8FF"
	If IsNull(DelBoardTitleGradientColor) Then DelBoardTitleGradientColor = "#80B8FF"
		
	If DelBoardPieTimerColor = "" Then DelBoardPieTimerColor = "000000"
	If IsNull(DelBoardPieTimerColor ) Then DelBoardPieTimerColor = "000000"
	Session("DelBoardPieTimerColor") = Replace(DelBoardPieTimerColor,"#","") ' Just this one for Javascript

	%>
	<script src="<%= BaseURL %>js/countdown/jquery.pietimer.js"></script>
	
    <script>
	   $(document).ready(function(){	

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
		    

    

			$('.btn-toggle').click(function() {
			
				  driverUserNo = $(this).attr("id");
				
				  if ($(this).find('.btn-nag-on').size() > 0) {
				  				    
					    $("#" + driverUserNo + "ON").removeClass('btn-nag-on');
					    $("#" + driverUserNo + "ON").addClass('btn-default');
					    $("#" + driverUserNo + "ON").addClass('active');
					    
					    $("#" + driverUserNo + "OFF").removeClass('btn-default');
					    $("#" + driverUserNo + "OFF").removeClass('active');
					    $("#" + driverUserNo + "OFF").addClass('btn-nag-off');
					    				    
					    $("#" + driverUserNo + "ON").html("ON");
					    $("#" + driverUserNo + "OFF").html("NAG OFF");	
					    
				    	$.ajax({
							type:"POST",
							url: "../../../inc/InSightFuncs_AjaxForRoutingModals.asp",
							cache: false,
							data: "action=TurnOnNagAlertsForDeliveryBoardDriverKiosk&driverUserNo=" + encodeURIComponent(driverUserNo),
							success: function(response)
							 {
				             }
						});
				  }
				  else {
	
					    $("#" + driverUserNo + "OFF").removeClass('btn-nag-off');
					    $("#" + driverUserNo + "OFF").addClass('btn-default');
					    $("#" + driverUserNo + "OFF").addClass('active');
					    
					    $("#" + driverUserNo + "ON").removeClass('btn-default');
					    $("#" + driverUserNo + "ON").removeClass('active');
					    $("#" + driverUserNo + "ON").addClass('btn-nag-on');
					    
					    $("#" + driverUserNo + "ON").html("NAG ON");
					    $("#" + driverUserNo + "OFF").html("OFF");			    
					    
				    	$.ajax({
							type:"POST",
							url: "../../../inc/InSightFuncs_AjaxForRoutingModals.asp",
							cache: false,
							data: "action=TurnOffNagAlertsForDeliveryBoardDriverKiosk&driverUserNo=" + encodeURIComponent(driverUserNo),
							success: function(response)
							 {
				             }
						});
				  }
				  
			});	
			
											
			$('#deliveryBoardInvoiceOptionsModal').on('show.bs.modal', function(e) {
			
			    //get data-id attribute of the clicked prospect
			    var myInvoiceNumber = $(e.relatedTarget).data('invoice-number');
			    var myCustomerName = $(e.relatedTarget).data('customer-name');	
			    var myCustID = $(e.relatedTarget).data('customer-id');
			    var myTruckNumber = $(e.relatedTarget).data('truck-number');
			    
			    //populate the textbox with the id of the clicked prospect
			    $(e.currentTarget).find('input[name="txtInvoiceNumber"]').val(myInvoiceNumber);
			    $(e.currentTarget).find('input[name="txtTruckNumber"]').val(myTruckNumber);
			    $(e.currentTarget).find('input[name="txtCustID"]').val(myCustID);
			    	    
			    var $modal = $(this);
		
	    		$modal.find('#deliveryBoardLabel').html("Delivery Options For " + myCustomerName + " - Invoice  #" + myInvoiceNumber);
	    		
		    	$.ajax({
					type:"POST",
					url: "../../../inc/InSightFuncs_AjaxForRoutingModals.asp",
					cache: false,
					data: "action=GetContentForDeliveryBoardOptionsModal&returnPage=directLaunch/kiosks/routing/DeliveryBoardKioskNoPaging.asp&invoiceNum=" + encodeURIComponent(myInvoiceNumber) + "&custID=" + encodeURIComponent(myCustID) + "&truckNum=" + encodeURIComponent(myTruckNumber),
					success: function(response)
					 {
		             	$modal.find('#deliveryBoardInvoiceOptionsModalContent').html(response);
		             },
		             failure: function(response)
					 {
					 	$modal.find('#deliveryBoardInvoiceOptionsModalContent').html("Failed");
		             }
				});
			    
			});
			
	
				
			$('#myDeliveryBoardCompletedOrSkippedModal').on('show.bs.modal', function(e) {
			
			    //get data-id attribute of the clicked prospect
			    var myInvoiceNumber = $(e.relatedTarget).data('invoice-number');
			    var myCustomerName = $(e.relatedTarget).data('customer-name');	
			    //populate the textbox with the id of the clicked prospect
			    $(e.currentTarget).find('input[name="txtInvoiceNumber"]').val(myInvoiceNumber);
			    	    
			    var $modal = $(this);
		
	    		$modal.find('#myDeliveryBoardCompletedOrSkippedLabel').html("Delivery Information for " + myCustomerName + " - Invoice #" + myInvoiceNumber);
	    		
		    	$.ajax({
					type:"POST",
					url: "../../../inc/InSightFuncs_AjaxForRoutingModals.asp",
					cache: false,
					data: "action=GetContentForCompletedOrSkippedInfoModal&myInvoiceNumber=" + encodeURIComponent(myInvoiceNumber),
					success: function(response)
					 {
		               	 $modal.find('#deliveryBoardCompletedOrSkippedModalContent').html(response);
		             },
		             failure: function(response)
					 {
					   $modal.find('#deliveryBoardCompletedOrSkippedModalContent').html("Failed");
		             }
				});
			    
			});
		
		var rgbcolor = '<%=Session("DelBoardPieTimerColor")%>';

		var pagetimer = new Timer(function() {
		   
		 if ('<%=MUV_Read("serviceModule")%>' == 'Enabled') {
		   window.location = "../../kiosks/service/fieldservicekiosknopaging.asp?pp=<%=Session("PassPhrase")%>&cl=<%=Session("ClientKey")%>&ri=<%=Session("RefreshInterval")%>";
		   }
		 if ('<%=MUV_Read("serviceModule")%>' != 'Enabled') {
		   location.reload();
		   }
		   
		},  <%=Session("RefreshInterval")%>*1000);
				
		$('#timer').pietimer({
			seconds: <%=Session("RefreshInterval")%>,
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
	
	
	
	      
</script>

	     
<!-- delivery board custom styles !-->
<style type="text/css">

	body{
		overflow-x:hidden;
		margin:10px;
		border: 2px solid <%=DelBoardTitleGradientColor%>;
		border-radius:5px;
		padding: 0px;
	}
	

	.wrapper{
		margin:0px;
		padding:0px;
	}
	
	.heading-legend{
		margin-top:15px;
		border-bottom:1px solid #eee;
		margin-bottom:20px;
		
	}
	
	.heading-legend h1{
		margin:0px;
	}
	
	.heading-legend h4{
		font-weight:bold;
		margin:0px;
		padding:0px;
		text-transform:uppercase;
		text-align:center;
	}
	   
   .navbar-inverse{
	  border: 0px;
	  border-top-left-radius: 5px;
	  border-top-right-radius: 5px;
	  border-bottom-left-radius: 0px;
	  border-bottom-right-radius: 0px;

   }
  
	.navbar {
	    position: relative;
	    margin-bottom: 0px;
	    border: 1px solid transparent;
	}
 
   .delivery-status{
	   margin-top: 0px;
	   color: #fff;
    }
   
   .navbar-logo{
	   /*position: absolute;
	   margin-top: 45px;*/
 	   max-width: 200px;
	   height: auto;
	   right:0;
    }
    
   .navbar-time{
		position: absolute;
		margin-top: 10px;
		margin-left: 10px;
		left: 0;
  		height: auto;
		font-size: 5em !important;
		text-align: center;
		font-family: 'Oswald';
		font-weight: 300;
		background-color: #484848;
		color: #fff;
		border: 1px solid #fff;
		padding-left:10px;
		padding-right:10px;
		padding-bottom:5px;
    }
    
  
   .delivery-status h2{
	   margin:8px 0px 0px 0px;
	   line-height:1;
   }
	
	.legend-title{
		margin: 0px;
		padding: 0px;
	}
	
	.legend-row{
		margin-bottom: 10px;
		margin-left: 0px;
		margin-right: 0px;
	 }

	.legend-box{
	 	padding-top: 10px;
		margin-bottom: 15px;
	}
	   
	.legend-complete{
		<% Response.Write("background:" & DelBoardCompletedColor  & ";") %>
	   padding:5px 5px 5px 5px;
	   color:#000;
	   text-align:center;
	}
	
	.legend-inprogress{
		<% Response.Write("background:" & DelBoardInProgressColor  & ";") %>
	   padding:5px 5px 5px 5px;
	   color:#000;
	   text-align:center;
	}

	.legend-nodelivery{
	   <% Response.Write("background:" & DelBoardSkippedColor  & ";")%>
	   padding:5px 5px 5px 5px;
	   color:#000;
	   text-align:center;
	}
	
	.legend-nextstop{
	   <% Response.Write("background:" & DelBoardNextStopColor & ";")%>
	   padding:5px 5px 5px 5px;
	   color:#000;
	   text-align:center;
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
	 	
	table thead a{
		color: #000;
	}

	.tr-inprogress{
		<% Response.Write("background:" & DelBoardInProgressColor & " !important;") %>
	}
	
	.tr-completed{
		<% Response.Write("background:" & DelBoardCompletedColor & " !important;") %>
	}
	
	.tr-nodelivery{
		<% Response.Write("background:" & DelBoardSkippedColor & " !important;") %>
	}
	
	.tr-nextstop{
		<% Response.Write("background:" & DelBoardNextStopColor & " !important;") %>
	}
		
	.tr-scheduled{
		<% Response.Write("background:" & DelBoardScheduledColor & ";") %>
		<% Response.Write("border-top: 1px solid #000000;") %>
		<% Response.Write("border-bottom: 1px solid #000000;") %>
		<% Response.Write("border-left: 1px solid #000000;") %>
		<% Response.Write("border-right: 1px solid #000000;") %>

	}
		
	.AM-border{
		<% Response.Write("border-top: 3px solid " & DelBoardAMColor & ";") %>
		<% Response.Write("border-bottom: 3px solid " & DelBoardAMColor & ";") %>
		<% Response.Write("border-left: 3px solid " & DelBoardAMColor & ";") %>
		<% Response.Write("border-right: 3px solid " & DelBoardAMColor & ";") %>
	}

		
	.Priority-border{
		<% Response.Write("border-top: 3px solid " & DelBoardPriorityColor & ";") %>
		<% Response.Write("border-bottom: 3px solid " & DelBoardPriorityColor & ";") %>
		<% Response.Write("border-left: 3px solid " & DelBoardPriorityColor & ";") %>
		<% Response.Write("border-right: 3px solid " & DelBoardPriorityColor & ";") %>
	}
	
	.tr-user-alert{
		<% Response.Write("background:" & DelBoardUserAlertColor & ";") %>
	}
	
	.tr-user-alert-top{
		<% Response.Write("border-top: 1px solid #000000;") %>
		<% Response.Write("border-left: 1px solid #000000;") %>
		<% Response.Write("border-right: 1px solid #000000;") %>
	}
	
	.tr-user-alert-bottom{
		<% Response.Write("border-bottom: 1px solid #000000;") %>
		<% Response.Write("border-left: 1px solid #000000;") %>
		<% Response.Write("border-right: 1px solid #000000;") %>
	}
 
	 .tr-border-line{
		 border-bottom: 1px solid #ccc;
	 }
	 
	 .tr-border-line-red{
		 border-bottom: 1px solid #999;
	 }
	
	.row{
		font-size:12px;
	}
			
	.table-condensed>tbody>tr>td, .table-condensed>tbody>tr>th, .table-condensed>tfoot>tr>td, .table-condensed>tfoot>tr>th, .table-condensed>thead>tr>td, .table-condensed>thead>tr>th{
		padding: 2px;
	}
	 
	.scrollable-table{
	 	overflow: hidden;
		border: 1px solid #ccc;
	 	font-size: 9px;
	 	border-bottom-left-radius: 5px;
	 	border-bottom-right-radius: 5px;
	 	
	}

	.table>tbody>tr>td, .table>tbody>tr>th, .table>tfoot>tr>td, .table>tfoot>tr>th, .table>thead>tr>td, .table>thead>tr>th{
	 	padding-top:5px;
	 	padding-bottom:5px;
	 	
	}

	.scrollable-title{
		border: 1px solid #ccc;
		padding: 10px;
		margin-bottom: -1px;
		background: #DCE6E9;
		font-size: 12px;
		border-top-left-radius: 5px;
		border-top-right-radius: 5px;
	}
	
	.scrollable-title strong{
		width:100%;
		display:block;
		white-space:normal;
	}
	 
	  
	[class^="col-"]{
		padding:2px;
	}
	   
	.col-lg-cust{
	   width:7%;
	   display:inline-block;
	   vertical-align:top;
	}
	   
	
	.table>tbody>tr>td, .table>tbody>tr>th, .table>tfoot>tr>td, .table>tfoot>tr>th, .table>thead>tr>td, .table>thead>tr>th{
		border:0px;
	}
	
	#sortableList .item{
		padding:2px;
	}
	
	.item-box{
		margin:2px;
		float:left;
		width:135px;
	}
	
	   
	.ui-state-highlight.item{height: 100px;}
	 
	.list-boxes{
		/*margin-left: 230px;*/
	}
	 
	.horizontal-layout{
		/*float:left;*/
		width:100%;
		margin-top:20px;
		padding-bottom:40px;
		margin-left:-60px;
	}
	
	.double-scroll{
		width:100%;
	}
	
	.trucknumber{
		width: 100%;
	    display: block;
	    white-space: normal;
	    height:25px;
	    font-weight:bold;
	}  
	
	
	.drivername{
		width: 100%;
	    display: block;
	    white-space: normal;
	    text-transform:uppercase;
	    line-height:12px;
	    height:30px;
	    font-weight:bold;
	    color:#00008B;
	}
		
	
	.fa-star {
		color:blue;
		float:right;
	}

	 .nag-on-off span{
	 	font-size: 10px;
	 	font-weight: normal;
	 	background: #459e44;
	 	color: #fff;
	 	display: inline-block;
	 	padding: 3px;
	 	border-radius: 2px;
 	 	position: absolute;
	 	top:7px;
	 	right:7px;
 	 }
	
	.btn-nag-on {
	  background-color:#449d44;
	  color:#FFF;
	}
	.btn-nag-off {
	  background-color:#ac2925;
	  color:#FFF;
	}
	
	.btn-nag-on:not(.active){
	  background-color:#449d44;
	  color:#FFF;
	}
	.btn-nag-off:not(.active){
	  background-color:#ac2925;
	  color:#FFF;
	}	


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

	.pause{
	   float:right;
	   /*margin:10px 30px 0px 0px;*/
	   margin:10px 30px 20px 10px;
	   color:#337ab7;
	   font-size:14px;
	}
	
	 .the-timer{
		float:right;
		margin-left:10px;
		margin-top: 5px;
	 }
</style>
<!-- end delivery board custom styles !-->
	
	
<script type="text/javascript" src="<%= BaseURL %>js/doublescroll/jquery.doubleScroll.js"></script>


<script type="text/javascript">
    $(document).ready(function(){

       $('.double-scroll').doubleScroll({
       		resetOnWindowResize: true
       	});
		
 		
    });
</script>
		
<%
Response.Write("<style type='text/css'>")
	Response.Write(".navbar-inverse{")
	Response.Write("background: " & DelBoardTitleGradientColor &"; /* For browsers that do not support gradients */" & "<br>")
    Response.Write("background: -webkit-linear-gradient(" & DelBoardTitleGradientColor & ", #fff); /* For Safari 5.1 to 6.0 */" & "<br>")
    Response.Write("background: -o-linear-gradient(" & DelBoardTitleGradientColor & ", #fff); /* For Opera 11.1 to 12.0 */" & "<br>")
    Response.Write("background: -moz-linear-gradient(" & DelBoardTitleGradientColor & ", #fff); /* For Firefox 3.6 to 15 */" & "<br>")
    Response.Write("background: linear-gradient(" & DelBoardTitleGradientColor & ", #fff); /* Standard syntax (must be last) */" & "<br>")
Response.Write("}")
Response.Write("</style>")
%> 

  </head>

<body>

 <!-- header !-->
<nav class="navbar navbar-inverse">
      <div class="container-fluid">
        <div class="navbar-header">
          
          <!-- row !-->
          <div class="row ">
	          
                     
          <!-- legend !-->
          <div class="col-lg-12 delivery-status">
          
          
          	<div class="col-lg-2" style="width: 240px;">
	      		<div class="row navbar-time" id="clocktext"></div>
	      	</div>
 
			<script type="text/javascript">
				"use strict";
				
				var textElem = document.getElementById("clocktext");
				var textNode = document.createTextNode("");
				textElem.appendChild(textNode);
				var targetWidth = 0.9;  // Proportion of full screen width
				var curFontSize = 20;  // Do not change
				
				function updateClock() {
					var d = new Date();
					var s = "";
					s += ((d.getHours() + 11) % 12 + 1) + ":";
					s += (10 > d.getMinutes() ? "0" : "") + d.getMinutes() + "\u00A0";
					s += d.getHours() >= 12 ? "pm" : "am";
					textNode.data = s;
					setTimeout(updateClock, 60000 - d.getTime() % 60000 + 100);
				}
								
				updateClock();

			</script>
            
            <div class="col-lg-10" style="text-align:center">    
             
		  		<h2>
		  			<font color="<%=DelBoardTitleTextFontColor%>"><%=DelBoardTitleText%></font>
		  			<img src="<%= BaseURL %>clientfiles/<%= MUV_Read("ClientID") %>/logos/logo.png" class="navbar-logo">
			  		<div class="pause">
						Pause Automatic Refresh&nbsp;&nbsp;
						<div class="material-switch pull-right">
							<input id="switchAutomaticRefresh" name="chkAutomaticRefresh" type="checkbox"/>
							<label for="switchAutomaticRefresh" class="label-primary"></label>
						</div>
					</div>
					<div id="timer" class="the-timer" style="height:30px;"></div>
		  		</h2>
		  		
		  		<div style="padding-top:20px;">
		  		
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
			    	
			        <!-- in progress !-->
			        <div class="col-lg-2">
			        	<div class="legend-inprogress">
			            	<h4>In Progress</h4>
			            </div>
			        </div>
			        <!-- eof in progress stop !-->
			        
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


			        <!-- in progress !-->
			        <div class="col-lg-2" style="width:19%">
			        </div>
		        	<!-- eof in progress stop !-->
			        
			        
				</div>

			</div>
            

		 </div>

        </div>
 
 
          </div>
          <!-- eof welcome !-->
          
          </div>
          <!-- eof row !-->
          
        </div>
         
      </div>
    </nav>
<!-- eof header !-->    



       
 <!-- eof side bar !-->

<!-- content area !-->
<div class="wrapper">
<div class="horizontal-layout">
	<div class="double-scroll">
		<table>
			<tr>
		    	<td>
		 			<div class='list-boxes' id='sortableList'>
					<%
					
					Set cnn_DeliveryBoardSum = Server.CreateObject("ADODB.Connection")
					cnn_DeliveryBoardSum.open (Session("ClientCnnString"))
					Set rs_DeliveryBoardSum = Server.CreateObject("ADODB.Recordset")
					rs_DeliveryBoardSum.CursorLocation = 3 
					Set rs_DeliveryBoardSum = cnn_DeliveryBoardSum.Execute(SQL)
					
					SQL_DeliveryBoardSum = "SELECT DISTINCT TruckNumber FROM RT_DeliveryBoard ORDER BY TruckNumber"
					Set rs_DeliveryBoardSum = cnn_DeliveryBoardSum.Execute(SQL_DeliveryBoardSum)
					
					If not rs_DeliveryBoardSum.EOF Then
						Do While Not rs_DeliveryBoardSum.Eof
						
							TruckNumber = rs_DeliveryBoardSum("TruckNumber")
						
							If DelBoardIgnoreThisRoute(TruckNumber) <> True Then 
							
								DriverUserNo = Trim(GetUserNumberByTruckNumber(TruckNumber))
							
								If userIsArchived(DriverUserNo) = False AND userIsEnabled(DriverUserNo) = True Then
									Call TruckNumberWrite(TruckNumber, GridColumn) 
								End If
								
							End If
								
							rs_DeliveryBoardSum.Movenext
						Loop
					End If
					%> 
				</div>
			</tr>
		</table>
	</div>
</div>

<%
Sub TruckNumberWrite(TruckNumber, GridColumn) %>
		<td class='item col-lg-cust' TruckNumber='<%= TruckNumber %>'>
		<div class='item-box'>
		<div class='scrollable-title' style='position: relative;'>
		
			<span class="trucknumber">Route:&nbsp;<%= TruckNumber %></span>
		
			<span class="drivername"><%= GetUserDisplayNameByUserNo(Trim(GetUserNumberByTruckNumber(TruckNumber))) %></span>
			

			<%
			'**************************************************************************
			'show nag alerts to admins or route managers only
			'**************************************************************************

			DriverUserNo = Trim(GetUserNumberByTruckNumber(TruckNumber))

			If DriverUserNo <> "*Not Found*" Then
			
				'First check to see if nags are off entirely for this user
				NagsON = False
				
				SQLUsers = "SELECT * FROM tblUsers Where UserNo = " & DriverUserNo 
				
				Set cnn_Users = Server.CreateObject("ADODB.Connection")
				cnn_Users.open (Session("ClientCnnString"))
				Set rsUsers = Server.CreateObject("ADODB.Recordset")
				rsUsers.CursorLocation = 3 
				'Response.write(SQLUsers)
				Set rsUsers = cnn_Users.Execute(SQLUsers)

				'ANY YES CONDITION TURNS THE BUTTON ON
				If Not rsUsers.EOF Then

					If rsUsers("userNextStopNagMessageOverride") = "Yes" Then NagsON = True
					If rsUsers("userNoActivityNagMessageOverride") = "Yes" Then NagsON = True
					
					If NagsON = False Then' only check if not already on
					
						If rsUsers("userNextStopNagMessageOverride") = "Use Global" or rsUsers("userNoActivityNagMessageOverride") = "Use Global" Then
						
							SQLGlobal = "SELECT * FROM Settings_Global "
							Set rsGlobal = Server.CreateObject("ADODB.Recordset")
							rsGlobal.CursorLocation = 3 
							Set rsGlobal = cnn_Users.Execute(SQLGlobal)
	
							If Not rsGlobal.EOF Then
								NoAct = rsGlobal("NoActivityNagMessageONOFF")
								NextSt = rsGlobal("NextStopNagMessageONOFF")
							End If
						
							Set rsGlobal = Nothing
						End If
						
						If NextSt  = 1 Then NagsON = True
						If NoAct = 1 Then NagsON = True
						
					End If
				
				End If
				
				Set rsUsers = Nothing
				cnn_Users.Close
				Set cnn_Users = Nothing
			End If

			
			If DriverUserNo <> "*Not Found*" Then
			
				If NagsOn = True Then
			
					If  DriverInNagSkipTable(DriverUserNo,"routingNoNextStop") = False AND DriverInNagSkipTable(DriverUserNo,"routingNoActivity") = False Then

						buttonClassGreen = "btn btn-xs btn-nag-on" 
						buttonClassRed= "btn btn-xs btn-default active"
					Else 
						buttonClassGreen = "btn btn-xs btn-default active" 
						buttonClassRed = "btn btn-xs btn-nag-off"
					End If
					%>	  
					<% If buttonClassGreen = "btn btn-xs btn-nag-on" Then %>
						  <div class="btn-group btn-toggle" id="<%= DriverUserNo %>">
						    <button class="<%= buttonClassGreen %>" id="<%= DriverUserNo %>ON">NAG ON</button>
						    <button class="<%= buttonClassRed %>" id="<%= DriverUserNo %>OFF">OFF</button>
						  </div>
					<% Else %>
						  <div class="btn-group btn-toggle" id="<%= DriverUserNo %>">
						    <button class="<%= buttonClassGreen %>" id="<%= DriverUserNo %>ON">ON</button>
						    <button class="<%= buttonClassRed %>" id="<%= DriverUserNo %>OFF">NAG OFF</button>
						  </div>						
					<% End If 
				
				Else
					%><div class="btn-group btn-toggle">Nags Off</div><%
				End If

			Else
				%><div class="btn-group btn-toggle">No User Setup</div><%
			End If

			'**************************************************************************
		
			%> 
			
			
			<!--<div style="color:#009900;"><i class="fa fa-play" aria-hidden="true"></i>&nbsp;Route Started</div>-->
		</div>

	        <div class='table-responsive scrollable-table'>
		        <% Response.Write("<table id='truck" & TruckNumber & "' name='truck" & TruckNumber & "' class='food_planner table table-condensed clickable'>") %>
					<!--<thead>
			        	<tr>
			        		<th class='sorttable_nosort'>Invoice</th>
			        		<th class='sorttable_nosort'><%=GetTerm("Customer")%></th>.
			        	</tr>
			        </thead>-->
			        <tbody class='searchable'>
			        	<%'Get all the tickets for this truck
						Set cnn_Tickets = Server.CreateObject("ADODB.Connection")
						cnn_Tickets.open (Session("ClientCnnString"))
						Set rs_DeliveryBoardDet = Server.CreateObject("ADODB.Recordset")
						rs_DeliveryBoardDet.CursorLocation = 3 
						
						SQL_Tickets = "SELECT * FROM RT_DeliveryBoard "
						SQL_Tickets = SQL_Tickets & "WHERE TruckNumber = '" & TruckNumber  & "' "
						
						If DelBoardDontUseStopSequencing() = False Then
	                        SQL_Tickets = SQL_Tickets & "Order By SequenceNumber, CustNum" 
	                    Else
   	                        SQL_Tickets = SQL_Tickets & "Order By CustNum" 
	                    End If

                        Set rs_DeliveryBoardDet = cnn_Tickets.Execute(SQL_Tickets)
                        
						If not rs_DeliveryBoardDet.Eof Then

							NumLines = 0
							Do While not rs_DeliveryBoardDet.Eof
							
								trclass = ""
								
								PriorityDelivery = rs_DeliveryBoardDet("Priority")
								InvoiceNumber = rs_DeliveryBoardDet("IvsNum")
								CustName = rs_DeliveryBoardDet("CustName")
								CustID = rs_DeliveryBoardDet("CustNum")
								TruckNumber = rs_DeliveryBoardSum("TruckNumber")
								AMorPM = rs_DeliveryBoardDet("AMorPM")
								DeliveryStatus = rs_DeliveryBoardDet("DeliveryStatus")
								
							
								If CustID = GetNextCustomerStopByTruck(TruckNumber) Then
									
									If AMorPM = "AM" Then
																			
										If rs_DeliveryBoardDet("DeliveryInProgress") = 1 Then
											%>
											<tr class="tr-inprogress AM-border" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;">
											<%
										Else
											%>
											<tr class="tr-nextstop AM-border" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;">
											<%	
										End If
										
										
										If len(CustName) > 19 then
											Response.Write("<td colspan='2'>" & left(CustName,19)) 
										Else
											Response.Write("<td colspan='2'>" & CustName) 
										End If
										
										Response.Write("</td></tr>")
										
									Else
									
										If rs_DeliveryBoardDet("DeliveryInProgress") = 1 Then
											
											If PriorityDelivery = 1 Then %>
												<tr class="tr-inprogress tr-scheduled Priority-border" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;">											
											<% Else %>
												<tr class="tr-inprogress tr-scheduled" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;">
											<% End If %>
											
										<%	
										Else											
											
											If PriorityDelivery = 1 Then %>
												<tr class="tr-nextstop tr-scheduled Priority-border" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;">											
											<% Else %>
												<tr class="tr-nextstop tr-scheduled" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;">
											<% End If %>
											
										<%	
										End If
										
										If len(CustName) > 19 then
											Response.Write("<td colspan='2'>" & left(CustName,19))
										Else
											Response.Write("<td colspan='2'>" & CustName)
										End If
										
										Response.Write("</td></tr>")
										
									End If
									
								Else
																	
									If rs_DeliveryBoardDet("DeliveryStatus") = "Delivered" Then
									
										If AMorPM = "AM" Then
											
											%><tr class="tr-completed AM-border" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#myDeliveryBoardCompletedOrSkippedModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;"><%
											
										Else
											
											If PriorityDelivery = 1 Then %>
												<tr class="tr-completed tr-scheduled Priority-border" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#myDeliveryBoardCompletedOrSkippedModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;">											
											<% Else %>
												<tr class="tr-completed tr-scheduled" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#myDeliveryBoardCompletedOrSkippedModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;">
											<% End If %>
											
										<%	
										End If
										
									ElseIf rs_DeliveryBoardDet("DeliveryStatus") = "No Delivery" Then
									
										If AMorPM = "AM" Then
											
											%><tr class="tr-nodelivery AM-border" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;"><%
											
										Else
											
											If PriorityDelivery = 1 Then %>
												<tr class="tr-nodelivery tr-scheduled Priority-border" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;">											
											<% Else %>
												<tr class="tr-nodelivery tr-scheduled" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;">
											<% End If %>
										<%											
											
										End If
										
									Else
									
										If AMorPM = "AM" Then
										
											If rs_DeliveryBoardDet("DeliveryInProgress") = 1 Then
												
												%><tr class="tr-inprogress AM-border" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;"><%
																					
											Else	
												%><tr class="tr-scheduled AM-border" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;"><%										
												
											End If
											
										Else
										
											If rs_DeliveryBoardDet("DeliveryInProgress") = 1 Then
																									
												If PriorityDelivery = 1 Then %>
													<tr class="tr-inprogress tr-scheduled Priority-border" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;">											
												<% Else %>
													<tr class="tr-inprogress tr-scheduled" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;">
												<% End If %>
												
											<%
											Else
																								
												If PriorityDelivery = 1 Then %>
													<tr class="tr-scheduled Priority-border" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;">											
												<% Else %>
													<tr class="tr-scheduled" data-toggle="modal" data-show="true" href="#" data-truck-number="<%= TruckNumber %>" data-customer-id="<%= CustID %>" data-invoice-number="<%= InvoiceNumber %>" data-customer-name="<%= CustName %>" data-target="#deliveryBoardInvoiceOptionsModal" data-tooltip="true" data-title="Show Delivery Options" style="cursor:pointer;">
												<% End If %>
												
											<%
												
											End If
											
										End If
										
									End If
									
									Response.Write(trclass)
									
									If len(CustName) > 19 then Cnam = left(CustName,19) Else Cnam = CustName
									
									If GetLastInvoiceMarkedByTruckNumber(TruckNumber) = InvoiceNumber Then
										Response.Write("<td colspan='2'>" & Cnam & "<i class='fa fa-star' aria-hidden='true'></i></td>")
									Else
										Response.Write("<td colspan='2'>" & Cnam & "</td>")
									End If
									
									Response.Write("</tr>")
 								End If
			
								
								rs_DeliveryBoardDet.movenext
								NumLines = NumLines + 1
								
							Loop
							
							'Make all boxes even
							If NumLines < MaxNumberOfDeliveries() Then
								For x = 1 to MaxNumberOfDeliveries() - NumLines
									Response.Write("<tr ><td>&nbsp;</td></tr>")
									Response.Write("<tr ><td>&nbsp;</td></tr>")
								Next
							End IF
							
						End IF%>
                        </td>
			        </tbody>
		        </table>
	        </div>
            </div>
        <%Response.Write("</div>")
		GridColumn = GridColumn +1
End Sub 

Set rs_DeliveryBoardSum = Nothing
cnn_DeliveryBoardSum.Close
Set cnn_DeliveryBoardSum = Nothing
%>	


<!-- same height titles !-->
<script type="text/javascript" src="<%= BaseURL %>js/grids.js"></script>

<script type="text/javascript">
	jQuery(function($) {
		$('.scrollable-title').responsiveEqualHeightGrid();	
	});
</script>
<!-- eof same height titles !-->


<!-- tooltip JS !-->
<script type="text/javascript">
	$(function () {
		$('[data-toggle="tooltip"]').tooltip()
	})
 </script>
<!-- eof tooltip JS !-->

  </div>
  <!-- eof content area !-->
  
</div>
<!-- dashboard ends here !-->

 
<script src="https://cdnjs.cloudflare.com/ajax/libs/highlight.js/8.3/highlight.min.js"></script>		

<!-- IE10 viewport hack for Surface/desktop Windows 8 bug -->
<script src="<%= BaseURL %>js/ie10-viewport-bug-workaround.js"></script>

<%'Subs and Funcs here

Sub Read_Settings_Global
	'Read delivery board settings
	SQL = "SELECT * FROM Settings_Global"
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	If not rs.EOF Then
			DelBoardNextStopColor = rs("DelBoardNextStopColor")
			DelBoardScheduledColor = rs("DelBoardScheduledColor")	
			DelBoardCompletedColor = rs("DelBoardCompletedColor")				
			DelBoardSkippedColor = rs("DelBoardSkippedColor")	
			DelBoardAMColor = rs("DelBoardAMColor")
			DelBoardPriorityColor = rs("DelBoardPriorityColor")
			DelBoardTitleText = rs("DelBoardTitleText")	
			DelBoardTitleTextFontColor = rs("DelBoardTitleTextFontColor")
			DelBoardTitleGradientColor = rs("DelBoardTitleGradientColor")
			DelBoardRoutesToIgnore = rs("DelBoardRoutesToIgnore")
			DelBoardUPSRoutes = rs("DelBoardUPSRoutes")
			DelBoardPieTimerColor = rs("DelBoardPieTimerColor")
	End If
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	If DelBoardNextStopColor = "" Then DelBoardNextStopColor = "#FFA500"
	If IsNull(DelBoardNextStopColor) Then DelBoardNextStopColor = "#FFA500"
	If DelBoardScheduledColor = "" Then DelBoardScheduledColor = "#F6F6F6"
	If IsNull(DelBoardScheduledColor) Then DelBoardScheduledColor = "#F6F6F6"
	If DelBoardCompletedColor = "" Then DelBoardCompletedColor = "#D8F9D1"
	If IsNull(DelBoardCompletedColor) Then DelBoardCompletedColor = "#D8F9D1"
	If DelBoardSkippedColor = "" Then DelBoardSkippedColor = "#FCB3B3"
	If IsNull(DelBoardSkippedColor) Then DelBoardSkippedColor = "#FCB3B3"
	If DelBoardAMColor = "" Then DelBoardAMColor = "#000000"
	If IsNull(DelBoardAMColor) Then DelBoardAMColor = "#000000"
	If DelBoardPriorityColor = "" Then DelBoardPriorityColor = "#000000"
	If IsNull(DelBoardPriorityColor) Then DelBoardPriorityColor = "#000000"
	If DelBoardTitleTextFontColor = "" Then DelBoardTitleTextFontColor = "#000000"
	If IsNull(DelBoardTitleTextFontColor) Then DelBoardTitleTextFontColor = "#000000"
	If DelBoardTitleText = "" Then DelBoardTitleText = "Delivery Status"
	If IsNull(DelBoardTitleText ) Then DelBoardTitleText = "Delivery Status"
	DelBoardTitleText = Replace(DelBoardTitleText,"'","")
	DelBoardTitleText = Replace(DelBoardTitleText,"~today~",FormatDateTime(Now(),2))
	DelBoardTitleText = Replace(DelBoardTitleText,"~dow~",WeekDayName(Datepart("w",Now())))
	If DelBoardTitleGradientColor = "" Then DelBoardTitleGradientColor = "#80B8FF"
	If IsNull(DelBoardTitleGradientColor) Then DelBoardTitleGradientColor = "#80B8FF"
	If DelBoardPieTimerColor = "" Then DelBoardPieTimerColor = "000000"
	If IsNull(DelBoardPieTimerColor ) Then DelBoardPieTimerColor = "000000"
	Session("DelBoardPieTimerColor") = Replace(DelBoardPieTimerColor,"#","") ' Just this one for Javascript
End Sub

%>

<!-- **************************************************************************************************************************** -->
<!-- MODALS FOR DELIVERY ALERTS BEGIN HERE !-->
<!-- **************************************************************************************************************************** -->

<!--#include file="../../../routing/deliveryBoardCommonModals.asp"-->

<!-- **************************************************************************************************************************** -->
<!-- MODALS FOR DELIVERY ALERTS END HERE !-->
<!-- **************************************************************************************************************************** -->

    
  </body>
</html>
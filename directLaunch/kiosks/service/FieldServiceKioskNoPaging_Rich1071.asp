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
<!--#include file="../../../inc/InSightFuncs_Service.asp"-->
<!--#include file="../../../inc/InSightFuncs_AR_AP.asp"-->
<%
UsageMessage = "http://www.mdsinsight.com/directLaunch/kiosks/service/FieldServiceKioskNoPaging.asp?pp={Your Passphrase}&cl={Your Client ID}&ri={Refresh Interval In Seconds}&rgn={Region List}"

'These must be declared here

Dim FSBoardKioskGlobalTitleText
Dim FSBoardKioskGlobalTitleTextFontColor
Dim FSBoardKioskGlobalTitleGradientColor
Dim FSBoardKioskGlobalColorPieTimer	
Dim FSBoardKioskGlobalColorUrgent
Dim FSBoardKioskGlobalColorAwaitingDispatch
Dim FSBoardKioskGlobalColorClosed
Dim FSBoardKioskGlobalColorEnRoute
Dim FSBoardKioskGlobalColorOnSite
Dim FSBoardKioskGlobalColorAwaitingAcknowledgement
Dim FSBoardKioskGlobalColorDispatchAcknowledged
Dim FSBoardKioskGlobalColorDispatchDeclined

PassPhrase = Request.QueryString("pp")
ClientKey = Request.QueryString("cl")


RegionList = Request.QueryString("rgn")
CurrentRegionToShow = Request.QueryString("rgnc")
RegionArray = Split(RegionList,",")
Session("RegionsToShow") = RegionList
Session("CurrentRegionToShow") = CurrentRegionToShow

'*****************************************************************************************
'When loading the page, determine if there are any regions to show
'If there are, set the current region to display
'*****************************************************************************************

If Session("RegionsToShow") <> "" Then
	For x = 0 to Ubound(RegionArray)
	
		If Session("CurrentRegionToShow") = "" Then
			Session("CurrentRegionToShow") = RegionArray(x)
			Exit For
		ElseIf cint(Session("CurrentRegionToShow")) = cint(RegionArray(x)) Then
			If x = Ubound(RegionArray) Then
				Session("CurrentRegionToShow") = RegionArray(0)
			Else
				Session("CurrentRegionToShow") = RegionArray(x+1)
			End If
			Exit For
		End If
		
	Next
End If

'*****************************************************************************************


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
	dummy = MUV_Write("routingModule",Recordset.Fields("routingModule"))
	If PassPhrase <>  Recordset.Fields("directLaunchPassphrase") Then
		Response.Write("Access Denied")
		Session.Abandon
		Response.End
	End If
	Recordset.close
	Connection.close
End If

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
    

	<%
	Call CheckTables
	
	 'Read field service settings
	

	
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
	End If
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	CurrentRegionToShowName = GetRegionNameByRegionIntRecID(Session("CurrentRegionToShow"))
		
	FSBoardKioskGlobalTitleText = Replace(FSBoardKioskGlobalTitleText,"'","")
	FSBoardKioskGlobalTitleText = Replace(FSBoardKioskGlobalTitleText,"~today~",FormatDateTime(Now(),2))
	FSBoardKioskGlobalTitleText = Replace(FSBoardKioskGlobalTitleText,"~dow~",WeekDayName(Datepart("w",Now())))
	
	If CurrentRegionToShowName <> "" then
		FSBoardKioskGlobalTitleText = "Region: " & CurrentRegionToShowName
	End If
	
	Session("FSBoardKioskGlobalColorPieTimer") = Replace(FSBoardKioskGlobalColorPieTimer,"#","") ' Just this one for Javascript

	%>

	<!-- pie timer -->
	<script src="<%= BaseURL %>js/countdown/jquery.pietimer.js"></script>
	
    <script>
	   $(document).ready(function(){		
	   
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
    	
										
		
		var rgbcolor = '<%=Session("FSBoardKioskGlobalColorPieTimer")%>';


		var pagetimer = new Timer(function() {
		
		 if ('<%= Session("RegionsToShow") %>' !== '') {
		 	//alert('<%=Session("CurrentRegionToShow")%>');
		   	window.location = "../../kiosks/service/FieldServiceKioskNoPaging.asp?pp=<%=Session("PassPhrase")%>&cl=<%=Session("ClientKey")%>&ri=<%=Session("RefreshInterval")%>&rgn=<%=Session("RegionsToShow")%>&rgnc=<%=Session("CurrentRegionToShow")%>";
		   }
		 else {
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

			
		$('#serviceBoardTicketOptionsModal').on('show.bs.modal', function(e) {
		
		    //get data-id attribute of the clicked prospect
		    var myTicketNumber = $(e.relatedTarget).data('invoice-number');
		    var myCustomerName = $(e.relatedTarget).data('customer-name');	
		    var myCustID = $(e.relatedTarget).data('customer-id');
		    var myUserNo = $(e.relatedTarget).data('user-no');
		    
		    //populate the textbox with the id of the clicked prospect
		    $(e.currentTarget).find('input[name="txtTicketNumber"]').val(myTicketNumber);
		    $(e.currentTarget).find('input[name="txtCustID"]').val(myCustID);
		    $(e.currentTarget).find('input[name="txtUserNo"]').val(myUserNo);
		    	    
		    var $modal = $(this);
	
    		$modal.find('#serviceBoardLabel').html("Service Options For " + myCustomerName + " - Ticket  #" + myTicketNumber);
    		
	    	$.ajax({
				type:"POST",
				url: "../inc/../../../inc/InSightFuncs_AjaxForServiceModals.asp",
				cache: false,
				data: "action=GetContentForServiceBoardTicketOptionsModal&returnURL=directLaunch/kiosks/service/FieldServiceKioskNoPaging_Rich1071.asp&memoNum=" + encodeURIComponent(myTicketNumber) + "&custID=" + encodeURIComponent(myCustID) + "&userNo=" + encodeURIComponent(myUserNo),
				success: function(response)
				 {
	             	$modal.find('#ServiceBoardTicketOptionsModalContent').html(response);
	             },
	             failure: function(response)
				 {
				 	$modal.find('#ServiceBoardTicketOptionsModalContent').html("Failed");
	             }
			});
		    
		});
		
			
		$('#serviceBoardSetAlertModal').on('show.bs.modal', function(e) {
		
			//close the service ticket options modal where we came from
			$('#serviceBoardTicketOptionsModal').modal('hide');
		    	
		    //get data-id attribute of the clicked prospect
		    var myInvoiceNumber = $(e.relatedTarget).data('invoice-number');
		    var myCustomerName = $(e.relatedTarget).data('customer-name');	
		    //populate the textbox with the id of the clicked prospect
		    $(e.currentTarget).find('input[name="txtInvoiceNumber"]').val(myInvoiceNumber);
		    	    
		    var $modal = $(this);
	
    		$modal.find('#serviceBoardSetAlertModalLabel').html("Delivery Information for " + myCustomerName + " - Invoice #" + myInvoiceNumber);
		});




		$('#serviceBoardXferModal').on('show.bs.modal', function(e) {
		
			//close the service ticket options modal where we came from
			$('#serviceBoardTicketOptionsModal').modal('hide');
		
		    //get data-id attribute of the clicked service ticket
		    var myTicketNumber = $(e.relatedTarget).data('invoice-number');
		    var myCustomerName = $(e.relatedTarget).data('customer-name');	
		    var myCustID = $(e.relatedTarget).data('customer-id');
		    var myUserNo = $(e.relatedTarget).data('user-no');
		    
		    //populate the textboxes with the id of the clicked service ticket
		    $(e.currentTarget).find('input[name="txtTicketNumber"]').val(myTicketNumber);
		    $(e.currentTarget).find('input[name="txtCustID"]').val(myCustID);
		    $(e.currentTarget).find('input[name="txtUserNo"]').val(myUserNo);
		    	    
		    var $modal = $(this);
	    		
	    	$.ajax({
				type:"POST",
				url: "../inc/../../../inc/InSightFuncs_AjaxForServiceModals.asp",
				cache: false,
				data: "action=GetTitleForServiceBoardTransferRedispatchModal&memoNum=" + encodeURIComponent(myTicketNumber) + "&custID=" + encodeURIComponent(myCustID) + "&userNo=" + encodeURIComponent(myUserNo),
				success: function(response)
				 {
	               	 $modal.find('#ServiceBoardXferModalTitle').html(response);
	             },
	             failure: function(response)
				 {
				 	$modal.find('#ServiceBoardXferModalTitle').html("Failed");
	             }
			});
    		
    		
	    	$.ajax({
				type:"POST",
				url: "../inc/../../../inc/InSightFuncs_AjaxForServiceModals.asp",
				cache: false,
				data: "action=GetContentForServiceBoardTransferRedispatchModal&memoNum=" + encodeURIComponent(myTicketNumber) + "&custID=" + encodeURIComponent(myCustID) + "&userNo=" + encodeURIComponent(myUserNo),
				success: function(response)
				 {
	             	$modal.find('#ServiceBoardXferModalContent').html(response);
	             },
	             failure: function(response)
				 {
				 	$modal.find('#ServiceBoardXferModalContent').html("Failed");
	             }
			});
		    
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
	     
 
</script>

	     
<!-- service board custom styles !-->
<style type="text/css">

	body{
		overflow-x:hidden;
		margin:10px;
		border: 2px solid <%=FSBoardKioskGlobalTitleGradientColor%>;
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
	   
	.legend-awaiting-dispatch{
		<% Response.Write("background:" & FSBoardKioskGlobalColorAwaitingDispatch & ";") %>
	   padding:5px 5px 5px 5px;
	   color:#000;
	   text-align:center;
	}
	
	.legend-closed{
		<% Response.Write("background:" & FSBoardKioskGlobalColorClosed & ";") %>
	   padding:5px 5px 5px 5px;
	   color:#000;
	   text-align:center;
	}

	.legend-enroute{
	   <% Response.Write("background:" & FSBoardKioskGlobalColorEnRoute & ";")%>
	   padding:5px 5px 5px 5px;
	   color:#000;
	   text-align:center;
	}
	
	.legend-onsite{
	   <% Response.Write("background:" & FSBoardKioskGlobalColorOnSite & ";")%>
	   padding:5px 5px 5px 5px;
	   color:#000;
	   text-align:center;
	}
		
	.legend-redo-swap{
	   <% Response.Write("background:" & FSBoardKioskGlobalColorRedoSwap & ";")%>
	   padding:5px 5px 5px 5px;
	   color:#000;
	   text-align:center;
	}

	.legend-redo-waitforparts{
	   <% Response.Write("background:" & FSBoardKioskGlobalColorRedoWaitForParts & ";")%>
	   padding:5px 5px 5px 5px;
	   color:#000;
	   text-align:center;
	}

	.legend-redo-followup{
	   <% Response.Write("background:" & FSBoardKioskGlobalColorRedoFollowUp & ";")%>
	   padding:5px 5px 5px 5px;
	   color:#000;
	   text-align:center;
	}

	.legend-redo-unabletowork{
	   <% Response.Write("background:" & FSBoardKioskGlobalColorRedoUnableToWork & ";")%>
	   padding:5px 5px 5px 5px;
	   color:#000;
	   text-align:center;
	}
		
	.legend-awaiting-acknowledgement{
	   <% Response.Write("background:" & FSBoardKioskGlobalColorAwaitingAcknowledgement & ";")%>
	   padding:5px 5px 5px 5px;
	   color:#000;
	   text-align:center;
	}
		
	.legend-dispatch-acknowledged{
	   <% Response.Write("background:" & FSBoardKioskGlobalColorDispatchAcknowledged & ";")%>
	   padding:5px 5px 5px 5px;
	   color:#000;
	   text-align:center;
	}

	.legend-declined{
	   <% Response.Write("background:" & FSBoardKioskGlobalColorDispatchDeclined & ";")%>
	   padding:5px 5px 5px 5px;
	   color:#000;
	   text-align:center;
	}


	.legend-urgent{  
		<% Response.Write("-webkit-box-shadow:inset 0px 0px 0px 3px " & FSBoardKioskGlobalColorUrgent & ";")%>
		<% Response.Write("-moz-box-shadow:inset 0px 0px 0px 3px " & FSBoardKioskGlobalColorUrgent & ";")%>
		<% Response.Write("box-shadow:inset 0px 0px 0px 3px " & FSBoardKioskGlobalColorUrgent & ";")%>    
		<% Response.Write("background-color:" & FSBoardKioskGlobalTitleGradientColor& ";")%>
	   padding:5px 5px 5px 5px;
	   color:#000;
	   text-align:center;
	}
		

		
	.alarm-bell{
		 position:absolute;
		 /*right:5px;*/
	}
	 
	.alarm-bell .alert-pop-up{
		display: none;
		background: #000;
		color: #fff;
		position: absolute;
		padding: 5px 10px 5px 10px;
		z-index: 900;
		margin:-17px 0px 0px 20px;
		font-weight: bold;
	}
 
	.alarm-bell:hover .alert-pop-up{
		 display: block;
	 }
	 	
	table thead a{
		color: #000;
	}

	
	.tr-awaiting-dispatch{
		<% Response.Write("background:" & FSBoardKioskGlobalColorAwaitingDispatch & ";") %>
	}
	

	.tr-closed{
		<% Response.Write("background:" & FSBoardKioskGlobalColorClosed & ";") %>
	}
		
	.tr-enroute{
		<% Response.Write("background:" & FSBoardKioskGlobalColorEnRoute & ";") %>
	}
	
	.tr-onsite{
		<% Response.Write("background:" & FSBoardKioskGlobalColorOnSite & ";") %>
	}
		
	.tr-redo-swap{
		<% Response.Write("background:" & FSBoardKioskGlobalColorRedoSwap & ";") %>
	}

	.tr-redo-waitforparts{
		<% Response.Write("background:" & FSBoardKioskGlobalColorRedoWaitForParts & ";") %>
	}

	.tr-redo-followup{
		<% Response.Write("background:" & FSBoardKioskGlobalColorRedoFollowUp & ";") %>
	}

	.tr-redo-unabletowork{
		<% Response.Write("background:" & FSBoardKioskGlobalColorRedoUnableToWork & ";") %>
	}

	.tr-awaiting-acknowledgement{
		<% Response.Write("background:" & FSBoardKioskGlobalColorAwaitingAcknowledgement & ";") %>
	}

	.tr-dispatch-acknowledged{
		<% Response.Write("background:" & FSBoardKioskGlobalColorDispatchAcknowledged& ";") %>
	}

	.tr-declined{
		<% Response.Write("background:" & FSBoardKioskGlobalColorDispatchDeclined& ";") %>
	}

	.tr-awaiting-dispatch-top{
		<% Response.Write("border-top: 1px solid #000000;") %>
		<% Response.Write("border-left: 1px solid #000000;") %>
		<% Response.Write("border-right: 1px solid #000000;") %>
	}
	
	.tr-awaiting-dispatch-bottom{
		<% Response.Write("border-bottom: 1px solid #000000;") %>
		<% Response.Write("border-left: 1px solid #000000;") %>
		<% Response.Write("border-right: 1px solid #000000;") %>
	}
	
	.Urgent-border-top{
		<% Response.Write("border-top: 3px solid " & FSBoardKioskGlobalColorUrgent & ";") %>
		<% Response.Write("border-left: 3px solid " & FSBoardKioskGlobalColorUrgent & ";") %>
		<% Response.Write("border-right: 3px solid " & FSBoardKioskGlobalColorUrgent & ";") %>
	}
	
	.Urgent-border-bottom{
		<% Response.Write("border-bottom: 3px solid " & FSBoardKioskGlobalColorUrgent & ";") %>
		<% Response.Write("border-left: 3px solid " & FSBoardKioskGlobalColorUrgent & ";") %>
		<% Response.Write("border-right: 3px solid " & FSBoardKioskGlobalColorUrgent & ";") %>
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
	   /* width:7%; */
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
		/*margin-left:-60px;
		margin-left: -475px;*/
	}
	
	.double-scroll{
		width:100%;
	}
	
	.fa-star{
		color:blue;
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
<!-- end service board custom styles !-->
	
	
	
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
	Response.Write("background: " & FSBoardKioskGlobalTitleGradientColor &"; /* For browsers that do not support gradients */" & "<br>")
    Response.Write("background: -webkit-linear-gradient(" & FSBoardKioskGlobalTitleGradientColor & ", #fff); /* For Safari 5.1 to 6.0 */" & "<br>")
    Response.Write("background: -o-linear-gradient(" & FSBoardKioskGlobalTitleGradientColor & ", #fff); /* For Opera 11.1 to 12.0 */" & "<br>")
    Response.Write("background: -moz-linear-gradient(" & FSBoardKioskGlobalTitleGradientColor & ", #fff); /* For Firefox 3.6 to 15 */" & "<br>")
    Response.Write("background: linear-gradient(" & FSBoardKioskGlobalTitleGradientColor & ", #fff); /* Standard syntax (must be last) */" & "<br>")
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
		  			<img src="<%= BaseURL %>clientfiles/<%= MUV_Read("ClientID") %>/logos/logo.png" class="navbar-logo">
		  			<font color="<%=FSBoardKioskGlobalTitleTextFontColor%>"><%=FSBoardKioskGlobalTitleText%></font>
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

			        <!-- in progress !-->
			        <div class="col-lg-2">
			        	<div class="legend-awaiting-dispatch">
			            	<h4>Awaiting <%= GetTerm("Dispatch") %></h4>
			            </div>
			        </div>
			        <!-- eof in progress stop !-->

			        <!-- in progress !-->
			        <div class="col-lg-2">
			        	<div class="legend-awaiting-acknowledgement">
			            	<h4><%= GetTerm("Dispatched") %></h4>
			            </div>
			        </div>
			        <!-- eof in progress stop !-->      

			        <!-- in progress !-->
			        <div class="col-lg-2">
			        	<div class="legend-awaiting-acknowledgement">
			            	<h4>Awaiting <%= GetTerm("Acknowledgement") %></h4>
			            </div>
			        </div>
			        <!-- eof in progress stop !-->

			        <!-- in declined !-->
			        <% If FS_TechCanDecline() Then %>
				        <div class="col-lg-2">
				        	<div class="legend-declined">
				            	<h4><%= GetTerm("Declined") %></h4>
				            </div>
				        </div>
				        <!-- eof in declined stop !-->
			        <% End If %>

			        <!-- in progress !-->
			        <div class="col-lg-2">
			        	<div class="legend-dispatch-acknowledged">
			            	<h4><%= GetTerm("Acknowledged") %></h4>
			            </div>
			        </div>
			        <!-- eof in progress stop !-->
		  		
			        <!-- no delivery !-->
			        <div class="col-lg-2">
			        	<div class="legend-enroute">
			            	<h4>En Route</h4>
			            </div>
			        </div>
			        <!-- eof no delivery !-->
			        
			        <!-- next stop !-->
			        <div class="col-lg-2">
			        	<div class="legend-onsite">
			            	<h4>On Site</h4>
			            </div>
			        </div>
			        <!-- eof next stop !-->
			    	
			        <!-- in progress !-->
			        <div class="col-lg-2">
			        	<div class="legend-closed">
			            	<h4>Closed</h4>
			            </div>
			        </div>
			        <!-- eof in progress stop !-->
			        
			        <!-- in progress !-->
			        <div class="col-lg-2">
			        	<div class="legend-redo-swap">
			            	<h4>Swap (<%= GetTerm("Redo") %>)</h4>
			            </div>
			        </div>
			        <!-- eof in progress stop !-->
			       			        
			        <!-- in progress !-->
			        <div class="col-lg-2">
			        	<div class="legend-redo-waitforparts">
			            	<h4>Wait For Parts (<%= GetTerm("Redo") %>)</h4>
			            </div>
			        </div>
			        <!-- eof in progress stop !-->

			        <!-- in progress !-->
			        <div class="col-lg-2">
			        	<div class="legend-redo-followup">
			            	<h4>Follow Up (<%= GetTerm("Redo") %>)</h4>
			            </div>
			        </div>
			        <!-- eof in progress stop !-->

			        <!-- in progress !-->
			        <div class="col-lg-2">
			        	<div class="legend-redo-unabletowork">
			            	<h4>Unable to Work (<%= GetTerm("Redo") %>)</h4>
			            </div>
			        </div>
			        <!-- eof in progress stop !-->
			        
			       			        
			        
			        <!-- in progress !-->
			        <div class="col-lg-2">
			        	<div class="legend-urgent">
			            	<h4><%= GetTerm("Urgent") %></h4>
			            </div>
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
		<table align="center">
			<tr>
		    	<td>
		 			<div class='list-boxes' id='sortableList'>
					<%
					
					Set cnn_FSBoardSum = Server.CreateObject("ADODB.Connection")
					cnn_FSBoardSum.open (Session("ClientCnnString"))
					Set rs_FSBoardSum = Server.CreateObject("ADODB.Recordset")
					Set rs_FSBoardSumForRegions = Server.CreateObject("ADODB.Recordset")
					rs_FSBoardSum.CursorLocation = 3 
					Set rs_FSBoardSum = cnn_FSBoardSum.Execute(SQL)
					
					'**************************************************************************************************					
					'SQL STMT to return only the field service technicians that have service call today
					'**************************************************************************************************
					
					SQL_FSBoardSum = "SELECT DISTINCT UserNoOfServiceTech FROM FS_ServiceMemosDetail "
					SQL_FSBoardSum = SQL_FSBoardSum & "WHERE MemoNumber IN "
                    SQL_FSBoardSum = SQL_FSBoardSum & "(SELECT MemoNumber FROM FS_ServiceMemos WHERE "
                    SQL_FSBoardSum = SQL_FSBoardSum & "(CurrentStatus = 'OPEN') OR "
                    SQL_FSBoardSum = SQL_FSBoardSum & "(CurrentStatus = 'CLOSE' AND "
                    
                    SQL_FSBoardSum = SQL_FSBoardSum & "YEAR(RecordCreatedDateTime) = YEAR(GetDate()) AND "
                    SQL_FSBoardSum = SQL_FSBoardSum & "MONTH(RecordCreatedDateTime) = MONTH(GetDate()) AND "
                    SQL_FSBoardSum = SQL_FSBoardSum & "DAY(RecordCreatedDateTime) = DAY(GetDate()) "
                    SQL_FSBoardSum = SQL_FSBoardSum & ")) AND UserNoOfServiceTech <> ''"
					SQL_FSBoardSum = SQL_FSBoardSum & " ORDER BY UserNoOfServiceTech"
					
					'Response.write(SQL_FSBoardSum & "<br><br>")
					

					Set rs_FSBoardSum = cnn_FSBoardSum.Execute(SQL_FSBoardSum)
					
					If not rs_FSBoardSum.EOF Then
					
						Do While Not rs_FSBoardSum.Eof

							If userIsArchived(rs_FSBoardSum("UserNoOfServiceTech")) = False AND userIsEnabled(rs_FSBoardSum("UserNoOfServiceTech")) = True Then	

								TechHasANYTicketsInRegion = False
						
								SQL_FSBoardSumForRegions = "SELECT DISTINCT CUstNum AS AccountNumber, MemoNumber FROM FS_ServiceMemosDetail "
								SQL_FSBoardSumForRegions = SQL_FSBoardSumForRegions & "WHERE MemoNumber IN "
			                    SQL_FSBoardSumForRegions = SQL_FSBoardSumForRegions & "(SELECT MemoNumber FROM FS_ServiceMemos WHERE "
			                    SQL_FSBoardSumForRegions = SQL_FSBoardSumForRegions & "(CurrentStatus = 'OPEN') OR "
			                    SQL_FSBoardSumForRegions = SQL_FSBoardSumForRegions & "(CurrentStatus = 'CLOSE' AND "
			                    
			                    SQL_FSBoardSumForRegions = SQL_FSBoardSumForRegions & "YEAR(RecordCreatedDateTime) = YEAR(GetDate()) AND "
			                    SQL_FSBoardSumForRegions = SQL_FSBoardSumForRegions & "MONTH(RecordCreatedDateTime) = MONTH(GetDate()) AND "
			                    SQL_FSBoardSumForRegions = SQL_FSBoardSumForRegions & "DAY(RecordCreatedDateTime) = DAY(GetDate()) "
			                    SQL_FSBoardSumForRegions = SQL_FSBoardSumForRegions & ")) AND UserNoOfServiceTech  = " & rs_FSBoardSum("UserNoOfServiceTech")
	
								'Response.Write(SQL_FSBoardSumForRegions & "<br>")
			                    
			                    Set rs_FSBoardSumForRegions = cnn_FSBoardSum.Execute(SQL_FSBoardSumForRegions)
			                    
								If not rs_FSBoardSumForRegions.EOF Then
								
									Do While NOT rs_FSBoardSumForRegions.EOF
									
										If GetServiceTicketCurrentStage(rs_FSBoardSumForRegions("MemoNumber")) <> "Received" Then 
									
											If LastTechUserNo(rs_FSBoardSumForRegions("MemoNumber")) = rs_FSBoardSum("UserNoOfServiceTech") Then
										
												If CurrentRegionToShow <> "" Then
												
													CustRegion = GetCustRegionIntRecIDByCustID(rs_FSBoardSumForRegions("AccountNumber"))
											
													If cint(CurrentRegionToShow) = cint(CustRegion) Then
														TechHasANYTicketsInRegion = True
													End IF
													
												End If
		
											End If
										
										End If
	
										If TechHasANYTicketsInRegion = True Then Exit Do
									
										rs_FSBoardSumForRegions.MoveNext
									Loop
								
								End IF
			                    
								If CurrentRegionToShow = "" Then TechHasANYTicketsInRegion = True
								
								'Only Write Route If The User Is Not Archived and Not Disabled
								If TechHasANYTicketsInRegion = True Then 
									Call TruckNumberWrite(rs_FSBoardSum("UserNoOfServiceTech")) 
								End IF
								
							End If
							
							rs_FSBoardSum.Movenext
						Loop
					End If
					%> 
				</div>
			</tr>
		</table>
	</div>
</div>

<%
Sub TruckNumberWrite(UserNoOfServiceTech) %>
		<td class='item col-lg-cust' TruckNumber='<%= UserNoOfServiceTech %>'>
		<div class='item-box'>
		<div class='scrollable-title' style='position: relative;'>
			<span class="trucknumber"><i class="fa fa-wrench" aria-hidden="true"></i>&nbsp;<%= UserNoOfServiceTech %></span>
			<span class="drivername"><%= GetUserDisplayNameByUserNo(UserNoOfServiceTech) %></span>
		</div>

	        <div class='table-responsive scrollable-table'>
		        <% Response.Write("<table id='truck" & UserNoOfServiceTech & "' name='truck" & UserNoOfServiceTech & "' class='food_planner table table-condensed clickable'>") %>
					<thead>
			        	<tr>
			        		<th class='sorttable_nosort'>Ticket #</th>
			        		<th class='sorttable_nosort'><%=GetTerm("Customer")%></th>.
			        	</tr>
			        </thead>
			        <tbody class='searchable'>
			        	<%'Get all the tickets for this truck
			        	
						Set cnn_Tickets = Server.CreateObject("ADODB.Connection")
						cnn_Tickets.open (Session("ClientCnnString"))
						Set rs_FSBoardDet = Server.CreateObject("ADODB.Recordset")
						rs_FSBoardDet.CursorLocation = 3 
						   	                    
						SQL_Tickets = "SELECT DISTINCT MemoNumber FROM FS_ServiceMemosDetail "
						SQL_Tickets = SQL_Tickets & "WHERE (MemoNumber IN "
                        SQL_Tickets = SQL_Tickets & "(SELECT MemoNumber FROM FS_ServiceMemos WHERE "
                        SQL_Tickets = SQL_Tickets & "(CurrentStatus = 'OPEN') OR "
                        SQL_Tickets = SQL_Tickets & "(CurrentStatus = 'CLOSE' AND "
	                    SQL_Tickets = SQL_Tickets & "YEAR(RecordCreatedDateTime) = YEAR(GetDate()) AND "
	                    SQL_Tickets = SQL_Tickets & "MONTH(RecordCreatedDateTime) = MONTH(GetDate()) AND "
	                    SQL_Tickets = SQL_Tickets & "DAY(RecordCreatedDateTime) = DAY(GetDate())) "
                        SQL_Tickets = SQL_Tickets & ")) AND (UserNoOfServiceTech = " & UserNoOfServiceTech &")"
   	                    
                        Set rs_FSBoardDet = cnn_Tickets.Execute(SQL_Tickets)
						If not rs_FSBoardDet.Eof Then

							NumLines = 0
							Do While not rs_FSBoardDet.Eof							
								
								If GetServiceTicketCurrentStage(rs_FSBoardDet("MemoNumber")) <> "Received" Then ' Need this in case something was undispatched
								
									If LastTechUserNo(rs_FSBoardDet("MemoNumber")) = UserNoOfServiceTech Then
									
										ServiceTicketCurrentStage = GetServiceTicketCurrentStage(rs_FSBoardDet("MemoNumber"))
										ServiceTicketCurrentStatus = GetServiceTicketStatus(rs_FSBoardDet("MemoNumber"))
									
										'Write first table row
										'**********************
										
										CustID = GetServiceTicketCust(rs_FSBoardDet("MemoNumber"))
																		
										If len(GetCustNameByCustNum(GetServiceTicketCust(rs_FSBoardDet("MemoNumber")))) > 19 then 
											Cnam = left(GetCustNameByCustNum(GetServiceTicketCust(rs_FSBoardDet("MemoNumber"))),19) 
										Else 
											Cnam = GetCustNameByCustNum(GetServiceTicketCust(rs_FSBoardDet("MemoNumber")))
										End If
										
																				
										If ServiceTicketCurrentStatus = "CLOSE" OR ServiceTicketCurrentStatus = "CANCEL" Then
											If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
												%>
												<tr class="tr-closed Urgent-border-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>" data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;"><%
											Else
												%><tr class="tr-closed tr-awaiting-dispatch-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;"><%
											End If
										
										End If
										
										If ServiceTicketCurrentStatus = "OPEN" AND ServiceTicketCurrentStage <> "En Route"_
										AND ServiceTicketCurrentStage <> "On Site"_
										AND ServiceTicketCurrentStage <> "Dispatched"_
										AND ServiceTicketCurrentStage <> "Dispatch Acknowledged"_
										AND ServiceTicketCurrentStage <> "Dispatch Declined" Then
										
																						
											If AwaitingRedispatch(rs_FSBoardDet("MemoNumber")) <> True Then
												If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
													%><tr class="tr-awaiting-dispatch Urgent-border-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;"><%
												Else
													%><tr class="tr-awaiting-dispatch tr-awaiting-dispatch-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;"><%
												End If
											Else
												If ServiceTicketCurrentStage = "Swap" Then
													className = "tr-redo-swap"
												ElseIf ServiceTicketCurrentStage = "Wait for parts" Then
													className = "tr-redo-waitforparts"
												ElseIf ServiceTicketCurrentStage = "Follow Up" Then
													className = "tr-redo-followup"
												ElseIf ServiceTicketCurrentStage = "Unable To Work" Then
													className = "tr-redo-unabletowork"
												End If

												If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
													%><tr class="<%= className %> Urgent-border-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;"><%
												Else
													%><tr class="<%= className %> tr-awaiting-dispatch-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;"><%
												End If
											End If
											
										End If
										
										
										
										If ServiceTicketCurrentStatus = "OPEN" AND ServiceTicketCurrentStage = "En Route" Then
										
											If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
												%><tr class="tr-enroute Urgent-border-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;"><%
											Else
												%><tr class="tr-enroute tr-awaiting-dispatch-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;"><%
											End If
										
										End If
		
										If ServiceTicketCurrentStatus = "OPEN" AND ServiceTicketCurrentStage = "On Site" Then
										
											If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
												%><tr class="tr-onsite Urgent-border-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;"><%
											Else
												%><tr class="tr-onsite tr-awaiting-dispatch-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;"><%
											End If
										
										End If
										
										
										If ServiceTicketCurrentStatus = "OPEN" AND ServiceTicketCurrentStage = "Dispatched" Then
										
											If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
												%><tr class="tr-awaiting-acknowledgement Urgent-border-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;"><%
											Else
												%><tr class="tr-awaiting-acknowledgement tr-awaiting-dispatch-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;"><%
											End If
										
										End If
										
										If ServiceTicketCurrentStatus = "OPEN" AND ServiceTicketCurrentStage = "Dispatch Acknowledged" Then
										
											If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
												%><tr class="tr-dispatch-acknowledged Urgent-border-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;"><%
											Else
												%><tr class="tr-dispatch-acknowledged tr-awaiting-dispatch-top" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;"><%
											End If
										
										End If
					
										If ServiceTicketCurrentStatus = "OPEN" AND ServiceTicketCurrentStage = "Dispatch Declined" Then
										
											If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
												%><tr class="tr-declined Urgent-border-top" style="cursor:pointer;"><%
											Else
												%><tr class="tr-declined tr-awaiting-dispatch-top" style="cursor:pointer;"><%
											End If
										
										End If
		
										%>
										
											<td><%= rs_FSBoardDet("MemoNumber") %></td>
											<td>
											<% If ServiceTicketCurrentStatus = "CLOSE" And ServiceTicketCurrentStage = "On Site" Then %>
												CLOSED
											<% Else %>
												<%= ServiceTicketCurrentStage %>
											<% End If %>
											
											</td>
										
										</tr>
										
										<%
		
																				
										'Write second table row
										'**********************
										
										If ServiceTicketCurrentStatus = "CLOSE" OR ServiceTicketCurrentStatus = "CANCEL" Then
																					
											If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
												%><tr class="tr-closed Urgent-border-bottom" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;"><%
											Else
												%><tr class="tr-closed tr-awaiting-dispatch-bottom" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;"><%
											End If
											
										End If
										
										
										If ServiceTicketCurrentStatus = "OPEN" AND ServiceTicketCurrentStage <> "En Route"_
										AND ServiceTicketCurrentStage <> "On Site"_
										AND ServiceTicketCurrentStage <> "Dispatched"_
										AND ServiceTicketCurrentStage <> "Dispatch Acknowledged"_
										AND ServiceTicketCurrentStage <> "Dispatch Declined" Then
											
											If AwaitingRedispatch(rs_FSBoardDet("MemoNumber")) <> True Then
												If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
													%><tr class="tr-awaiting-dispatch Urgent-border-bottom" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;"><%
												Else
													%><tr class="tr-awaiting-dispatch tr-awaiting-dispatch-bottom" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;"><%
												End If
											Else
												If ServiceTicketCurrentStage = "Swap" Then
													className = "tr-redo-swap"
												ElseIf ServiceTicketCurrentStage = "Wait for parts" Then
													className = "tr-redo-waitforparts"
												ElseIf ServiceTicketCurrentStage = "Follow Up" Then
													className = "tr-redo-followup"
												ElseIf ServiceTicketCurrentStage = "Unable To Work" Then
													className = "tr-redo-unabletowork"
												End If

												If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
													%><tr class="<%= className %> Urgent-border-bottom" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;"><%
												Else
													%><tr class="<%= className %> tr-awaiting-dispatch-bottom" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;"><%
												End If
											End If

										End If
										
										If ServiceTicketCurrentStatus = "OPEN" AND ServiceTicketCurrentStage = "En Route" Then
										
											If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
												%><tr class="tr-enroute Urgent-border-bottom" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;"><%
											Else
												%><tr class="tr-enroute tr-awaiting-dispatch-bottom" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;"><%
											End If
										
										End If
		
										If ServiceTicketCurrentStatus = "OPEN" AND ServiceTicketCurrentStage = "On Site" Then
										
											If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
												%><tr class="tr-onsite Urgent-border-bottom" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;"><%
											Else
												%><tr class="tr-onsite tr-awaiting-dispatch-bottom" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;"><%
											End If
										
										End If
										
		
										If ServiceTicketCurrentStatus = "OPEN" AND ServiceTicketCurrentStage = "Dispatched" Then
										
											If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
												%><tr class="tr-awaiting-acknowledgement Urgent-border-bottom" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;"><%
											Else
												%><tr class="tr-awaiting-acknowledgement tr-awaiting-dispatch-bottom" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;"><%
											End If
										
										End If
		
		
										If ServiceTicketCurrentStatus = "OPEN" AND ServiceTicketCurrentStage = "Dispatch Acknowledged" Then
										
											If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
												%><tr class="tr-dispatch-acknowledged Urgent-border-bottom" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;"><%
											Else
												%><tr class="tr-dispatch-acknowledged tr-awaiting-dispatch-bottom" data-toggle="modal" data-show="true" href="#" data-invoice-number="<%= rs_FSBoardDet("MemoNumber") %>" data-customer-id="<%= CustID %>"  data-customer-name="<%= Cnam %>" data-user-no="<%= UserNoOfServiceTech %>"  data-target="#serviceBoardTicketOptionsModal" data-tooltip="true" data-title="Service Ticket Options" style="cursor:pointer;"><%
											End If
										
										End If

										If ServiceTicketCurrentStatus = "OPEN" AND ServiceTicketCurrentStage = "Dispatch Declined" Then

											If TicketIsUrgent(rs_FSBoardDet("MemoNumber")) Then
												%><tr class="tr-declined Urgent-border-bottom" style="cursor:pointer;"><%
											Else
												%><tr class="tr-declined tr-awaiting-dispatch-bottom" style="cursor:pointer;"><%
											End If
										
										End If
																		
										If len(GetCustNameByCustNum(GetServiceTicketCust(rs_FSBoardDet("MemoNumber")))) > 19 then 
											Cnam = left(GetCustNameByCustNum(GetServiceTicketCust(rs_FSBoardDet("MemoNumber"))),19) 
										Else 
											Cnam = GetCustNameByCustNum(GetServiceTicketCust(rs_FSBoardDet("MemoNumber")))
										End If
		
										%>
										
											<td colspan="2"><span class="tooltip-button" data-toggle="tooltip" data-placement="bottom" title="Account #<%= CustID %>"><%= Cnam %></span></td>
										
										</tr>
									
	 							<%
	 								End If ' last tech user no
	 							End If ' status not received
								rs_FSBoardDet.movenext
								NumLines = NumLines + 1
								
							Loop
							
						End IF%>
                        </td>
			        </tbody>
		        </table>
	        </div>
            </div>
        <%Response.Write("</div>")
End Sub 

Set rs_FSBoardSum = Nothing
cnn_FSBoardSum.Close
Set cnn_FSBoardSum = Nothing

Sub CheckTables

	Set cnnCheckTables = Server.CreateObject("ADODB.Connection")
	cnnCheckTables.open (Session("ClientCnnString"))
	Set rsCheckTables = Server.CreateObject("ADODB.Recordset")
	rsCheckTables.CursorLocation = 3 
	
	SQL_CheckTables = "SELECT COL_LENGTH('Settings_Global', 'FSBoardKioskGlobalColorDispatchDeclined') AS IsItThere"
	Set rsCheckTables = cnnCheckTables.Execute(SQL_CheckTables)
	If IsNull(rsCheckTables("IsItThere")) Then
		SQL_CheckTables = "ALTER TABLE Settings_Global ADD FSBoardKioskGlobalColorDispatchDeclined varchar(255) NULL"
		Set rsCheckTables = cnnCheckTables.Execute(SQL_CheckTables)
		SQL_CheckTables = "UPDATE Settings_Global SET FSBoardKioskGlobalColorDispatchDeclined = '#4a86e8'"
		Set rsCheckTables = cnnCheckTables.Execute(SQL_CheckTables)
	End If

	SQL_CheckTables = "SELECT COL_LENGTH('Settings_Global', 'FS_TechCanDecline') AS IsItThere"
	Set rsCheckTables = cnnCheckTables.Execute(SQL_CheckTables)
	If IsNull(rsCheckTables("IsItThere")) Then
		SQL_CheckTables = "ALTER TABLE Settings_Global ADD FS_TechCanDecline int NULL"
		Set rsCheckTables = cnnCheckTables.Execute(SQL_CheckTables)
		SQL_CheckTables = "UPDATE Settings_Global SET FS_TechCanDecline= 0"
		Set rsCheckTables = cnnCheckTables.Execute(SQL_CheckTables)
	End If

	
	
	Set rsCheckTables = Nothing
	cnnCheckTables.Close
	Set cnnCheckTables = Nothing

End Sub

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

<!-- **************************************************************************************************************************** -->
<!-- MODALS FOR DELIVERY ALERTS BEGIN HERE !-->
<!-- **************************************************************************************************************************** -->

<!--#include file="../../../service/serviceBoardCommonModals.asp"-->

<!-- **************************************************************************************************************************** -->
<!-- MODALS FOR DELIVERY ALERTS END HERE !-->
<!-- **************************************************************************************************************************** -->
   
  </body>
</html>
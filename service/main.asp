<!--#include file="../inc/header.asp"-->

<script src="<%= baseURL %>js/moment.min.js" type="text/javascript"></script>
<script type="text/javascript" src="<%= BaseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.min.js"></script>
<link href="<%= BaseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.css" rel="stylesheet">

<link rel="stylesheet" href="<%= BaseURL %>js/bootstrap-multiselect/bootstrap-multiselect.css" type="text/css">
<script type="text/javascript" src="<%= BaseURL %>js/bootstrap-multiselect/bootstrap-multiselect.js"></script>

<link rel="stylesheet" href="https://cdn.datatables.net/1.10.20/css/jquery.dataTables.min.css" />
<script type="text/javascript" src="https://cdn.datatables.net/1.10.20/js/jquery.dataTables.min.js"></script>

<!-- Add fancyBox main JS and CSS files -->
<script type="text/javascript" src="<%= BaseURL %>js/jquery-lightbox/jquery.fancybox.js?v=2.1.5"></script>
<link rel="stylesheet" href="<%= BaseURL %>js/jquery-lightbox/jquery.fancybox.css?v=2.1.5" type="text/css" media="screen" />

<!-- time picker !-->
<link rel="stylesheet" href="<%= BaseURL %>js/timepicker/timepicker/ui-1.10.0/ui-lightness/jquery-ui-1.10.0.custom.min.css" type="text/css" />
<link rel="stylesheet" href="<%= BaseURL %>js/timepicker/timepicker/jquery.ui.timepicker.css?v=0.3.3" type="text/css" />
<script type="text/javascript" src="<%= BaseURL %>js/timepicker/timepicker/ui-1.10.0/jquery.ui.core.min.js"></script>
<script type="text/javascript" src="<%= BaseURL %>js/timepicker/timepicker/ui-1.10.0/jquery.ui.widget.min.js"></script>
<script type="text/javascript" src="<%= BaseURL %>js/timepicker/timepicker/ui-1.10.0/jquery.ui.tabs.min.js"></script>
<script type="text/javascript" src="<%= BaseURL %>js/timepicker/timepicker/ui-1.10.0/jquery.ui.position.min.js"></script>
<script type="text/javascript" src="<%= BaseURL %>js/timepicker/timepicker/jquery.ui.timepicker.js?v=0.3.3"></script>
<!-- eof time picker !-->




<%

NumberOfMinutesInServiceDayVar = GetNumberOfMinutesInServiceDay()
'PeriodSeqBeingEvaluated = GetLastClosedPeriodSeqNum() ' Believe it or not, we do need this - actually, I dont think we do

'Read field service board settings


SQL = "SELECT * FROM Settings_Global"
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
If not rs.EOF Then
	ServiceColorsOn = rs("ServiceColorsOn")
	ServicePriorityColor = rs("ServicePriorityColor")
	ServiceNormalAlertColor = rs("ServiceNormalAlertColor")
End If
set rs = Nothing
cnn8.close
set cnn8 = Nothing

SQL = "SELECT * FROM Settings_FieldService"
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
If not rs.EOF Then
	FilterChangeIndicatorAndButtonColor = rs("FilterChangeIndicatorAndButtonColor")	
	ShowSeparateFilterChangesTabOnServiceScreen = rs("ShowSeparateFilterChangesTabOnServiceScreen")	
	ServiceTicketScreenShowHoldTab = rs("ServiceTicketScreenShowHoldTab")
	FSBoardKioskGlobalUseRegions = rs("FSBoardKioskGlobalUseRegions")
	FSBoardKioskGlobalTitleText = rs("FSBoardKioskGlobalTitleText")
	FSBoardKioskGlobalTitleTextFontColor = rs("FSBoardKioskGlobalTitleTextFontColor")
	FSBoardKioskGlobalTitleGradientColor = rs("FSBoardKioskGlobalTitleGradientColor")
	FSBoardKioskGlobalColorPieTimer = rs("FSBoardKioskGlobalColorPieTimer")
	FSBoardKioskGlobalColorAwaitingDispatch = rs("FSBoardKioskGlobalColorAwaitingDispatch")
	FSBoardKioskGlobalColorAwaitingAcknowledgement = rs("FSBoardKioskGlobalColorAwaitingAcknowledgement")
	FSBoardKioskGlobalColorDispatchAcknowledged = rs("FSBoardKioskGlobalColorDispatchAcknowledged")
	FSBoardKioskGlobalColorDispatchDeclined = rs("FSBoardKioskGlobalColorDispatchDeclined")
	FSBoardKioskGlobalColorEnRoute = rs("FSBoardKioskGlobalColorEnRoute")
	FSBoardKioskGlobalColorOnSite = rs("FSBoardKioskGlobalColorOnSite")
	FSBoardKioskGlobalColorRedoSwap = rs("FSBoardKioskGlobalColorRedoSwap")
	FSBoardKioskGlobalColorRedoWaitForParts = rs("FSBoardKioskGlobalColorRedoWaitForParts")
	FSBoardKioskGlobalColorRedoFollowUp = rs("FSBoardKioskGlobalColorRedoFollowUp")
	FSBoardKioskGlobalColorRedoUnableToWork = rs("FSBoardKioskGlobalColorRedoUnableToWork")
	FSBoardKioskGlobalColorClosed = rs("FSBoardKioskGlobalColorClosed")
	FSBoardKioskGlobalColorUrgent = rs("FSBoardKioskGlobalColorUrgent")		
End If
set rs = Nothing
cnn8.close
set cnn8 = Nothing

If FilterChangeIndicatorAndButtonColor = "" Then FilterChangeIndicatorAndButtonColor = "#dddd53"
If IsNull(FilterChangeIndicatorAndButtonColor) Then FilterChangeIndicatorAndButtonColor = "#dddd53"	

FSBoardKioskGlobalTitleText = Replace(FSBoardKioskGlobalTitleText,"'","")
FSBoardKioskGlobalTitleText = Replace(FSBoardKioskGlobalTitleText,"~today~",FormatDateTime(Now(),2))
FSBoardKioskGlobalTitleText = Replace(FSBoardKioskGlobalTitleText,"~dow~",WeekDayName(Datepart("w",Now())))

If FSBoardKioskGlobalTitleGradientColor = "" Then FSBoardKioskGlobalTitleGradientColor = "#80B8FF"
If IsNull(FSBoardKioskGlobalTitleGradientColor) Then FSBoardKioskGlobalTitleGradientColor = "#80B8FF"
	
Session("FSBoardKioskGlobalColorPieTimer") = Replace(FSBoardKioskGlobalColorPieTimer,"#","") ' Just this one for Javascript
%>

<!--#include file="../css/fa_animation_styles.css"-->
<!--#include file="../inc/jquery_table_search.asp"-->
<!--#include file="../inc/mail.asp"-->
<!--#include file="../inc/InsightFuncs_Service.asp"-->
<!--#include file="../inc/InSightFuncs_BizIntel.asp"-->
<!--#include file="../inc/InsightFuncs_Equipment.asp"-->
<!--#include file="../inc/InsightFuncs_AR_AP.asp"-->

<!-----------------IMPORTANT FILE FOR SERVICE BOARD HEADER ------------------------------------------->
<!-- JavaScript Cookie Files To Save State of Dismissed Alerts -->
<script src="<%= BaseURL %>js/js.cookie.js"></script>
<!-- End JavaScript Cookie -->
<!-----------------END IMPORTANT FILE FOR SERVICE BOARD HEADER ---------------------------------------->
<%
'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)
%>


<script type="text/javascript">
//**********************************************
$.ajaxSetup ({
    // Disable caching of AJAX responses
    cache: false
});

$(window).on("load",function() {
	var activeTab=$(".rowtext .tab-pane.active");
	console.log(activeTab.attr("data-source"));
	activeTab.load(activeTab.attr("data-source"),
		function(response, status, xhr) {
			if (status == "error") {
				var msg = "Sorry but there was an error: ";
				console.log(msg + xhr.status + " " + xhr.statusText);

			}
		});
	$('a[data-toggle="tab"]').on("show.bs.tab", function (e) {
	$(".waitdiv").removeClass("d-none");
			var selectedTabID=$(e.target).attr("href");
			console.log(selectedTabID);
		    var activeTab=$(".rowtext .tab-pane"+selectedTabID);
			activeTab.load(activeTab.attr("data-source"),
			function(response, status, xhr) {
				if (status == "error") {
					var msg = "Sorry but there was an error: ";
					console.log(msg + xhr.status + " " + xhr.statusText);

				}
				else {
					console.log("loaded");
				}
				$(".waitdiv").addClass("d-none");
			});
		    
		});
});


//***************************
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
	
	$(document).ready(function() {
	
		$("#PleaseWaitPanel").hide();	

		$('[data-toggle="tooltip"]').tooltip();
		
		
		//************************************************************
		//Code to save current tab when pie time reloads the page
				
		//$('a[data-toggle="tab"]').click(function (e) {
		//    e.preventDefault();
		//    $(this).tab('show');
		//});
		//*************************************************************
		
		
		var selectedTab = localStorage.getItem('selectedTab');
		if (selectedTab != null) {
		    $('a[data-toggle="tab"][href="' + selectedTab + '"]').tab('show');
		}	
	
		
		//End code to save current tab when pie time reloads the page	
		//************************************************************
		
		
		var rgbcolor = '<%= Session("FSBoardKioskGlobalColorPieTimer") %>';

		var pagetimer = new Timer(function() {
			$(".waitdiv").removeClass("d-none");
		    var activeTab=$(".rowtext .tab-pane.active");
			console.log(activeTab.attr("data-source"));
			activeTab.load(activeTab.attr("data-source"),
			function(response, status, xhr) {
			if (status == "error") {
				var msg = "Sorry but there was an error: ";
				console.log(msg + xhr.status + " " + xhr.statusText);

			}
			else console.log("loaded");
			$(".waitdiv").addClass("d-none");
			pagetimer.resume();
			$('#timer').pietimer('start');
		});
		},  120*1000);
				
		$('#timer').pietimer({
			seconds: 120,
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
		

		//************************************************************
		//Code to manage active tabs when page reloads

  		currentlySelectedTab = window.location.hash;
  		//alert(currentlySelectedTab);
  		
  		if (currentlySelectedTab == '') {
  			$('.nav-tabs li').removeClass('active');
  			//$('a[data-toggle="tab"][href="#awaitingdispatch"]').addClass('active');
  			$('.nav-tabs a[href="#awaitingdispatch"]').tab('show');
  		}
  
		$('a[data-toggle="tab"]').on('click', function (e) {
		    //get id of tab just clicked
		    currActiveTab = $(this).attr("id")
		    //remove active class on all tabs
		    $('.nav-tabs li').removeClass('active');
		    //set active class on recently clicked tab
		    $('a[data-toggle="tab"][href="' + currActiveTab + '"]').addClass('active');
		})
		
		//End code to manage active tabs when page reloads	
		//************************************************************

		
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

	   	$('#serviceBoardChangeTypeModal').on('show.bs.modal', function () {
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

	   	$('#modalCreateNewServiceTicketForClient').on('show.bs.modal', function () {
			//if (Cookies.get('service-board-pause-autorefresh') == 'false' ){
				$('#switchAutomaticRefresh').prop('checked', true).trigger("change");
			//}
	    });

	   	$('#modalEditExistingServiceTicketForClient').on('show.bs.modal', function (e) {

		    //get data-id attribute of the clicked service ticket
		    var passedMemoNumber = $(e.relatedTarget).data('memo-number');
		    var passedCustID = $(e.relatedTarget).data('cust-id');
		    
		    //populate the textboxes with the id of the clicked service ticket
		    $(e.currentTarget).find('input[name="txtMemoNumberCloseCancel"]').val(passedMemoNumber);
			$(e.currentTarget).find('input[name="txtCustIDCloseCancel"]').val(passedCustID);
			$(e.currentTarget).find('input[name="txtReturnPathCloseCancel"]').val("ServiceMain");

 			//alert("passedMemoNumber: " + passedMemoNumber);		
 			    	    
		    var $modal = $(this);

			$.ajax({
				type:"POST",
				url: "../inc/InSightFuncs_AjaxForServiceModals.asp",
				data: "action=GetContentForEditServiceTicketModalTitle&memo="+encodeURIComponent(passedMemoNumber)+ "&custID=" + encodeURIComponent(passedCustID),
				success: function(response){
					$("#modalEditExistingServiceTicketForClientTitle").html(response);
	             },
	            failure: function(response)
				 {
				  	$modal.find('#modalEditExistingServiceTicketForClientTitle').html("Failed");
	             }
			});
	
	    	$.ajax({
				type:"POST",
				url: "../inc/InSightFuncs_AjaxForServiceModals.asp",
				cache: false,
				data: "action=GetContentForEditServiceTicketModal&memo="+encodeURIComponent(passedMemoNumber)+ "&custID=" + encodeURIComponent(passedCustID),
				success: function(response)
				 {
	               	 $modal.find("#selectedTicketNumberInformation").html(response);
	               	 $modal.find("#btnEditExistingServiceTicketForClientSave").show();               	 
	             },
	             failure: function(response)
				 {
				  	$modal.find('#selectedTicketNumberInformation').html("Failed");
	             }
			});
    				
			$('#switchAutomaticRefresh').prop('checked', true).trigger("change");

	    });
	    
	    
	    
	   	$('#modalViewOpenClosedServiceTicketDetailsForClient').on('show.bs.modal', function (e) {

		    //get data-id attribute of the clicked service ticket
		    var passedMemoNumber = $(e.relatedTarget).data('memo-number');
		    var passedCustID = $(e.relatedTarget).data('cust-id');
		    
		    //populate the textboxes with the id of the clicked service ticket
		    $(e.currentTarget).find('input[name="txtMemoNumberView"]').val(passedMemoNumber);
			$(e.currentTarget).find('input[name="txtCustIDView"]').val(passedCustID);

 	    
		    var $modal = $(this);
	
	    	$.ajax({
				type:"POST",
				url: "../inc/InSightFuncs_AjaxForServiceModals.asp",
				cache: false,
				data: "action=GetContentForViewOpenClosedServiceTicketModal&memo="+encodeURIComponent(passedMemoNumber)+ "&custID=" + encodeURIComponent(passedCustID),
				success: function(response)
				 {
	               	 $modal.find("#selectedOpenClosedTicketNumberInformation").html(response);              	 
	             },
	             failure: function(response)
				 {
				  	$modal.find('#selectedOpenClosedTicketNumberInformation').html("Failed");
	             }
			});
    				
			$('#switchAutomaticRefresh').prop('checked', true).trigger("change");

	    });
	    

 		//**************************************************************************
 		//Special code here******
 		//The service ticket options modal leads to the opening of other modals
 		//So when we hide this modal, we want to keep the pause button paused
 		
	    $('#serviceBoardTicketOptionsModal').on('hidden.bs.modal', function () {
		   // if (Cookies.get('service-board-pause-autorefresh') == 'false' ){
		       $('#switchAutomaticRefresh').prop('checked', true).trigger("change");
		    //}
	    });
	    //**************************************************************************
 
 
 
		//These other three modals appear when a user has clicked a button on the
		//first modal, the options modal. So when these modals are closed, if the user
		//has not said to keep the board paused, then we can start the timer again.
		      	
		$('#serviceBoardXferModal').on('hidden.bs.modal', function () {
		    //if (Cookies.get('service-board-pause-autorefresh') == 'false' ){
		       $('#switchAutomaticRefresh').prop('checked', false).trigger("change");
		   // }	
    	}); 	

		$('#serviceBoardChangeTypeModal').on('hidden.bs.modal', function () {
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
   	
	   	$('#modalCreateNewServiceTicketForClient').on('hidden.bs.modal', function () {
			//if (Cookies.get('service-board-pause-autorefresh') == 'false' ){
				$('#switchAutomaticRefresh').prop('checked', false).trigger("change");
			//}
	    });

	   	$('#modalEditExistingServiceTicketForClient').on('hidden.bs.modal', function () {
			//if (Cookies.get('service-board-pause-autorefresh') == 'false' ){
				$('#switchAutomaticRefresh').prop('checked', false).trigger("change");
			//}
	    });

	   	$('#modalViewOpenClosedServiceTicketDetailsForClient').on('hidden.bs.modal', function () {
			//if (Cookies.get('service-board-pause-autorefresh') == 'false' ){
				$('#switchAutomaticRefresh').prop('checked', false).trigger("change");
			//}
	    });
			
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
				url: "../inc/InSightFuncs_AjaxForServiceModals.asp",
				cache: false,
				data: "action=GetContentForServiceBoardTicketOptionsModal&memoNum=" + encodeURIComponent(myTicketNumber) + "&custID=" + encodeURIComponent(myCustID) + "&userNo=" + encodeURIComponent(myUserNo),
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
				url: "../inc/InSightFuncs_AjaxForServiceModals.asp",
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
				url: "../inc/InSightFuncs_AjaxForServiceModals.asp",
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
		

		$('#serviceBoardChangeTypeModal').on('show.bs.modal', function(e) {
		
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
				url: "../inc/InSightFuncs_AjaxForServiceModals.asp",
				cache: false,
				data: "action=GetTitleForServiceBoardChangeTypeModal&memoNum=" + encodeURIComponent(myTicketNumber) + "&custID=" + encodeURIComponent(myCustID) + "&userNo=" + encodeURIComponent(myUserNo),
				success: function(response)
				 {
	               	 $modal.find('#ServiceBoardChangeTypeModalTitle').html(response);
	             },
	             failure: function(response)
				 {
				 	$modal.find('#ServiceBoardChangeTypeModalTitle').html("Failed");
	             }
			});
    		
    		
	    	$.ajax({
				type:"POST",
				url: "../inc/InSightFuncs_AjaxForServiceModals.asp",
				cache: false,
				data: "action=GetContentForServiceBoardChangeTypeModal&memoNum=" + encodeURIComponent(myTicketNumber) + "&custID=" + encodeURIComponent(myCustID) + "&userNo=" + encodeURIComponent(myUserNo),
				success: function(response)
				 {
	             	$modal.find('#ServiceBoardChangeTypeModalContent').html(response);
	             },
	             failure: function(response)
				 {
				 	$modal.find('#ServiceBoardChangeTypeModalContent').html("Failed");
	             }
			});
		    
		});


///////////////////////////////////////////
		$('#serviceBoardDispatchModal').on('show.bs.modal', function(e) {
		
		
		    //get data-id attribute of the clicked service ticket
		    var myTicketNumber = $(e.relatedTarget).data('service-ticket-number');
		    
		    //populate the textboxes with the id of the clicked service ticket
		    $(e.currentTarget).find('input[name="txtTicketNumber"]').val(myTicketNumber);
		    	    
		    var $modal = $(this);
	    		
	    	$.ajax({
				type:"POST",
				url: "../inc/InSightFuncs_AjaxForServiceModals.asp",
				cache: false,
				data: "action=GetTitleForServiceBoardDispatchModal&memoNum=" + encodeURIComponent(myTicketNumber),
				success: function(response)
				 {
	               	 $modal.find('#ServiceBoardDispatchModalTitle').html(response);
	             },
	             failure: function(response)
				 {
				 	$modal.find('#ServiceBoardDispatchModalTitle').html("Failed");
	             }
			});
    		
    		
	    	$.ajax({
				type:"POST",
				url: "../inc/InSightFuncs_AjaxForServiceModals.asp",
				cache: false,
				data: "action=GetContentForServiceBoardDispatchModal&memoNum=" + encodeURIComponent(myTicketNumber),
				success: function(response)
				 {
	             	$modal.find('#ServiceBoardDispatchModalContent').html(response);
	             },
	             failure: function(response)
				 {
				 	$modal.find('#ServiceBoardDispatchModalContent').html("Failed");
	             }
			});
		    
		});
	

		$('#modalEquipmentVPC').on('show.bs.modal', function(j) {

		    //get data-id attribute of the clicked order
		    var CustID = $(j.relatedTarget).data('cust-id');
		    var LCPGP = $(j.relatedTarget).data('lcp-gp');
	 
		    //populate the textbox with the id of the clicked order
		    $(j.currentTarget).find('input[name="txtCustIDToPass"]').val(CustID);
		    $(j.currentTarget).find('input[name="txtLastClosedPeriodGP"]').val(LCPGP);
		    	    
		    var $modal = $(this);
		    //$modal.find('#PleaseWaitPanelModal').show();  
	
	    	$.ajax({
				type:"POST",
				url: "../inc/InSightFuncs_AjaxForBizIntelModals.asp",
				cache: false,
				data: "action=GetTitleForEquipmentVPCModal&CustID="+encodeURIComponent(CustID)+"&LCPGP="+encodeURIComponent(LCPGP),
				success: function(response)
				 {
	               	 $modal.find('#modalEquipmentVPCTitle').html(response);            	 
	             },
	             failure: function(response)
				 {
				  	$modal.find('#modalEquipmentVPCTitle').html("Failed");
	             }
			});
			

	    	$.ajax({
				type:"POST",
				url: "../inc/InSightFuncs_AjaxForBizIntelModals.asp",
				cache: false,
				data: "action=GetContentForEquipmentVPCModal&CustID="+encodeURIComponent(CustID),
			  	beforeSend: function() {
			     	$('#PleaseWaitPanelModal').show();
			     	$modal.find('#modalCategoryVPCContent').html('');
			  	},
			  	complete: function(){
			     	$('#PleaseWaitPanelModal').hide();
			  	},
				success: function(response)
				 {
	               	 $("#PleaseWaitPanelModal").hide();
	               	 $modal.find('#modalCategoryVPCContent').html(response);    
          	 
	             },
	             failure: function(response)
				 {
				  	$modal.find('#modalCategoryVPCContent').html("Failed");
	             }
			});
		});



		$('#modalEditServiceTicketNotes').on('show.bs.modal', function(e) {
	
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
				url: "../inc/InSightFuncs_AjaxForServiceModals.asp",
				cache: false,
				data: "action=GetContentForServiceTicketNotesModal&memoNum=" + encodeURIComponent(myTicketNumber) + "&custID=" + encodeURIComponent(myCustID) + "&userNo=" + encodeURIComponent(myUserNo),
				success: function(response)
				 {
	               	 $modal.find('#modalEditServiceTicketNotesContent').html(response);               	 
	             },
	             failure: function(response)
				 {
				  	$modal.find('#modalEditServiceTicketNotesContent').html("Failed");
		            //var height = $(window).height() - 600;
		            //$(this).find(".modal-body").css("max-height", height);
	             }
			});
			
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
				buttonWidth: '425px',
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
		
	
   	
	}); //end document.ready() function
	

	function showAlert(Msg)
	{
	
		var Msg1 = Msg
		
		swal({
			title: "Cancelled While En Route",
			text: Msg1,
			type: 'warning',
			timer: 300000,
			confirmButtonText: 'OK'
		});
	
	} 
	
	
	function ajaxRowMode(type, id, mode) {
	
		$('#ajaxRow'+type+'-'+id).attr("class", "ajaxRow"+mode);
		if(id==0){
			$('#ajaxRow'+type+'-' + 0 + '').remove();
		}	
	
		 $(".ajaxRowEdit").find('input[disabled="true"]').each(function () {
		     $(this).removeAttr("disabled");
		 });
		
	}

	
</script>


<!-- modal scroll !-->
<script type="text/javascript">
  $(document).ready(ajustamodal);
  $(window).resize(ajustamodal);
  function ajustamodal() {
    var altura = $(window).height() - 155; //value corresponding to the modal heading + footer
    $(".ativa-scroll").css({"height":altura,"overflow-y":"auto"});
  }
</script>
<!-- eof modal scroll !-->


<%
If advancedDispatchIsOn() Then
	If Session("MultiUseVar") <> "" Then
		If UserIsServiceManager(Session("UserNo")) = True Then
			Response.write("<script type='text/javascript'>showAlert('" & Session("MultiUseVar") & "');</script>")
		End If
		Session("MultiUseVar") = ""
	End If
End If
%>


<!-- DYNAMIC FORM !-->
<style type="text/css">

	 body{
		 overflow-x:hidden;
	 }
	 
 	.email-table{
		width:46%;
	}
		
	.bs-example-modal-lg-customize .row{
		margin-bottom: 10px;
	 	width: 100%;
		overflow: hidden;
	}
	
	.bs-example-modal-lg-customize .left-column{
		background: #eaeaea;
		padding-bottom: 1000px;
	    margin-bottom: -1000px;
	}
	
	.bs-example-modal-lg-customize .left-column h4{
		margin-top: 0px;
	}
	
	.bs-example-modal-lg-customize .right-column{
		background: #fff;
		padding-bottom: 1000px;
	    margin-bottom: -1000px;
	}
	
	
		table.sortable th:not(.sorttable_sorted):not(.sorttable_sorted_reverse):not(.sorttable_nosort):after { 
	    content: " \25B4\25BE" 
	}
	
	table thead a{
		color: #000;
	}
	
	.tr-even{
		background: #f6f6f6;
	}
	
	.tr-odd{
		background: #fff;
	}
	 
	.date-range label{
		font-weight: normal;
		margin-right: 10px;
		margin-top: 10px;
	}
	
	.data-range-box{
		border:1px solid #ccc;
		padding-top: 5px;
	}
	
	.btn-link{
		padding: 0px;
		text-align: left;
	}
	
	.date-time-hidden-value{
		display:none;
	}
	
	.rowtext{
		font-size:12px;
	}
	
	.fa-exclamation-triangle{
	 	color:#ddcd1e;
	 	cursor:pointer;
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
	 
	
	.yesbtn{
		background: transparent;
		border: 0px;
		color: green;
	}
	
	.nobtn{
		background: transparent;
		border: 0px;
		color: red;
	}
	
	.table-info{
		padding: 5px;
		border: 1px solid #eaeaea;
	}
	
	.table-info .table>tbody>tr>td, .table>tbody>tr>th, .table>tfoot>tr>td, .table>tfoot>tr>th, .table>thead>tr>td, .table>thead>tr>th{
		border: 0px;
		font-weight: bold;
		line-height: 1;
	}
	 
	 
	 .page-header{
		 border-bottom:0px;
	 }
	
	.heading-legend{
		border-bottom:1px solid #eee;
		margin-bottom:20px;
		margin-top: 35px;
	 }
	
	.heading-legend h1{
		margin:0px;
	}
	
	.custom-table{
	 	font-size: 11px;
	}
	
	#td-padding{
		padding:5px 25px 5px 5px;
		display: block;
	} 

	.modal.modal-wide .modal-dialog {
	  width: 50%;
	}
	.modal-wide .modal-body {
	  overflow-y: auto;
	}
	
	.modal.modal-xwide .modal-dialog {
	  width: 70%;
	}
	.modal-xwide .modal-body {
	  overflow-y: auto;
	  max-height:600px;
	}

	.modal.modal-wide-autocomplete .modal-dialog {
	  width: 50%;
	}
	.modal-wide-autocomplete .modal-body {
	  /*overflow-y: auto;*/
	}
	
	.ativa-scroll{
		 max-height: 300px
	 }
	
	.mark {
	    background-color: yellow;
	    color: black;
	}
	
	
	.pause-timer{
		margin:0px 0px 0px 0px;
		
	}

  .pause{
 	   margin:10px 30px 0px 0px;
	   color:#337ab7;
	   float:left
   }
   .material-switch{
	   display: inline-block;
	   }

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
		
	#PleaseWaitPanel{
		position: fixed;
		left: 470px;
		top: 275px;
		width: 975px;
		height: 300px;
		z-index: 9999;
		background-color: #fff;
		opacity:1.0;
		text-align:center;
	}    
	
	#PleaseWaitPanelModal{
		position: relative;
		z-index: 9999;
		background-color: #fff;
		opacity:1.0;
		text-align:center;
	} 
	
	#PleaseWaitPanelModalService{
		position: relative;
		z-index: 9999;
		background-color: #fff;
		opacity:1.0;
		text-align:center;
	}  
 	
	#PleaseWaitPanelModalServiceCloseCancel{
		position: relative;
		z-index: 9999;
		background-color: #fff;
		opacity:1.0;
		text-align:center;
	}  

 
 
	.large-table {
	    font-size: 12px;
	}
	
	.nav-tabs {
	    border-bottom: 0px;
	}
	.nav-tabs>li>a{
		background: #f5f5f5;
		border: 1px solid #ccc;
		color: #000;
	}

	.nav-tabs>li>a:hover{
		border: 1px solid #ccc;
	}

	.nav-tabs>li.active>a, .nav-tabs>li.active>a:focus, .nav-tabs>li.active>a:hover{
		color: #fff;
		border: 1px solid #ccc;
		background:#337ab7;
	}
	
	.labelAwaitingDispatch {
	    display: inline-block;
	    padding: .4em .6em .4em;
	    font-size: 100%;
	    font-weight: 400;
	    line-height: 1;
	    color: #fff;
	    text-align: center;
	    white-space: nowrap;
	    vertical-align: baseline;
	    border-radius: .25em;
	    margin-top:5px;
	    margin-bottom:3px;	    
	    <% Response.Write("background:" & FSBoardKioskGlobalColorAwaitingDispatch & ";")%>
	}	

	.labelAwaitingAcknowledgement {
	    display: inline-block;
	    padding: .4em .6em .4em;
	    font-size: 100%;
	    font-weight: 400;
	    line-height: 1;
	    color: #fff;
	    text-align: center;
	    white-space: nowrap;
	    vertical-align: baseline;
	    border-radius: .25em;
	    margin-top:5px;
	    margin-bottom:3px;	    
	    <% Response.Write("background:" & FSBoardKioskGlobalColorAwaitingAcknowledgement & ";")%>
	}	


	.labelDispatchAcknowledged {
	    display: inline-block;
	    padding: .4em .6em .4em;
	    font-size: 100%;
	    font-weight: 400;
	    line-height: 1;
	    color: #fff;
	    text-align: center;
	    white-space: nowrap;
	    vertical-align: baseline;
	    border-radius: .25em;
	    margin-top:5px;
	    margin-bottom:3px;	
	    <% Response.Write("background:" & FSBoardKioskGlobalColorDispatchAcknowledged & ";")%>
	}	


	.labelEnRoute {
	    display: inline-block;
	    padding: .4em .6em .4em;
	    font-size: 100%;
	    font-weight: 400;
	    line-height: 1;
	    color: #fff;
	    text-align: center;
	    white-space: nowrap;
	    vertical-align: baseline;
	    border-radius: .25em;
	    margin-top:5px;
	    margin-bottom:3px;	    
	    <% Response.Write("background:" & FSBoardKioskGlobalColorEnRoute & ";")%>
	}	


	.labelOnSite {
	    display: inline-block;
	    padding: .4em .6em .4em;
	    font-size: 100%;
	    font-weight: 400;
	    line-height: 1;
	    color: #fff;
	    text-align: center;
	    white-space: nowrap;
	    vertical-align: baseline;
	    border-radius: .25em;
	    margin-top:5px;
	    margin-bottom:3px;	   
	    <% Response.Write("background:" & FSBoardKioskGlobalColorOnSite & ";")%>
	}	


	.labelSwap {
	    display: inline-block;
	    padding: .4em .6em .4em;
	    font-size: 100%;
	    font-weight: 400;
	    line-height: 1;
	    color: #fff;
	    text-align: center;
	    white-space: nowrap;
	    vertical-align: baseline;
	    border-radius: .25em;
	    margin-top:5px;
	    margin-bottom:3px;	   
	    <% Response.Write("background:" & FSBoardKioskGlobalColorRedoSwap & ";")%>
	}	


	.labelWaitForParts {
	    display: inline-block;
	    padding: .4em .6em .4em;
	    font-size: 100%;
	    font-weight: 400;
	    line-height: 1;
	    color: #fff;
	    text-align: center;
	    white-space: nowrap;
	    vertical-align: baseline;
	    border-radius: .25em;
	    margin-top:5px;
	    margin-bottom:3px;	    
	    <% Response.Write("background:" & FSBoardKioskGlobalColorRedoWaitForParts & ";")%>
	}	


	.labelFollowUp {
	    display: inline-block;
	    padding: .4em .6em .4em;
	    font-size: 100%;
	    font-weight: 400;
	    line-height: 1;
	    color: #fff;
	    text-align: center;
	    white-space: nowrap;
	    vertical-align: baseline;
	    border-radius: .25em;
	    margin-top:5px;
	    margin-bottom:3px;	    
	    <% Response.Write("background:" & FSBoardKioskGlobalColorRedoFollowUp & ";")%>
	}	


	.labelUnableToWork {
	    display: inline-block;
	    padding: .4em .6em .4em;
	    font-size: 100%;
	    font-weight: 400;
	    line-height: 1;
	    color: #fff;
	    text-align: center;
	    white-space: nowrap;
	    vertical-align: baseline;
	    border-radius: .25em;
	    margin-top:5px;
	    margin-bottom:3px;
	    <% Response.Write("background:" & FSBoardKioskGlobalColorRedoUnableToWork & ";")%>
	}	

	.label-default{
	    display: inline-block;
	    padding: .4em .6em .4em;
	    font-size: 100%;
	    font-weight: 400;
	    line-height: 1;
	    color: #fff;
	    text-align: center;
	    white-space: nowrap;
	    vertical-align: baseline;
	    border-radius: .25em;
	    margin-top:5px;
	    margin-bottom:3px;
		background-color: #777;
	}	
	
	.labelFilterChangeIndicatorAndButtonColor {
	    display: inline-block;
	    padding: .4em .6em .4em;
	    font-size: 100%;
	    font-weight: 400;
	    line-height: 1;
	    color: #000;
	    text-align: center;
	    white-space: nowrap;
	    vertical-align: baseline;
	    border-radius: .25em;
	    margin-top:5px;
	    margin-bottom:3px;
	    <% Response.Write("background:" & FilterChangeIndicatorAndButtonColor & ";")%>
	}

	.labelFilterChangeIndicatorAndButtonColorOverDue {
	    display: inline-block;
	    padding: .4em .6em .4em;
	    font-size: 100%;
	    font-weight: 400;
	    line-height: 1;
	    color: #000;
	    text-align: center;
	    white-space: nowrap;
	    vertical-align: baseline;
	    border-radius: .25em;
	    margin-top:5px;
	    margin-bottom:3px;
	    <% Response.Write("background:#FF0000;")%>
	}


	span.filtercircle {
		<% Response.Write("background:" & FilterChangeIndicatorAndButtonColor & ";")%>
		border-radius: 3em;
		-moz-border-radius: 3em;
		-webkit-border-radius: 3em;
		color: #000;
		display: inline-block;
		font-weight: bold;
		line-height: 1.5em;
		margin-right: 8px;
		text-align: center;
		width: 2em;
		font-size: .9em;
		padding: 0.3em;
	}


	.buttonFilterChangeIndicatorAndButtonColor {
	    <% Response.Write("background:" & FilterChangeIndicatorAndButtonColor & ";")%>
	}
		
	.legendAwaitingDispatch {    
	    <% Response.Write("background:" & FSBoardKioskGlobalColorAwaitingDispatch & ";")%>
	    color: #fff;
	}	

	.legendAwaitingAcknowledgement {	    
	    <% Response.Write("background:" & FSBoardKioskGlobalColorAwaitingAcknowledgement & ";")%>
	    color: #fff;
	}	


	.legendDispatchAcknowledged {
	    <% Response.Write("background:" & FSBoardKioskGlobalColorDispatchAcknowledged & ";")%>
	    color: #fff;
	}	


	.legendEnRoute {    
	    <% Response.Write("background:" & FSBoardKioskGlobalColorEnRoute & ";")%>
	    color: #fff;
	}	


	.legendOnSite {   
	    <% Response.Write("background:" & FSBoardKioskGlobalColorOnSite & ";")%>
	    color: #fff;
	}	


	.legendSwap {	   
	    <% Response.Write("background:" & FSBoardKioskGlobalColorRedoSwap & ";")%>
	    color: #fff;
	}	


	.legendWaitForParts {    
	    <% Response.Write("background:" & FSBoardKioskGlobalColorRedoWaitForParts & ";")%>
	    color: #fff;
	}	


	.legendFollowUp {    
	    <% Response.Write("background:" & FSBoardKioskGlobalColorRedoFollowUp & ";")%>
	    color: #fff;
	}	


	.legendUnableToWork {
	    <% Response.Write("background:" & FSBoardKioskGlobalColorRedoUnableToWork & ";")%>
	    color: #fff;
	}	
	

	.service-ticket-note {
		display:block;
		/*float:right;*/
		cursor: pointer;
	}
	
	.service-ticket-note a{
		color:#FFF;
		cursor:pointer;
		
	}
	
	.service-ticket-note a:hover{
		color:#23527c;
		cursor:pointer;	
	}
	
	.ajaxRowView .visibleRowEdit, .ajaxRowEdit .visibleRowView { display: none; }

	[data-tooltip]:hover:before, [data-tooltip]:hover:after {
	  display: block;
	  position: absolute;
	  font-size: 1.1em;
	  color: white;
	}
	[data-tooltip]:hover:before {
	  border-radius: 0.2em;
	  content: attr(title);
	  background-color: rgba(0, 0, 0, 0.9);
	  margin-top: -2.5em;
	  padding: 0.3em;
	}
	[data-tooltip]:hover:after {
	  content: '';
	  margin-top: -1.8em;
	  margin-left: 1em;
	  border-style: solid;
	  border-color: transparent;
	  border-top-color: rgba(0, 0, 0, 0.9);
	  border-width: 0.5em 0.5em 0 0.5em;
	}		
		
</style>


<%
'This is here so we only open it once for the whole page
Set cnn_CheckAlerts = Server.CreateObject("ADODB.Connection")
cnn_CheckAlerts.open (Session("ClientCnnString"))
Set rs_CheckAlerts = Server.CreateObject("ADODB.Recordset")
rs_CheckAlerts.CursorLocation = 3 
SQL_CheckAlerts = "SELECT * FROM Settings_EmailService "
Set rs_CheckAlerts = cnn_CheckAlerts.Execute(SQL_CheckAlerts)
If not rs_CheckAlerts.EOF Then
	RealtimeAlertsOn = rs_CheckAlerts("RealtimeAlertsOn")
	SendAlertToServiceManagers = rs_CheckAlerts("SendAlertToServiceManagers")
	SendAlertToAdditionalEmails = rs_CheckAlerts("SendAlertToAdditionalEmails")
	SendAlertHours = rs_CheckAlerts("SendAlertHours")
	SendAlertsSkipDispatched = rs_CheckAlerts("SendAlertsSkipDispatched")
	EscalationAlertsOn = rs_CheckAlerts("EscalationAlertsOn")
	EscalationAlertToEmails = rs_CheckAlerts("EscalationAlertToEmails")
	EscalationAlertHours = rs_CheckAlerts("EscalationAlertHours")
	EscalationAlertsSkipDispatched = rs_CheckAlerts("EscalationAlertsSkipDispatched")
	AlertsDuringBusinessHoursOnly = rs_CheckAlerts("AlertsDuringBizHoursOnly")
	HoldAlertsOn = rs_CheckAlerts("HoldAlertsOn")
	SendHoldAlertToFinanceManagers = rs_CheckAlerts("SendHoldAlertToFinanceManagers")
	SendHoldAlertToAdditionalEmails = rs_CheckAlerts("SendHoldAlertToAdditionalEmails")
	SendHoldAlertHours = rs_CheckAlerts("SendHoldAlertHours")
	EscalationAlertHours = rs_CheckAlerts("EscalationAlertHours")
	HoldEscalationAlertsOn = rs_CheckAlerts("HoldEscalationAlertsOn")
	HoldEscalationAlertToEmails = rs_CheckAlerts("HoldEscalationAlertToEmails")
	HoldEscalationAlertHours = rs_CheckAlerts("HoldEscalationAlertHours")
Else
	RealtimeAlertsOn = vbFalse
	EscalationAlertsOn = vbFalse
	HoldAlertsOn = vbFalse
	HoldEscalationAlertsOn = vbFalse
End If
Set rs_CheckAlerts = Nothing
cnn_CheckAlerts.Close
Set cnn_CheckAlerts = Nothing

Session("MemoNumber") = ""
Session("ServiceCustID") = ""

%>

<!-- on/off scripts !-->

<style type="text/css">
<%	
		
	Response.Write(".high-priority{")
	Response.Write("	background:" & ServicePriorityColor & ";")
	Response.Write("}")

	Response.Write(".urgent-priority{")
	Response.Write("	background:" & FSBoardKioskGlobalColorUrgent & ";")
	Response.Write("}")

	Response.Write(".alert-priority{")
	Response.Write("	background:" & ServiceNormalAlertColor & ";")
	Response.Write("}")

	Response.Write(".alert-high-priority{")
	Response.Write("background:" & ServicePriorityAlertColor & ";")
	Response.Write("}")

%> 

</style>

<%
	Response.Write("<div id=""PleaseWaitPanel"">")
	Response.Write("<br><br>Loading Today's Service Ticket Status Screen<br><br>This may take up to a full minute, please wait...<br><br>")
	Response.Write("<img src=""" &  baseURL & "/img/loading.gif"" />")
	Response.Write("</div>")
	Response.Flush()
%>

 
<div class="row rowtext heading-legend">
	
	<div class="col-lg-3">
		<h1 class="page-header"><i class="fa fa-wrench"></i> Service Tickets</h1>
	</div>
	
	<!-- accordion line starts here !-->
	<div class="col-lg-9">
	<div class="panel-group" id="accordion" role="tablist" aria-multiselectable="true">
  <div class="panel panel-default">
    <div class="panel-heading" role="tab" id="headingOne">
      <h4 class="panel-title">
        <a role="button" data-toggle="collapse" data-parent="#accordion" href="#collapseOne" aria-expanded="false" aria-controls="collapseOne">
         Legend 
        </a>
      </h4>
    </div>
    <div id="collapseOne" class="panel-collapse collapse" role="tabpanel" aria-labelledby="headingOne">
      <div class="panel-body">

	<!-- legends !-->
  
<% If UserIsServiceManager(Session("UserNo")) or UserIsAdmin(Session("UserNo")) Then %>		
  	<div class="col-lg-4 legend-box">
	  	
		<%If ServiceColorsOn Then %>
		  	<!-- Normal Account - Alert Sent!-->
			<div class="row rowtext legend-row">
				<div class="col-lg-2 col-md-2 col-sm-2 col-xs-2 alert-priority">&nbsp;</div>
				<div class="col-lg-10 col-md-10 col-sm-10 col-xs-10"><h6 class="legend-title">Normal Account - Alert Sent</h6></div>
			</div>
			<!-- eof line !-->
			
			<!-- Priority Account!-->
			<div class="row rowtext legend-row">
				<div class="col-lg-2 col-md-2 col-sm-2 col-xs-2 high-priority">&nbsp;</div>
				<div class="col-lg-10 col-md-10 col-sm-10 col-xs-10"><h6 class="legend-title">Priority Account</h6></div>
			</div>
			<!-- eof line !-->
			
			<!-- Priority Account - Alert Sent!-->
			<div class="row rowtext legend-row">
				<div class="col-lg-2 col-md-2 col-sm-2 col-xs-2 alert-high-priority">&nbsp;</div>
				<div class="col-lg-10 col-md-10 col-sm-10 col-xs-10"><h6 class="legend-title">Priority Account - Alert Sent</h6></div>
			</div>
			<!-- eof line !-->
		<%End If%>

		<!--(Always available) Urgent!-->
		<div class="row rowtext legend-row">
			<div class="col-lg-2 col-md-2 col-sm-2 col-xs-2 urgent-priority">&nbsp;</div>
			<div class="col-lg-10 col-md-10 col-sm-10 col-xs-10"><h6 class="legend-title">Urgent</h6></div>
		</div>
		<!-- eof line !-->
    
		<script type="text/javascript">
			function toggleChevron(e) {
	    $(e.target)
	        .prev('.panel-heading')
	        .find("i.indicator")
	        .toggleClass('glyphicon-chevron-down glyphicon-chevron-up');
				}
				$('#accordion').on('hidden.bs.collapse', toggleChevron);
				$('#accordion').on('shown.bs.collapse', toggleChevron);
		</script>	
	</div>	
<%End If%>
 
<% If advancedDispatchIsOn() Then %>
	<!-- legend stages !-->
	<div class="col-lg-8">
		
		<div class="table-info">
		    <div class="table-responsive">
				<table class="table custom-table">
							
							<tbody>
								<tr>
									<td><u>Status</u></td>
									<td>&nbsp;&nbsp;</td>
									<td><u>Stage</u></td>
									<td>&nbsp;&nbsp;</td>
									<td>&nbsp;&nbsp;</td>
									<td>&nbsp;&nbsp;</td>								
								</tr>
	
								<tr>
									<td>&nbsp;&nbsp;</td>
									<td>&nbsp;&nbsp;</td>
									<td>Received</td>
									<td>&nbsp;&nbsp;</td>
									<td>&nbsp;&nbsp;</td>
									<td>&nbsp;&nbsp;</td>								
								</tr>
								<tr>
									<td>HOLD</td>
									<td>&nbsp;&nbsp;</td>
									<td colspan="2">Under review</td>
									<td>&nbsp;&nbsp;</td>
									<td>&nbsp;&nbsp;</td>
									<td>&nbsp;&nbsp;</td>								
								</tr>
								<tr>
									<td>OPEN</td>
									<td>&nbsp;&nbsp;</td>
									<td class="legendAwaitingDispatch">Awaiting Dispatch</td>
									<td class="legendAwaitingAcknowledgement">Awaiting ACK</td>
									<td class="legendDispatchAcknowledged">Dispatch ACK</td>
									<td class="legendEnRoute">En Route</td>
									
								</tr>
								<tr>
									<td>&nbsp;&nbsp;</td>
									<td>&nbsp;&nbsp;</td>
									<td class="legendOnSite">On Site</td>
									<td class="legendSwap">Swap</td>
									<td class="legendWaitForParts">Wait For Parts</td>
									<td class="legendUnableToWork">Unable to work</td>
									
								</tr>
								<tr>
									<td>&nbsp;&nbsp;</td>
									<td>&nbsp;&nbsp;</td>
									<td class="legendFollowUp">Follow Up</td>
									<td>&nbsp;&nbsp;</td>
									<td>&nbsp;&nbsp;</td>
									<td>&nbsp;&nbsp;</td>
								</tr>								
								<tr>
									<td align="left" colspan="3">CLOSE /CANCEL</td>
									<td>&nbsp;&nbsp;</td>
									<td>&nbsp;&nbsp;</td>
									<td>&nbsp;&nbsp;</td>								
								</tr>
							</tbody>	
	  			</table>
			</div>
			       		
			 
	
			</div>
	</div>
	<!-- eof legend stages !-->	
<% End If %>


  </div>
    </div>
  </div>
	  	</div>
	</div>
<!-- accordion line ends here !-->		
		
	</div>
	<!-- eof legends !-->

 <div class="row rowtext">
 		
	<div class="col-lg-4">
		<p>
			<% If userCanCreateNewServiceTicket(Session("UserNo")) = true Then %>
				<button type="button" class="btn btn-success" data-toggle="modal" data-target="#modalCreateNewServiceTicketForClient">New Service Ticket</button>
			<% End If %>
			
			<% If userCanAccessServiceDispatchCenter(Session("UserNo")) = true Then %>
				<a href="dispatchcenter/main.asp">
					<button type="button" class="btn btn-warning">Dispatch Center</button>
				</a>
			<% End If %>
				
			<a href="serviceBoard.asp">
				<button type="button" class="btn btn-danger">Service Board</button>
			</a>		
		</p>
	</div>

	<!-- search box -->
	<div class="col-lg-3">
		 
		<div class="input-group"> <span class="input-group-addon">Search Tickets</span>
		    <input id="filter" type="text" class="form-control " placeholder="Type here...">
		</div>
		<br>
	</div>
	<!-- eof search box -->


	<!-- pause and timer !-->
	<div class="col-lg-2 pause-timer">
 		<div class="pause">
	        Pause Automatic Refresh&nbsp;&nbsp;
            <div class="material-switch">
                <input id="switchAutomaticRefresh" name="chkAutomaticRefresh" type="checkbox"/>
                <label for="switchAutomaticRefresh" class="label-primary"></label>
            </div>
		</div>

		<div id="timer"  style="height:30px;"></div>
	</div>
 	<!-- eof pause and timer !-->
 	
 	
 	<div class="col-lg-3">
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

	
<!-- row !-->
<div class="row rowtext">

 	<div class="col-lg-12" style="margin-left: -15px; border-bottom: 1px solid #ddd;"id="tabsHolder">
 	
		<div class="col-lg-8">
		    <!-- Nav tabs -->
		    
		    <% If ShowSeparateFilterChangesTabOnServiceScreen = 1 Then %>
		    
			    <ul class="nav nav-tabs pull-left" role="tablist" style="float:left !important;">
			    	  <li role="presentation"><a href="#awaitingdispatch" id="#awaitingdispatch" role="tab" data-toggle="tab">Awaiting Dispatch (<%= GetNumberOfServiceCallsAwaitingDispatch() %>)</a></li>
				      <li role="presentation"><a href="#awaitingacknowledgment" id="#awaitingacknowledgment" role="tab" data-toggle="tab">Awaiting Ack (<%= GetNumberOfServiceCallsAwaitingAcknowledgement() %>)</a></li>
				      <li role="presentation"><a href="#acknowledged" id="#acknowledged" role="tab" data-toggle="tab">Acknowledged (<%= GetNumberOfServiceCallsAcknowledged() %>)</a></li>
				      <li role="presentation"><a href="#enrouteonsite" id="#enrouteonsite" role="tab" data-toggle="tab">En Route/On Site (<%= GetNumberOfServiceCallsEnRouteOnSite() %>)</a></li>
				      <li role="presentation"><a href="#redo" role="tab" id="#redo" data-toggle="tab">REDO (<%= GetNumberOfServiceCallsRedo() %>)</a></li>
				      
				      	<%
				      	
							'******************************************************************************
							'Obtain Dates of Rolling Last Five Work Days
							'******************************************************************************
	
							'*******************************************************************************
							'Obtain the first working day to start counting 5 days backwards from
							'Start with today. If today is a weekend or a closed company holiday, 
							'then subtract one day and repeat the check. Once we find a valid workday,
							'stop looping and start counting back 5 valid business days from the
							'starting date
							'*******************************************************************************
							
							firstDateOfPastFiveDaysFound = False
							firstDateOfPastFiveDays = Now()
							
							Do While firstDateOfPastFiveDaysFound = False
							
								SQL = "SELECT * FROM Settings_CompanyCalendar WHERE MonthNum='" & Month(firstDateOfPastFiveDays) & "' AND DayNum ='" & Day(firstDateOfPastFiveDays) & "' AND YearNum='" & Year(firstDateOfPastFiveDays) & "'"
							
								Set cnn9 = Server.CreateObject("ADODB.Connection")
								cnn9.open (Session("ClientCnnString"))
								Set rs9 = Server.CreateObject("ADODB.Recordset")
								rs9.CursorLocation = 3 
								Set rs9 = cnn9.Execute(SQL)
									
								If not rs9.EOF Then
									Select Case rs9("OpenClosedCloseEarly")
									Case "Closed"
										firstDateOfPastFiveDaysFound = False
									Case "Close Early"
										firstDateOfPastFiveDaysFound = True
									Case Else
										firstDateOfPastFiveDaysFound = True
									End Select
								Else
									firstDateOfPastFiveDaysFound = True
								End If
											
								'******************************************************
								'Make sure that that the date is also not a weekend
								'If it is, set the control variable to false
								'And go back one calendar day to test as the next day
								'******************************************************
								If firstDateOfPastFiveDaysFound = True Then			
									If Weekday(firstDateOfPastFiveDays,vbMonday) <= 5 Then
										firstDateOfPastFiveDaysFound = True
									Else
										firstDateOfPastFiveDaysFound = False
										firstDateOfPastFiveDays = DateAdd("d", -1, firstDateOfPastFiveDays)
									End If
								Else
									firstDateOfPastFiveDays = DateAdd("d", -1, firstDateOfPastFiveDays)
								End If	
							Loop
							
							
							
							'*************************************************************
							'At this point we have found the ending business day
							'Now we need to loop through the days prior to this day and
							'come up with the last 5 business days starting date
							'*************************************************************
							
							'Response.Write("firstDateOfPastFiveDays : " & firstDateOfPastFiveDays & "<br>")
							lastDateOfPastFiveDays = DateAdd("d", -1, firstDateOfPastFiveDays)
							validDaysGoneBackSoFar = 1
							
							Do While validDaysGoneBackSoFar < 4
							
								lastDateOfPastFiveDaysFound = False
							
								SQL = "SELECT * FROM Settings_CompanyCalendar WHERE MonthNum='" & Month(lastDateOfPastFiveDays) & "' AND DayNum ='" & Day(lastDateOfPastFiveDays) & "' AND YearNum='" & Year(lastDateOfPastFiveDays) & "'"
							
								Set cnn9 = Server.CreateObject("ADODB.Connection")
								cnn9.open (Session("ClientCnnString"))
								Set rs9 = Server.CreateObject("ADODB.Recordset")
								rs9.CursorLocation = 3 
								Set rs9 = cnn9.Execute(SQL)
									
								If not rs9.EOF Then
									Select Case rs9("OpenClosedCloseEarly")
									Case "Closed"
										lastDateOfPastFiveDaysFound = False
									Case "Close Early"
										lastDateOfPastFiveDaysFound = True
									Case Else
										lastDateOfPastFiveDaysFound = True
									End Select
								Else
									lastDateOfPastFiveDaysFound = True
								End If
											
								'******************************************************
								'Make sure that that the date is also not a weekend
								'If it is, set the control variable to false
								'And go back one calendar day to test as the next day
								'******************************************************
						
								If lastDateOfPastFiveDaysFound = True Then				
									If Weekday(lastDateOfPastFiveDays,vbMonday) <= 5 Then
										lastDateOfPastFiveDaysFound = True
										validDaysGoneBackSoFar = validDaysGoneBackSoFar + 1
										lastDateOfPastFiveDays = DateAdd("d", -1, lastDateOfPastFiveDays)
									Else
										lastDateOfPastFiveDaysFound = False
										lastDateOfPastFiveDays = DateAdd("d", -1, lastDateOfPastFiveDays)
									End If
								Else
									lastDateOfPastFiveDays = DateAdd("d", -1, lastDateOfPastFiveDays)
								End If	
	
								'Response.Write("lastDateOfPastFiveDays : " & lastDateOfPastFiveDays & "<br>")
								'Response.Write("validDaysGoneBackSoFar : " & validDaysGoneBackSoFar & "<br>")
							Loop
						
						
							set rs9 = Nothing
							cnn9.close
							set cnn9 = Nothing
							
							
							lastDateOfPastFiveDaysDisplay = padDate(MONTH(lastDateOfPastFiveDays),2) & "/" & padDate(DAY(lastDateOfPastFiveDays),2) & "/" & padDate(RIGHT(YEAR(lastDateOfPastFiveDays),2),2)
							firstDateOfPastFiveDaysDisplay = padDate(MONTH(firstDateOfPastFiveDays),2) & "/" & padDate(DAY(firstDateOfPastFiveDays),2) & "/" & padDate(RIGHT(YEAR(firstDateOfPastFiveDays),2),2)
							
							'******************************************************************************
				      	
				      	%>
				      <li role="presentation"><a href="#closed" id="#closed" role="tab" data-toggle="tab">Closed&nbsp;<%= lastDateOfPastFiveDaysDisplay %>-<%= firstDateOfPastFiveDaysDisplay %>&nbsp;(<%= GetNumberOfServiceCallsClosedRolling5Days() %>)</a></li>
				      <% If filterChangeModuleOn() Then %>
				      	<li role="presentation"><a href="#filters" role="tab" id="#filters" data-toggle="tab">Filter Changes (<%= GetNumberOfServiceCallsFilterChanges() %>)</a></li>
				      	<li role="presentation"><a href="#filtersredo" role="tab" id="#filtersredo" data-toggle="tab">Filter Redo (<%= GetNumberOfServiceCallsRedoFiltersOnly() %>)</a></li>
				      <% End If %>
				      <% If ServiceTicketScreenShowHoldTab = 1 Then %>
					      <li role="presentation"><a href="#onhold" role="tab" id="#onhold" data-toggle="tab">On Hold (<%= GetNumberOfServiceCallsOnHold() %>)</a></li>
				      <% End If %>
			    </ul>
			    
			<% Else %>
			
			
			
			    <ul class="nav nav-tabs pull-left" role="tablist" style="float:left !important;">
			    	  <li role="presentation"><a href="#awaitingdispatch" id="#awaitingdispatch" role="tab" data-toggle="tab">Awaiting Dispatch (<%= GetNumberOfServiceCallsAwaitingDispatchWithFilters() %>)</a></li>
				      <li role="presentation"><a href="#awaitingacknowledgment" id="#awaitingacknowledgment" role="tab" data-toggle="tab">Awaiting Ack (<%= GetNumberOfServiceCallsAwaitingAcknowledgementWithFilters() %>)</a></li>
				      <li role="presentation"><a href="#acknowledged" id="#acknowledged" role="tab" data-toggle="tab">Acknowledged (<%= GetNumberOfServiceCallsAcknowledgedWithFilters() %>)</a></li>
				      <li role="presentation"><a href="#enrouteonsite" id="#enrouteonsite" role="tab" data-toggle="tab">En Route/On Site (<%= GetNumberOfServiceCallsEnRouteOnSiteWithFilters() %>)</a></li>
				      <li role="presentation"><a href="#redo" role="tab" id="#redo" data-toggle="tab">REDO (<%= GetNumberOfServiceCallsRedoWithFilters() %>)</a></li>
				      
				      	<%
				      	
							'******************************************************************************
							'Obtain Dates of Rolling Last Five Work Days
							'******************************************************************************
	
							'*******************************************************************************
							'Obtain the first working day to start counting 5 days backwards from
							'Start with today. If today is a weekend or a closed company holiday, 
							'then subtract one day and repeat the check. Once we find a valid workday,
							'stop looping and start counting back 5 valid business days from the
							'starting date
							'*******************************************************************************
							
							firstDateOfPastFiveDaysFound = False
							firstDateOfPastFiveDays = Now()
							
							Do While firstDateOfPastFiveDaysFound = False
							
								SQL = "SELECT * FROM Settings_CompanyCalendar WHERE MonthNum='" & Month(firstDateOfPastFiveDays) & "' AND DayNum ='" & Day(firstDateOfPastFiveDays) & "' AND YearNum='" & Year(firstDateOfPastFiveDays) & "'"
							
								Set cnn9 = Server.CreateObject("ADODB.Connection")
								cnn9.open (Session("ClientCnnString"))
								Set rs9 = Server.CreateObject("ADODB.Recordset")
								rs9.CursorLocation = 3 
								Set rs9 = cnn9.Execute(SQL)
									
								If not rs9.EOF Then
									Select Case rs9("OpenClosedCloseEarly")
									Case "Closed"
										firstDateOfPastFiveDaysFound = False
									Case "Close Early"
										firstDateOfPastFiveDaysFound = True
									Case Else
										firstDateOfPastFiveDaysFound = True
									End Select
								Else
									firstDateOfPastFiveDaysFound = True
								End If
											
								'******************************************************
								'Make sure that that the date is also not a weekend
								'If it is, set the control variable to false
								'And go back one calendar day to test as the next day
								'******************************************************
								If firstDateOfPastFiveDaysFound = True Then			
									If Weekday(firstDateOfPastFiveDays,vbMonday) <= 5 Then
										firstDateOfPastFiveDaysFound = True
									Else
										firstDateOfPastFiveDaysFound = False
										firstDateOfPastFiveDays = DateAdd("d", -1, firstDateOfPastFiveDays)
									End If
								Else
									firstDateOfPastFiveDays = DateAdd("d", -1, firstDateOfPastFiveDays)
								End If	
							Loop
							
							
							
							'*************************************************************
							'At this point we have found the ending business day
							'Now we need to loop through the days prior to this day and
							'come up with the last 5 business days starting date
							'*************************************************************
							
							'Response.Write("firstDateOfPastFiveDays : " & firstDateOfPastFiveDays & "<br>")
							lastDateOfPastFiveDays = DateAdd("d", -1, firstDateOfPastFiveDays)
							validDaysGoneBackSoFar = 1
							
							Do While validDaysGoneBackSoFar < 4
							
								lastDateOfPastFiveDaysFound = False
							
								SQL = "SELECT * FROM Settings_CompanyCalendar WHERE MonthNum='" & Month(lastDateOfPastFiveDays) & "' AND DayNum ='" & Day(lastDateOfPastFiveDays) & "' AND YearNum='" & Year(lastDateOfPastFiveDays) & "'"
							
								Set cnn9 = Server.CreateObject("ADODB.Connection")
								cnn9.open (Session("ClientCnnString"))
								Set rs9 = Server.CreateObject("ADODB.Recordset")
								rs9.CursorLocation = 3 
								Set rs9 = cnn9.Execute(SQL)
									
								If not rs9.EOF Then
									Select Case rs9("OpenClosedCloseEarly")
									Case "Closed"
										lastDateOfPastFiveDaysFound = False
									Case "Close Early"
										lastDateOfPastFiveDaysFound = True
									Case Else
										lastDateOfPastFiveDaysFound = True
									End Select
								Else
									lastDateOfPastFiveDaysFound = True
								End If
											
								'******************************************************
								'Make sure that that the date is also not a weekend
								'If it is, set the control variable to false
								'And go back one calendar day to test as the next day
								'******************************************************
						
								If lastDateOfPastFiveDaysFound = True Then				
									If Weekday(lastDateOfPastFiveDays,vbMonday) <= 5 Then
										lastDateOfPastFiveDaysFound = True
										validDaysGoneBackSoFar = validDaysGoneBackSoFar + 1
										lastDateOfPastFiveDays = DateAdd("d", -1, lastDateOfPastFiveDays)
									Else
										lastDateOfPastFiveDaysFound = False
										lastDateOfPastFiveDays = DateAdd("d", -1, lastDateOfPastFiveDays)
									End If
								Else
									lastDateOfPastFiveDays = DateAdd("d", -1, lastDateOfPastFiveDays)
								End If	
	
								'Response.Write("lastDateOfPastFiveDays : " & lastDateOfPastFiveDays & "<br>")
								'Response.Write("validDaysGoneBackSoFar : " & validDaysGoneBackSoFar & "<br>")
							Loop
						
						
							set rs9 = Nothing
							cnn9.close
							set cnn9 = Nothing
							
							
							lastDateOfPastFiveDaysDisplay = padDate(MONTH(lastDateOfPastFiveDays),2) & "/" & padDate(DAY(lastDateOfPastFiveDays),2) & "/" & padDate(RIGHT(YEAR(lastDateOfPastFiveDays),2),2)
							firstDateOfPastFiveDaysDisplay = padDate(MONTH(firstDateOfPastFiveDays),2) & "/" & padDate(DAY(firstDateOfPastFiveDays),2) & "/" & padDate(RIGHT(YEAR(firstDateOfPastFiveDays),2),2)
							
							'******************************************************************************
				      	
				      	%>
				      <li role="presentation"><a href="#closed" id="#closed" role="tab" data-toggle="tab">Closed&nbsp;<%= lastDateOfPastFiveDaysDisplay %>-<%= firstDateOfPastFiveDaysDisplay %>&nbsp;(<%= GetNumberOfServiceCallsClosedRolling5Days() %>)</a></li>
   				      <% If ServiceTicketScreenShowHoldTab = 1 Then %>
					      <li role="presentation"><a href="#onhold" role="tab" id="#onhold" data-toggle="tab">On Hold (<%= GetNumberOfServiceCallsOnHold() %>)</a></li>
				      <% End If %>

			    </ul>


			<% End If %>
						
		</div>
		<div class="col-lg-4">
		    <!-- Nav tabs -->
		    <ul class="nav nav-tabs pull-right" role="tablist"  style="float:right !important;">
			      <li role="presentation"><a href="#0to8" id="#0to8" role="tab" data-toggle="tab">Day 1 (<%= GetNumberOfServiceTicketsInTimeRange(-999, NumberOfMinutesInServiceDayVar) %>)</a></li>
			      <li role="presentation"><a href="#8to16" id="#8to16" role="tab" data-toggle="tab">Day 2 (<%= GetNumberOfServiceTicketsInTimeRange(NumberOfMinutesInServiceDayVar+1, NumberOfMinutesInServiceDayVar*2) %>)</a></li>
			      <li role="presentation"><a href="#16to40" id="#16to40" role="tab" data-toggle="tab">3-5 Days (<%= GetNumberOfServiceTicketsInTimeRange((NumberOfMinutesInServiceDayVar*2)+1, NumberOfMinutesInServiceDayVar*5) %>)</a></li>
   			      <li role="presentation"><a href="#over40" id="#over40" role="tab" data-toggle="tab">Over 5 Days (<%= GetNumberOfServiceTicketsInTimeRange((NumberOfMinutesInServiceDayVar*5)+1, 99999) %>)</a></li> 
			      <li role="presentation"><a href="#all" id="#all" role="tab" data-toggle="tab">ALL (<%= GetNumberOfServiceTicketsInTimeRange(-999, 99999) %>)</a></li>
		    </ul>
		</div> 	

</div>
</div>

<!-- row !-->
<div class="row rowtext">
 	<div class="col-lg-12">
 	
	    <!-- Tab panes -->
	    <div class="tab-content">
	    
			<% ' Leave this here, the include files below will use it
	    	Set cnnCustInfo = Server.CreateObject("ADODB.Connection")
			cnnCustInfo.open Session("ClientCnnString")
			Set rsCustInfo  = Server.CreateObject("ADODB.Recordset") 
			Set cnn8 = Server.CreateObject("ADODB.Connection")
			cnn8.open (Session("ClientCnnString"))
			Set rs = Server.CreateObject("ADODB.Recordset")
			 %>

		      <div class="tab-pane fade active in" id="awaitingdispatch" data-source="maintabs/mainTabAwaitingDispatch.asp"></div>
		      
		      <div class="tab-pane fade in" id="awaitingacknowledgment" data-source="maintabs/mainTabAwaitingAcknowledgment.asp"></div>
		      
		      <div class="tab-pane fade in" id="acknowledged" data-source="maintabs/mainTabAcknowledged.asp"></div>
		      
		      <div class="tab-pane fade in" id="enrouteonsite" data-source="maintabs/mainTabEnRouteOnSite.asp"></div>

		      <div class="tab-pane fade in" id="redo" data-source="maintabs/mainTabRedo.asp"></div>
		      
		      <div class="tab-pane fade in" id="closed" data-source="maintabs/mainTabClosed.asp"></div>
		      
			  <% If filterChangeModuleOn() Then %>
			      <div class="tab-pane fade in" id="filters" data-source="maintabs/mainTabFilters.asp"></div>
			      <div class="tab-pane fade in" id="filtersredo" data-source="maintabs/mainTabFiltersRedo.asp"></div>
		      <% End If %>
		      
   			  <% If ServiceTicketScreenShowHoldTab = 1 Then %>
			      <div class="tab-pane fade in" id="onhold" data-source="maintabs/mainTabOnHold.asp"></div>
		      <% End If %>

		      
		      <div class="tab-pane fade in" id="0to8" data-source="maintabs/mainTabRange1.asp"></div>
		      
		      <div class="tab-pane fade in" id="8to16" data-source="maintabs/mainTabRange2.asp"></div>
		      
		      <div class="tab-pane fade in" id="16to40" data-source="maintabs/mainTabRange3.asp"></div>
		      
		      <div class="tab-pane fade in" id="over40" data-source="maintabs/mainTabRange4.asp"></div>
		      
   		      <div class="tab-pane fade in" id="all" data-source="maintabs/mainTabAll.asp"></div>
		      
   			<% ' Leave this here, the include files below will use it
				set rsCustInfo = Nothing
				cnnCustInfo.Close	
				set cnnCustInfo = Nothing

				set rs = Nothing
				cnn8.close
				set cnn8 = Nothing

			%>
	    </div>
	    

    </div>	
    

</div>
<!-- eof row !-->  



<!-- countdown script !-->
<script src="<%= BaseURL %>js/countdown/jquery.pietimer.js"></script>
<!-- eof countdown script !-->






<!-- **************************************************************************************************************************** -->
<!-- MODALS FOR DELIVERY ALERTS BEGIN HERE !-->
<!-- **************************************************************************************************************************** -->

<!--#include file="serviceBoardCommonModals.asp"-->

<!-- modal window add symptoms -->
<!--#include file="onthefly_SymptomCodes.asp"--> 
<!-- end modal window add symptoms -->

<!-- modal window add problems -->
<!--#include file="onthefly_ProblemCodes.asp"--> 
<!-- end modal window add problems -->

<!-- modal window add resolutions -->
<!--#include file="onthefly_ResolutionCodes.asp"--> 
<!-- end modal window add resolutions -->

<!-- Equipment Modal -->
<div class="modal modal-xwide fade" id="modalEquipmentVPC" tabindex="-1" role="dialog" aria-labelledby="modalEquipmentVPCLabel">

	<div class="modal-dialog" role="document">
	
	<style>
		.modal-header {
		    padding: 15px;
		    border-bottom: 1px solid #e5e5e5;
		    /*min-height: 135px;*/
		}
		
		.tile {
		  width: 6%;
		  display: inline-block;
		  box-sizing: border-box;
		  background: #fff;
		  padding-top: 10px;
		  padding-top: 10px;
		  padding-left: 5px;
		  color:#FFF;
		}
		
		.tile .title {
		  margin-top: 0px;
		}
		
		.tile.red {
		  background: #AC193D;
		  color:#FFF;
		}
		
		.tile.red:hover {
		  background: #7f132d;
		  color:#FFF;
		}
		
		.tile.blue {
		  background: #2672EC;
		  color:#FFF;
		}
		
		.tile.blue:hover {
		  background: #125acd;
		  color:#FFF;
		}
		
		.equip_qty{
			font-weight:bold;
			font-size:1.2em;
	
		}
		.d-none {display:none;}
	</style>
						
		<div class="modal-content">	
		
			<input type="hidden" name="txtCustIDToPass" id="txtCustIDToPass">
			<input type="hidden" name="txtLastClosedPeriodGP" id="txtLastClosedPeriodGP">
			
			<div class="modal-header">
				<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
				<h4 class="modal-title" id="modalEquipmentVPCTitle"><!-- Content generated by Sub GetTitleForEquipmentVPCModal() in InsightFuncs_AjaxForBizIntelModals.asp --></h4>
			</div>
			
			<div class="modal-body modalResponsiveTable">
			
				<div id="PleaseWaitPanelModal">
				<br><br><strong>Analyzing Equipment, please wait...</strong><br><br>
				<img src="<%= BaseURL %>img/loading.gif">
				</div>	
			
				<div id="modalCategoryVPCContent">
				<!-- Content for the modal will be generated and written here -->
				<!-- Content generated by Sub GetContentForEquipmentVPCModal() in InsightFuncs_AjaxForBizIntelModals.asp -->
				</div>
			</div>
			
			<div class="modal-footer">
				<button type="button" class="btn btn-default" data-dismiss="modal">Close Window</button>
			</div>

		</div>
		<!-- eof modal content !-->
	</div>
	<!-- eof modal dialog !-->
</div>
<!-- eof modal !-->


<!-- **************************************************************************************************************************** -->
<!-- MODALS FOR DELIVERY ALERTS END HERE !-->
<!-- **************************************************************************************************************************** -->

<div class="waitdiv d-none" style="position:fixed;z-index: 999999999; top: 0px; left: 0px; width: 100%; height:80%; background-color:transparent; text-align: center; padding-top: 20%; filter: alpha(opacity=0); opacity:0; "></div>
    <div id="waitdiv" class="waitdiv d-none small" style="padding-bottom: 90px;text-align: center; vertical-align:middle;padding-top:50px;background-color:#ebebeb;width:300px;height:100px;margin: 0 auto; top:40%; left:40%;position:absolute;-webkit-box-shadow: 0 5px 10px rgba(0, 0, 0, 0.2); -moz-box-shadow: 0 5px 10px rgba(0, 0, 0, 0.2); box-shadow: 0 5px 10px rgba(0, 0, 0, 0.2); z-index:999999999;">
        <img src="/img/loader.gif" alt="" /><br />Request data from server. <br /> Please wait...
    </div>
<!--#include file="../inc/footer-main.asp"-->
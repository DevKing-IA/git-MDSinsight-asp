<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/jquery_table_search.asp"-->
<!--#include file="../../inc/mail.asp"-->
<!--#include file="../../inc/InsightFuncs_Service.asp"-->
<!--#include file="../../inc/InsightFuncs_AR_AP.asp"-->
<script src="<%= baseURL %>js/moment.min.js" type="text/javascript"></script>
<script type="text/javascript" src="<%= BaseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.min.js"></script>
<link href="<%= BaseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.css" rel="stylesheet">

<link rel="stylesheet" href="https://cdn.datatables.net/1.10.20/css/jquery.dataTables.min.css" />
<script type="text/javascript" src="https://cdn.datatables.net/1.10.20/js/jquery.dataTables.min.js"></script>

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
'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)
%>

								
<SCRIPT LANGUAGE="JavaScript">
<!--
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
 
// -->
</SCRIPT> 


<script>
  function myFunction(num)
	  {   

		  var  memnum=num;
				
		   if(num!='')
		   {
		    $.ajax({
		   type:'post',
		      url:'toggleUrgent.asp',
		          data:{memnum: memnum},
					success: function(msg){
						window.location = "main.asp";
					}
		 });
		  }
	}
</script>

<!-- DYNAMIC FORM !-->
<style type="text/css">
	 .ativa-scroll{
	 max-height: 300px
 }
</style>

<style type="text/css">


span.servicecircle {
  	background: #183E5B;
	border-radius: 3em;
	-moz-border-radius: 3em;
	-webkit-border-radius: 3em;
	color: #fff;
	display: inline-block;
	font-weight: bold;
	line-height: 1.5em;
	margin-right: 8px;
	text-align: center;
	width: 2em;
	font-size: .9em;
	padding: 0.3em;
}

span.filtercircle {
  	background: #FECF00;
	border-radius: 3em;
	-moz-border-radius: 3em;
	-webkit-border-radius: 3em;
	color: #fff;
	display: inline-block;
	font-weight: bold;
	line-height: 1.5em;
	margin-right: 8px;
	text-align: center;
	width: 2em;
	font-size: .9em;
	padding: 0.3em;
}

mark {
    background-color: yellow;
    color: black;
}

.pause-timer{
	margin:5px 0px 0px -80px;
	
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

<%	
SQL = "SELECT * FROM Settings_Global"
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
		
		Response.Write(".high-priority{")
		Response.Write("	background:" & rs("ServicePriorityColor") & ";")
		Response.Write("}")

		Response.Write(".urgent-priority{")
		Response.Write("	background:" & rs("FSBoardUrgentColor") & ";")
		Response.Write("}")

		Response.Write(".alert-priority{")
		Response.Write("	background:" & rs("ServiceNormalAlertColor") & ";")
		Response.Write("}")

		Response.Write(".alert-high-priority{")
		Response.Write("background:" & rs("ServicePriorityAlertColor") & ";")
		Response.Write("}")
		
		If rs("ServiceColorsOn") = 1 Then ServiceColorsOn = True Else ServiceColorsOn = False
		
		DelBoardPieTimerColor = rs("DelBoardPieTimerColor")
		If DelBoardPieTimerColor = "" Then DelBoardPieTimerColor = "000000"
		If IsNull(DelBoardPieTimerColor ) Then DelBoardPieTimerColor = "000000"
		Session("DelBoardPieTimerColor") = Replace(DelBoardPieTimerColor,"#","") ' Just this one for Javascript


End If
set rs = Nothing
cnn8.close
set cnn8 = Nothing
%> 
	
</style>

<!-- modal scroll !-->
<script type="text/javascript">
  $(document).ready(ajustamodal);
  $(window).resize(ajustamodal);
  function ajustamodal() {
    var altura = $(window).height() - 155; //value corresponding to the modal heading + footer
    $(".ativa-scroll").css({"height":altura,"overflow-y":"auto"});
  }
  
  $(document).ready(function() {
  
  	   	$('#modalEditExistingServiceTicketForClient').on('show.bs.modal', function (e) {

		    //get data-id attribute of the clicked service ticket
		    var passedMemoNumber = $(e.relatedTarget).data('memo-number');
		    var passedCustID = $(e.relatedTarget).data('cust-id');
		    
		    //populate the textboxes with the id of the clicked service ticket
		    $(e.currentTarget).find('input[name="txtMemoNumberCloseCancel"]').val(passedMemoNumber);
			$(e.currentTarget).find('input[name="txtCustIDCloseCancel"]').val(passedCustID);
			$(e.currentTarget).find('input[name="txtReturnPathCloseCancel"]').val("DispatchCenter");
			

 			//alert("passedMemoNumber: " + passedMemoNumber);		
 			    	    
		    var $modal = $(this);
	
	    	$.ajax({
				type:"POST",
				url: "../../inc/InSightFuncs_AjaxForServiceModals.asp",
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
	    
	    
	    
	    $('#modalEditExistingServiceTicketForClient').on('hidden.bs.modal', function () {
			//if (Cookies.get('service-board-pause-autorefresh') == 'false' ){
				$('#switchAutomaticRefresh').prop('checked', false).trigger("change");
			//}
	    });
	    
	});


</script>
<!-- eof modal scroll !-->
<!-- END DYNAMIC FORM !-->


<%If Request.QueryString("ses")="" Then Session("RefreshOptions")= "" ' just blank it out if we didnt come from there
 %>

<%

Session("MemoNumber") = ""
Session("ServiceCustID") = ""
ViewType="Condensed"

WHERE_CLAUSE=""
	
%>

<!-- on/off scripts !-->

 
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
 
.btn-link{
	padding: 0px;
	text-align: left;
}

.date-time-hidden-value{
	display:none;
}

.row{
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
 }

.heading-legend h1{
	margin:0px;
}

.custom-table{
 	font-size: 11px;
}

.btn-dispatch{
	font-size: 11px;
	padding: 5px;
}

.scrollable-table{
	height: 200px;
	overflow-y: auto;
	border: 1px solid #ccc;
	padding: 10px;
	font-size: 9px;
}

.row-line > div{
	margin-bottom: 25px;
}

.scrollable-title{
	border: 1px solid #ccc;
	padding: 10px;
	margin-bottom: -1px;
	background: #DCE6E9;
	font-size: 12px;
}

.scrollable-title-awaiting-dispatch {
    border: 1px solid #ccc;
    padding: 10px;
    margin-bottom: -1px;
    background: #FDF478;
    font-size: 12px;
}
 
 .tooltip-button{
	 padding: 0px;
	 border: 0px;
	 background: transparent;
	 font-size: 9px;
	 vertical-align: top;
 }
 
  .tooltip-button:hover{
	  background: transparent;
  }
  
  h1{
	  margin-top:0px;
  }
 .ui-state-highlight.item{height: 238.15px;}
 /*==================================  */
.recent-comments {
    margin-bottom: 0;
	
}

table.food_planner {margin-bottom: 5px;}
.list-group {
    margin-bottom: 20px;
    padding-left: 0;
	min-height:60%;
}
.recent-comments li.comment-success:hover {
    border-left-color: #8ea83f;
}
.list-group .list-group-item:hover {
    background-color: #f3f3f3;
}

.recent-comments li.comment-success {
    border-left-color: #b1c86b;
}
.recent-comments li:hover {
    background-repeat: repeat-x;
	 border-color: #848484;
	 background-image: linear-gradient(to bottom, #f3f3f3 0%, #ffffff 100%);
 }
.list-group .list-group-item {
   cursor:pointer;
}
.list-group-item:first-child {
    border-top-right-radius: 2px;
    border-top-left-radius: 2px;
}
.recent-comments li {
    border-radius: 2px;
    border-left-width: 4px;
    margin-bottom: 2px;
     background-repeat: repeat-x;
	 background-image: linear-gradient(to bottom, #ffffff 0%, #f3f3f3 100%);
	 border-left-color: #c4c4c4;
	 border-right-color: #c4c4c4;
	 border-top-color: #c4c4c4;
	 border-bottom-color: #c4c4c4;
 }
.recent-comments li a, .recent-comments li, .customer-data {
font-size: 9px;
}
.list-group-item {
    position: relative;
    display: block;
    padding: 2px;
    margin-bottom: -1px;
    background-color: #ffffff;
    border: 1px solid #c4c4c4;
}
.pl-1 {padding-left:1px;}
.pr-1 {padding-right:1px;}
.pt-5 {padding-top:5px;}
.ticket-number {text-align:center;}
.tickets-list-header {margin-bottom:1px;}
.tickets-list-header div {font-weight:bold;font-size: 9px; text-align:center;}
<%	
SQL = "SELECT * FROM Settings_Global"
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
		
		Response.Write(".high-priority{")
		Response.Write("	background:" & rs("ServicePriorityColor") & ";")
		Response.Write("}")

		Response.Write(".urgent-priority{")
		Response.Write("	background:" & rs("FSBoardUrgentColor") & ";")
		Response.Write("}")

		Response.Write(".alert-priority{")
		Response.Write("	background:" & rs("ServiceNormalAlertColor") & ";")
		Response.Write("}")

		Response.Write(".alert-high-priority{")
		Response.Write("background:" & rs("ServicePriorityAlertColor") & ";")
		Response.Write("}")

End If
set rs = Nothing
cnn8.close
set cnn8 = Nothing
%> 
 
 </style>

<!--- eof on/off scripts !-->


 
<div class="row heading-legend">
	
	<!-- heading !-->
	<div class="col-lg-5">
 		<h1 class="page-header">
 		
	 		<i class="fa fa-briefcase"></i> Dispatch Center 
	
			<a href="../main.asp">
				<button type="button" class="btn btn-warning">Service Tickets</button>
			</a>
			<a href="../serviceBoard.asp">
				<button type="button" class="btn btn-danger">Service Board</button>
			</a>			
		
		</h1>
	</div>
	<!-- eof heading !-->
	
	<!-- pause and timer !-->
	<div class="col-lg-7 pause-timer">
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
	
</div>
<div class='row row-line' id='sortableList'>

<%

dim fs,t,userorder
userorder=0
set fs=Server.CreateObject("Scripting.FileSystemObject")
filename = Server.MapPath(".")&"\userorder\"&Session("Userno")&".txt"
if fs.FileExists(filename) then
	set t=fs.OpenTextFile(filename,1,false)
	userorder=t.ReadLine
	t.close
end if

Dim userorder_arr
userorder_arr=split(userorder,",")

GridColumn = 1
DynamicFormCounter = 1


Set cnnAwaitingDispatch = Server.CreateObject("ADODB.Connection")
cnnAwaitingDispatch.open(Session("ClientCnnString"))
Set rsAwaitingDispatch= Server.CreateObject("ADODB.Recordset")
rsAwaitingDispatch.CursorLocation = 3 

SQLAwaitingDispatch = "SELECT * FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN' AND RecordSubType = 'OPEN' "
SQLAwaitingDispatch = SQLAwaitingDispatch & "ORDER BY submissionDateTime DESC"


Set rsAwaitingDispatch = cnnAwaitingDispatch.Execute(SQLAwaitingDispatch)

If not rsAwaitingDispatch.EOF Then


	TicketCount = 0
	FilterCount = 0
	
	Set cnn_Tickets = Server.CreateObject("ADODB.Connection")
	cnn_Tickets.open (Session("ClientCnnString"))
	Set rs_Tickets = Server.CreateObject("ADODB.Recordset")

	
	'Count all the tickets for the header
	SQL_Tickets = "SELECT DISTINCT MemoNumber FROM FS_ServiceMemosDetail "
	SQL_Tickets = SQL_Tickets & "WHERE (MemoNumber IN "
    SQL_Tickets = SQL_Tickets & "(SELECT MemoNumber FROM FS_ServiceMemos WHERE "
    SQL_Tickets = SQL_Tickets & "(CurrentStatus = 'OPEN')))"
    
    Set rs_Tickets = cnn_Tickets.Execute(SQL_Tickets)
	If not rs_Tickets.Eof Then
		Do While not rs_Tickets.Eof	
		
			GetServiceTicketCurrentStageVar = GetServiceTicketCurrentStage(rs_Tickets("MemoNumber"))
							
			If (GetServiceTicketCurrentStageVar = "Received" OR GetServiceTicketCurrentStageVar = "Released" OR GetServiceTicketCurrentStageVar = "Declined") Then

							
				If TicketIsFilterChange(rs_Tickets("MemoNumber")) Then
					 FilterCount = FilterCount + 1
				Else
					 TicketCount = TicketCount + 1										
				End If
   			End If
			rs_Tickets.movenext
		Loop
	End IF


	Response.Write("<div class='technitian-area item col-lg-2' userNo='1'>")
	Response.Write("<div class='scrollable-title-awaiting-dispatch' style='position: relative;'><strong>AWAITING DISPATCH")
	If TicketCount > 0 OR FilterCount > 0 Then Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;")
	If TicketCount > 0 Then	Response.Write("<span class='servicecircle' title='" & TicketCount & " Service Tickets' >" & TicketCount & "</span>")
	If FilterCount > 0 Then Response.Write("<span class='filtercircle' title='" & FilterCount & " Filter Tickets' >" & FilterCount  & "</span>")
	Response.Write("</strong>")
	Response.Write("<a class='btn-move' href='#' style='position: absolute; top: 5px; right: 7px;'><i class='fa fa-arrows'></i></a></div>")
        
    %>
    <div class='table-responsive scrollable-table'>

	<div id="await1" name="await1" class="food_planner table table-condensed sortable container-fluid tickets-list-header">
		<div class="row">
			<div class="col-xs-2">
				Ticket
			</div>
			<div class="col-xs-10">
				<%=GetTerm("Customer")%>
			</div>
		</div>
	</div>
	<%

	Do While Not rsAwaitingDispatch.EOF

		GetServiceTicketCurrentStageVar = GetServiceTicketCurrentStage(rsAwaitingDispatch("MemoNumber"))
	
		ShowThisRec = True
		
		Set cnnUserRegionsForServiceBoard = Server.CreateObject("ADODB.Connection")
		cnnUserRegionsForServiceBoard.open (Session("ClientCnnString"))
		Set rsUserRegionsForServiceBoard = Server.CreateObject("ADODB.Recordset")
		rsUserRegionsForServiceBoard.CursorLocation = 3 
		
		SQLUserRegionsForServiceBoard = "SELECT UserRegionsToViewService FROM tblUsers WHERE UserNo = " & Session("UserNo")
		Set rsUserRegionsForServiceBoard = cnnUserRegionsForServiceBoard.Execute(SQLUserRegionsForServiceBoard)
	
		If IsNull(rsUserRegionsForServiceBoard("UserRegionsToViewService")) Then 
			UserRegionList  = ""
		Else
			UserRegionList = rsUserRegionsForServiceBoard("UserRegionsToViewService")
		End If
		
		set rsUserRegionsForServiceBoard = Nothing
		cnnUserRegionsForServiceBoard.close
		set cnnUserRegionsForServiceBoard = Nothing
		
		
		If UserRegionList <> "" Then
		
			CustRegion = GetCustRegionIntRecIDByCustID(rsAwaitingDispatch("AccountNumber"))
			ShowThisRec = False
			
			RegionArray = Split(UserRegionList,",")
			
			For x = 0 to Ubound(RegionArray)
				If cint(RegionArray(x)) = cint(CustRegion) Then
					ShowThisRec = True
					Exit For
				End IF
			Next
		End If
	
	
		If ShowThisRec = True AND rsAwaitingDispatch("RecordSubType") <> "HOLD" AND (GetServiceTicketCurrentStageVar = "Received" OR GetServiceTicketCurrentStageVar = "Released" OR GetServiceTicketCurrentStageVar = "Declined") Then
			
			
			If rsAwaitingDispatch("CurrentStatus") = rsAwaitingDispatch("RecordSubType") Then ' Show only 1 line per memo, the most current status
				
				Call AwaitingDispatchWrite(DynamicFormCounter,rsAwaitingDispatch("MemoNumber")) 
			
			End If
			
		
		End If 'End Awaiting Dispatch Check 
		
		rsAwaitingDispatch.movenext
		
		
	loop
	
	%></div>
	</div><%

End If


GridColumn = 2

Set cnn_Users = Server.CreateObject("ADODB.Connection")
cnn_Users.open (Session("ClientCnnString"))
Set rs_Users = Server.CreateObject("ADODB.Recordset")
rs_Users.CursorLocation = 3 

'Write ordered users
For Each userNo In userorder_arr
	'Fixit
	' cheap fix to let adam henchel see service stuff wihtout being a service manager
	SQL_Users = "SELECT * FROM tblUsers WHERE userNo="&userNo&" AND (UserType='Field Service' or userType = 'Service Manager' OR userNo=56) AND userEnabled = 1"
	Set rs_Users = cnn_Users.Execute(SQL_Users)
	If not rs_Users.EOF Then
		Do While Not rs_Users.Eof
			If GridColumn = 7 Then
				GridColumn = 2	
			End If
			Call TruckNumberWrite(rs_Users("userNo"), rs_Users("userDisplayName"), GridColumn, DynamicFormCounter) 
			rs_Users.Movenext
		Loop
	End If
Next


'Lets get all the field technicians
	'Fixit
	' cheap fix to let adam henchel see service stuff wihtout being a service manager


SQL_Users = "SELECT * FROM tblUsers WHERE userNo not in ("&userorder&") AND (UserType='Field Service' or userType = 'Service Manager' OR userNo=56) AND userEnabled = 1 Order By UserType,userDisplayName"
Set rs_Users = cnn_Users.Execute(SQL_Users)
If not rs_Users.EOF Then
	Do While Not rs_Users.Eof
		If GridColumn = 7 Then
			GridColumn = 2
		End If
		Call TruckNumberWrite(rs_Users("userNo"), rs_Users("userDisplayName"), GridColumn, DynamicFormCounter) 
		rs_Users.Movenext
	Loop
End If


Sub TruckNumberWrite(userNo, userDisplayName, GridColumn, DynamicFormCounter)

		TicketCount = 0
		FilterCount = 0
		
		Set cnn_Tickets = Server.CreateObject("ADODB.Connection")
		cnn_Tickets.open (Session("ClientCnnString"))
		Set rs_Tickets = Server.CreateObject("ADODB.Recordset")

		'TicketCount = GetNumberOfOPENServiceTicketsForTech(userNo)
		'If filterChangeModuleOn() Then FilterCount = GetNumberOfOPENFilterTicketsForTech(userNo)
		
		'Count all the tickets for the header
		SQL_Tickets = "SELECT DISTINCT MemoNumber FROM FS_ServiceMemosDetail "
		SQL_Tickets = SQL_Tickets & "WHERE (MemoNumber IN "
        SQL_Tickets = SQL_Tickets & "(SELECT MemoNumber FROM FS_ServiceMemos WHERE "
        SQL_Tickets = SQL_Tickets & "(CurrentStatus = 'OPEN'))) AND (UserNoOfServiceTech = " & rs_Users("userNo") &")"

        Set rs_Tickets = cnn_Tickets.Execute(SQL_Tickets)
		If not rs_Tickets.Eof Then
			Do While not rs_Tickets.Eof	
								
				If AwaitingRedispatch(rs_Tickets("MemoNumber")) <> True AND LastTechUserNo(rs_Tickets("MemoNumber")) =  rs_Users("userNo") Then	
								
					If TicketIsFilterChange(rs_Tickets("MemoNumber")) Then
						 FilterCount = FilterCount + 1
					Else
						 TicketCount = TicketCount + 1										
					End If
       			End If
				rs_Tickets.movenext
			Loop
		End IF


		Response.Write("<div class='technitian-area item col-lg-2' userNo='" & rs_Users("userNo") & "'>")
		Response.Write("<div class='scrollable-title' style='position: relative;'><strong>" & rs_Users("userDisplayName") & "&nbsp;&nbsp;&nbsp;&nbsp;" & getUserCellNumber(rs_Users("userNo")) )
		If TicketCount > 0 OR FilterCount > 0 Then Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;")
		If TicketCount > 0 Then	Response.Write("<span class='servicecircle' title='" & TicketCount & " Service Tickets' >" & TicketCount & "</span>")
		If FilterCount > 0 Then Response.Write("<span class='filtercircle' title='" & FilterCount & " Filter Tickets' >" & FilterCount  & "</span>")
		Response.Write("</strong>")
		Response.Write("<a class='btn-move' href='#' style='position: absolute; top: 5px; right: 7px;'><i class='fa fa-arrows'></i></a></div>")
		%> 
	        <div class='table-responsive scrollable-table'>
				<div id="tech<%=rs_Users("userNo")%>" name="tech<%=rs_Users("userNo")%>" class="food_planner table table-condensed sortable container-fluid tickets-list-header">
					<div class="row">
						<div class="col-xs-2">
							Ticket
						</div>
						<div class="col-xs-7">
							<%=GetTerm("Customer")%>
						</div>
						<div class="col-xs-3">
							<%=GetTerm("Stage")%>
						</div>
					</div>
					
					

		        </div>
				<ul class="list-group recent-comments">
				<%'Get all the tickets for this person
						SQL_Tickets = "SELECT DISTINCT MemoNumber FROM FS_ServiceMemosDetail "
						SQL_Tickets = SQL_Tickets & "WHERE (MemoNumber IN "
                        SQL_Tickets = SQL_Tickets & "(SELECT MemoNumber FROM FS_ServiceMemos WHERE "
                        SQL_Tickets = SQL_Tickets & "(CurrentStatus = 'OPEN'))) AND (UserNoOfServiceTech = " & rs_Users("userNo") &")"
                        'Response.Write(SQL_Tickets)
                        Set rs_Tickets = cnn_Tickets.Execute(SQL_Tickets)
						If not rs_Tickets.Eof Then
							Do While not rs_Tickets.Eof	
								'Response.Write(LastTechUserNo(rs_Tickets("MemoNumber")))
								'Response.Write(":" & userNo & "::")
								
								If AwaitingRedispatch(rs_Tickets("MemoNumber")) <> True AND LastTechUserNo(rs_Tickets("MemoNumber")) =  rs_Users("userNo") Then	
								
									If ServiceColorsOn = True Then ' Only need to do this extra color stuff if the colors are on
										GetCustTypeCodeByCustIDVar = GetCustTypeCodeByCustID(GetServiceTicketCust(rs_Tickets("MemoNumber")))
										If GetCustTypeCodeByCustIDVar  = "1" or GetCustTypeCodeByCustIDVar  = "2" or GetCustTypeCodeByCustIDVar  = "3" Then 
												If AlertEmailSent(rs_Tickets("MemoNumber")) <> True  Then
													If TicketIsUrgent(rs_Tickets("MemoNumber")) = True Then Response.Write("<li class=""list-group-item clearfix urgent-priority"">") Else Response.Write("<li class=""list-group-item clearfix high-priority"">")
												Else
													If TicketIsUrgent(rs_Tickets("MemoNumber")) = True Then Response.Write("<li class=""list-group-item clearfix urgent-priority"">") Else Response.Write("<li class=""list-group-item clearfix alert-high-priority"">")
												End If
										Else 
											'Not high priority but see if an alert was ever sent
											If AlertEmailSent(rs_Tickets("MemoNumber")) <> True Then
												If LineX Mod 2 = 0 then
													'THESE ARE EVEN LINES
													If TicketIsUrgent(rs_Tickets("MemoNumber")) = True Then Response.Write("<li class=""list-group-item clearfix urgent-priority"">") Else Response.Write("<li class=""list-group-item clearfix tr-even"">")
												Else
													'THESE ARE ODD LINE
													If TicketIsUrgent(rs_Tickets("MemoNumber")) = True Then Response.Write("<li class=""list-group-item clearfix urgent-priority"">") Else Response.Write("<li class=""list-group-item clearfix tr-odd"">")
												End If
											Else
												If TicketIsUrgent(rs_Tickets("MemoNumber")) = True Then Response.Write("<li class=""list-group-item clearfix urgent-priority"">") Else Response.Write("<li class=""list-group-item clearfix alert-priority"">")
											End If
										 End If	
									 Else
										If LineX Mod 2 = 0 then
											'THESE ARE EVEN LINES
											If TicketIsUrgent(rs_Tickets("MemoNumber")) = True Then Response.Write("<li class=""list-group-item clearfix urgent-priority"">") Else Response.Write("<li class=""list-group-item clearfix tr-even"">")
										Else
											'THESE ARE ODD LINE
											If TicketIsUrgent(rs_Tickets("MemoNumber")) = True Then Response.Write("<li class=""list-group-item clearfix urgent-priority"">") Else Response.Write("<li class=""list-group-item clearfix tr-odd"">")
										End If
									 End If
									 
									Response.Write("<div class=""container-fluid""><div class=""row"">")
									Response.Write("<div class=""col-xs-2 pl-1 pr-1 ticket-number"" data-tickets-number="""+rs_Tickets("MemoNumber")+""">")
									
									'Response.Write("<a href='../editServiceMemo.asp?memo=" & rs_Tickets("MemoNumber") & "'>" & rs_Tickets("MemoNumber")& "</a>")
									%>
									<% If userCanAccessServiceCloseCancelButton(Session("UserNo")) = true Then %>
										<a data-toggle="modal" data-show="true" href="#" data-target="#modalEditExistingServiceTicketForClient" data-memo-number="<%= rs_Tickets("MemoNumber") %>" data-cust-id="<%= GetServiceTicketCust(rs_Tickets("MemoNumber")) %>" data-tooltip="true" data-title="Close/Cancel Service Ticket" style="cursor:pointer;"><%= rs_Tickets("MemoNumber") %></a>
									<% End If 
									
									If TicketIsFilterChange(rs_Tickets("MemoNumber")) Then
										 imagetoshow = baseURL & "img/general/F.png"
										 ToolText="Filter Change"
									Else
										imagetoshow = baseURL & "img/general/S.png" 
										 ToolText="Service Ticket"										
									End If
									Response.Write("<br><img src='" & imagetoshow  & "' height='10' width='10' title='" & ToolText & "'>")
									
									Response.Write("</div>")
									Response.Write("<div class=""col-xs-7 pl-1 pr-1 customer-data pt-5"">")
									Response.Write("<button type='button' class='btn btn-default tooltip-button' data-toggle='tooltip' data-placement='bottom' title='" & GetCustNameByCustNum(GetServiceTicketCust(rs_Tickets("MemoNumber"))) & "'>" & GetServiceTicketCust(rs_Tickets("MemoNumber")) & "</button>&nbsp;&nbsp;" & GetCustNameByCustNum(GetServiceTicketCust(rs_Tickets("MemoNumber"))))
									Response.Write("</div>")
									Response.Write("<div class=""col-xs-3 pl-1 pr-1 pt-5"">")
									StageText = GetServiceTicketCurrentStage(rs_Tickets("MemoNumber"))
									If StageText="Dispatch Acknowledged" Then 
										StageText="Dispatch<br>Ack"
									END IF
									Response.Write("<button type='button' class='btn btn-default tooltip-button' data-toggle='tooltip' data-container='body' data-placement='bottom' title='" & GetServiceTicketSTAGEDateTime(rs_Tickets("MemoNumber"),GetServiceTicketCurrentStage(rs_Tickets("MemoNumber"))) & "'>" & StageText & "</button>" )
									Response.Write("</div>")
									Response.Write("</div></div>")
				        			Response.Write("</li>")
				        			DynamicFormCounter = DynamicFormCounter + 1
			        			End If
								rs_Tickets.movenext
							Loop
						End IF
						
						%>
						
	        </div>
        <%Response.Write("</div>")
		GridColumn = GridColumn +1
		DynamicFormCounter = DynamicFormCounter + 1
End Sub 




Sub AwaitingDispatchWrite(DynamicFormCounter,MemoNumber)

		%>		
		<ul class="list-group recent-comments">
		
		<%'Get all the tickets awaiting dispatch for every customer
		
	   SQL_Tickets = "SELECT DISTINCT MemoNumber FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN' AND MemoNumber = '" & MemoNumber & "' "

        Set rs_Tickets = cnn_Tickets.Execute(SQL_Tickets)
        
		If not rs_Tickets.Eof Then
		
			'Do While not rs_Tickets.Eof	
				
			GetServiceTicketCurrentStageVar = GetServiceTicketCurrentStage(rs_Tickets("MemoNumber"))
			
			If GetServiceTicketCurrentStageVar = "Received" OR GetServiceTicketCurrentStageVar = "Released" OR GetServiceTicketCurrentStageVar = "Declined" Then

				If ServiceColorsOn = True Then ' Only need to do this extra color stuff if the colors are on
					GetCustTypeCodeByCustIDVar = GetCustTypeCodeByCustID(GetServiceTicketCust(rs_Tickets("MemoNumber")))
					If GetCustTypeCodeByCustIDVar  = "1" or GetCustTypeCodeByCustIDVar  = "2" or GetCustTypeCodeByCustIDVar  = "3" Then 
							If AlertEmailSent(rs_Tickets("MemoNumber")) <> True  Then
								If TicketIsUrgent(rs_Tickets("MemoNumber")) = True Then Response.Write("<li class=""list-group-item clearfix urgent-priority"">") Else Response.Write("<li class=""list-group-item clearfix high-priority"">")
							Else
								If TicketIsUrgent(rs_Tickets("MemoNumber")) = True Then Response.Write("<li class=""list-group-item clearfix urgent-priority"">") Else Response.Write("<li class=""list-group-item clearfix alert-high-priority"">")
							End If
					Else 
						'Not high priority but see if an alert was ever sent
						If AlertEmailSent(rs_Tickets("MemoNumber")) <> True Then
							If LineX Mod 2 = 0 then
								'THESE ARE EVEN LINES
								If TicketIsUrgent(rs_Tickets("MemoNumber")) = True Then Response.Write("<li class=""list-group-item clearfix urgent-priority"">") Else Response.Write("<li class=""list-group-item clearfix tr-even"">")
							Else
								'THESE ARE ODD LINE
								If TicketIsUrgent(rs_Tickets("MemoNumber")) = True Then Response.Write("<li class=""list-group-item clearfix urgent-priority"">") Else Response.Write("<li class=""list-group-item clearfix tr-odd"">")
							End If
						Else
							If TicketIsUrgent(rs_Tickets("MemoNumber")) = True Then Response.Write("<li class=""list-group-item clearfix urgent-priority"">") Else Response.Write("<li class=""list-group-item clearfix alert-priority"">")
						End If
					 End If	
				 Else
					If LineX Mod 2 = 0 then
						'THESE ARE EVEN LINES
						If TicketIsUrgent(rs_Tickets("MemoNumber")) = True Then Response.Write("<li class=""list-group-item clearfix urgent-priority"">") Else Response.Write("<li class=""list-group-item clearfix tr-even"">")
					Else
						'THESE ARE ODD LINE
						If TicketIsUrgent(rs_Tickets("MemoNumber")) = True Then Response.Write("<li class=""list-group-item clearfix urgent-priority"">") Else Response.Write("<li class=""list-group-item clearfix tr-odd"">")
					End If
				 End If
				 
				Response.Write("<div class=""container-fluid""><div class=""row"">")
				Response.Write("<div class=""col-xs-3 pl-1 pr-1 ticket-number"" data-tickets-number="""+rs_Tickets("MemoNumber")+""">")
				
				%>
				<% If userCanAccessServiceCloseCancelButton(Session("UserNo")) = true Then %>
					<a data-toggle="modal" data-show="true" href="#" data-target="#modalEditExistingServiceTicketForClient" data-memo-number="<%= rs_Tickets("MemoNumber") %>" data-cust-id="<%= GetServiceTicketCust(rs_Tickets("MemoNumber")) %>" data-tooltip="true" data-title="Close/Cancel Service Ticket" style="cursor:pointer;"><%= rs_Tickets("MemoNumber") %></a>
				<% End If 
				
				If TicketIsFilterChange(rs_Tickets("MemoNumber")) Then
					 imagetoshow = baseURL & "img/general/F.png"
					 ToolText="Filter Change"
				Else
					imagetoshow = baseURL & "img/general/S.png" 
					 ToolText="Service Ticket"										
				End If
				Response.Write("<br><img src='" & imagetoshow  & "' height='10' width='10' title='" & ToolText & "'>")
				
				Response.Write("</div>")
				
				Response.Write("<div class=""col-xs-7 pl-1 pr-1 customer-data pt-5"">")
				Response.Write("<button type='button' class='btn btn-default tooltip-button' data-toggle='tooltip' data-placement='bottom' title='" & GetCustNameByCustNum(GetServiceTicketCust(rs_Tickets("MemoNumber"))) & "'>" & GetServiceTicketCust(rs_Tickets("MemoNumber")) & "</button>&nbsp;&nbsp;" & GetCustNameByCustNum(GetServiceTicketCust(rs_Tickets("MemoNumber"))))
				Response.Write("</div>")
				
				Response.Write("</div></div>")
    			Response.Write("</li>")
    			DynamicFormCounter = DynamicFormCounter + 1
			End If
			
			'rs_Tickets.movenext
		'Loop
		End IF
						
		DynamicFormCounter = DynamicFormCounter + 1
		
End Sub 





Set rs_Users = Nothing
cnn_Users.Close
Set cnn_Users = Nothing
%>	
</div>


<!--#include file="../serviceBoardCommonModals.asp"-->

<!--#include file="../../inc/footer-main.asp"-->

<script type="text/javascript">
	function setSortable() {
		$("#sortableList").sortable({ placeholder: "ui-state-highlight item col-lg-2", handle: ".btn-move", scrollSensitivity: 40, scrollSpeed: 60, update: function (event, ui) { saveSelection(); } });
		
		
		$("#sortableList").disableSelection();
	}
	function saveSelection() {
		var list = "";
		try {
			var sep = "";
			$("#sortableList .item").each(function () {
				list += "" + sep + $(this).attr("userNo");
				sep = ",";
			});
		}
		catch (ex) {
			alert(ex);
			return;
		}
		
		var url = "userorder/save.asp";
		var jsondata = {};
		jsondata.userorder = list;
		$.ajax({
			type: "POST",
			url: url,
			dataType: "json",
			data: jsondata,
			success: function (data) {
			}
		});
	}
	$(function () {
		setSortable();
	});
</script>

<!-- countdown script !-->
<script src="<%= BaseURL %>js/countdown/jquery.pietimer.js"></script>
<script type="text/javascript">

	
	$(document).ready(function() {

	    $('.modal').on('show.bs.modal', function () {
			//console.log( "opened modal!" );
			$('#switchAutomaticRefresh').prop('checked', true).trigger("change");
	    });
	    	          	
		$('.modal').on('hidden.bs.modal', function () {
			//console.log( "closed modal!" );
			$('#switchAutomaticRefresh').prop('checked', false).trigger("change");   		
    	}); 
		$(".recent-comments").sortable( {
			connectWith: ".recent-comments"
		});
		$( ".recent-comments" ).on( "sortreceive", function( event, ui ) {
			
			console.log($(ui.item).find(".ticket-number").attr("data-tickets-number"));
			console.log($(ui.item).closest(".technitian-area").attr("userno"));
			$.ajax({
				url: "dispatchCenter_ajax_SaveValues.asp",
				data: "selFieldTech="+$(ui.item).closest(".technitian-area").attr("userno")+"&txtServiceTicketNumber="+$(ui.item).find(".ticket-number").attr("data-tickets-number"),
				success: function (data) {
					result=JSON.parse(data);
					swal({
						  title: 'Send a text message?',
						  text: "You have successfully re-assigned ticket #"+result.ServiceTicketNumber+" to "+result.UserDisplayName+". Send text message to this tech?",
						  type: 'warning',
						  showCancelButton: true,
						  confirmButtonColor: '#3085d6',
						  cancelButtonColor: '#d33',
						  cancelButtonText: 'No',
						  confirmButtonText: 'Yes, send!'
						},
						function(isConfirm){
							
						  if (isConfirm) {
								$.ajax({
									url: "dispatchCenter_ajax_sendtxt.asp",
									data: "selFieldTech="+result.UserToDispatch+"&txtServiceTicketNumber="+result.ServiceTicketNumber,
									success: function (data) {
										swal({
										title: 'Text sent!',
										text: "Text message was sent to "+result.UserDisplayName,
										type: 'success'
										});
									} 
								});
							}
						}
					);
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


	$(function(){

		var rgbcolor = '<%=Session("DelBoardPieTimerColor")%>';	
		
		var pagetimer = new Timer(function() {
		    location.reload();
		}, 45*1000);
				
		$('#timer').pietimer({
			seconds: 45,
			color: hexToRgb(rgbcolor),
			height: 35,
			width: 35,
			is_reversed: true
		});
		
		
		$('#timer').pietimer('start');


		$('#switchAutomaticRefresh').on('change', function() {
		   if (this.checked) {
		        $('#timer').pietimer('pause');
		        pagetimer.pause();
		        return false;
		   }
		   else {
				$('#timer').pietimer('start');
				pagetimer.resume();
				return false;
		    }

		})	
	});
</script>
<!-- eof countdown script !-->

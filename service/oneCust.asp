<!--#include file="../inc/header.asp"-->
<!--#include file="../inc/InsightFuncs_Service.asp"-->
<!--#include file="../inc/jquery_table_search.asp"-->
<!--#include file="../inc/mail.asp"-->

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


<script>

	$(document).ready(function() {

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
    				
	    });
	    
	 });
	 
</script>

<% CustToFind = Request.QueryString("cust") %>


<!-- on/off scripts !-->

 
 <style type="text/css">
 	.email-table{
		width:46%;
	}
	
	table.sortable th:not(.sorttable_sorted):not(.sorttable_sorted_reverse):not(.sorttable_nosort):after { 
    	content: " \25B4\25BE" 
	}

	table thead a{
		color: #000;
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
		border: 1px solid #eaeaea;
		padding-top: 10px;
		margin-bottom: 15px;
	}
	 
	.high-priority{
		background:#fad5d5;
	}
	
	.alert-priority{
		background:#faf99d;
	}
	
	.alert-high-priority{
		background:#fa9090;
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

 </style>

<!--- eof on/off scripts !-->


 

<h1 class="page-header"><i class="fa fa-wrench"></i> Service tickets for <%=GetTerm("account")%>: <%=CustToFind%></h1>

 

<!-- row !-->
<div class="row-data">	
	
	<div class="col-lg-3 col-md-3 col-sm-12 col-xs-12 alert-success alert">
		 			<%Response.Write(FormattedCustInfoByCustNum(CustToFind))%>
 	</div>
	
	<form method="post" action="main.asp" name="frmMain">
	 
		<div class="col-lg-3">

			<div class="row">
			
			<!-- date !-->
			<div class="col-lg-12">
 					<div class="row">
	 					
				<div class="col-lg-6">
			</div>
			
			<div class="col-lg-6">
			</div>
			     
 					</div>
  			</div>
 			<!-- eof date !-->
			
		
		
		</div>
		</div>
		
		
	</form>
 	<div class="col-lg-3 legend-box">
		
		<!-- line !-->
		<div class="row legend-row">
			<div class="col-lg-2 col-md-2 col-sm-2 col-xs-2 alert-priority">&nbsp;</div>
			<div class="col-lg-10 col-md-10 col-sm-10 col-xs-10"><h6 class="legend-title">Normal Account - Alert Sent</h6></div>
		</div>
		<!-- eof line !-->
		
		<!-- line !-->
		<div class="row legend-row">
			<div class="col-lg-2 col-md-2 col-sm-2 col-xs-2 high-priority">&nbsp;</div>
			<div class="col-lg-10 col-md-10 col-sm-10 col-xs-10"><h6 class="legend-title">Priority Account</h6></div>
		</div>
		<!-- eof line !-->
		
				<!-- line !-->
		<div class="row legend-row">
			<div class="col-lg-2 col-md-2 col-sm-2 col-xs-2 alert-high-priority">&nbsp;</div>
			<div class="col-lg-10 col-md-10 col-sm-10 col-xs-10"><h6 class="legend-title">Priority Account - Alert Sent</h6></div>
		</div>
		<!-- eof line !-->

		
	</div>
  <!-- eof legend !-->

  
 </div>
<!-- eof row !-->
	
<!-- row !-->
<div class="row">
 	<div class="col-lg-12">
 	
    

 
    	
        <div class="table-responsive">
            <table  id="tableSuperSum"    class="food_planner table  table-condensed  sortable">
              <thead>
                <tr>
	              <th class="sorttable_numeric">Date</th>
	              <th>Ticket #</th>	              
                  <th>Status</th>
                  <th class="sorttable_nosort">Description</th>
                  <th>Dispatched</th>
                  <th class="sorttable_numeric">Elapsed<br>Time</th>
                  <th>Details</th>
                  <th>Submitted Via</th>
                </tr>
              </thead>
              
              <tbody class='searchable'>
              
			<%
			
			SQL = "SELECT * FROM FS_ServiceMemos WHERE AccountNumber ='" & CustToFind & "' AND CurrentStatus = 'OPEN' AND RecordSubType = 'OPEN' "
			SQL = SQL & " order by submissionDateTime desc"


			Set cnn8 = Server.CreateObject("ADODB.Connection")
			cnn8.open (Session("ClientCnnString"))
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.CursorLocation = 3 
			Set rs = cnn8.Execute(SQL)

			If not rs.EOF Then

				Do While Not rs.EOF
				
						If rs.Fields("CurrentStatus") = rs.Fields("RecordSubType") Then ' Show only 1 line per memo, the most current status
				        %>
							<!-- table line !-->
							<% If GetCustTypeCodeByCustID(rs.Fields("AccountNumber")) = "1" or GetCustTypeCodeByCustID(rs.Fields("AccountNumber")) = "2" or GetCustTypeCodeByCustID(rs.Fields("AccountNumber")) = "3" Then 
									If Isnull(rs.Fields("AlertEmailSent")) Then
										%>
										<tr class="high-priority">
									<%Else%>
										<tr class="alert-high-priority">
									<%End If%>
							<% Else 
								'Not high priority but see if an alert was ever sent
								If Isnull(rs.Fields("AlertEmailSent")) Then
									%>
									<tr class="low-priority">
								<%Else%>
									<tr class="alert-priority">
								<%End If%>
							<% End If	
							Response.write("<td sorttable_customkey=" & FormatAsSortableDateTime(rs("submissionDateTime")) & ">" & FormatDateTime(rs("submissionDateTime")) & "</td>")%>
							<td><%= rs.Fields("MemoNumber")%></td>
							<td><%= rs.Fields("RecordSubType") %></td>
							<td><%= rs.Fields("ProblemDescription") %></td>
							<td sorttable_customkey="<%= rs.Fields("Dispatched") %>">
							<%If rs.Fields("Dispatched") = vbTrue then 
								Response.Write("Yes")
							Else 
								Response.Write("No")
							End If%>
							</td>
							<%			
							If ElapsedTimeCalcMethod() = "Actual" Then
								If rs.Fields("CurrentStatus") = "CLOSE" or rs.Fields("CurrentStatus") = "CANCEL" Then
									elapsedMinutes = datediff("n",GetServiceTicketOpenDateTime(rs.Fields("MemoNumber")),GetServiceTicketCloseDateTime(rs.Fields("MemoNumber")))
									elapsedMinutesForSorting = elapsedMinutes
									elapsedString = ""
									minutesInServiceDay = GetNumberOfMinutesInServiceDay()
									elapsedDays = 	elapsedMinutes \ minutesInServiceDay
									If int(elapsedDays) > 0 Then
										elapsedMinutes = elapsedMinutes - (int(elapsedDays) * minutesInServiceDay)
										elapsedString = elapsedDays & "d "
									End If
									elapsedHours = 	elapsedMinutes \ 60
									If int(elapsedHours) > 0 Then 
										elapsedMinutes = elapsedMinutes - (int(elapsedHours) * 60)
										elapsedString = elapsedString  & elapsedHours & "h "
									End IF	
									If int(elapsedMinutes) > 0 Then
										elapsedString = elapsedString  & elapsedMinutes & "m"
									End If
									elapsedMinutes = datediff("n",GetServiceTicketOpenDateTime(rs.Fields("MemoNumber")),GetServiceTicketCloseDateTime(rs.Fields("MemoNumber")))
									Response.Write("<td sorttable_customkey='" & elapsedMinutesForSorting & "'>" & elapsedString & "<br>") 
									'response.Write(elapsedMinutes)
								ElseIf rs.Fields("CurrentStatus") = "OPEN" Then 
									elapsedMinutes = datediff("n",GetServiceTicketOpenDateTime(rs.Fields("MemoNumber")),Now())
									elapsedMinutesForSorting = elapsedMinutes
									elapsedString = ""
									minutesInServiceDay = GetNumberOfMinutesInServiceDay()
									elapsedDays = 	elapsedMinutes \ minutesInServiceDay
									If int(elapsedDays) > 0 Then
										elapsedMinutes = elapsedMinutes - (int(elapsedDays) * minutesInServiceDay)
										elapsedString = elapsedDays & "d "
									End If
									elapsedHours = 	elapsedMinutes \ 60
									If int(elapsedHours) > 0 Then 
										elapsedMinutes = elapsedMinutes - (int(elapsedHours) * 60)
										elapsedString = elapsedString  & elapsedHours & "h "
									End IF	
									If int(elapsedMinutes) > 0 Then
										elapsedString = elapsedString  & elapsedMinutes & "m"
									End If
									elapsedMinutes = datediff("n",GetServiceTicketOpenDateTime(rs.Fields("MemoNumber")),Now())
									Response.Write("<td sorttable_customkey='" & elapsedMinutesForSorting & "'>" & elapsedString & "<br>") 
									'response.Write(elapsedMinutes)
								Elseif rs.Fields("CurrentStatus") = "HOLD" Then
									'Response.Write("<td sorttable_customkey='" & 0 & "'>" & "Hold<br>") 
									elapsedMinutes = datediff("n",GetServiceTicketOpenDateTime(rs.Fields("MemoNumber")),Now())
									elapsedMinutesForSorting = elapsedMinutes
									elapsedString = ""
									minutesInServiceDay = GetNumberOfMinutesInServiceDay()
									elapsedDays = 	elapsedMinutes \ minutesInServiceDay
									If int(elapsedDays) > 0 Then
										elapsedMinutes = elapsedMinutes - (int(elapsedDays) * minutesInServiceDay)
										elapsedString = elapsedDays & "d "
									End If
									elapsedHours = 	elapsedMinutes \ 60
									If int(elapsedHours) > 0 Then 
										elapsedMinutes = elapsedMinutes - (int(elapsedHours) * 60)
										elapsedString = elapsedString  & elapsedHours & "h "
									End IF	
									If int(elapsedMinutes) > 0 Then
										elapsedString = elapsedString  & elapsedMinutes & "m"
									End If
									elapsedMinutes = datediff("n",GetServiceTicketOpenDateTime(rs.Fields("MemoNumber")),Now())
									Response.Write("<td sorttable_customkey='" & elapsedMinutesForSorting & "'>" & elapsedString & "<br>") 
								End If
							Else
								If rs.Fields("CurrentStatus") = "CLOSE" or rs.Fields("CurrentStatus") = "CANCEL" Then
									elapsedMinutes = ServiceCallElapsedMinutes(rs.Fields("MemoNumber"))
									elapsedMinutesForSorting = elapsedMinutes 
									elapsedString = ""
									elapsedHours = 	elapsedMinutes \ 60
									If int(elapsedHours) > 0 Then 
										elapsedMinutes = elapsedMinutes - (int(elapsedHours) * 60)
										elapsedString = elapsedString  & elapsedHours & "h "
									End IF	
									If int(elapsedMinutes) > 0 Then
										elapsedString = elapsedString  & elapsedMinutes & "m"
									End If
									If elapsedMinutes = 0 Then elapsedString = "0"
									Response.Write("<td sorttable_customkey='" & elapsedMinutesForSorting & "'>" & elapsedString & "<br>") 
									'response.Write(elapsedMinutes)
								ElseIf rs.Fields("CurrentStatus") = "OPEN" Then 
									elapsedMinutes = ServiceCallElapsedMinutes(rs.Fields("MemoNumber"))
									elapsedMinutesForSorting = elapsedMinutes
									elapsedString = ""
									elapsedHours = 	elapsedMinutes \ 60
									If int(elapsedHours) > 0 Then 
										elapsedMinutes = elapsedMinutes - (int(elapsedHours) * 60)
										elapsedString = elapsedString  & elapsedHours & "h "
									End IF	
									If int(elapsedMinutes) > 0 Then
										elapsedString = elapsedString  & elapsedMinutes & "m"
									End If
									Response.Write("<td sorttable_customkey='" & elapsedMinutesForSorting & "'>" & elapsedString & "<br>") 
									'response.Write(elapsedMinutes)
								Elseif rs.Fields("CurrentStatus") = "HOLD" Then
									'Response.Write("<td sorttable_customkey='" & 0 & "'>" & "Hold<br>") 
									elapsedMinutes = ServiceCallElapsedMinutes(rs.Fields("MemoNumber"))
									elapsedMinutesForSorting = elapsedMinutes
									elapsedString = ""
									elapsedHours = 	elapsedMinutes \ 60
									If int(elapsedHours) > 0 Then 
										elapsedMinutes = elapsedMinutes - (int(elapsedHours) * 60)
										elapsedString = elapsedString  & elapsedHours & "h "
									End IF	
									If int(elapsedMinutes) > 0 Then
										elapsedString = elapsedString  & elapsedMinutes & "m"
									End If
									Response.Write("<td sorttable_customkey='" & elapsedMinutesForSorting & "'>" & elapsedString & "<br>") 
								End If
							End If
							%>
							</td>
							
							<% If userIsServiceManager(Session("userno")) = True Then 
								If rs.Fields("CurrentStatus")="OPEN" Then %>
									<td>
										<!-- new close/cancel button !-->
										<% If userCanAccessServiceCloseCancelButton(Session("UserNo")) = true Then %>
											<button type="button" class="btn btn-danger btn-sm" style="margin-top:-4px" data-toggle="modal" data-show="true" href="#" data-target="#modalEditExistingServiceTicketForClient" data-memo-number="<%= rs.Fields("MemoNumber") %>" data-cust-id="<%= CustToFind %>" data-tooltip="true" data-title="Close/Cancel Service Ticket" style="cursor:pointer;">Close/Cancel</button>
										<% End If %>
										<!-- new close/cancel button !-->
									</td>
								<% Else %>
									<td><button type="button" class="btn btn-success btn-sm" data-toggle="modal" data-show="true" href="#" data-memo-number="<%= rs.Fields("MemoNumber") %>" data-cust-id="<%= CustToFind %>" data-target="#modalViewOpenClosedServiceTicketDetailsForClient" data-tooltip="true" data-title="View Service Ticket Details" style="cursor:pointer;"><i class="fas fa-eye"></i></button></td>
								<% End If %>
							<% Else %>
								<td><button type="button" class="btn btn-success btn-sm" data-toggle="modal" data-show="true" href="#" data-memo-number="<%= rs.Fields("MemoNumber") %>" data-cust-id="<%= CustToFind %>" data-target="#modalViewOpenClosedServiceTicketDetailsForClient" data-tooltip="true" data-title="View Service Ticket Details" style="cursor:pointer;"><i class="fas fa-eye"></i></button></td>
							<% End If %>
							
							<td><%= rs.Fields("SubmissionSource") %></td>
							</tr>
							<!-- eof table line !-->
						<%
						
						
					
						
				
						End If
						
						rs.movenext
				loop
				
			End If
	
			set rs = Nothing
			cnn8.close
			set cnn8 = Nothing

            %>
              
              
              
              
              </tbody>
            </table>
          </div>

    </div>	
    

</div>
<!-- eof row !-->    

<!--#include file="serviceBoardCommonModals.asp"-->

<!--#include file="../inc/footer-main.asp"-->
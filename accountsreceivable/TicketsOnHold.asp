<!--#include file="../inc/header-accounts-receivable.asp"-->
<!--#include file="../inc/jquery_table_search.asp"-->
<!--#include file="../inc/InsightFuncs_AR_AP.asp"-->
<!--#include file="../inc/mail.asp"-->

<% ''If it comes in here with a querystring called ses=1 then it was an auto refresh, otherwise it was direct navigation %>
<meta http-equiv="refresh" content="120, URL=TicketsOnHold.asp?ses=1"> 


<%If Request.QueryString("ses")="" Then Session("RefreshOptions")= "" ' just blank it out if we didnt come from there 


'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)


Session("MemoNumber") = ""
Session("ServiceCustID") = ""
'This code for auto refresh
If Request.QueryString("ses") <> ""  Then 
	If Instr(Session("RefreshOptions"),"Condensed") <> 0 Then ViewType="Condensed"
	If Instr(Session("RefreshOptions"),"Normal") <> 0 Then ViewType="Normal"
Else
	ViewType=Request.Form("optView")
	If ViewType = "" Then ViewType="Condensed"
	Session("RefreshOptions") = ViewType
End If
 

If Request.form("selDtRange") <> "" Then
	'Construct WHERE_CLAUSE variable
	Select Case Request.form("selDtRange")
		Case "Last 48 Hours"
			WHERE_CLAUSE ="Where SubmissionDateTime >= DATEADD(HOUR, -48, GETDATE())"
		Case "Last 24 Hours"
			WHERE_CLAUSE ="Where SubmissionDateTime >= DATEADD(HOUR, -24, GETDATE())"
		Case "All Dates"
			WHERE_CLAUSE =""
			WHERE_CLAUSE ="Where SubmissionDateTime> DateAdd(d,-5,getdate()) "
		Case "Today"
			WHERE_CLAUSE ="Where DATEPART(dayofyear,SubmissionDateTime) = DATEPART(dayofyear,getdate()) AND DATEPART(year,SubmissionDateTime) = DATEPART(year,getdate()) "
		Case "This Week"
			WHERE_CLAUSE ="Where DATEPART(week,SubmissionDateTime) = DATEPART(week,getdate()) AND DATEPART(year,SubmissionDateTime) = DATEPART(year,getdate()) "
		Case "This Month"
			WHERE_CLAUSE ="Where DATEPART(month,SubmissionDateTime) = DATEPART(month,getdate()) AND DATEPART(year,SubmissionDateTime) = DATEPART(year,getdate()) "
		Case "This Quarter"
			WHERE_CLAUSE ="Where DATEPART(quarter,SubmissionDateTime) = DATEPART(quarter,getdate()) AND DATEPART(year,SubmissionDateTime) = DATEPART(year,getdate()) "
		Case "Last 3 Days"
			WHERE_CLAUSE ="Where SubmissionDateTime> DateAdd(d,-3,getdate()) "
		Case "Last 10 Days"
			WHERE_CLAUSE ="Where SubmissionDateTime> DateAdd(d,-10,getdate()) "
		Case "Last 30 Days"
			WHERE_CLAUSE ="Where SubmissionDateTime> DateAdd(d,-30,getdate()) "
		Case "Last 60 Days"
			WHERE_CLAUSE ="Where SubmissionDateTime> DateAdd(d,-60,getdate()) "
		Case "Last 90 Days"
			WHERE_CLAUSE ="Where SubmissionDateTime> DateAdd(d,-90,getdate()) "
		Case Else
			WHERE_CLAUSE=""
	End Select
	If Instr(Session("RefreshOptions"),Request.form("selDtRange") & "|") = 0 Then 'stop uncontrolled growth
		Session("RefreshOptions") = Session("RefreshOptions") & Request.form("selDtRange") & "|"
	End If
else
	If Instr(Session("RefreshOptions"),"All Dates|") = 0 Then 'stop uncontrolled growth
		Session("RefreshOptions") = Session("RefreshOptions") & "All Dates|"
	End IF
End IF

'This code to handle auto refreshes
chkHoldValue ="on"


%>

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
	  width: 75%;
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
	
	.modalResponsiveTable {
		margin-left: 25px;
		margin-right: 25px;
	}
	
	
	.ajaxRowView .visibleRowEdit, .ajaxRowEdit .visibleRowView { display: none; }
	
	
	.ajax-loading {
	    position: relative;
	}
	.ajax-loading::after {
	    background-image: url("/img/loading.gif");
	    background-position: center top;
	    background-repeat: no-repeat;
	    content: "";
	    display: block;
	    height: 100%;
	    min-height: 100px;
	    position: absolute;
	    top: 0;
	    width: 100%;
	}
	
	.bs-example-modal-lg-customize .row{
		margin-bottom: 10px;
	 	width: 100%;
		overflow: hidden;
		font-size:11px;
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

 </style>

<script type="text/javascript">

	
	$(document).ready(function() {
	
	    $("#PleaseWaitPanel").hide();
		
		$('#modalEditCustomerNotes').on('show.bs.modal', function(e) {
	
		    //get data-id attribute of the clicked order
		    var CustID = $(e.relatedTarget).data('cust-id');
		    var CategoryID = $(e.relatedTarget).data('category-id');
		    
		    //populate the textbox with the id of the clicked order
		    $(e.currentTarget).find('input[name="txtCustIDToPassToGenerateNotes"]').val(CustID);
		    $(e.currentTarget).find('input[name="txtCustIDToPass"]').val(CustID);
		    $(e.currentTarget).find('input[name="txtCategoryID"]').val(CategoryID);
		    	    
		    var $modal = $(this);
	
	    	$.ajax({
				type:"POST",
				url: "../../../inc/InSightFuncs_AjaxForARAP.asp",
				cache: false,
				data: "action=GetContentForCustomerNotesModal&CustID="+encodeURIComponent(CustID),
				success: function(response)
				 {
	               	 $modal.find('#modalEditCustomerNotesContent').html(response);
	               	 //alert(response);               	 
	             },
	             failure: function(response)
				 {
				  	$modal.find('#modalEditCustomerNotesContent').html("Failed");
		            //var height = $(window).height() - 600;
		            //$(this).find(".modal-body").css("max-height", height);
	             }
			});
		});
			
	
	});

</script>

 
<%
	Response.Write("<div id=""PleaseWaitPanel"" class=""container"">")
	Response.Write("<br><br>Loading Service Tickets On Hold <br><br>Please wait...<br><br>")
	Response.Write("<img src='" & baseURL & "/img/loading.gif'/>")
	Response.Write("</div>")
	Response.Flush()
%>

<h1 class="page-header"><i class="fa fa-wrench"></i> Service Tickets On Hold</h1>

<div class="row">
 	<div class="col-lg-12">
	 	<p>
	 	</p>
 	</div>
</div>

<!-- row !-->
<div class="row row-data">	
	
	<div class="col-lg-4 col-md-4 col-sm-12 col-xs-12">
		<br>
		<div class="input-group"> <span class="input-group-addon">Narrow Results</span>
		    <input id="filter" type="text" class="form-control " placeholder="Type here...">
		</div>
		<br>
	</div>
	
	<form method="post" action="../accountsreceivable/TicketsOnHold.asp" name="frmMain">
	 
		<div class="col-lg-4">
			<div class="row  data-range-box">
			
			<!-- date !-->
			<div class="col-lg-12">
 					<div class="row">
	 					
				<div class="col-lg-6">
			<strong >Date Range</strong>
	      	<select class="form-control" name="selDtRange">
			      	<option <%If Request.form("selDtRange") = "Last 48 Hours" Then %>selected<%End If%>>Last 48 Hours</option>
			      	<option <%If Request.form("selDtRange") = "Last 24 Hours" Then %>selected<%End If%>>Last 24 Hours</option>
		         	<option <%If Request.form("selDtRange") = "All Dates" or Request.form("selDtRange") = "" Then %>selected<%End If%>>All Dates</option>
		         	<option <%If Request.form("selDtRange") = "Today" Then %>selected<%End If%>>Today</option>
					<option <%If Request.form("selDtRange") = "This Week" Then %>selected<%End If%>>This Week</option>
					<option <%If Request.form("selDtRange") = "This Month" Then %>selected<%End If%>>This Month</option>
					<option <%If Request.form("selDtRange") = "This Quarter" Then %>selected<%End If%>>This Quarter</option>
					<option <%If Request.form("selDtRange") = "Last 3 Days" Then %>selected<%End If%>>Last 3 Days</option>
					<option <%If Request.form("selDtRange") = "Last 10 Days" Then %>selected<%End If%>>Last 10 Days</option>
					<option <%If Request.form("selDtRange") = "Last 30 Days" Then %>selected<%End If%>>Last 30 Days</option>
					<option <%If Request.form("selDtRange") = "Last 60 Days" Then %>selected<%End If%>>Last 60 Days</option>
					<option <%If Request.form("selDtRange") = "Last 90 Days" Then %>selected<%End If%>>Last 90 Days</option>
			</select> 
			</div>
			
			<div class="col-lg-6">
				<br>
 			<a href="#" onClick="document.frmMain.submit()"><button type="button" class="btn btn-primary">Apply</button></a>
			</div>
			     
 					</div>
  			</div>
 			<!-- eof date !-->
			
			<!-- checkboxes !-->
			<div class="col-lg-12 date-range">
			  
			  <label>
			  <input type="radio" name="optView" id="optView" value="Normal" <%If ViewType = "Normal" Then Response.Write("checked ") %> > Normal view
			  </label>
			  
			  <label>
			  <input type="radio" name="optView" id="optView" value="Condensed" <%If ViewType = "Condensed" Then Response.Write("checked ") %> > Condensed view
			  </label>
			</div>
			<!-- eof checkboxes !-->
		
		
		</div>
		</div>
		
		<!-- service tickets / dispatch !-->
		<div class="col-lg-4">
		Service tickets on hold: <%= NumberOfHoldServiceCalls() %><br>&nbsp;
		</div>
		<!-- eof service tickets / dispatch !-->
		
	</form>
 </div>
<!-- eof row !-->
	
<!-- row !-->
<div class="row">
 	<div class="col-lg-12">
 	
    

 
    	
        <div class="table-responsive">
            <table  id="tableSuperSum"    class="food_planner table  table-condensed  table-striped sortable">
              <thead>
                <tr>
	              <th class="sorttable_numeric">Date</th>
	              <th>Ticket #</th>	              
                  <th>Hold Reason</th>
                  <th class="sorttable_nosort">&nbsp;</th>
                  <th><%=GetTerm("Account")%>#</th>
                  <th>Company</th>
                  <th class="sorttable_nosort">Description</th>
	              <th>Stage</th>
                  <th class="sorttable_numeric">Time On<br>Hold</th>
                  <th class="sorttable_nosort">Email</th>
                  <th class="sorttable_nosort">Notes</th>
                </tr>
              </thead>
              
              <tbody class='searchable'>
              
			<%
			
			SQL = "SELECT * FROM FS_ServiceMemos WHERE CurrentStatus = 'HOLD' AND RecordSubType = 'HOLD'"
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
							<tr class="low-priority">
							<% Response.write("<td sorttable_customkey=" & FormatAsSortableDateTime(rs("submissionDateTime")) & ">" & FormatDateTime(rs("submissionDateTime")) & "</td>")%>
							<td>
								<a href='../accountsreceivable/ReleaseServiceMemo.asp?memo=<%= rs.Fields("MemoNumber")%>'><%= rs.Fields("MemoNumber")%></a>
							</td>
							<td><%= rs.Fields("HoldReason") %></td>
							<td>
							</td>
							<td><%= rs.Fields("AccountNumber") %>
							</td>
							<td><%= rs.Fields("Company") %></td>
							<% If ViewType="Normal" Then %>
							<td><%= rs.Fields("ProblemDescription") %></td>
							<% Else %>
							<td>
							<%
								CompressLen = 27
								'See if there are linefeeds in there that need to come out
								If Instr(rs.Fields("ProblemDescription"),"<br>") <> 0 Then CompressLen = Instr(rs.Fields("ProblemDescription"),"<br>")
								If CompressLen > 27 Then CompressLen = 27
								If len(rs.Fields("ProblemDescription")) > CompressLen Then Response.Write(Left(rs.Fields("ProblemDescription"),CompressLen)) Else Response.Write(rs.Fields("ProblemDescription"))%>
							</td>
							<%End If
							Response.Write("<td><b>" & GetServiceTicketCurrentStage(rs.Fields("MemoNumber")) & "</b>") 
							Response.Write("<br>")
							Response.Write(GetServiceTicketSTAGEUser(rs.Fields("MemoNumber"),GetServiceTicketCurrentStage(rs.Fields("MemoNumber"))) & "</td>")

							If ElapsedTimeCalcMethod() = "Actual" Then
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
							Else
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
							%>
							</td>
							<td><a href="#"><i class="fa fa-envelope-o"></i></a></td>
							<td><a data-toggle="modal" data-target="#modalEditCustomerNotes" data-category-id="-2" data-cust-id="<%= rs.Fields("AccountNumber") %>" class="ole" rel="tooltip" style="cursor:pointer;"><i class="fa fa-file-text-o" aria-hidden="true"></i></a></td>
							</tr>
							<!-- eof table line !-->
						<%
						
							'***********************************************************						
						
				
						End If
						
						rs.movenext
				loop
				
			End If
	
			set rs = Nothing
			cnn8.close
			set cnn8 = Nothing

			Call Check_HOLD_Alerts
            %>
              
              
              
              
              
              </tbody>
            </table>
          </div>

    </div>	
    

</div>
<!-- eof row !-->    


<!-- pencil Modal -->
<div class="modal modal-wide fade" id="modalEditCustomerNotes" tabindex="-1" role="dialog" aria-labelledby="modalEditCustomerNotesLabel">

	<style>
	.modal-header {
	    padding: 15px;
	    border-bottom: 1px solid #e5e5e5;
	    min-height: 35px !important;
	}
	</style>
	
	<div class="modal-dialog" role="document">
						
		<div class="modal-content">	

			<input type="hidden" name="txtCategoryID" id="txtCategoryID">
			<input type="hidden" name="txtCustIDToPassToGenerateNotes" id="txtCustIDToPassToGenerateNotes">
			    
			<div id="modalEditCustomerNotesContent">
				<!-- Content for the modal will be generated and written here -->
				<!-- Content generated by Sub GetContentForCustomerNotesModal() in InsightFuncs_AjaxForBizIntelModals.asp -->
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


<!--#include file="../inc/footer-main.asp"-->
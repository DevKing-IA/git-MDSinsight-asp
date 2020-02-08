<!--#include file="..\..\..\inc\header.asp"-->
<%
'Check to see if there is a querystring value for 's'
'Any value at all indicates a failure to read the 
'seetings from UNIX
UnixOK = True
If Request.QueryString("s") <> "" Then UnixOK = False
%> 

<SCRIPT LANGUAGE="JavaScript">
<!--
    function showSuccess()
    {
    	
	        swal({
		        title: "Success",
		        text: "The night batch schedule has been submitted",
		        type: 'success',
		        confirmButtonText: 'OK'
		        });

      }  
 
// -->
</SCRIPT> 
<SCRIPT LANGUAGE="JavaScript">
<!--
    function showFailure()
    {
    	
	        swal({
		        title: "Error",
		        text: "An error was encountered while trying to send the night batch settings to UNIX. Please try again and contact support if the error persists.",
		        type: 'error',
		        confirmButtonText: 'OK'
		        });

      }  
 
// -->
</SCRIPT> 


<%
If MUV_INSPECT("NIGHTBATCHOK") = True Then
	If MUV_READ("NIGHTBATCHOK") = 1 Then
		Response.Write("<script language=javascript>showSuccess();</script>")
	End IF
	If MUV_READ("NIGHTBATCHOK") <> 1 Then
		Response.Write("<script language=javascript>showFailure();</script>")
	End IF
	MUV_REMOVE("NIGHTBATCHOK")
End If

Cronvar = MUV_READ("CRONQUERY")

'Init the vars
SundayOn = vbFalse
MondayOn = vbFalse
TuesdayOn = vbFalse
WednesdayOn = vbFalse
ThursdayOn = vbFalse
FridayOn = vbFalse
SaturdayOn = vbFalse
SundayStartTime = ""
MondayStartTime = "" 
TuesdayStartTime = ""
WednesdayStartTime = ""
ThursdayStartTime = ""
FridayStartTime = ""
SaturdayStartTime = ""

IF UnixOK = True Then 
	For x = 1 to Len(Cronvar)/5
		If Cronvar <> "" Then
			heldvar = left(Cronvar,5)
			Cronvar = Replace(Cronvar,heldvar,"")
			Select case Left(heldvar,1)
				Case 0
					SundayOn = vbTrue
					SundayStartTime = Mid(heldvar,2,2) & ":" & right(heldvar,2)
					SundayStartTime = FormatDateTime(SundayStartTime,4)
					dummy = MUV_Write("Orig_SundayOn",SundayOn)
					dummy = MUV_Write("Orig_SundayStartTime",SundayStartTime)
				Case 1
					MondayOn = vbTrue
					MondayStartTime = Mid(heldvar,2,2) & ":" & right(heldvar,2)
					MondayStartTime = FormatDateTime(MondayStartTime,4)
					dummy = MUV_Write("Orig_MondayOn",MondayOn)
					dummy = MUV_Write("Orig_MondayStartTime",MondayStartTime)
				Case 2
					TuesdayOn = vbTrue
					TuesdayStartTime = Mid(heldvar,2,2) & ":" & right(heldvar,2)
					TuesdayStartTime = FormatDateTime(TuesdayStartTime,4)
					dummy = MUV_Write("Orig_TuesdayOn",TuesdayOn)
					dummy = MUV_Write("Orig_TuesdayStartTime",TuesdayStartTime)
				Case 3
					WednesdayOn = vbTrue
					WednesdayStartTime = Mid(heldvar,2,2) & ":" & right(heldvar,2)
					WednesdayStartTime = FormatDateTime(WednesdayStartTime,4)
					dummy = MUV_Write("Orig_WednesdayOn",WednesdayOn)
					dummy = MUV_Write("Orig_WednesdayStartTime",WednesdayStartTime)
				Case 4
					ThursdayOn = vbTrue
					ThursdayStartTime = Mid(heldvar,2,2) & ":" & right(heldvar,2)
					ThursdayStartTime = FormatDateTime(ThursdayStartTime,4)
					dummy = MUV_Write("Orig_ThursdayOn",ThursdayOn)
					dummy = MUV_Write("Orig_ThursdayStartTime",ThursdayStartTime)
				Case 5
					FridayOn = vbTrue
					FridayStartTime = Mid(heldvar,2,2) & ":" & right(heldvar,2)
					FridayStartTime = FormatDateTime(FridayStartTime,4)
					dummy = MUV_Write("Orig_FridayOn",FridayOn)
					dummy = MUV_Write("Orig_FridayStartTime",FridayStartTime)
				Case 6
					SaturdayOn = vbTrue
					SaturdayStartTime = Mid(heldvar,2,2) & ":" & right(heldvar,2)
					SaturdayStartTime = FormatDateTime(SaturdayStartTime,4)
					dummy = MUV_Write("Orig_SaturdayOn",SaturdayOn)
					dummy = MUV_Write("Orig_SaturdayStartTime",SaturdayStartTime)
			End Select
		End If
	next

	' Run/Don't run report settings are in a different table
	SQL = "SELECT NightBatchRunReportTime, NightBatchRunReportEmail,NightBatchRunReportOn FROM Settings_Global"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	If not rs.EOF Then
		NightBatchRunReportTime = rs("NightBatchRunReportTime")
		NightBatchRunReportEmail = rs("NightBatchRunReportEmail")
		NightBatchRunReportOn = rs("NightBatchRunReportOn")
	End If
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
End If
%>

 

<!-- time picker !-->
<link rel="stylesheet" href="<%= BaseURL %>js/timepicker/timepicker/ui-1.10.0/ui-lightness/jquery-ui-1.10.0.custom.min.css" type="text/css" />
<link rel="stylesheet" href="<%= BaseURL %>js/timepicker/timepicker/jquery.ui.timepicker.css?v=0.3.3" type="text/css" />
<!--<script type="text/javascript" src="<%= BaseURL %>js/timepicker/timepicker/jquery-1.9.0.min.js"></script>-->
<script type="text/javascript" src="<%= BaseURL %>js/timepicker/timepicker/ui-1.10.0/jquery.ui.core.min.js"></script>
<script type="text/javascript" src="<%= BaseURL %>js/timepicker/timepicker/ui-1.10.0/jquery.ui.widget.min.js"></script>
<script type="text/javascript" src="<%= BaseURL %>js/timepicker/timepicker/ui-1.10.0/jquery.ui.tabs.min.js"></script>
<script type="text/javascript" src="<%= BaseURL %>js/timepicker/timepicker/ui-1.10.0/jquery.ui.position.min.js"></script>
<script type="text/javascript" src="<%= BaseURL %>js/timepicker/timepicker/jquery.ui.timepicker.js?v=0.3.3"></script>
<!-- eof time picker !-->


<style type="text/css">
	 
	 body{
		 overflow-x:hidden;
	 }
	 .page-header{
		 margin-top: 0px;
	 }
	 
	 
	  
	.days-hours  .table>tbody>tr>td, .table>tbody>tr>th, .table>tfoot>tr>td, .table>tfoot>tr>th, .table>thead>tr>td, .table>thead>tr>th{
		 border: 0px;
 	  }
 	  
 	  .red-text{
	 	  color:red;
 	  }
	 
	 h3{
		 margin: 0px;
		 padding: 0px;
		 line-height: 1;
	 }
	 
	 .ui-timepicker-table td a{
		padding: 3px;
		width:auto;
		text-align: left;
		font-size: 11px;
	}	
	
	.ui-timepicker-table .ui-timepicker-title{
		font-size: 13px;
	}
	
	.ui-timepicker-table th.periods{
		font-size: 13px;
	}
	
	.ui-widget-header{
		background: #193048;
		border: 1px solid #193048;
	}
	 
		[class^="col-"]{
		 margin-bottom:25px;
	  } 
	  
	  .custompick{
		  width: 45%;
	  }
	  
	  .custom-row{
		  margin-top: 10px;
	  }
	  
	  .modal-link{
	cursor: pointer;
}

.table-history .table>thead>tr>th {
    vertical-align: bottom;
    border-bottom: 2px solid #ddd;
}

table.sortable th:not(.sorttable_sorted):not(.sorttable_sorted_reverse):not(.sorttable_nosort):after { 
    content: " \25B4\25BE" 
}
</style>

<br>
<h1 class="page-header"><i class="fa fa-clock-o"></i> Night Batch</h1>
	
<!-- content starts here !-->
<div class="row">

			<!-- tabs start here !-->
		<div class="global-tabs">
			<ul class="nav nav-tabs responsive-tabs ">
			    <li role="presentation" class="active"><a href="#schedule" aria-controls="manage" role="tab" data-toggle="tab">Schedule</a></li>
			    <li role="presentation"><a href="#history" aria-controls="manage" role="tab" data-toggle="tab">History</a></li>
			</ul>

			<!-- Shcedule tab !-->
			<div class="tab-content">
			
				<div class="tab-pane active" id="schedule">
				<% If UnixOK = True Then %>
						<form method="post" action="submitNBschedule.asp" name="frmNBSchedule">
						<!-- left column with days / hours starts here !-->
						<div class="col-lg-6 days-hours">
							<div class="table-responsive"> 
								<br><font color="red">*Note: The cutoff for invoice posting dates is <strong>7:00AM</strong></font><br>
								<font color="red">If your night batch program does not include posting invoices</font><br>
								<font color="red">you can ignore the posting date info shown.</font><br>
								<table class="table">
									
									<!-- titles !-->
									<thead>
										<tr>
											<td>&nbsp;</td>
											<td width="20%"><strong>Start Time</strong></td>
											<td><strong>On/Off</strong></td>
											<td><strong>Posting Date</strong></td>
										</tr>
									</thead>
									<!-- eof titles !-->
									
									<tbody>
					 	
										<!-- line !-->
										<tr>
											<td>Monday</td>
											<td><input type="text" id="txtmonday" name="txtmonday" class="form-control" value="<%=MondayStartTime%>"  onchange="monChanged1()"/></td>
											<td>
												<% If MondayOn = vbTrue Then %>
													<input type="checkbox" id="chkmonday" name="chkmonday" onclick='monChanged()' checked>
												<% Else %>
													<input type="checkbox" id="chkmonday" name="chkmonday" onclick='monChanged()'>
												<% End If%>
											</td>											
												<td>
													<div id="pnlMonday" style="display: none;">
														<label name="lblMon" id="lblMon" >post date</label>.
													</div>
												</td>
										</tr>
										<!-- eof line !-->
					 	
										<!-- line !-->
										<tr>
											<td>Tuesday</td>
											<td><input type="text" id="txttuesday"  name="txttuesday"   class="form-control" value="<%=TuesdayStartTime%>" onchange="tueChanged1()"/></td>
											<td>
												<% If TuesdayOn = vbTrue Then %>
													<input type="checkbox" id="chktuesday" name="chktuesday" onclick='tueChanged()' checked>
												<% Else %>
													<input type="checkbox" id="chktuesday" name="chktuesday" onclick='tueChanged()' >
												<% End If%>
											</td>
											
											<td>
												<div id="pnlTuesday" style="display: none;">
													<label name="lblTues" id="lblTues" >post date</label>
												</div>
											</td>
											
											
										</tr>
										<!-- eof line !-->
					 	
										<!-- line !-->
										<tr>
											<td>Wednesday</td>
											<td><input type="text" id="txtwednesday"  name="txtwednesday"   class="form-control" value="<%=WednesdayStartTime%>" onchange="wedChanged1()"/></td>
											<td>
												<% If WednesdayOn = vbTrue Then %>
													<input type="checkbox" id="chkwednesday" name="chkwednesday" onclick='wedChanged()' checked>
												<% Else %>
													<input type="checkbox" id="chkwednesday" name="chkwednesday" onclick='wedChanged()' >
												<% End If%>
											</td>
											
											<td>
												<div id="pnlWednesday" style="display: none;">
													<label name="lblWed" id="lblWed" >post date</label>
												</div>
											</td>
	
											
											
										</tr>
										<!-- eof line !-->
					 	
										<!-- line !-->
										<tr>
											<td>Thursday</td>
											<td><input type="text" id="txtthursday"  name="txtthursday"   class="form-control" value="<%=ThursdayStartTime%>" onchange="thuChanged1()" />	</td>
											<td>
												<% If ThursdayOn = vbTrue Then %>
													<input type="checkbox" id="chkthursday" name="chkthursday"  onclick='thuChanged()' checked>
												<% Else %>
													<input type="checkbox" id="chkthursday" name="chkthursday" onclick='thuChanged()' >
												<% End If%>
											</td>
											
											<td>
												<div id="pnlThursday" style="display: none;">
													<label name="lblThu" id="lblThu" >post date</label>
												</div>
											</td>
											
											
										</tr>
										<!-- eof line !-->
					 	
										<!-- line !-->
										<tr>
											<td>Friday</td>
											<td><input type="text" id="txtfriday"  name="txtfriday"   class="form-control" value="<%=FridayStartTime%>" onchange="friChanged1()"/></td>
											<td>
												<% If FridayOn = vbTrue Then %>
													<input type="checkbox" id="chkfriday" name="chkfriday"  onclick='friChanged()' checked>
												<% Else %>
													<input type="checkbox" id="chkfriday" name="chkfriday" onclick='friChanged()' >
												<% End If%>
											</td>
											
											<td>
												<div id="pnlFriday" style="display: none;">
													<label name="lblFri" id="lblFri" >post date</label>
												</div>
											</td>
	
											
											
										</tr>
										<!-- eof line !-->
					 	
										<!-- line !-->
										<tr>
											<td>Saturday</td>
											<td><input type="text" id="txtsaturday"  name="txtsaturday"   class="form-control" value="<%=SaturdayStartTime%>" onchange="satChanged1()"/>
											<td>
												<% If SaturdayOn = vbTrue Then %>
													<input type="checkbox" id="chkSaturday" name="chkSaturday"  onclick='satChanged()' checked>
												<% Else %>
													<input type="checkbox" id="chkSaturday" name="chkSaturday" onclick='satChanged()' >
												<% End If%>
											</td>
											
											<td>
												<div id="pnlSaturday" style="display: none;">
													<label name="lblSat" id="lblSat" >post date</label>
												</div>
											</td>
	
											
											
										</tr>
										<!-- eof line !-->
					 	
										<!-- line !-->
										<tr>
											<td>Sunday</td>
											<td><input type="text" id="txtsunday"  name="txtsunday"   class="form-control" value="<%=SundayStartTime%>" onchange="sunChanged1()"/></td>
											<td>
												<% If SundayOn = vbTrue Then %>
													<input type="checkbox" id="chkSunday" name="chkSunday"  onclick='sunChanged()' checked>
												<% Else %>
													<input type="checkbox" id="chkSunday" name="chkSunday" onclick='sunChanged()' >
												<% End If%>
											</td>
											
											<td>
												<div id="pnlSunday" style="display: none;">
													<label name="lblSun" id="lblSun" >post date</label>
												</div>
											</td>
	
										
										</tr>
										<!-- eof line !-->
					 	
								 	</tbody>
							 	</table>
						 	</div>
						 	
						 	<p >
			 	   			    <a href="#">
							    	<button type="button" class="btn btn-default">Cancel</button>			    
				 				</a>
								<button type="submit" class="btn btn-primary"><i class="far fa-save"></i> Save</button>
						 	</p>
					 	</div>
	
					 	<div class="col-lg-6">
		 					<br><br>
						 	<div class="row">		 	
						 	</div>
					 </div>
				</form>
			<% Else  'Else for UnixOK %>
					<div class="col-lg-6 days-hours">
						<div class="table-responsive"> 
							<br>The schedule tab is unavailable becuase Insight was<br>
							unable to read the night batch settings from <%=GetTerm("Backend")%>.<br>
							Please try again and contact techincal support if<br> the problem continues.<br><br>
							<strong>You can still view the night batch logs via the history tab.</strong><br>
						</div>	
					</div>
			<% End If%>
		</div>
		
		<!-- history tab !-->
		<div class="tab-pane" id="history">
					
				<div class="table-responsive table-history">
		            <table    class="table table-striped table-hover sortable">
		              <thead>
		                <tr>
		                  <th>Log Received</th>
		                  <th class="sorttable_nosort">Log Data</th>
		                </tr>
		              </thead>
		              <tbody>
              
						<%
			
						SQL = "SELECT * FROM SC_NightBatchLogs Order by RecordCreationDateTime desc"
		
						Set cnn8 = Server.CreateObject("ADODB.Connection")
						cnn8.open (Session("ClientCnnString"))
						Set rs = Server.CreateObject("ADODB.Recordset")
						rs.CursorLocation = 3 
						Set rs = cnn8.Execute(SQL)
				
						If not rs.EOF Then
						
							DynamicFormCounter= 0
		
							Do While Not rs.EOF
				
					        %>
								<!-- table line !-->
								<tr>
									<td sorttable_customkey=' & FormatAsSortableDateTime(rs("RecordCreationDateTime"))'><%= rs.Fields("RecordCreationDateTime")%></td>
									<td>
										<%
										DisplayDataArray = Split(rs.Fields("NightBatchLogData1"),vbCRLF)
										If Ubound(DisplayDataArray) < 3 Then TopLimit = Ubound(DisplayDataArray) Else TopLimit = 3
										For x = 1 to TopLimit
											Response.Write(DisplayDataArray(x) & "<br>")
										Next
										%>
										<a class="modal-link" data-toggle="modal" data-target=".bs-example-modal-lg-customize<%=DynamicFormCounter%>"><strong>Click here to view the full log file</strong></a>
										<div class="modal fade bs-example-modal-lg-customize<%=DynamicFormCounter%>" tabindex="-1" role="dialog" aria-labelledby="myLargeModalLabel" aria-hidden="true">
										<!--#include file="history_modal.asp"-->
									</td>
							   	</tr>
								<%
								
								DynamicFormCounter = DynamicFormCounter + 1 
								rs.movenext
							loop
						Else
							Response.Write("<tr><td>There is currently no night batch history<br></td></tr>")
						End If
						set rs = Nothing
						cnn8.close
						set cnn8 = Nothing
			            %>
					</tbody>
				</table>
				</div>
			</div>
		<!-- eof history tab !-->
	</div> 	

<!-- time picker js !-->
<script type="text/javascript">
	$('#txtmonday').timepicker();
	$('#txttuesday').timepicker();
	$('#txtwednesday').timepicker();
	$('#txtthursday').timepicker();
	$('#txtfriday').timepicker();
	$('#txtsaturday').timepicker();
	$('#txtsunday').timepicker();
	$('#txtemailtime').timepicker();
</script>
<!-- eof time picker js !-->



<!--#include file="dayText.js"-->
 
<!--#include file="..\..\..\inc\footer-main.asp"-->
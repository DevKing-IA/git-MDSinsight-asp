<!--#include file="../inc/header.asp"-->


<% MemoNumber = Request.QueryString("memo") 
If MemoNumber = "" Then Response.Redirect(BaseURL)
%>

<link rel="stylesheet" type="text/css" href="<%= BaseURL %>css/tabs.css" />

          

<!-- time picker !-->
 	<link rel="stylesheet" href="<%= BaseURL %>js/timepicker/timepicker/ui-1.10.0/ui-lightness/jquery-ui-1.10.0.custom.min.css" type="text/css" />
    <link rel="stylesheet" href="<%= BaseURL %>js/timepicker/timepicker/jquery.ui.timepicker.css?v=0.3.3" type="text/css" />
    <script type="text/javascript" src="<%= BaseURL %>js/timepicker/timepicker/ui-1.10.0/jquery.ui.core.min.js"></script>
    <script type="text/javascript" src="<%= BaseURL %>js/timepicker/timepicker/ui-1.10.0/jquery.ui.widget.min.js"></script>
    <script type="text/javascript" src="<%= BaseURL %>js/timepicker/timepicker/ui-1.10.0/jquery.ui.tabs.min.js"></script>
    <script type="text/javascript" src="<%= BaseURL %>js/timepicker/timepicker/ui-1.10.0/jquery.ui.position.min.js"></script>
    <script type="text/javascript" src="<%= BaseURL %>js/timepicker/timepicker/jquery.ui.timepicker.js?v=0.3.3"></script>
 <!-- eof time picker !-->


<!-- date picker !-->
	<link rel="stylesheet" href="<%= baseURL %>css/datepicker/BeatPicker.min.css"/>
	<script src="<%= baseURL %>js/datepicker/BeatPicker.min.js"></script>
<!-- eof date picker !-->
   
   
 
<style>
	
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

.beatpicker-clear{
	display: block;
	text-indent:-9999em;
	line-height: 0;
	visibility: hidden;
} 

.beatpicker ul{
    display: block;
    list-style-type: none;
    -webkit-margin-before: 0px;
    -webkit-margin-after: 0px;
    -webkit-margin-start: 0px;
    -webkit-margin-end: 0px;
    -webkit-padding-start: 0px;
} 

.beatpicker li.cell{
	margin: 2%;
} 

 

 	.alert{
 		padding: 6px 12px;
 		margin-bottom: 0px;
	}
	
	.form-control{
		margin-bottom: 20px;
	}
	
	a:hover{
		text-decoration: none;
	}
	
	[class^="col-"]{
	 margin-bottom:25px;
  } 
  
  .custom-hr{
height: 3px;
margin-left: auto;
margin-right: auto;
background-color:#183049;
color:#183049;
border: 0 none;
}

.control-label{
	padding-top: 5px;
}

.do-not-send-alert{
 	padding: 5px 10px 5px 10px;
	display: inline-block;
	background: #fff200;
}

.table-info .table>tbody>tr>td, .table>tbody>tr>th, .table>tfoot>tr>td, .table>tfoot>tr>th, .table>thead>tr>td, .table>thead>tr>th{
	border: 0px;
	font-weight: bold;
	line-height: 0.8;
}

.the-information{
	font-size:12px;
} 

.date-col{
	width:15%;
} 

.stage-col{
	width:10%;
} 

.user-col{
	width:10%;
} 


 
.alert{
	display:block;
	float:left;
	margin:0px 5px 5px 0px;
}

	</style>


<h1 class="page-header"><i class="fa fa-wrench"></i> Close or Cancel Service Memo</h1>

 

<%
SQL = "SELECT * FROM FS_ServiceMemos where MemoNumber = '" & MemoNumber  & "'"
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	ServiceMemoRecNumber = rs("ServiceMemoRecNumber")
	CurrentStatus = rs("CurrentStatus")
	RecordSubType = rs("RecordSubType")
	SubmittedByName = rs("SubmittedByName")
	AccountNumber = rs("AccountNumber")
	Company = rs("Company")
	ProblemLocation = rs("ProblemLocation")
	SubmittedByPhone = rs("SubmittedByPhone")
	SubmittedByEmail = rs("SubmittedByEmail")
	SubmissionDateTime = rs("SubmissionDateTime")
	ProblemDescription = rs("ProblemDescription")
	Mode = rs("Mode")
	SubmissionSource = rs("SubmissionSource")
	UserNoOfServiceTech = rs("UserNoOfServiceTech")
	ReleasedDateTime = rs("ReleasedDateTime")
	ReleasedByUserNo = rs("ReleasedByUserNo")
	ReleasedNotes = rs("ReleasedNotes")
End If
	
set rs = Nothing
cnn8.close
set cnn8 = Nothing

If SubmittedByName = "" Then SubmittedByName = "Not provided"
If SubmittedByPhone = "" Then SubmittedByPhone = "Not provided"
If SubmittedByEmail = "" Then SubmittedByEmail = "Not provided"
If ProblemLocation = "" Then ProblemLocation = "Not provided"
If ProblemDescription = "" Then ProblemDescription = "Not provided"

%>


	
	<form method="POST" action="editservicememo_submit.asp" name="frmEditServiceMemo">		    
      

        <input type="hidden" id="txtPrintedName" name="txtPrintedName" value="N/A Closed From MDS Insight"  class="form-control last-run-inputs">

        
 	        <!-- row !-->		
	        <div class="row ">
		        

		        <!--account number !-->
		        <div class="col-lg-6 col-md-4 col-sm-12 col-xs-12">
		        	<%SelectedCustomer = AccountNumber %>
					<!--#include file="../inc/commonCustomerDisplay.asp"-->
			    </div>
		        <!-- eof account number !-->
		        
		        <!-- company name !-->
		        <div class="col-lg-4 col-md-4 col-sm-12 col-xs-12">
			        
			        
				        
			        <div class="alert alert-info" role="alert"><strong>Ticket#: <%= MemoNumber %></strong></div>
				        
				        <div class="alert alert-warning" role="alert"> 
			        <% If advancedDispatchIsOn() Then %>
						<strong>Stage: <%= GetServiceTicketCurrentStage(MemoNumber)%></strong>
					<%End If%>
			        <input type="hidden" id="txtMemoNumber" name="txtMemoNumber" value="<%= MemoNumber %>"  class="form-control last-run-inputs">
				        </div>
				        
				        </div>
 		        
		        <!-- eof company name !-->
		        
		        
	 		        
		        <!-- the information !-->
		         <div class="col-lg-4 col-md-4 col-sm-12 col-xs-12 the-information">
		         <div class="row">
		         
		        	<!-- Contact Name !-->
			    <div class="col-lg-12">
			        <strong>Contact Name</strong><br>
			        <% =SubmittedByName %>
			        </div>
			    	<!-- Contact Name !-->
	
			    	<!-- Contact Phone !-->
			    	  <div class="col-lg-12">
				    	 <strong>Contact Phone</strong><br>
				    	 <% =SubmittedByPhone %>
			        </div>
			    	<!-- Contact Phone !-->
			    	

<!-- Problem Location !-->
			   <div class="col-lg-12">
			        <strong>Problem Location</strong><br>
					<% =ProblemLocation %>

			        </div>
			    	<!-- Problem Location !-->
	  					    	
			    	<!-- Description of problem !-->
			    	 <div class="col-lg-12">
 					<strong>Problem Description</strong><br>
					<% =ProblemDescription %>
			    	 </div>
 			    	<!-- Description of problem !-->
</div>
</div>
<!-- eof the information !-->						        	
						        	
						        		
		        </div>
 <!-- eof row !-->

 <!-- main row !-->
 <div class="row">
	 
	
			
			
			
		        		         
		</div>
		<!-- eof main row !-->
		
		<div class="row">
			<div class="col-lg-12">
				<hr class="custom-hr">
			</div>
		</div>
		
		<!-- main row !-->
 <div class="row">
	 
	 <!-- left col !-->
	 <div class="col-lg-4 col-md-4 col-sm-12 col-xs-12">
	
	 
 		        <!-- row !-->			
			    <div class="row">

			    	<!-- Close or Cancel !-->
					<div class="col-lg-12">
						<%If userIsServiceManager(Session("userNo")) or userIsAdmin(Session("userNo")) Then%> 
							<label><input type="radio" name="optradio1" id="close" value="Close" checked>  Close</label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<% 'We need to see if advanced dispatching is on & if so, which stage are we in
							If advancedDispatchIsOn() Then
								cStage = GetServiceTicketCurrentStage(MemoNumber)
								'Can only be cancelled in the following stages
								If cStage = "Received" or cStage = "Released" or cStage = "Under Review" or cStage = "Dispatched" or cStage = "Dispatch Acknowledged" or cStage = "Dispatch Declined" or cStage = "En Route" Then%>
									<label><input type="radio" name="optradio1" id="cancel" value="Cancel"> Cancel</label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								<%End If
							Else%>
								<label><input type="radio" name="optradio1" id="cancel" value="Cancel"> Cancel</label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;							
							<%End If%>
					
							<input type='checkbox' class='check' id='chkDoNotEmail' name='chkDoNotEmail'>&nbsp;<strong class="do-not-send-alert">Do not send a close email to the customer</strong>
						<%Else%>
							<label><input type="radio" name="optradio1" id="cancel" value="Cancel" checked readonly> Cancel</label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<input type='checkbox' class='check' id='chkDoNotEmail' name='chkDoNotEmail'>&nbsp;<strong class="do-not-send-alert">Do not send a close email to the customer</strong>
						<%End If%>
			        </div>
 			    	<!-- Close or Cancel !-->
 			    	
 			    	<!-- service notes !-->
 			    	<div class="col-lg-12">
	 			    	<label>Service Notes</label>
	 			    	<textarea name="ServiceNotes" id="ServiceNotes" spellcheck="True" class="form-control" rows="6"></textarea>
 			    	</div>
 			    	<!-- eof service notes !-->
 			    	
 			    	<% If advancedDispatchIsOn() <> True Then ' if advanced dispacth is on, we already know the tech%>
	 			    	<!-- field tech !-->
	 			    	<%If userIsServiceManager(Session("userNo")) or userIsAdmin(Session("userNo")) Then%> 
	 			    	<div class="col-lg-12">
	    					<label>Field Tech</label>
							<select name="selFieldTech" id="selFieldTech" class="form-control">
							<option>Select Field Tech </option>
								<%	
									'Fixit
								' cheap fix to let adam henchel see service stuff wihtout being a service manager


								SQL = "SELECT * FROM " & MUV_Read("SQL_Owner") & ".tblUsers WHERE (userType = 'Field Service' OR userType = 'Service Manager' OR userNo=56) and userArchived <> 1 Order By userLastName"
								
								Set cnn8 = Server.CreateObject("ADODB.Connection")
								cnn8.open (Session("ClientCnnString"))
								Set rs = Server.CreateObject("ADODB.Recordset")
								rs.CursorLocation = 3 
								Set rs = cnn8.Execute(SQL)
		
								If not rs.EOF Then
	
									Do While Not rs.EOF
										userFirstName = rs("userFirstName")
										userLastName = rs("userLastName")
										userDisplayName = rs("userDisplayName")
										userEmail = rs("userEmail")
										userNo = rs("UserNo")
										
										%><option value='<%=userNo%>'><%=userFirstName%>&nbsp;<%=userLastName%>&nbsp;---<%=userDisplayName%>&nbsp;---<%=userEmail%></option><%
										
										rs.MoveNext
									Loop
	
								End If
								%>
							</select>
	  					 </div>
	  					 <%End If%>
	 			    	<!-- eof field tech !-->
	 			    	<%End If%>
 			    	
 			    		
			    	
			    	</div>
			    <!-- eof row !-->
			    	
			    	   
	 </div>
		    	 <!-- eof left col !-->
		    	 
		    	 
 		       
  	
  			<!-- right col !-->
  			<div class="col-lg-3 col-md-3 col-sm-12 col-xs-12">  
	  			<%If userIsServiceManager(Session("userNo")) or userIsAdmin(Session("userNo")) Then%> 
					<div class="row">
					
						<div class="col-lg-12">
							<strong>Asset Location</strong><br>
							<small>To update the location of an asset, fill in the info below.</small>
						</div>
						
						<!--asset !-->
						<div class="form-group ">
							<label class="col-sm-4 control-label">Asset Tag #</label>
							<div class="col-sm-8">
								<input type="text" name="txtAssetTagNumber" id="txtAssetTagNumber" class="form-control" >
							</div>
						</div>
						<!-- eof asset !-->	  
						
						<!--location !-->
						<div class="form-group ">
							<label class="col-sm-4 control-label">Location #</label>
							<div class="col-sm-8">
								<input type="text"  name="txtAssetLocation" id="txtAssetLocation" class="form-control" >
							</div>
						</div>
						<!-- eof location !-->					    	
					
					</div>
		  		<%End IF%>
  			</div>
			<!-- eof right col !-->
			
			<!-- date and time !-->
			<div class="col-lg-4">
				<div class="row">
					
					<!--date !-->
					<div class="form-group ">
						<label class="col-sm-4 control-label">Close/Cancel Date</label>
						<div class="col-sm-8">
							<input type="text" id="txtCloseCancelDate" name="txtCloseCancelDate" value="<%=Date() %>"  class="form-control last-run-inputs" data-beatpicker="true" data-beatpicker-format="['MM','DD','YYYY'],separator:'/'">
						</div>
					</div>
 			    	<!-- eof date !-->
 			    	
					<!--time !-->
					<div class="form-group ">
						<label class="col-sm-4 control-label">Close/Cancel Time</label>
						<div class="col-sm-3">
							
								<input type="text" id="timepicker"  name="timepicker" class="form-control" value="<%=Hour(Time()) & ":" & Minute(Time()) %>" />  
														
							<!--<input type="text" id="txtCloseCancelTime" name="txtCloseCancelTime" value="<%=Time() %>"  class="form-control" >!-->
						</div>
					</div>
					<!-- eof time !-->
					
				</div></div>
					<!-- eof date and time !-->
		        		         
		</div>
		<!-- eof main row !-->


			<% MDG_MemoNumber = MemoNumber %>
			<!--#include file="memo_details_grid.asp"-->





			<div class="row">
			
			<div class="col-lg-12">	
			    <% If Instr(ucase(Request.ServerVariables ("HTTP_REFERER")),"CUSTOMERSERVICE") <> 0 Then %>
    			    <a href="<%=Request.ServerVariables("HTTP_REFERER")%>">
			    	<button type="button" class="btn btn-default">&lsaquo; Go Back To <%=GetTerm("Customer")%> Notes</button>
			    	<input type="hidden" id="txtReturnPath" name="txtReturnPath" value="Customer Service"  class="form-control last-run-inputs">
			    <% Elseif Instr(ucase(Request.ServerVariables ("HTTP_REFERER")),"DISPATCHCENTER") <> 0 Then%>	
    			    <a href="<%= BaseURL %>service/dispatchcenter/main.asp">
			    	<button type="button" class="btn btn-default">&lsaquo; Go Back To Dispatch Center</button>
			    	<input type="hidden" id="txtReturnPath" name="txtReturnPath" value="DispatchCenter"  class="form-control last-run-inputs">
			    <% Elseif Instr(ucase(Request.ServerVariables ("HTTP_REFERER")),"SERVICEBOARD.ASP") <> 0 Then%>	
    			    <a href="<%= BaseURL %>service/serviceBoard.asp">
			    	<button type="button" class="btn btn-default">&lsaquo; Go Back To Service Board</button>
			    	<input type="hidden" id="txtReturnPath" name="txtReturnPath" value="ServiceBoard"  class="form-control last-run-inputs">
			    <% Else %>
	   			    <a href="<%= BaseURL %>service/main.asp">
			    	<button type="button" class="btn btn-default">&lsaquo; Go Back To Service Screen</button>
			    	<input type="hidden" id="txtReturnPath" name="txtReturnPath" value=""  class="form-control last-run-inputs">			    
			    <%End IF%>
				</a>
				<button type="submit" class="btn btn-primary"><i class="fa fa-upload"></i> Submit</button>

			</div>
			
 			
			</div>
			<!-- eof row !-->    

		</form>

 
<!-- time picker js !-->
 	<script type="text/javascript">
  $('#timepicker').timepicker();
  </script>
	    
	 
 <!-- eof time picker js !-->


   
<!--#include file="../inc/footer-main.asp"-->

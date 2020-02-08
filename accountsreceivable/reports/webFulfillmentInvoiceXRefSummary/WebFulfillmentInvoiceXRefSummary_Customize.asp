<SCRIPT LANGUAGE="JavaScript">
<!--
    function validateCustomizeForm()
    {
    	var WebOrderOrMDSInvoiceFilter = $("#selWebOrderOrMDSInvoiceFilter option:selected").val();
    	var DatesOrPeriods = document.frmWebFulfillInvXRef_Customize.optDatesorPeriods.value;
    	
    	if (WebOrderOrMDSInvoiceFilter == '' && DatesOrPeriods != '') 				
    	{		
			swal("Please select to filter by OCS Web Order Date or MDS Invoice Date.")
			return false;
		}
    		
     	if (WebOrderOrMDSInvoiceFilter != '' && DatesOrPeriods == '') 				
    	{		
			swal("Please select either the Dates or Periods option for date filtering.")
			return false;
		}
   		
		//************************************************************************
		//These are the checks for the text boxes for OCS Web Orders
		//************************************************************************
    	//If they chose the option button for dates, make sure the dates are filled-in
		if (DatesOrPeriods == "Dates")
		{
			if (document.frmWebFulfillInvXRef_Customize.txtRangeStartDate.value == "" || 
				document.frmWebFulfillInvXRef_Customize.txtRangeEndDate.value == "") 
				{		
					swal("Please make sure date range is complete for date filtering.")
					return false;
				}

	    	// First start range specified but no end range
	        if (document.frmWebFulfillInvXRef_Customize.txtRangeStartDate.value != "")
	        {
				if (document.frmWebFulfillInvXRef_Customize.txtRangeEndDate.value == "")
				{
					swal("You specified a range start date but not a range end date for date filtering.");
					return false;
				}
			}
				
	    	// First end range specified but no start range
	        if (document.frmWebFulfillInvXRef_Customize.txtRangeEndDate.value != "")
	        {
				if (document.frmWebFulfillInvXRef_Customize.txtRangeStartDate.value == "")
				{
					swal("You specified a range end date but not a range start date for date filtering.");
					return false;
				}
			}

			// First range dates, start is greater than end
	        if (document.frmWebFulfillInvXRef_Customize.txtRangeStartDate.value != "")
	        {
	        	var FStart = new Date(document.frmWebFulfillInvXRef_Customize.txtRangeEndDate.value);
	        	var FEnd = new Date(document.frmWebFulfillInvXRef_Customize.txtRangeStartDate.value);
				if (FStart < FEnd)
				{
					swal("The start date specified must be older than the end date for date filtering.");
					return false;
				}
	        }
	
		}
		

		//********************************************
		//These are the checks for the drop down boxes
		//********************************************
    	//If they chose the option button for period, make sure they selected 2 periods
		if (DatesOrPeriods == "Periods")
		{
	    	ocsPeriodStartValue = document.frmWebFulfillInvXRef_Customize.selPeriodStart.value;
	    	ocsPeriodEndValue = document.frmWebFulfillInvXRef_Customize.selPeriodEnd.value;
	    	
	    	// Start period empty
	        if (ocsPeriodStartValue == "")
				{
					swal("Please select a starting period from the drop-down for period filtering.");
					return false;
				}	
	    	// End period empty
	        if (ocsPeriodEndValue == "")
				{
					swal("Please select an ending period from the drop-down for period filtering.");
					return false;
				}	
				
	    	// Check for start period less than end period
	        if (ocsPeriodStartValue !== "" && ocsPeriodEndValue !== "")
				{
					if (ocsPeriodStartValue > ocsPeriodEndValue)
					{
						swal("The starting period is greater than the ending period for period filtering. Please select a greater end period.");
						return false;
					}
				}					    	
		}
				
        return true;  
    }
// -->
</SCRIPT>

<style type="text/css">
	 .ativa-scroll{
	 max-height: 300px
 }
 
 .sellperiodform{
	 margin-top: 20px;
 }
 
 
 
 .categories-checkboxes{
	 font-size: 12px;
 }
 
  .categories-checkboxes input{
	  margin-right: 6px;
  }
  
.container-modal{
	  border-bottom: 1px solid #e5e5e5;
	  margin-bottom: 10px;
   }
	</style>

<!-- modal scroll !-->
<script type="text/javascript">

  $(document).ready(ajustamodal);
  $(window).resize(ajustamodal);
  function ajustamodal() {
    var altura = $(window).height(); //value corresponding to the modal heading + footer
    $(".ativa-scroll").css({"height":"500px","overflow-y":"auto"});
  }
  
  $(document).ready(function() {
  
  		$("#resetModalFormFields").on('click',function(){
			$('#frmWebFulfillInvXRef_Customize').trigger("reset");
			$('option').attr('selected', false);
			$(':radio').prop('checked', false);
			$('input:checkbox').removeAttr('checked');
			$(':input').val('');
			$('.left-column').removeClass("activefilter");
			$("#InvoiceReportRange").val(''); 
		});	
		
		$("#btnClearDateFilters").click(function() {
			$('#selWebOrderOrMDSInvoiceFilter').children().removeProp('selected');
			$("#input[name='optDatesorPeriods']").removeAttr("checked");
			$("input[name='optDatesorPeriods']").prop("checked", false);
			$("#txtRangeStartDate").val('');
			$("#txtRangeStartDate").val('');
			$("#InvoiceReportRange").val('');
			$('#selPeriodStart').children().removeProp('selected');
			$('#selPeriodEnd').children().removeProp('selected');
		});	

		$("#btnClearClassCodeFilter").click(function() {
			$("input[name='chkClassCode']").removeAttr("checked");
		});	
		
		$("#btnSelectAllClassCodes").click(function() {
			$('input[name="chkClassCode"]').attr('checked','checked');
			$('input[name="chkClassCode"]').prop("checked", true);
		});	
		

		$("#btnClearCustomerTypeFilter").click(function() {
			$("input[name='chkCustomerType']").removeAttr("checked");
		});	
		
		$("#btnSelectAllCustomerTypes").click(function() {
			$('input[name="chkCustomerType"]').attr('checked','checked');
			$('input[name="chkCustomerType"]').prop("checked", true);
		});	

		$("#btnClearInvoicedFilter").click(function() {
			$("#input[name='chkShowOrdersThatAreInvoiced']").removeAttr("checked");
			$('input[name="chkShowOrdersThatAreInvoiced"]').prop("checked", false);
			$("#input[name='chkShowOrdersThatAreNotInvoiced']").removeAttr("checked");
			$('input[name="chkShowOrdersThatAreNotInvoiced"]').prop("checked", false);
		});	
		$("#btnSelectAllInvoices").click(function() {
			$('input[name="chkShowOrdersThatAreInvoiced"]').attr('checked','checked');
			$('input[name="chkShowOrdersThatAreInvoiced"]').prop("checked", true);
			$('input[name="chkShowOrdersThatAreNotInvoiced"]').attr('checked','checked');
			$('input[name="chkShowOrdersThatAreNotInvoiced"]').prop("checked", true);
		});	
		

		$("#btnClearRemarksFilter").click(function() {
			$("#input[name='chkShowOrdersWithRemarks']").removeAttr("checked");
			$('input[name="chkShowOrdersWithRemarks"]').prop("checked", false);
			$("#input[name='chkShowOrdersWithoutRemarks']").removeAttr("checked");
			$('input[name="chkShowOrdersWithoutRemarks"]').prop("checked", false);
		});	
		$("#btnSelectAllRemarks").click(function() {
			$('input[name="chkShowOrdersWithRemarks"]').attr('checked','checked');
			$('input[name="chkShowOrdersWithRemarks"]').prop("checked", true);
			$('input[name="chkShowOrdersWithoutRemarks"]').attr('checked','checked');
			$('input[name="chkShowOrdersWithoutRemarks"]').prop("checked", true);
		});	
		

		$("#btnClearHiddenFilter").click(function() {
			$("#input[name='chkShowOrdersThatAreHidden']").removeAttr("checked");
			$('input[name="chkShowOrdersThatAreHidden"]').prop("checked", false);
		});	
	
		
	    $(document).on('change','[name="selPeriodStart"]',function(){

			var selectedPeriod = $(this).find(":selected").val();
			
			if (selectedPeriod != '') {
				$('input[type=radio]').removeAttr('checked'); 
				$('input[name="optDatesorPeriods"][value="Periods"]').attr('checked','checked');
				$('input[name="optDatesorPeriods"][value="Dates"]').removeAttr('checked');
				$('input[name="optDatesorPeriods"][value="Periods"]').prop("checked", true);
			}
			else {
				$('input[type=radio]').removeAttr('checked'); 
				$('input[name="optDatesorPeriods"][value="Periods"]').removeAttr('checked');
				$('input[name="optDatesorPeriods"][value="Dates"]').attr('checked','checked');
				$('input[name="optDatesorPeriods"][value="Dates"]').prop("checked", true);
			}
			
		}); 
	        
		$('#InvoiceReportRange').on('show.daterangepicker', function(ev, picker) {
			$('input[type=radio]').removeAttr('checked'); 
			$('input[name="optDatesorPeriods"][value="Periods"]').removeAttr('checked');
			$('input[name="optDatesorPeriods"][value="Dates"]').attr('checked','checked');
			$('input[name="optDatesorPeriods"][value="Dates"]').prop("checked", true);
		}); 
		
		$('#InvoiceReportRange').on('apply.daterangepicker', function(ev, picker) {
			$('input[type=radio]').removeAttr('checked'); 
			$('input[name="optDatesorPeriods"][value="Periods"]').removeAttr('checked');
			$('input[name="optDatesorPeriods"][value="Dates"]').attr('checked','checked');
			$('input[name="optDatesorPeriods"][value="Dates"]').prop("checked", true);
			
		});  
 
   });
   
</script>
<!-- eof modal scroll !-->

<%
'************************
'Read Settings_Reports
'************************

SQLReportName = Replace(MUV_READ("CRMVIEWSTATE"),"'","''")

SQL = "SELECT * FROM Settings_Reports WHERE ReportNumber = 1700 AND UserNo = " & Session("userNo")
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs= cnn8.Execute(SQL)
UseSettings_Reports = False
If NOT rs.EOF Then
	UseSettings_Reports = True
	ReportSpecificData1  = rs("ReportSpecificData1")
	ReportSpecificData2  = rs("ReportSpecificData2")
	ReportSpecificData3  = rs("ReportSpecificData3")
	ReportSpecificData4  = rs("ReportSpecificData4")
	ReportSpecificData5  = rs("ReportSpecificData5")
	ReportSpecificData6  = rs("ReportSpecificData6")
	ReportSpecificData7  = rs("ReportSpecificData7")
	ReportSpecificData8  = rs("ReportSpecificData8")
	ReportSpecificData9  = rs("ReportSpecificData9")
	ReportSpecificData10  = rs("ReportSpecificData10")
	ReportSpecificData11  = rs("ReportSpecificData11")
	ReportSpecificData12  = rs("ReportSpecificData12")
	ReportSpecificData13 = rs("ReportSpecificData13")
End If
'****************************
'End Read Settings_Reports
'****************************
%>

	
<!-- modal box !-->
<div class="modal fade bs-example-modal-lg-customize" tabindex="-1" role="dialog" aria-labelledby="myLargeModalLabel" aria-hidden="true">
	<div class="modal-dialog modal-lg modal-height">
		<div class="modal-content">
			
			<div class="modal-header">
				<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
				<h4 class="modal-title" id="myModalLabel" align="center">Web Fulfillment and Invoice Cross Reference (Summary)</h4>
			</div>

			<form method="post" action="WebFulfillmentInvoiceXRefSummary_Customize_SaveValues.asp" name="frmWebFulfillInvXRef_Customize" id="frmWebFulfillInvXRef_Customize" onsubmit="return validateCustomizeForm();">

			      <!-- insert content in here !-->
			      <div class="modal-body ativa-scroll">
			      
			      
			      
		 	      	<!-- date ranges !-->
			      	<div class="container-fluid container-modal">
			      	
				      	<div class="row">
		      	
		 		      	<!-- left column !-->
		 		      	<div class="col-lg-2 col-md-2 col-sm-12 col-xs-12 left-column">
			 		      	<h4><br>Clear</h4>
		 		      	</div>
		 		      	<!-- eof left column !-->
		      	
		 		      	<!-- right column !-->
		 		      	<div class="col-lg-10 col-md-10 col-sm-12 col-xs-12 right-column">
			      	
				      	<!-- row !-->
					      	<div class="row">
						        <!-- First Date !-->
						    	<div class="col-xs-8 col-sm-1 col-md-12 col-lg-8">
									<button type="button" class="btn btn-warning btn-lg btn-block" id="resetModalFormFields">
										<i class="fa fa-refresh" aria-hidden="true"></i>&nbsp;Clear All Selected Options/Filters
									</button>
						        </div>
		    			  	</div>
		    			  	<!-- eof row !-->
		      	
		 		      	</div>
		 		      	<!-- eof right column !-->
			      	</div>
		      	</div>
		      	<!-- eof date ranges !-->  
		      	
 			      
 	      	
		 	      	<!-- date ranges !-->
			      	<div class="container-fluid container-modal">
			      	
				      	<div class="row">
 		      	
		 		      	<!-- left column !-->
		 		      	<div class="col-lg-2 col-md-3 col-sm-12 col-xs-12 left-column <% If ReportSpecificData1 <> "" Then Response.Write("activefilter") %>">
			 		      	<h4><br>Filter By<br><br>OCS Web Order Date <br><br>or<br><br>MDS Invoice Date<br><br><button type="button" class="btn btn-info btn-xs btn-block" id="btnClearDateFilters">clear date filters</button></h4>
		 		      	</div>
		 		      	<!-- eof left column !-->
 		      	
		 		      	<!-- right column !-->
		 		      	<div class="col-lg-10 col-md-9 col-sm-12 col-xs-12 right-column">
	 		      	
				      	<!-- row !-->
				      	
				      	
				      	
					      	<div class="row" style="margin-top:20px;margin-bottom:30px;">
							    <div class="col-xs-12 col-sm-12 col-md-12 col-lg-12">
							    	<strong>Select Filter By Date Type</strong>:
									<select class="form-control" name="selWebOrderOrMDSInvoiceFilter" id="selWebOrderOrMDSInvoiceFilter">
										<option value="" <% If OCSWebOrderOrMDSInvoice = "" Then Response.Write("selected") %>>Select Date Type For Filtering</option>
										<option value="OCS" <% If OCSWebOrderOrMDSInvoice = "OCS" Then Response.Write("selected") %>>Filter By OCS Web Order Date</option>
										<option value="MDS" <% If OCSWebOrderOrMDSInvoice = "MDS" Then Response.Write("selected") %>>Filter By MDS Invoice Date</option>
									</select>
		 					     </div>
					      	</div>
					      
					      	
					      	<div class="row">
						        <!-- First Date !-->
						    	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-6">
							    	<% If DatesOrPeriods = "Dates" Then %>
								    	<input type="radio" id="optDatesorPeriodsDates" name="optDatesorPeriods" value="Dates" checked="checked"><strong> Use Dates</strong>
								    <% Else %>
								    	<input type="radio" id="optDatesorPeriodsDates" name="optDatesorPeriods" value="Dates"><strong> Use Dates</strong>
								    <% End If %>

							    	<br><br>
									<div class="form-group">
										<input type="hidden" id="txtRangeStartDate" name="txtRangeStartDate" value="<%= RangeStartDateCustomize %>">
										<input type="hidden" id="txtRangeEndDate" name="txtRangeEndDate" value="<%= RangeEndDateCustomize %>">
										Date Range<br>
										<input type="text" id="InvoiceReportRange" class="form-control">
	                            		<i class="glyphicon glyphicon-calendar fa fa-calendar invoicerangedatepicker"></i>	

									</div>
						        </div>

		  					    <!-- First Period !-->
						    	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-6"> 
							    	<% If DatesOrPeriods = "Periods" Then %>
								    	<input type="radio" id="optDatesorPeriodsPeriods" name="optDatesorPeriods" value="Periods" checked="checked"><strong> Use Periods</strong>
								    <% Else %>
								    	<input type="radio" id="optDatesorPeriodsPeriods" name="optDatesorPeriods" value="Periods"><strong> Use Periods</strong>
								    <% End If %>
							    	
							    	<br><br>
							    	 Starting Period
								    <select class="form-control" name="selPeriodStart" id="selPeriodStart">
			 					      	<option selected value="">--none--</option>
								      	<%'Dont go past the last closed period
								      	 
								      	'Get all period info
							      	  	
							      	  	SQL = "SELECT * FROM " & MUV_Read("SQL_Owner") & ".Settings_CompanyPeriods "
							      	  	SQL = SQL & "WHERE InternalRecordIdentifier <= " & GetLastClosedReportPeriodIntRecID() + 1
							      	  	SQL = SQL & " ORDER BY InternalRecordIdentifier ASC"
							      	  	

										Set cnn8 = Server.CreateObject("ADODB.Connection")
										cnn8.open (Session("ClientCnnString"))
										Set rs = Server.CreateObject("ADODB.Recordset")
										rs.CursorLocation = 3 
										Set rs = cnn8.Execute(SQL)
									
										If not rs.EOF Then
											Do
												If StartPeriodBeingEvaluatedCustomize <> "" Then
													If cInt(StartPeriodBeingEvaluatedCustomize) = cInt(rs("InternalRecordIdentifier")) Then
														Response.Write("<option value='" & rs("InternalRecordIdentifier") & "' selected>" & rs("Year") & " - P" & rs("Period") & " - " & FormatDateTime(rs("BeginDate")) & " - " & FormatDateTime(rs("EndDate")) & "</option>")
													Else
														Response.Write("<option value='" & rs("InternalRecordIdentifier") & "'>" & rs("Year") & " - P" & rs("Period") & " - " & FormatDateTime(rs("BeginDate")) & " - " & FormatDateTime(rs("EndDate")) & "</option>")
													End If
												Else
													Response.Write("<option value='" & rs("InternalRecordIdentifier") & "'>" & rs("Year") & " - P" & rs("Period") & " - " & FormatDateTime(rs("BeginDate")) & " - " & FormatDateTime(rs("EndDate")) & "</option>")
												End If
												rs.movenext
											Loop until rs.eof
										End If
										set rs = Nothing
										cnn8.close
										set cnn8 = Nothing
								      	%>
								    </select>
								    
							  
							    	<br>Ending Period
								    <select class="form-control" name="selPeriodEnd" id="selPeriodEnd">
			 					      	<option selected value="">--none--</option>
								      	<%'Dont go past the last closed period
								      	 
								      	'Get all period info
							      	  	
							      	  	SQL = "SELECT * FROM " & MUV_Read("SQL_Owner") & ".Settings_CompanyPeriods "
							      	  	SQL = SQL & "WHERE InternalRecordIdentifier <= " & GetLastClosedReportPeriodIntRecID() + 1
							      	  	SQL = SQL & " ORDER BY InternalRecordIdentifier ASC"
							      	  	

										Set cnn8 = Server.CreateObject("ADODB.Connection")
										cnn8.open (Session("ClientCnnString"))
										Set rs = Server.CreateObject("ADODB.Recordset")
										rs.CursorLocation = 3 
										Set rs = cnn8.Execute(SQL)
									
										If not rs.EOF Then
											Do
												If EndPeriodBeingEvaluatedCustomize<> "" Then
													If cInt(EndPeriodBeingEvaluatedCustomize) = cInt(rs("InternalRecordIdentifier")) Then
														Response.Write("<option value='" & rs("InternalRecordIdentifier") & "' selected>" & rs("Year") & " - P" & rs("Period") & " - " & FormatDateTime(rs("BeginDate")) & " - " & FormatDateTime(rs("EndDate")) & "</option>")
													Else
														Response.Write("<option value='" & rs("InternalRecordIdentifier") & "'>" & rs("Year") & " - P" & rs("Period") & " - " & FormatDateTime(rs("BeginDate")) & " - " & FormatDateTime(rs("EndDate")) & "</option>")
													End If
												Else
													Response.Write("<option value='" & rs("InternalRecordIdentifier") & "'>" & rs("Year") & " - P" & rs("Period") & " - " & FormatDateTime(rs("BeginDate")) & " - " & FormatDateTime(rs("EndDate")) & "</option>")
												End If
												rs.movenext
											Loop until rs.eof
										End If
										set rs = Nothing
										cnn8.close
										set cnn8 = Nothing
								      	%>
								    </select>
								    
						        </div>
						    	<!-- First Period !-->

		    			  	</div>
		    			  	<!-- eof row !-->
			 		      	</div>
		 		      	<!-- eof right column !-->
			      	</div>
		      	</div>
	 	      	<!-- eof date ranges !-->    
	 	      	

	 	      	
	 	      	
 	      	 	      	
	 	      	<!-- categories !-->
		      	<div class="container-fluid container-modal">
			      	<div class="row">
 		      	
		 		      	<!-- left column !-->
		 		      	<div class="col-lg-2 col-md-3 col-sm-12 col-xs-12 left-column <% If ReportSpecificData7 <> "" Then Response.Write("activefilter") %>">
			 		      	<h4><br>Customer Class<br><br><button type="button" class="btn btn-info btn-xs btn-block" id="btnClearClassCodeFilter">clear selections</button></h4>
 				      	</div>
 				      	<!-- eof left column !-->
 		      	
 				      	<!-- right column !-->
 				      	<div class="col-lg-10 col-md-9 col-sm-12 col-xs-12 right-column check-row">
	 		      	
				      	<!-- row !-->
					      	<div class="row categories-checkboxes ">
   
								<div class="checkbox">
			  
					     			<%'This is where we get all the categories
					     				CustomerClassArrayForTab = ""
										CustomerClassArrayForTab = Split(DefaultSelectedCustomerClassesForInvoiceReport,",")

					     				SQL = "SELECT * FROM AR_CustomerClass ORDER BY ClassCode"
						
										Set cnn8 = Server.CreateObject("ADODB.Connection")
										cnn8.open (Session("ClientCnnString"))
										Set rs = Server.CreateObject("ADODB.Recordset")
										rs.CursorLocation = 3 
										Set rs = cnn8.Execute(SQL)
												
										If NOT rs.EOF Then
											Do
												%>
										      	<div class="col-lg-6 col-md-6 col-sm-12 col-xs-12 ">
											      	<label>
											      	<%
	      									      	ResponseLine = ""
											      	ResponseLine = ResponseLine & "<input type='checkbox' class='check' "
											      	For x = 0 to Ubound(CustomerClassArrayForTab)
											      		If rs("ClassCode") = CustomerClassArrayForTab(x) Then ResponseLine = ResponseLine & " checked "
											      	Next 
											    	ResponseLine = ResponseLine & "id='chk" & rs("ClassCode") & "' name='chkClassCode' value='" & rs("ClassCode") & "'>" & rs("ClassDescription") & " (" & rs("ClassCode") & ")<br>"
											    	Response.Write(ResponseLine)
	
													%>
											      	</label>
										      	</div>   
												<%
												rs.MoveNext
											Loop until rs.EOF
										End If
					     			%>
			    			

								</div> 		      			      	
					      	</div>
					      	<!-- eof row !-->
		 		      	</div>
		 		      	<!-- eof right column !-->
			      	</div>
		      	</div>
 		      	<!-- eof categories !-->      
 		      	
 		      	
 		      	
  	      	 	      	
	 	      	<!-- customer type !-->
		      	<div class="container-fluid container-modal">
			      	<div class="row">
 		      	
		 		      	<!-- left column !-->
		 		      	<div class="col-lg-2 col-md-3 col-sm-12 col-xs-12 left-column <% If ReportSpecificData13 <> "" Then Response.Write("activefilter") %>">
			 		      	<h4><br>Customer Type<br><br><button type="button" class="btn btn-info btn-xs btn-block" id="btnClearCustomerTypeFilter">clear selections</button></h4>
 				      	</div>
 				      	<!-- eof left column !-->
 		      	
 				      	<!-- right column !-->
 				      	<div class="col-lg-10 col-md-9 col-sm-12 col-xs-12 right-column check-row">
	 		      	
				      	<!-- row !-->
					      	<div class="row categories-checkboxes ">
   
								<div class="checkbox">
			  
					     			<%'This is where we get all the customer types
					     			
					     				CustomerTypeArrayForTab = ""
										CustomerTypeArrayForTab = Split(DefaultSelectedCustomerTypesForInvoiceReport,",")

					     				SQL = "SELECT DISTINCT(CustType) FROM AR_Customer ORDER BY CustType"
						
										Set cnn8 = Server.CreateObject("ADODB.Connection")
										cnn8.open (Session("ClientCnnString"))
										Set rs = Server.CreateObject("ADODB.Recordset")
										rs.CursorLocation = 3 
										Set rs = cnn8.Execute(SQL)
												
										If NOT rs.EOF Then
											Do
												%>
										      	<div class="col-lg-6 col-md-6 col-sm-12 col-xs-12 ">
											      	<label>
											      	<%
	      									      	ResponseLine = ""
											      	ResponseLine = ResponseLine & "<input type='checkbox' class='check' "
											      	For x = 0 to Ubound(CustomerTypeArrayForTab)
											      		If rs("CustType") = CustomerTypeArrayForTab(x) Then ResponseLine = ResponseLine & " checked "
											      	Next 
											    	ResponseLine = ResponseLine & "id='chk" & rs("CustType") & "' name='chkCustomerType' value='" & rs("CustType") & "'>" & GetCustTypeByCustTypeNum(rs("CustType")) & " (Cust Type " & rs("CustType") & ")<br>"
											    	Response.Write(ResponseLine)
	
													%>
											      	</label>
										      	</div>   
												<%
												rs.MoveNext
											Loop until rs.EOF
										End If
					     			%>
			    			

								</div> 		      			      	
					      	</div>
					      	<!-- eof row !-->
		 		      	</div>
		 		      	<!-- eof right column !-->
			      	</div>
		      	</div>
 		      	<!-- eof customer type!-->    
 		      	  
		      	
	 	      	<!-- categories !-->
		      	<div class="container-fluid container-modal">
			      	<div class="row">
 		      	
		 		      	<!-- left column !-->
		 		      	<div class="col-lg-2 col-md-3 col-sm-12 col-xs-12 left-column <% If (ReportSpecificData10 <> "false" OR ReportSpecificData11 <> "false") AND (ReportSpecificData10 <> "" OR ReportSpecificData11 <> "") Then Response.Write("activefilter") %>">
			 		      	<h4><br>Invoiced Orders<br><br><button type="button" class="btn btn-info btn-xs btn-block" id="btnClearInvoicedFilter">clear selections</button></h4>
 				      	</div>
 				      	<!-- eof left column !-->
 		      	
 				      	<!-- right column !-->
 				      	<div class="col-lg-10 col-md-9 col-sm-12 col-xs-12 right-column check-row">
	 		      	
				      	<!-- row !-->
					      	<div class="row categories-checkboxes ">
   
								<div class="checkbox">
							      	<div class="col-lg-12 col-md-12 col-sm-12 col-xs-12 ">
								      	<label><input type="checkbox" class="checkInvoice" id="chkShowOrdersThatAreInvoiced" name="chkShowOrdersThatAreInvoiced" <% If ShowOrdersThatAreInvoiced = 1 Then Response.Write("checked") %>>Show Orders That Have Been Invoiced</label><br>
								      	<label><input type="checkbox" class="checkInvoice" id="chkShowOrdersThatAreNotInvoiced" name="chkShowOrdersThatAreNotInvoiced" <% If ShowOrdersThatAreNotInvoiced = 1 Then Response.Write("checked") %>>Show Orders That Have Not Been Invoiced</label><br>
							      	</div>   
								</div> 		      			      	
					      	</div>
					      	<!-- eof row !-->
		 		      	</div>
		 		      	<!-- eof right column !-->
			      	</div>
		      	</div>
 		      	<!-- eof categories !-->      
 		      	
  		      	
 		      	
	 	      	<!-- categories !-->
		      	<div class="container-fluid container-modal">
			      	<div class="row">
 		      	
		 		      	<!-- left column !-->
		 		      	<div class="col-lg-2 col-md-3 col-sm-12 col-xs-12 left-column <% If (ReportSpecificData8 <> "false" OR ReportSpecificData9 <> "false") AND (ReportSpecificData8 <> "" OR ReportSpecificData9 <> "") Then Response.Write("activefilter") %>">
			 		      	<h4><br>Order Remarks<br><br><button type="button" class="btn btn-info btn-xs btn-block" id="btnClearRemarksFilter">clear selections</button></h4>
 				      	</div>
 				      	<!-- eof left column !-->
 		      	
 				      	<!-- right column !-->
 				      	<div class="col-lg-10 col-md-9 col-sm-12 col-xs-12 right-column check-row">
	 		      	
				      	<!-- row !-->
					      	<div class="row categories-checkboxes ">
   
								<div class="checkbox">
							      	<div class="col-lg-12 col-md-12 col-sm-12 col-xs-12 ">
								      	<label><input type="checkbox" class="checkInvoice" id="chkShowOrdersWithRemarks" name="chkShowOrdersWithRemarks" <% If ShowOrdersWithRemarks = 1 Then Response.Write("checked") %>>Show Orders With Remarks</label><br>
								      	<label><input type="checkbox" class="checkInvoice" id="chkShowOrdersWithoutRemarks" name="chkShowOrdersWithoutRemarks" <% If ShowOrdersWithoutRemarks = 1 Then Response.Write("checked") %>>Show Orders Without Remarks</label><br>
							      	</div>   
								</div> 		      			      	
					      	</div>
					      	<!-- eof row !-->
		 		      	</div>
		 		      	<!-- eof right column !-->
			      	</div>
		      	</div>
 		      	<!-- eof categories !-->      	
 		      	
  		      	
 		      	
	 	      	<!-- categories !-->
		      	<div class="container-fluid container-modal">
			      	<div class="row">
 		      	
		 		      	<!-- left column !-->
		 		      	<div class="col-lg-2 col-md-3 col-sm-12 col-xs-12 left-column <% If ReportSpecificData12 <> "false" AND ReportSpecificData12 <> "" Then Response.Write("activefilter") %>">
			 		      	<h4><br>Hidden Orders<br><br><button type="button" class="btn btn-info btn-xs btn-block" id="btnClearHiddenFilter">clear selections</button></h4>
 				      	</div>
 				      	<!-- eof left column !-->
 		      	
 				      	<!-- right column !-->
 				      	<div class="col-lg-10 col-md-9 col-sm-12 col-xs-12 right-column check-row">
	 		      	
				      	<!-- row !-->
					      	<div class="row categories-checkboxes ">
   
								<div class="checkbox">
							      	<div class="col-lg-12 col-md-12 col-sm-12 col-xs-12 ">
								      	<label><input type="checkbox" class="checkInvoice" id="chkShowOrdersThatAreHidden" name="chkShowOrdersThatAreHidden" <% If ShowOrdersThatAreHidden = 1 Then Response.Write("checked") %>>Show Orders That Are Hidden</label><br>
							      	</div>   
								</div> 		      			      	
					      	</div>
					      	<!-- eof row !-->
		 		      	</div>
		 		      	<!-- eof right column !-->
			      	</div>
		      	</div>
 		      	<!-- eof categories !-->      	
		      	
	
</div>
      <!-- eof content insertion !-->
      
     <div class="modal-footer">
	     <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
	     <button type="submit" class="btn btn-primary">Run Report</button>
     </div>
			</form>
		</div>
	</div>
</div>
	<!-- eof modal box !-->
	
		<!-- eof select all / deselect all checkboxes !-->
		
		
		
<style type="text/css">
	.datepicker.dropdown-menu {right: auto;}
</style>

<script type="text/javascript" src="http://cdn.jsdelivr.net/momentjs/latest/moment.min.js"></script>
<script type="text/javascript" src="http://cdn.jsdelivr.net/bootstrap.daterangepicker/2/daterangepicker.js"></script>
<link rel="stylesheet" type="text/css" href="http://cdn.jsdelivr.net/bootstrap.daterangepicker/2/daterangepicker.css">

<script type="text/javascript">

		invoiceRangeStartDate = moment(Date.parse($('#txtRangeStartDate').val()));
		invoiceRangeEndDate = moment(Date.parse($('#txtRangeEndDate').val()));
		
        $('#InvoiceReportRange').daterangepicker({
                opens: 'right',
                startDate: invoiceRangeStartDate,
                endDate: invoiceRangeEndDate,
                alwaysShowCalendars: true,
                timePicker: false,
                linkedCalendars: false,
                autoUpdateInput:true,
                autoApply:true,
                showWeekNumbers: true,
                showClear: true,
                ranges: {
                    'Last 7 Days': [moment().subtract('days', 6), moment()],
                    'Last 30 Days': [moment().subtract('days', 29), moment()],
                    'This Month': [moment().startOf('month'), moment().endOf('month')],
                    'Last Month': [moment().subtract('month', 1).startOf('month'), moment().subtract('month', 1).endOf('month')]
                },
                buttonClasses: ['btn'],
                applyClass: 'green',
                cancelClass: 'default',
                format: 'MM/DD/YYYY',
                separator: ' to ',
                locale: {
                    applyLabel: 'Apply',
                    fromLabel: 'From',
                    toLabel: 'To',
                    customRangeLabel: 'Custom Range',
                    daysOfWeek: ['Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa'],
                    monthNames: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
                    firstDay: 1
                }
            },
            function (start, end) {
            	//$('#InvoiceReportRange span').html(start.format('MMMM D, YYYY') + ' - ' + end.format('MMMM D, YYYY'));
                $('#txtRangeStartDate').val(start.format('MM/DD/YYYY'));
                $('#txtRangeEndDate').val(end.format('MM/DD/YYYY'));  
                
            },
            
        );
</script>


	
 
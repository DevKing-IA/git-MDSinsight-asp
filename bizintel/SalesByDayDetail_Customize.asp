<SCRIPT LANGUAGE="JavaScript">
<!--
    function validateCustomizeForm()
    {
    		
		//***************************************
		//These are the checks for the text boxes
		//***************************************
    	//If they chose the option button for dates, make sure the dates are filled-in
		if (document.frmSalesByDay_Customize.optDatesorPeriods.value == "Dates")
		{
			if (document.frmSalesByDay_Customize.txtRangeStartDate.value == "" || 
				document.frmSalesByDay_Customize.txtRangeEndDate.value == "") 
				{		
					swal("Please make sure date range is complete.")
					return false;
				}

	    	// First start range specified but no end range
	        if (document.frmSalesByDay_Customize.txtRangeStartDate.value != "")
	        {
				if (document.frmSalesByDay_Customize.txtRangeEndDate.value == "")
				{
					swal("You specified a range start date but not a range end date.");
					return false;
				}
			}
				
	    	// First end range specified but no start range
	        if (document.frmSalesByDay_Customize.txtRangeEndDate.value != "")
	        {
				if (document.frmSalesByDay_Customize.txtRangeStartDate.value == "")
				{
					swal("You specified a range end date but not a range start date.");
					return false;
				}
			}

			// First range dates, start is greater than end
	        if (document.frmSalesByDay_Customize.txtRangeStartDate.value != "")
	        {
	        	var FStart = new Date(document.frmSalesByDay_Customize.txtRangeEndDate.value);
	        	var FEnd = new Date(document.frmSalesByDay_Customize.txtRangeStartDate.value);
				if (FStart < FEnd)
				{
					swal("The start date specified must be older than the end date.");
					return false;
				}
	        }
	
		}
		

		//********************************************
		//These are the checks for the drop down boxes
		//********************************************
    	//If they chose the option button for period, make sure they selected 2 periods
		if (document.frmSalesByDay_Customize.optDatesorPeriods.value == "Periods")
		{
	    	// Both periods empty
	        if (document.frmSalesByDay_Customize.selPeriod.value == "")
				{
					swal("Please select the period to include from the drop-down.");
					return false;
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
    var altura = $(window).height() - 155; //value corresponding to the modal heading + footer
    $(".ativa-scroll").css({"height":altura,"overflow-y":"auto"});
  }
</script>
<!-- eof modal scroll !-->
	
<!-- modal box !-->
<div class="modal fade bs-example-modal-lg-customize" tabindex="-1" role="dialog" aria-labelledby="myLargeModalLabel" aria-hidden="true">
	<div class="modal-dialog modal-lg modal-height">
		<div class="modal-content">
		
			
			<script>
			
				$(document).ready(function() {
				
					//$('input[name="optDatesorPeriods"][value="Dates"]').attr('checked',true);
				    
				    $(document).on('change','[name="selPeriod"]',function(){
	
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
					
				        
					$('#reportrange').on('show.daterangepicker', function(ev, picker) {
					  //do something, like clearing an input
						$('input[type=radio]').removeAttr('checked'); 
						$('input[name="optDatesorPeriods"][value="Periods"]').removeAttr('checked');
						$('input[name="optDatesorPeriods"][value="Dates"]').attr('checked','checked');
						$('input[name="optDatesorPeriods"][value="Dates"]').prop("checked", true);
			
					}); 
					
					$('#reportrange').on('apply.daterangepicker', function(ev, picker) {
						$('input[type=radio]').removeAttr('checked'); 
						$('input[name="optDatesorPeriods"][value="Periods"]').removeAttr('checked');
						$('input[name="optDatesorPeriods"][value="Dates"]').attr('checked','checked');
						$('input[name="optDatesorPeriods"][value="Dates"]').prop("checked", true);
					});     
						  
				}); //end document.ready() function
			
			</script>

			<div class="modal-header">
				<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
				<h4 class="modal-title" id="myModalLabel" align="center">Sales By Day and Customer Class (Detail)</h4>
			</div>

			<form method="post" action="SalesByDayDetail_Customize_SaveValues.asp" name="frmSalesByDay_Customize" onsubmit="return validateCustomizeForm();">

			      <!-- insert content in here !-->
			      <div class="modal-body ativa-scroll">
 	      	
		 	      	<!-- date ranges !-->
			      	<div class="container-fluid container-modal">
			      	
				      	<div class="row">
 		      	
		 		      	<!-- left column !-->
		 		      	<div class="col-lg-2 col-md-3 col-sm-12 col-xs-12 left-column">
			 		      	<h4><br>Date Range</h4>
		 		      	</div>
		 		      	<!-- eof left column !-->
 		      	
		 		      	<!-- right column !-->
		 		      	<div class="col-lg-10 col-md-9 col-sm-12 col-xs-12 right-column">
	 		      	
				      	<!-- row !-->
					      	<div class="row">
					      	
						        <!-- First Date !-->
						    	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-6">
							    	<% If DatesOrPeriods = "Dates" OR DatesOrPeriods = "" Then %>
								    	<input type="radio" id="optDatesorPeriods" name="optDatesorPeriods" value="Dates"  checked="checked"><strong> Use Dates</strong>
								    <% Else %>
								    	<input type="radio" id="optDatesorPeriods" name="optDatesorPeriods" value="Dates"><strong> Use Dates</strong>
								    <% End If %>
							    	<br><br>
									<div class="form-group">
										<input type="hidden" id="txtRangeStartDate" name="txtRangeStartDate" value="<%= RangeStartDateCustomize %>">
										<input type="hidden" id="txtRangeEndDate" name="txtRangeEndDate" value="<%= RangeEndDateCustomize %>">
										Date Range<br>
										<div class="btn btn-default" id="reportrange">
											<i class="fa fa-calendar"></i> &nbsp;
											<span></span>
											<b class="fa fa-angle-down"></b>
										</div>
									</div>
						        </div>

		  					    <!-- First Period !-->
						    	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-6"> 
							    	<% If DatesOrPeriods = "Periods" Then %>
								    	<input type="radio" id="optDatesorPeriods" name="optDatesorPeriods" value="Periods"  checked="checked"><strong> Use Periods</strong>
								    <% Else %>
								    	<input type="radio" id="optDatesorPeriods" name="optDatesorPeriods" value="Periods"><strong> Use Periods</strong>
								    <% End If %>

							    	<br><br>
							    	 Period
								    <select class="form-control" name="selPeriod" id="selPeriod">
			 					      	<option selected value="">--none--</option>
								      	<%'Dont go past the last closed period
								      	 
								      	'Get all period info							      	  	
							      	  	SQL = "SELECT * FROM " & MUV_Read("SQL_Owner") & ".Settings_CompanyPeriods "
							      	  	SQL = SQL & "WHERE InternalRecordIdentifier <= " & GetLastClosedReportPeriodIntRecID() + 1
							      	  	SQL = SQL & " ORDER BY [Year] DESC, Period DESC"

										Set cnn8 = Server.CreateObject("ADODB.Connection")
										cnn8.open (Session("ClientCnnString"))
										Set rs = Server.CreateObject("ADODB.Recordset")
										rs.CursorLocation = 3 
										Set rs = cnn8.Execute(SQL)
									
										If not rs.EOF Then
											Do
												If PeriodBeingEvaluatedCustomize = rs("InternalRecordIdentifier") Then
													Response.Write("<option value='" & rs("InternalRecordIdentifier") & "' selected>" & rs("Year") & " - P" & rs("Period") & " - " & FormatDateTime(rs("BeginDate")) & " - " & FormatDateTime(rs("EndDate")) & "</option>")
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
		 		      	<div class="col-lg-2 col-md-3 col-sm-12 col-xs-12 left-column">
			 		      	<h4><br>Customer Class</h4>
 				      	</div>
 				      	<!-- eof left column !-->
 		      	
 				      	<!-- right column !-->
 				      	<div class="col-lg-10 col-md-9 col-sm-12 col-xs-12 right-column check-row">
	 		      	
				      	<!-- row !-->
					      	<div class="row categories-checkboxes ">
   
								<div class="checkbox">
			  
					     			<%'This is where we get all the categories
					     				CustomerClassArrayForTab = ""
										CustomerClassArrayForTab = Split(DefaultSelectedCustomerClassesForSalesReport,",")

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
 		      	
	 	      	<!-- categories !-->
		      	<div class="container-fluid container-modal">
			      	<div class="row">
 		      	
		 		      	<!-- left column !-->
		 		      	<div class="col-lg-2 col-md-3 col-sm-12 col-xs-12 left-column">
			 		      	<h4><br>Invoice Type</h4>
 				      	</div>
 				      	<!-- eof left column !-->
 		      	
 				      	<!-- right column !-->
 				      	<div class="col-lg-10 col-md-9 col-sm-12 col-xs-12 right-column check-row">
	 		      	
				      	<!-- row !-->
					      	<div class="row categories-checkboxes ">
   
								<div class="checkbox">
							      	<div class="col-lg-6 col-md-6 col-sm-12 col-xs-12 ">
								      	<label><input type="checkbox" class="checkInvoice" id="chkBackorder" name="chkBackorder" <% If InvoiceTypeBackOrder = "B" Then Response.Write("checked") %>>Backorder (B)</label><br>
								      	<label><input type="checkbox" class="checkInvoice" id="chkCreditMemo" name="chkCreditMemo" <% If InvoiceTypeCreditMemo = "C" Then Response.Write("checked") %>>Credit Memo (C)</label><br>
								      	<label><input type="checkbox" class="checkInvoice" id="chkARDebit" name="chkARDebit" <% If InvoiceTypeARDebit = "E" Then Response.Write("checked") %>>Simple A/R Debit (E)</label><br>
								      	<label><input type="checkbox" class="checkInvoice" id="chkRental" name="chkRental" <% If InvoiceTypeRental = "G" Then Response.Write("checked") %>>Rental (G)</label><br>
							      	</div>   
							      	<div class="col-lg-6 col-md-6 col-sm-12 col-xs-12 ">
								      	<label><input type="checkbox" class="checkInvoice" id="chkRouteInvoicing" name="chkRouteInvoicing" <% If InvoiceTypeRouteInvoicing = "I" Then Response.Write("checked") %>>Route Invoicing (I)</label><br>
								      	<label><input type="checkbox" class="checkInvoice" id="chkInterest" name="chkInterest" <% If InvoiceTypeInterest = "O" Then Response.Write("checked") %>>Interest (O)</label><br>
								      	<label><input type="checkbox" class="checkInvoice" id="chkTelselInvoicing" name="chkTelselInvoicing" <% If InvoiceTypeTelselInvoicing = "T" Then Response.Write("checked") %>>Telsel Invoicing (T)</label><br>
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

	
		rangeStartDate = moment(Date.parse($('#txtRangeStartDate').val()));
		rangeEndDate = moment(Date.parse($('#txtRangeEndDate').val()));
		
        $('#reportrange').daterangepicker({
                opens: 'right',
                startDate: rangeStartDate,
                endDate: rangeEndDate,
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
            	$('#reportrange span').html(start.format('MMMM D, YYYY') + ' - ' + end.format('MMMM D, YYYY'));
				
                $('#txtRangeStartDate').val(start.format('MM/DD/YYYY'));
                $('#txtRangeEndDate').val(end.format('MM/DD/YYYY'));  
    	
                //$('#reportrange span').html(startDate.format('MM/DD/YYYY') + ' - ' + endDate.format('MM/DD/YYYY'));
                //$('#txtRangeStartDate').val(rangeStartDate.format('MM/DD/YYYY'));
                //$('#txtRangeEndDate').val(rangeEndDate.format('MM/DD/YYYY'));      
          
            },
            
        );
  
</script>


	
 
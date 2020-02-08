<SCRIPT LANGUAGE="JavaScript">
<!--
    function validateCustomizeForm()
    {
    					    	
		var selectedBasePeriods1 = $("#selBasePeriodsRange0").val(); 
		var selectedComparisonPeriods1 = $("#selPeriodsForComparison0").val();
		
		var selectedBasePeriods2 = $("#selBasePeriodsRange1").val(); 
		var selectedComparisonPeriods2 = $("#selPeriodsForComparison1").val();

		var selectedBasePeriods3 = $("#selBasePeriodsRange2").val(); 
		var selectedComparisonPeriods3 = $("#selPeriodsForComparison2").val();

		var selectedBasePeriods4 = $("#selBasePeriodsRange3").val(); 
		var selectedComparisonPeriods4 = $("#selPeriodsForComparison3").val();

		var selectedBasePeriods5 = $("#selBasePeriodsRange4").val(); 
		var selectedComparisonPeriods5 = $("#selPeriodsForComparison4").val();

		var selectedBasePeriods6 = $("#selBasePeriodsRange5").val(); 
		var selectedComparisonPeriods6 = $("#selPeriodsForComparison5").val();
		
		//********************************************
		//These are the checks for the select boxes
		//********************************************
		if ((selectedBasePeriods1 == "" && selectedComparisonPeriods1 !== "") ||
		  (selectedBasePeriods2 == "" && selectedComparisonPeriods2 !== "") ||
		  (selectedBasePeriods3 == "" && selectedComparisonPeriods3 !== "") ||
		  (selectedBasePeriods4 == "" && selectedComparisonPeriods4 !== "") ||
		  (selectedBasePeriods5 == "" && selectedComparisonPeriods5 !== "") ||
		  (selectedBasePeriods6 == "" && selectedComparisonPeriods6 !== "")) { 		
			swal("You have a comparison period selected without selecting a base period.");
			return false;	
		}

		if ((selectedBasePeriods1 !== "" && selectedComparisonPeriods1 == "") ||
		  (selectedBasePeriods2 !== "" && selectedComparisonPeriods2 == "") ||
		  (selectedBasePeriods3 !== "" && selectedComparisonPeriods3 == "") ||
		  (selectedBasePeriods4 !== "" && selectedComparisonPeriods4 == "") ||
		  (selectedBasePeriods5 !== "" && selectedComparisonPeriods5 == "") ||
		  (selectedBasePeriods6 !== "" && selectedComparisonPeriods6 == "")) {
			swal("You have a base period selected without selecting a comparison period.");
			return false;	
		}
		
		var basePeriod1Array = selectedBasePeriods1.split("*");
		var comparisonPeriod1Array = selectedComparisonPeriods1.split("*");
		var basePeriod1EndDate = new Date(basePeriod1Array[1]);
		var comparisonPeriod1EndDate = new Date(comparisonPeriod1Array[1]);
		
		if (basePeriod1EndDate < comparisonPeriod1EndDate)
		{
			swal("Your 1st selected comparison period must occur prior to the selected base period.");
			return false;			
		}
				
		
		var basePeriod2Array = selectedBasePeriods2.split("*");
		var comparisonPeriod2Array = selectedComparisonPeriods2.split("*");
		var basePeriod2EndDate = new Date(basePeriod2Array[2]);
		var comparisonPeriod2EndDate = new Date(comparisonPeriod2Array[2]);
		
		if (basePeriod2EndDate < comparisonPeriod2EndDate)
		{
			swal("Your 2nd selected comparison period must occur prior to the selected base period.");
			return false;			
		}
		
		
		var basePeriod3Array = selectedBasePeriods3.split("*");
		var comparisonPeriod3Array = selectedComparisonPeriods3.split("*");
		var basePeriod3EndDate = new Date(basePeriod3Array[3]);
		var comparisonPeriod3EndDate = new Date(comparisonPeriod3Array[3]);
		
		if (basePeriod3EndDate < comparisonPeriod3EndDate)
		{
			swal("Your 3rd selected comparison period must occur prior to the selected base period.");
			return false;			
		}
		
		
		var basePeriod4Array = selectedBasePeriods4.split("*");
		var comparisonPeriod4Array = selectedComparisonPeriods4.split("*");
		var basePeriod4EndDate = new Date(basePeriod4Array[4]);
		var comparisonPeriod4EndDate = new Date(comparisonPeriod4Array[4]);
		
		if (basePeriod4EndDate < comparisonPeriod4EndDate)
		{
			swal("Your 4th selected comparison period must occur prior to the selected base period.");
			return false;			
		}	
		
		
		var basePeriod5Array = selectedBasePeriods5.split("*");
		var comparisonPeriod5Array = selectedComparisonPeriods5.split("*");
		var basePeriod5EndDate = new Date(basePeriod5Array[5]);
		var comparisonPeriod5EndDate = new Date(comparisonPeriod5Array[5]);
		
		if (basePeriod5EndDate < comparisonPeriod5EndDate)
		{
			swal("Your 5th selected comparison period must occur prior to the selected base period.");
			return false;			
		}



		var basePeriod6Array = selectedBasePeriods6.split("*");
		var comparisonPeriod6Array = selectedComparisonPeriods6.split("*");
		var basePeriod6EndDate = new Date(basePeriod6Array[6]);
		var comparisonPeriod6EndDate = new Date(comparisonPeriod6Array[6]);
		
		if (basePeriod6EndDate < comparisonPeriod6EndDate)
		{
			swal("Your 6th selected comparison period must occur prior to the selected base period.");
			return false;			
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
			
			<div class="modal-header">
				<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
				<h4 class="modal-title" id="myModalLabel" align="center">Sales By Period and Customer Class (Summary)</h4>
			</div>

			<form method="post" action="SalesByPeriodSummary_Customize_SaveValues.asp" name="frmSalesByPeriod_Customize" onsubmit="return validateCustomizeForm();">

			      <!-- insert content in here !-->
			      <div class="modal-body ativa-scroll">
			      
			      
			      
	 	      	<!-- date ranges !-->
		      	<div class="container-fluid container-modal">
		      	
			      	<div class="row">
	      	
	 		      	<!-- left column !-->
	 		      	<div class="col-lg-12 col-md-12 col-sm-12 col-xs-12 left-column">
		 		      	<h4>Select Up To Six Periods To Analyze/Compare:</h4>
	 		      	</div>
	 		      	<!-- eof left column !-->
	 		      	
	 		      </div>
	 		    </div>
      	
 	      	
	 	      	<!----------------------------------------------------------------------------------------------------------------------------->
	 	      	<!--BEGIN LOOP WRITING PERIOD COMPARISON SELECT BOXES-------------------------------------------------------------------------->
	 	      	<!----------------------------------------------------------------------------------------------------------------------------->
  	      		<% For i = 0 to 5 %>
 	      	
		 	      	<!-- date ranges !-->
			      	<div class="container-fluid container-modal">
			      	
				      	<div class="row">
 		      	
		 		      	<!-- left column !-->
		 		      	<div class="col-lg-2 col-md-3 col-sm-12 col-xs-12 left-column">
			 		      	<h4><br>[<%= i + 1 %>] Periods To Compare</h4>
		 		      	</div>
		 		      	<!-- eof left column !-->
 		      	
		 		      	<!-- right column !-->
		 		      	<div class="col-lg-10 col-md-9 col-sm-12 col-xs-12 right-column">
	 		      	
				      	<!-- row !-->
					      	<div class="row">
					      	
						        <!-- First Date !-->
						    	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-6">
						    		<strong> Select Base Period</strong>
							    	<br><br>

								    <select class="form-control" name="selBasePeriodsRange<%= i %>" id="selBasePeriodsRange<%= i %>">
			 					      	
			 					        <% If i <= uBound(BasePeriodsRangeIntRecIDsArray) Then %>
				 					        <% If BasePeriodsRangeIntRecIDsArray(i) = "" Then %>
				 					      		<option value="" selected="selected">--none--</option>
				 					      	<% Else %>
				 					      		<option value="">--none--</option>
				 					      	<% End If %>
				 					    <% Else %>
				 					      	<option value="" selected="selected">--none--</option>
				 					    <% End If %>
			 					      	

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
												periodSelected = False
												
												If i <= uBound(BasePeriodsRangeIntRecIDsArray) Then
													If cInt(BasePeriodsRangeIntRecIDsArray(i)) = cInt(rs("InternalRecordIdentifier")) Then
														periodSelected = True
													End If
												End If
												
												If periodSelected = True Then
													Response.Write("<option value='" & rs("InternalRecordIdentifier") & "*" & rs("EndDate") & "' selected>" & rs("Year") & " - P" & rs("Period") & " - " & FormatDateTime(rs("BeginDate")) & " - " & FormatDateTime(rs("EndDate")) & "</option>")
												Else
													Response.Write("<option value='" & rs("InternalRecordIdentifier") & "*" & rs("EndDate") & "'>" & rs("Year") & " - P" & rs("Period") & " - " & FormatDateTime(rs("BeginDate")) & " - " & FormatDateTime(rs("EndDate")) & "</option>")
												End If
												
												rs.MoveNext
											Loop until rs.EOF
										End If
										
										set rs = Nothing
										cnn8.close
										set cnn8 = Nothing
								      	%>
								    </select>
						        </div>

		  					    <!-- First Period !-->
						    	<div class="col-xs-6 col-sm-1 col-md-1 col-lg-6"> 

									<strong> Compare Base Period To:</strong>
									
							    	<br><br>
							    
								    <select class="form-control" name="selPeriodsForComparison<%= i %>" id="selPeriodsForComparison<%= i %>">
			 					        <% If i <= uBound(ComparisonPeriodsRangeIntRecIDsArray) Then %>
				 					        <% If ComparisonPeriodsRangeIntRecIDsArray(i) = "" Then %>
				 					      		<option value="" selected="selected">--none--</option>
				 					      	<% Else %>
				 					      		<option value="">--none--</option>
				 					      	<% End If %>
				 					    <% Else %>
				 					      	<option value="" selected="selected">--none--</option>
				 					    <% End If %>
				 					    
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
												periodSelected = False
												
												If i <= uBound(ComparisonPeriodsRangeIntRecIDsArray) Then
													If cInt(ComparisonPeriodsRangeIntRecIDsArray(i)) = cInt(rs("InternalRecordIdentifier")) Then
														periodSelected = True
													End If
												End If
												
												If periodSelected = True Then
													Response.Write("<option value='" & rs("InternalRecordIdentifier") & "*" & rs("EndDate") & "' selected>" & rs("Year") & " - P" & rs("Period") & " - " & FormatDateTime(rs("BeginDate")) & " - " & FormatDateTime(rs("EndDate")) & "</option>")
												Else
													Response.Write("<option value='" & rs("InternalRecordIdentifier") & "*" & rs("EndDate") & "'>" & rs("Year") & " - P" & rs("Period") & " - " & FormatDateTime(rs("BeginDate")) & " - " & FormatDateTime(rs("EndDate")) & "</option>")
												End If
												
												rs.MoveNext
											Loop until rs.EOF
											
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
	 	      	
	 	      	<% Next %>
	 	      	
	 	      	<!----------------------------------------------------------------------------------------------------------------------------->
	 	      	<!--END LOOP WRITING PERIOD COMPARISON SELECT BOXES---------------------------------------------------------------------------->
	 	      	<!----------------------------------------------------------------------------------------------------------------------------->
 	      	 	      	
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
		

	
 
<SCRIPT LANGUAGE="JavaScript">
<!--
    function validateCustomizeForm()
    {
    					    	
	
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

			<form method="post" action="SalesByMonthDetail_Customize_SaveValues.asp" name="frmSalesByPeriod_Customize" onsubmit="return validateCustomizeForm();">

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
				    	<div class="col-xs-12 col-sm-1 col-md-1 col-lg-12">
				    		<strong> Select Month of Year (CTRL + CLICK To Select Multiple)</strong>
					    	<br><br>

						    <select class="form-control" name="selMonthYearCombinations" id="selMonthYearCombinations" multiple="multiple" style="height:200px; width:200px;">
	 					      	
	 					        <% If i <= uBound(MonthYearCombinationsArray) Then %>
		 					        <% If MonthYearCombinationsArray(i) = "" Then %>
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
									endReportYear = rs("Year")
									endReportMonthNum = Month(rs("EndDate"))
								End If
								
								startReportYear = GetFirstReportPeriodYear()
								
								'**************************************************************************
								'First write the valid month/year combintations for the current report year
								'**************************************************************************
								For i = endReportMonthNum to 1 step-1
								
										yearMonthComboSelected = False
										yearLoopValue = endReportYear
										monthLoopValue = i
										monthLoopName = MonthName(i)
										
										If uBound(MonthYearCombinationsArray) >= 0 Then
										
											For x = 0 to uBound(MonthYearCombinationsArray)
												If MonthYearCombinationsArray(x) = monthLoopValue & "*" & yearLoopValue Then
													yearMonthComboSelected = True
												End If
											Next
											
										End If
										
										If yearMonthComboSelected = True Then
											Response.Write("<option value='" & monthLoopValue & "*" & yearLoopValue & "' selected>" & monthLoopName & " " & yearLoopValue & "</option>")
										Else
											Response.Write("<option value='" & monthLoopValue & "*" & yearLoopValue & "'>" & monthLoopName & " " & yearLoopValue & "</option>")
										End If
								Next	
								
								'***********************************************************************************
								'Then write the remaining month/year combintations for the remaining report years
								'***********************************************************************************
								
							
								For i = (endReportYear-1) to startReportYear step-1
								
									For z = 12 to 1 step-1
	
										yearMonthComboSelected = False
										yearLoopValue = i
										monthLoopValue = z
										monthLoopName = MonthName(z)
										
										If uBound(MonthYearCombinationsArray) >= 0 Then
										
											For x = 0 to uBound(MonthYearCombinationsArray)
												If MonthYearCombinationsArray(x) = monthLoopValue & "*" & yearLoopValue Then
													yearMonthComboSelected = True
												End If
											Next
											
										End If
										
										If yearMonthComboSelected = True Then
											Response.Write("<option value='" & monthLoopValue & "*" & yearLoopValue & "' selected>" & monthLoopName & " " & yearLoopValue & "</option>")
										Else
											Response.Write("<option value='" & monthLoopValue & "*" & yearLoopValue & "'>" & monthLoopName & " " & yearLoopValue & "</option>")
										End If
									Next
								Next	
								
								
								set rs = Nothing
								cnn8.close
								set cnn8 = Nothing
						      	%>
						    </select>
				        </div>

			      	</div>
		      	</div>
	 	      	<!-- eof date ranges !-->    
	 	      	
				
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
		

	
 
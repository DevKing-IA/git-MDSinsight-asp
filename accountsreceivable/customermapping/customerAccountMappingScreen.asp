<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/jquery_table_search.asp"-->
<!--#include file="../../inc/InsightFuncs_AR_AP.asp"-->

<script type="text/javascript">
		
 	
	function saveCustomerIDToEquivalentsTable(inputEnteredValue, inputID) {
	
		    //alert("Entered Value: " + inputEnteredValue + "-------Input ID: " + inputID);
		    
	   		//When the user blurs off of an input box, make an Ajax post to InSightFuncs_AjaxForARAP.asp
	   		
	   		$.ajax({
				type:"POST",
				url: "../../inc/InSightFuncs_AjaxForARAP.asp",
				cache: false,
				data: "action=SaveEquivalentCustomerAccount&id=" + encodeURIComponent(inputID) + "&equivID=" + encodeURIComponent(inputEnteredValue),
				success: function(response)
				 {
				 	if (response == "Success") {
				 		//If successfully saved, change the style of the input box, so the user gets a visual cue that the SKU was saved   
				 		//alert("#" + inputID);
				 		$("#" + inputID).addClass("saved"); 
				 	}
				 	else {
				 	 alert(response);
					}       	 
	             }	
			});	//end ajax post to data: "action=saveCustomerIDToEquivalentsTable"
		
			
	}	

</script>

 <style type="text/css">
	

	.container{
		max-width:1000px !important;
		margin-left:auto 0;
	}
	.panel-table {
	  width:1000px;
	}
	
	.panel-table .panel-body{
	  padding:0;
	}
	
	.panel-table .panel-body .table-bordered{
	  border-style: none;
	  margin:0;
	}
	
	.panel-table .panel-body .table-bordered > thead > tr > th:first-of-type {
	    text-align:center;
	    width: 100px;
	}
	
	.panel-table .panel-body .table-bordered > thead > tr > th:last-of-type,
	.panel-table .panel-body .table-bordered > tbody > tr > td:last-of-type {
	  border-right: 0px;
	}
	
	.panel-table .panel-body .table-bordered > thead > tr > th:first-of-type,
	.panel-table .panel-body .table-bordered > tbody > tr > td:first-of-type {
	  border-left: 0px;
	}
	
	.panel-table .panel-body .table-bordered > tbody > tr:first-of-type > td{
	  border-bottom: 0px;
	}
	
	.panel-table .panel-body .table-bordered > thead > tr:first-of-type > th{
	  border-top: 0px;
	}
	
	.panel-table .panel-footer .pagination{
	  margin:0; 
	}

	/*
	used to vertically center elements, may need modification if you're not using default sizes.
	*/
	.panel-table .panel-footer .col{
	 line-height: 34px;
	 height: 34px;
	}
	
	.panel-table .panel-heading .col h3{
	 line-height: 30px;
	 height: 30px;
	}
	
	.panel-table .panel-body .table-bordered > tbody > tr > td{
	  line-height: 34px;
	  text-align:left;
	}
	
	.description{
		line-height:14px !important;
		font-size:11px;
	}	

	.saved{
		background-color:#aeeaae !important;
	}

	.row-filter{
		margin-bottom: 15px;
	}

	table.sortable th:not(.sorttable_sorted):not(.sorttable_sorted_reverse):not(.sorttable_nosort):after { 
    content: " \25B4\25BE" 
	}
 </style>

<!--- eof on/off scripts !-->

<div class="searchable">
<%
						
	PartnerInternalRecordIdentifier = Request.QueryString("i")
	FirstLetter = Request.QueryString("letter")

%>
<% If FirstLetter <> "all" Then %>
	<h1 class="page-header"><i class="fa fa-map-marker" aria-hidden="true"></i> Map Customer Accounts Beginning with the Letter "<%= FirstLetter %>"
<% Else %>
	<h1 class="page-header"><i class="fa fa-map-marker" aria-hidden="true"></i> Map Customer Accounts - ALL
<% End If %>

<a href="selectCustomerToEditByAlphabet.asp?i=<%= PartnerInternalRecordIdentifier %>"><button type="button" class="btn btn-lg btn-success btn-create">Back To Customer List</button></a></h1>

	<!-- tabs start here !-->
	<div class="container">

		 <!-- narrow search -->
              <div class="row row-filter">
              <div class="col-lg-4">
		<div class="input-group"> <span class="input-group-addon">Find Customer</span>
		    <input id="filter" type="text" class="form-control " placeholder="Type here...">
		</div>
	</div>
			</div>
              <!-- eof narrow search -->

		<div class="panel panel-default panel-table">
              <div class="panel-heading">
                <div class="row">
                  <div class="col col-xs-10">
                    <h3 class="panel-title">NOTE: Accounts will be automatically saved when you type or click out of the box.</h3>
                  </div>
                  <div class="col col-xs-2text-right">
                    <!--<button type="button" class="btn btn-lg btn-primary btn-create">Save All Changes</button>-->
                  </div>
                </div>
              </div>

             

              <div class="panel-body">
                <table class="table table-striped table-bordered table-list sortable">
                  <thead>
                    <tr>
                        <th>Our Cust ID</th>
                        <th>Account Name</th>
                        <th>Account Address</th>
                        <th  class="sorttable_nosort">Equivalent Cust ID</th>
                    </tr> 
                  </thead>
                  <tbody>
						<%
						Set cnnCustomerTable = Server.CreateObject("ADODB.Connection")
						cnnCustomerTable.open (Session("ClientCnnString"))
						Set rsCustomerTable = Server.CreateObject("ADODB.Recordset")
						rsCustomerTable.CursorLocation = 3 
					
						Set cnnEquivalentCustomers = Server.CreateObject("ADODB.Connection")
						cnnEquivalentCustomers.open (Session("ClientCnnString"))
						Set rsEquivalentCustomers = Server.CreateObject("ADODB.Recordset")
						rsEquivalentCustomers.CursorLocation = 3 
						
						If FirstLetter = "all" Then				
							SQLCustomersTable = "SELECT * FROM AR_Customer WHERE AcctStatus = 'A' ORDER BY Name ASC"
						Else
							SQLCustomersTable = "SELECT * FROM AR_Customer WHERE LEFT(Name,1) = '" & FirstLetter & "' AND AcctStatus = 'A' ORDER BY Name, CustNum ASC"
						End If
						
						'Response.write(SQLCustomersTable)
				
						Set rsCustomerTable = cnnCustomerTable.Execute(SQLCustomersTable)
				
						If NOT rsCustomerTable.EOF Then
													

							Do While Not rsCustomerTable.EOF
									
								customerID = rsCustomerTable("CustNum")
								customerName = rsCustomerTable("Name") 
								customerAddr1 = rsCustomerTable("Addr1") 
								customerAddr2 = rsCustomerTable("Addr2") 
								customerCityStateZip = rsCustomerTable("CityStateZip") 
								customerPhone = rsCustomerTable("Phone")
								
								SQLEquivalentCustomers = "SELECT * FROM AR_CustomerMapping WHERE "
								SQLEquivalentCustomers = SQLEquivalentCustomers & "partnerRecID = " & PartnerInternalRecordIdentifier & " AND "
								SQLEquivalentCustomers = SQLEquivalentCustomers & "ourCustID = '" & customerID & "'"
								
								'Response.Write(SQLEquivalentCustomers)	
								
								Set rsEquivalentCustomers = cnnEquivalentCustomers.Execute(SQLEquivalentCustomers)
								
								If NOT rsEquivalentCustomers.EOF Then
									partnerEquivalentCustID = rsEquivalentCustomers("partnerCustID")
								Else
									partnerEquivalentCustID = ""
								End If
						        %>
	                          
								<tr>
		                            <td><%= customerID %></td>
		                            <td><strong><%= customerName %></strong></td>
		                            <td class="description">
		                            	<%= customerAddr1 %>
		                            	<% If customerAddr2 <> "" Then Response.Write("<br>" & customerAddr2 & "<br>") %>
		                            	<%= customerCityStateZip %><br>
		                            	<%= customerPhone %></td>
		                            <td>

		                            	 <input type="text" id="txtPartnerEquivalentCustomer*<%= customerID %>*<%= PartnerInternalRecordIdentifier %>" name="txtPartnerEquivalentCustomer<%= customerID %>" value="<%= partnerEquivalentCustID %>" class="equiv-sku-input form-control" onChange="saveCustomerIDToEquivalentsTable(this.value, this.id)"> </td>
	                            </tr>		                            
	                            		                          
			        			<%
								rsCustomerTable.movenext
							Loop
							
						End If
						
						set rsCustomerTable = Nothing
						cnnCustomerTable.close
						set cnnCustomerTable = Nothing
						
						set rsEquivalentCustomers = Nothing
						cnnEquivalentCustomers.close
						set cnnEquivalentCustomers = Nothing
						
			            %>
 
                  
                        </tbody>
                </table>
            
              </div>
              <div class="panel-footer">
                <div class="row">
                  <div class="col col-xs-6">
                    <h3 class="panel-title"><strong>END ALL CUSTOMERS</strong> Account Listing</h3>
                  </div>
                  <div class="col col-xs-6 text-right">
                    <!--<<button type="button" class="btn btn-lg btn-primary btn-create">Save All Changes</button>-->
                  </div>
                </div>
            </div>              
						<!-- table line !-->
		</div>

</div>
<!-- eof row !-->
</div>								

<!--#include file="../../inc/footer-main.asp"-->
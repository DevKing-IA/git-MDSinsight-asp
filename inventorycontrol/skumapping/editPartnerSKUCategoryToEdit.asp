<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/jquery_table_search.asp"-->
<!--#include file="../../inc/InsightFuncs_InventoryControl.asp"-->

<script type="text/javascript">
		
 	
	function saveSKUToProductEquivalentsTable(inputEnteredValue, inputID) {
	
		    //alert("Entered Value: " + inputEnteredValue + "-------Input ID: " + inputID);
		    
	   		//When the user blurs off of an input box, make an Ajax post to InSightFuncs_AjaxForInventoryControlModals.asp
	   		
	   		$.ajax({
				type:"POST",
				url: "../../inc/InSightFuncs_AjaxForInventoryControlModals.asp",
				cache: false,
				data: "action=SaveEquivalentSKUandUM&id=" + encodeURIComponent(inputID) + "&sku=" + encodeURIComponent(inputEnteredValue),
				success: function(response)
				 {
				 	if (response == "Success") {
				 		//If successfully saved, change the style of the input box, so the user gets a visual cue that the SKU was saved   
				 		$("#" + inputID).addClass("saved"); 
				 	}       	 
	             }	
			});	//end ajax post to data: "action=SaveEquivalentSKUandUM"
		
			
	}	

</script>

 

 <style type="text/css">
	

	.container{
		max-width:1600px !important;
		margin-left:100px;
	}
	.panel-table {
	  width:1600px;
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
	  text-align:center;
	}
	
	.description{
		line-height:14px !important;
		font-size:10px;
	}	

	.saved{
		background-color:#aeeaae;
	}

	.row-filter{
		margin-bottom: 15px;
	}

	table.sortable th:not(.sorttable_sorted):not(.sorttable_sorted_reverse):not(.sorttable_nosort):after { 
    content: " \25B4\25BE" 
	}
 </style>

<!--- eof on/off scripts !-->
<%
						
	PartnerInternalRecordIdentifier = Request.QueryString("i")
	CategoryID = Request.QueryString("c")
	
	If CategoryID = "all" Then
		CategoryName = "ALL CATEGORIES"
	Else
		CategoryName = GetCategoryByID(CategoryID)
	End If

%>


<div class="searchable">

<h1 class="page-header"><i class="fa fa-map-marker" aria-hidden="true"></i> Map SKUs in <%= CategoryName %> <a href="selectPartnerSKUCategoryToEdit.asp?i=<%= PartnerInternalRecordIdentifier %>"><button type="button" class="btn btn-lg btn-success btn-create">Back To Categories</button></a></h1>

	<!-- tabs start here !-->
	<div class="container">

		 <!-- narrow search -->
              <div class="row row-filter">
              <div class="col-lg-4">
		<div class="input-group"> <span class="input-group-addon">Find Product</span>
		    <input id="filter" type="text" class="form-control " placeholder="Type here...">
		</div>
	</div>
			</div>
              <!-- eof narrow search -->


		<div class="panel panel-default panel-table">
              <div class="panel-heading">
                <div class="row">
                  <div class="col col-xs-6">
                    <h3 class="panel-title">NOTE: SKUs will be automatically saved when you type or click out of the box.</h3>
                  </div>
                  <div class="col col-xs-6 text-right">
                    <!--<button type="button" class="btn btn-lg btn-primary btn-create">Save All Changes</button>-->
                  </div>
                </div>
              </div>
              <div class="panel-body">
                <table class="table table-striped table-bordered table-list sortable">
                  <thead>
                    <tr>
                        <th>My SKU</th>
                        <th>My DESC</th>
                        <th>My UM</th>
                        <th class="sorttable_nosort">Equivalent SKU 1</th>
                        <th class="sorttable_nosort">Equivalent SKU 2</th>
                        <th class="sorttable_nosort">Equivalent SKU 3</th>
                        <th class="sorttable_nosort">Equivalent SKU 4</th>
                        <th class="sorttable_nosort">Equivalent SKU 5</th>
                        <th class="sorttable_nosort">Equivalent SKU 6</th>
                        <th><i class="fas fa-trash-alt"></i></th>
                    </tr> 
                  </thead>
                  <tbody>
						<%
						Set cnnProductsTable = Server.CreateObject("ADODB.Connection")
						cnnProductsTable.open (Session("ClientCnnString"))
						Set rsProductsTable = Server.CreateObject("ADODB.Recordset")
						rsProductsTable.CursorLocation = 3 
					
						Set cnnEquivalentSKUs = Server.CreateObject("ADODB.Connection")
						cnnEquivalentSKUs.open (Session("ClientCnnString"))
						Set rsEquivalentSKUs = Server.CreateObject("ADODB.Recordset")
						rsEquivalentSKUs.CursorLocation = 3 
						
						If CategoryID = "all" Then				
							SQLProductsTable = "SELECT * FROM Product ORDER BY PartNo ASC"
						Else
							SQLProductsTable = "SELECT * FROM Product WHERE Category = " & CategoryID & " ORDER BY PartNo ASC"
						End If
				
						Set rsProductsTable = cnnProductsTable.Execute(SQLProductsTable)
				
						If NOT rsProductsTable.EOF Then
													

							Do While Not rsProductsTable.EOF
	
								CategoryIDToPass = rsProductsTable("Category")
								SKUFromProductsTable = rsProductsTable("PartNo")
								UMFromProductsTable = rsProductsTable("CasePricing")
								DESCFromProductsTable = rsProductsTable("Description") 

								
								If UMFromProductsTable = "U" OR UMFromProductsTable = "C" Then 
									UnitDESCFromProductsTable = rsProductsTable("Description") 
									CaseDESCFromProductsTable = rsProductsTable("CaseDescription") 
								End If
								
								SQLEquivalentSKUs = "SELECT * FROM IC_ProductMapping WHERE SKU = '" & SKUFromProductsTable & "' AND partnerIntRecID = " & PartnerInternalRecordIdentifier
								
								'Response.Write(SQLEquivalentSKUs & "<br>")
								
								Set rsEquivalentSKUs = cnnEquivalentSKUs.Execute(SQLEquivalentSKUs)
								
									
								partnerEquivalentSKU1Unit = ""
								partnerEquivalentSKU2Unit = ""
								partnerEquivalentSKU3Unit = ""
								partnerEquivalentSKU4Unit = ""
								partnerEquivalentSKU5Unit = ""
								partnerEquivalentSKU6Unit = ""
								SKUInternalRecordIdentifierUnit = ""
										
								partnerEquivalentSKU1Case = ""
								partnerEquivalentSKU2Case = ""
								partnerEquivalentSKU3Case = ""
								partnerEquivalentSKU4Case = ""
								partnerEquivalentSKU5Case = ""
								partnerEquivalentSKU6Case = ""
								SKUInternalRecordIdentifierCase = ""

								partnerEquivalentSKU1N = ""
								partnerEquivalentSKU2N = ""
								partnerEquivalentSKU3N = ""
								partnerEquivalentSKU4N = ""
								partnerEquivalentSKU5N = ""
								partnerEquivalentSKU6N = ""
								SKUInternalRecordIdentifierN = ""
									
								
								If NOT rsEquivalentSKUs.EOF Then
									
								
									Do While Not rsEquivalentSKUs.EOF
										
										UM = Trim(rsEquivalentSKUs("UM"))
										
										If UM = "U" Then
										
											partnerEquivalentSKU1Unit = rsEquivalentSKUs("partnerEquivalentSKU1")
											partnerEquivalentSKU2Unit = rsEquivalentSKUs("partnerEquivalentSKU2")
											partnerEquivalentSKU3Unit = rsEquivalentSKUs("partnerEquivalentSKU3")
											partnerEquivalentSKU4Unit = rsEquivalentSKUs("partnerEquivalentSKU4")
											partnerEquivalentSKU5Unit = rsEquivalentSKUs("partnerEquivalentSKU5")
											partnerEquivalentSKU6Unit = rsEquivalentSKUs("partnerEquivalentSKU6")
											
											SKUInternalRecordIdentifierUnit = rsEquivalentSKUs("InternalRecordIdentifier")
											
											If partnerEquivalentSKU1Unit <> "" Then partnerEquivalentSKU1Unit = Replace(partnerEquivalentSKU1Unit, """", "&quot;")
											If partnerEquivalentSKU2Unit <> "" Then partnerEquivalentSKU2Unit = Replace(partnerEquivalentSKU2Unit, """", "&quot;")
											If partnerEquivalentSKU3Unit <> "" Then partnerEquivalentSKU3Unit = Replace(partnerEquivalentSKU3Unit, """", "&quot;")
											If partnerEquivalentSKU4Unit <> "" Then partnerEquivalentSKU4Unit = Replace(partnerEquivalentSKU4Unit, """", "&quot;")
											If partnerEquivalentSKU5Unit <> "" Then partnerEquivalentSKU5Unit = Replace(partnerEquivalentSKU5Unit, """", "&quot;")
											If partnerEquivalentSKU6Unit <> "" Then partnerEquivalentSKU6Unit = Replace(partnerEquivalentSKU6Unit, """", "&quot;")
											
										
										ElseIf UM = "C" Then
										
											partnerEquivalentSKU1Case = rsEquivalentSKUs("partnerEquivalentSKU1")
											partnerEquivalentSKU2Case = rsEquivalentSKUs("partnerEquivalentSKU2")
											partnerEquivalentSKU3Case = rsEquivalentSKUs("partnerEquivalentSKU3")
											partnerEquivalentSKU4Case = rsEquivalentSKUs("partnerEquivalentSKU4")
											partnerEquivalentSKU5Case = rsEquivalentSKUs("partnerEquivalentSKU5")
											partnerEquivalentSKU6Case = rsEquivalentSKUs("partnerEquivalentSKU6")
											
											SKUInternalRecordIdentifierCase = rsEquivalentSKUs("InternalRecordIdentifier")
											
											If partnerEquivalentSKU1Case <> "" Then partnerEquivalentSKU1Case = Replace(partnerEquivalentSKU1Case, """", "&quot;")
											If partnerEquivalentSKU2Case <> "" Then partnerEquivalentSKU2Case = Replace(partnerEquivalentSKU2Case, """", "&quot;")
											If partnerEquivalentSKU3Case <> "" Then partnerEquivalentSKU3Case = Replace(partnerEquivalentSKU3Case, """", "&quot;")
											If partnerEquivalentSKU4Case <> "" Then partnerEquivalentSKU4Case = Replace(partnerEquivalentSKU4Case, """", "&quot;")
											If partnerEquivalentSKU5Case <> "" Then partnerEquivalentSKU5Case = Replace(partnerEquivalentSKU5Case, """", "&quot;")
											If partnerEquivalentSKU6Case <> "" Then partnerEquivalentSKU6Case = Replace(partnerEquivalentSKU6Case, """", "&quot;")
											
											
										ElseIf UM = "N" Then
										
											partnerEquivalentSKU1N = rsEquivalentSKUs("partnerEquivalentSKU1")
											partnerEquivalentSKU2N = rsEquivalentSKUs("partnerEquivalentSKU2")
											partnerEquivalentSKU3N = rsEquivalentSKUs("partnerEquivalentSKU3")
											partnerEquivalentSKU4N = rsEquivalentSKUs("partnerEquivalentSKU4")
											partnerEquivalentSKU5N = rsEquivalentSKUs("partnerEquivalentSKU5")
											partnerEquivalentSKU6N = rsEquivalentSKUs("partnerEquivalentSKU6")
											
											SKUInternalRecordIdentifierN = rsEquivalentSKUs("InternalRecordIdentifier")
											
											If partnerEquivalentSKU1N <> "" Then partnerEquivalentSKU1N = Replace(partnerEquivalentSKU1N, """", "&quot;")
											If partnerEquivalentSKU2N <> "" Then partnerEquivalentSKU2N = Replace(partnerEquivalentSKU2N, """", "&quot;")
											If partnerEquivalentSKU3N <> "" Then partnerEquivalentSKU3N = Replace(partnerEquivalentSKU3N, """", "&quot;")
											If partnerEquivalentSKU4N <> "" Then partnerEquivalentSKU4N = Replace(partnerEquivalentSKU4N, """", "&quot;")
											If partnerEquivalentSKU5N <> "" Then partnerEquivalentSKU5N = Replace(partnerEquivalentSKU5N, """", "&quot;")
											If partnerEquivalentSKU6N <> "" Then partnerEquivalentSKU6N = Replace(partnerEquivalentSKU6N, """", "&quot;")
			
										End If
										
										rsEquivalentSKUs.movenext
									Loop
									
								End If
						        %>
		                          
								<% If UMFromProductsTable = "U" OR UMFromProductsTable = "C" Then %>
									<tr>
			                            <td><%= SKUFromProductsTable %></td>
			                            <td class="description"><%= UnitDESCFromProductsTable %></td>
			                            <td>U</td>
			                            <td><input type="text" id="txtpartnerEquivalentSKU1*U*<%= SKUFromProductsTable %>*<%= PartnerInternalRecordIdentifier %>*<%= CategoryIDToPass %>" name="txtpartnerEquivalentSKU1*U*<%= SKUFromProductsTable %>" value="<%= partnerEquivalentSKU1Unit %>"  class="equiv-sku-input form-control" onChange="saveSKUToProductEquivalentsTable(this.value, this.id)"></td>
			                            <td><input type="text" id="txtpartnerEquivalentSKU2*U*<%= SKUFromProductsTable %>*<%= PartnerInternalRecordIdentifier %>*<%= CategoryIDToPass %>" name="txtpartnerEquivalentSKU2*U*<%= SKUFromProductsTable %>" value="<%= partnerEquivalentSKU2Unit %>"  class="equiv-sku-input form-control" onChange="saveSKUToProductEquivalentsTable(this.value, this.id)"></td>
			                            <td><input type="text" id="txtpartnerEquivalentSKU3*U*<%= SKUFromProductsTable %>*<%= PartnerInternalRecordIdentifier %>*<%= CategoryIDToPass %>" name="txtpartnerEquivalentSKU3*U*<%= SKUFromProductsTable %>" value="<%= partnerEquivalentSKU3Unit %>"  class="equiv-sku-input form-control" onChange="saveSKUToProductEquivalentsTable(this.value, this.id)"></td>
			                            <td><input type="text" id="txtpartnerEquivalentSKU4*U*<%= SKUFromProductsTable %>*<%= PartnerInternalRecordIdentifier %>*<%= CategoryIDToPass %>" name="txtpartnerEquivalentSKU4*U*<%= SKUFromProductsTable %>" value="<%= partnerEquivalentSKU4Unit %>"  class="equiv-sku-input form-control" onChange="saveSKUToProductEquivalentsTable(this.value, this.id)"></td>
			                            <td><input type="text" id="txtpartnerEquivalentSKU5*U*<%= SKUFromProductsTable %>*<%= PartnerInternalRecordIdentifier %>*<%= CategoryIDToPass %>" name="txtpartnerEquivalentSKU5*U*<%= SKUFromProductsTable %>" value="<%= partnerEquivalentSKU5Unit %>"  class="equiv-sku-input form-control" onChange="saveSKUToProductEquivalentsTable(this.value, this.id)"></td>
			                            <td><input type="text" id="txtpartnerEquivalentSKU6*U*<%= SKUFromProductsTable %>*<%= PartnerInternalRecordIdentifier %>*<%= CategoryIDToPass %>" name="txtpartnerEquivalentSKU6*U*<%= SKUFromProductsTable %>" value="<%= partnerEquivalentSKU6Unit %>"  class="equiv-sku-input form-control" onChange="saveSKUToProductEquivalentsTable(this.value, this.id)"></td>
		                            	
		                            	<% If SKUInternalRecordIdentifierUnit <> "" Then %>
		                            		<% If CategoryID = "all" Then %>
		                            			<td><a href="deletePartnerSKUandUMQues.asp?i=<%= SKUInternalRecordIdentifierUnit %>&x=<%= PartnerInternalRecordIdentifier %>&p=<%= SKUFromProductsTable %>&u=U&c=all"><i class="fa fa-eraser" aria-hidden="true"></i></a></td>
		                            		<% Else %>
		                            			<td><a href="deletePartnerSKUandUMQues.asp?i=<%= SKUInternalRecordIdentifierUnit %>&x=<%= PartnerInternalRecordIdentifier %>&p=<%= SKUFromProductsTable %>&u=U&c=<%= CategoryIDToPass %>"><i class="fa fa-eraser" aria-hidden="true"></i></a></td>
		                            		<% End If %>
		                            	<% Else %>
		                            		<td>&nbsp;</td>
		                            	<% End If %>
		                            	
		                            </tr> 
									<tr>
			                            <td><%= SKUFromProductsTable %></td>
			                            <td class="description"><%= CaseDESCFromProductsTable %></td>
			                            <td>C</td>
			                            <td><input type="text" id="txtpartnerEquivalentSKU1*C*<%= SKUFromProductsTable %>*<%= PartnerInternalRecordIdentifier %>*<%= CategoryIDToPass %>" name="txtpartnerEquivalentSKU1*C*<%= SKUFromProductsTable %>" value="<%= partnerEquivalentSKU1Case %>"  class="equiv-sku-input form-control" onChange="saveSKUToProductEquivalentsTable(this.value, this.id)"></td>
			                            <td><input type="text" id="txtpartnerEquivalentSKU2*C*<%= SKUFromProductsTable %>*<%= PartnerInternalRecordIdentifier %>*<%= CategoryIDToPass %>" name="txtpartnerEquivalentSKU2*C*<%= SKUFromProductsTable %>" value="<%= partnerEquivalentSKU2Case %>"  class="equiv-sku-input form-control" onChange="saveSKUToProductEquivalentsTable(this.value, this.id)"></td>
			                            <td><input type="text" id="txtpartnerEquivalentSKU3*C*<%= SKUFromProductsTable %>*<%= PartnerInternalRecordIdentifier %>*<%= CategoryIDToPass %>" name="txtpartnerEquivalentSKU3*C*<%= SKUFromProductsTable %>" value="<%= partnerEquivalentSKU3Case %>"  class="equiv-sku-input form-control" onChange="saveSKUToProductEquivalentsTable(this.value, this.id)"></td>
			                            <td><input type="text" id="txtpartnerEquivalentSKU4*C*<%= SKUFromProductsTable %>*<%= PartnerInternalRecordIdentifier %>*<%= CategoryIDToPass %>" name="txtpartnerEquivalentSKU4*C*<%= SKUFromProductsTable %>" value="<%= partnerEquivalentSKU4Case %>"  class="equiv-sku-input form-control" onChange="saveSKUToProductEquivalentsTable(this.value, this.id)"></td>
			                            <td><input type="text" id="txtpartnerEquivalentSKU5*C*<%= SKUFromProductsTable %>*<%= PartnerInternalRecordIdentifier %>*<%= CategoryIDToPass %>" name="txtpartnerEquivalentSKU5*C*<%= SKUFromProductsTable %>" value="<%= partnerEquivalentSKU5Case %>"  class="equiv-sku-input form-control" onChange="saveSKUToProductEquivalentsTable(this.value, this.id)"></td>
			                            <td><input type="text" id="txtpartnerEquivalentSKU6*C*<%= SKUFromProductsTable %>*<%= PartnerInternalRecordIdentifier %>*<%= CategoryIDToPass %>" name="txtpartnerEquivalentSKU6*C*<%= SKUFromProductsTable %>" value="<%= partnerEquivalentSKU6Case %>"  class="equiv-sku-input form-control" onChange="saveSKUToProductEquivalentsTable(this.value, this.id)"></td>
		                            	
		                            	<% If SKUInternalRecordIdentifierCase <> "" Then %>
		                            		<% If CategoryID = "all" Then %>
		                            			<td><a href="deletePartnerSKUandUMQues.asp?i=<%= SKUInternalRecordIdentifierCase %>&x=<%= PartnerInternalRecordIdentifier %>&p=<%= SKUFromProductsTable %>&u=U&c=all"><i class="fa fa-eraser" aria-hidden="true"></i></a></td>
		                            		<% Else %>
		                            			<td><a href="deletePartnerSKUandUMQues.asp?i=<%= SKUInternalRecordIdentifierCase %>&x=<%= PartnerInternalRecordIdentifier %>&p=<%= SKUFromProductsTable %>&u=U&c=<%= CategoryIDToPass %>"><i class="fa fa-eraser" aria-hidden="true"></i></a></td>
		                            		<% End If %>
		                            	<% Else %>
		                            		<td>&nbsp;</td>
		                            	<% End If %>
		                            	
		                            </tr>		                            
	                            <% ElseIf UMFromProductsTable = "N" Then %>
									<tr>
			                            <td><%= SKUFromProductsTable %></td>
			                            <td class="description"><%= DESCFromProductsTable %></td>
			                            <td>N</td>
			                            <td><input type="text" id="txtpartnerEquivalentSKU1*N*<%= SKUFromProductsTable %>*<%= PartnerInternalRecordIdentifier %>*<%= CategoryIDToPass %>" name="txtpartnerEquivalentSKU1*N*<%= SKUFromProductsTable %>" value="<%= partnerEquivalentSKU1N %>"  class="equiv-sku-input form-control" onChange="saveSKUToProductEquivalentsTable(this.value, this.id)"></td>
			                            <td><input type="text" id="txtpartnerEquivalentSKU2*N*<%= SKUFromProductsTable %>*<%= PartnerInternalRecordIdentifier %>*<%= CategoryIDToPass %>" name="txtpartnerEquivalentSKU2*N*<%= SKUFromProductsTable %>" value="<%= partnerEquivalentSKU2N %>"  class="equiv-sku-input form-control" onChange="saveSKUToProductEquivalentsTable(this.value, this.id)"></td>
			                            <td><input type="text" id="txtpartnerEquivalentSKU3*N*<%= SKUFromProductsTable %>*<%= PartnerInternalRecordIdentifier %>*<%= CategoryIDToPass %>" name="txtpartnerEquivalentSKU3*N*<%= SKUFromProductsTable %>" value="<%= partnerEquivalentSKU3N %>"  class="equiv-sku-input form-control" onChange="saveSKUToProductEquivalentsTable(this.value, this.id)"></td>
			                            <td><input type="text" id="txtpartnerEquivalentSKU4*N*<%= SKUFromProductsTable %>*<%= PartnerInternalRecordIdentifier %>*<%= CategoryIDToPass %>" name="txtpartnerEquivalentSKU4*N*<%= SKUFromProductsTable %>" value="<%= partnerEquivalentSKU4N %>"  class="equiv-sku-input form-control" onChange="saveSKUToProductEquivalentsTable(this.value, this.id)"></td>
			                            <td><input type="text" id="txtpartnerEquivalentSKU5*N*<%= SKUFromProductsTable %>*<%= PartnerInternalRecordIdentifier %>*<%= CategoryIDToPass %>" name="txtpartnerEquivalentSKU5*N*<%= SKUFromProductsTable %>" value="<%= partnerEquivalentSKU5N %>"  class="equiv-sku-input form-control" onChange="saveSKUToProductEquivalentsTable(this.value, this.id)"></td>
			                            <td><input type="text" id="txtpartnerEquivalentSKU6*N*<%= SKUFromProductsTable %>*<%= PartnerInternalRecordIdentifier %>*<%= CategoryIDToPass %>" name="txtpartnerEquivalentSKU6*N*<%= SKUFromProductsTable %>" value="<%= partnerEquivalentSKU6N %>"  class="equiv-sku-input form-control" onChange="saveSKUToProductEquivalentsTable(this.value, this.id)"></td>
	
		                            	<% If SKUInternalRecordIdentifierN <> "" Then %>
		                            		<% If CategoryID = "all" Then %>
		                            			<td><a href="deletePartnerSKUandUMQues.asp?i=<%= SKUInternalRecordIdentifierN %>&x=<%= PartnerInternalRecordIdentifier %>&p=<%= SKUFromProductsTable %>&u=U&c=all"><i class="fa fa-eraser" aria-hidden="true"></i></a></td>
		                            		<% Else %>
		                            			<td><a href="deletePartnerSKUandUMQues.asp?i=<%= SKUInternalRecordIdentifierN %>&x=<%= PartnerInternalRecordIdentifier %>&p=<%= SKUFromProductsTable %>&u=U&c=<%= CategoryIDToPass %>"><i class="fa fa-eraser" aria-hidden="true"></i></a></td>
		                            		<% End If %>
		                            	<% Else %>
		                            		<td>&nbsp;</td>
		                            	<% End If %>
		                            		                            		
		                            </tr>		                            
	                            <% End If %>
		                          
			        			<%
								rsProductsTable.movenext
							Loop
							
						End If
						
						set rsProductsTable = Nothing
						cnnProductsTable.close
						set cnnProductsTable = Nothing
						
						set rsEquivalentSKUs = Nothing
						cnnEquivalentSKUs.close
						set cnnEquivalentSKUs = Nothing
						
			            %>
 
                  
                        </tbody>
                </table>
            
              </div>
              <div class="panel-footer">
                <div class="row">
                  <div class="col col-xs-6">
                    <h3 class="panel-title"><strong>END ALL CATEGORIES</strong> SKU Listing</h3>
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
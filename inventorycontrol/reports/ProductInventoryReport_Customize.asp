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
				<h4 class="modal-title" id="myModalLabel" align="center">Product UPC Report Customization Settings</h4>
			</div>

			<form method="post" action="ProductInventoryReport_Customize_SaveValues.asp" name="frmProductInventoryReport_Customize" >

			      <!-- insert content in here !-->
			      <div class="modal-body ativa-scroll">
 	      	
		 	      	<!-- date ranges !-->
			      	<div class="container-fluid container-modal">
			      	
				      	<div class="row">
 		      	
		 		      	<!-- left column !-->
		 		      	<div class="col-lg-2 col-md-2 col-sm-12 col-xs-12 left-column">
			 		      	<h4><br>Unit UPC</h4>
		 		      	</div>
		 		      	<!-- eof left column !-->
 		      	
		 		      	<!-- right column !-->
		 		      	<div class="col-lg-4 col-md-4 col-sm-12 col-xs-12 right-column">
	 		      	
				      	<!-- row !-->
					      	<div class="row">
					      	
						        <!-- First Date !-->
						  
							    	<% If UnitUPCData = "NOTEMPTY" Then %>
								    	<input type="radio" id="optUnitUPCData" name="optUnitUPCData" value="NOTEMPTY"  checked="checked"> Unit UPC NOT Empty
								    <% Else %>
								    	<input type="radio" id="optUnitUPCData" name="optUnitUPCData" value="NOTEMPTY"> Unit UPC NOT Empty
								    <% End If %>
								    
								    
								    <br><br>					        
					        
							    	<% If UnitUPCData = "EMPTY" Then %>
								    	<input type="radio" id="optUnitUPCData" name="optUnitUPCData" value="EMPTY"  checked="checked"> Unit UPC Is Empty
								    <% Else %>
								    	<input type="radio" id="optUnitUPCData" name="optUnitUPCData" value="EMPTY"> Unit UPC Is Empty
								    <% End If %>
								    
								    <br><br>
		  					    
		  					    					        						      						        							    				<% If UnitUPCData = "" Then %>
								    	<input type="radio" id="optUnitUPCData" name="optUnitUPCData" value=""  checked="checked"> Show All Unit UPC's
								    <% Else %>
								    	<input type="radio" id="optUnitUPCData" name="optUnitUPCData" value=""> Show All Unit UPC's
								    <% End If %>					        						     
		  					    


		    			  	</div>
		    			  	<!-- eof row !-->
			 		      </div>
		 		      	<!-- eof right column !-->

 		      	
		 		      	<!-- left column !-->
		 		      	<div class="col-lg-2 col-md-2 col-sm-12 col-xs-12 left-column">
			 		      	<h4><br>Case UPC</h4>
		 		      	</div>
		 		      	<!-- eof left column !-->
 		      	
		 		      	<!-- right column !-->
		 		      	<div class="col-lg-4 col-md-4 col-sm-12 col-xs-12 right-column">
	 		      	
				      	<!-- row !-->
					      	<div class="row">

								<% If CaseUPCData = "NOTEMPTY" Then %>
									<input type="radio" id="optCaseUPCData" name="optCaseUPCData" value="NOTEMPTY"  checked="checked"> Case UPC NOT Empty
								<% Else %>
									<input type="radio" id="optCaseUPCData" name="optCaseUPCData" value="NOTEMPTY"> Case UPC NOT Empty
								<% End If %>


								<br>
								<br>					        

								<% If CaseUPCData = "EMPTY" Then %>
									<input type="radio" id="optCaseUPCData" name="optCaseUPCData" value="EMPTY"  checked="checked"> Case UPC Is Empty
								<% Else %>
									<input type="radio" id="optCaseUPCData" name="optCaseUPCData" value="EMPTY"> Case UPC Is Empty
								<% End If %>

								<br><br>		  					    					        						      						        							    				<% If CaseUPCData = "" Then %>
								    	<input type="radio" id="optCaseUPCData" name="optCaseUPCData" value=""  checked="checked"> Show All Case UPC's
								    <% Else %>
								    	<input type="radio" id="optCaseUPCData" name="optCaseUPCData" value=""> Show All Case UPC's
								<% End If %>					        						     


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
			 		      	<h4><br>Inventoried Items</h4>
 				      	</div>
 				      	<!-- eof left column !-->
 		      	
 			 		      	<!-- right column !-->
		 		      	<div class="col-lg-10 col-md-9 col-sm-12 col-xs-12 right-column">
	 		      	
				      	<!-- row !-->
					      	<div class="row">


								<% If InventoriedItem = "YES" Then %>
									<input type="radio" id="optInventoriedItem" name="optInventoriedItem" value="YES"  checked="checked"> Inventoried Item = YES
								<% Else %>
									<input type="radio" id="optInventoriedItem" name="optInventoriedItem" value="YES"> Inventoried Item = YES
								<% End If %>
								
								<br><br>
								
								<% If InventoriedItem = "NO" Then %>
									<input type="radio" id="optInventoriedItem" name="optInventoriedItem" value="NO"  checked="checked">  Inventoried Item = NO
								<% Else %>
									<input type="radio" id="optInventoriedItem" name="optInventoriedItem" value="NO"> Inventoried Item = NO
								<% End If %>

								<br><br>
								
								<% If InventoriedItem = "" Then %>
									<input type="radio" id="optInventoriedItem" name="optInventoriedItem" value=""  checked="checked"> Show All Inventory Statuses
								<% Else %>
									<input type="radio" id="optInventoriedItem" name="optInventoriedItem" value="">Show All Inventory Statuses
								<% End If %>																								


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
			 		      	<h4><br>Pickable Items</h4>
 				      	</div>
 				      	<!-- eof left column !-->
 		      	
 			 		      	<!-- right column !-->
		 		      	<div class="col-lg-10 col-md-9 col-sm-12 col-xs-12 right-column">
	 		      	
				      	<!-- row !-->
					      	<div class="row">


								<% If PickableItem = "YES" Then %>
									<input type="radio" id="optPickableItem" name="optPickableItem" value="YES"  checked="checked"> Pickable Item = YES
								<% Else %>
									<input type="radio" id="optPickableItem" name="optPickableItem" value="YES"> Pickable Item = YES
								<% End If %>
								
								<br><br>
								
								<% If PickableItem = "NO" Then %>
									<input type="radio" id="optPickableItem" name="optPickableItem" value="NO"  checked="checked">  Pickable Item = NO
								<% Else %>
									<input type="radio" id="optPickableItem" name="optPickableItem" value="NO"> Pickable Item = NO
								<% End If %>

								<br><br>
								
								<% If PickableItem = "" Then %>
									<input type="radio" id="optPickableItem" name="optPickableItem" value=""  checked="checked"> Show All Pickable Statuses
								<% Else %>
									<input type="radio" id="optPickableItem" name="optPickableItem" value="">Show All Pickable Statuses
								<% End If %>																								


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
			 		      	<h4><br>Product Categories</h4>
 				      	</div>
 				      	<!-- eof left column !-->
 		      	
 				      	<!-- right column !-->
 				      	<div class="col-lg-10 col-md-9 col-sm-12 col-xs-12 right-column check-row">
	 		      	
				      	<!-- row !-->
					      	<div class="row categories-checkboxes ">
   
								<div class="checkbox">

							<%
								ProductCategoryArrayForTab = "" 
								ProductCategoryArrayForTab = Split(ProductCategoriesForInventoryReport,",")

								SQL = "SELECT DISTINCT(prodCategory) FROM IC_Product ORDER BY prodCategory"

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
											For x = 0 to Ubound(ProductCategoryArrayForTab)
												If rs("prodCategory") = ProductCategoryArrayForTab(x) Then ResponseLine = ResponseLine & " checked "
											Next 
											ResponseLine = ResponseLine & "id='chk" & rs("prodCategory") & "' name='chkCategoryNum' value='" & rs("prodCategory") & "'>" & GetCategoryByID(rs("prodCategory")) & "<br>"
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
		

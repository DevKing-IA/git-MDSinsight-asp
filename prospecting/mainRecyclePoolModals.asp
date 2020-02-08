<!-- ******************************************************************************************************************************** -->
<!-- MODAL WINDOW DESIGN AND DEFINITIONS -->
<!-- ******************************************************************************************************************************** -->

<!------------------------------------------------------------------------------>	
<!-- modal for filtering data in the prospecting table/grid view !-->
<!------------------------------------------------------------------------------>

<!--#include file="mainRecyclePoolModalCustomizeDataFilterValues.asp"-->

<!------------------------------------------------------------------------------>	
<!-- END modal for filtering data in the prospecting table/grid view !-->
<!------------------------------------------------------------------------------>


<div class="modal fade bs-modal-show-hide-columns-recycle-pool" tabindex="-1" role="dialog" aria-labelledby="myLargeModalLabel" aria-hidden="true">
<div class="modal-dialog modal-lg modal-height">
<div class="modal-content">

	<style type="text/css">
	.ativa-scroll{
		max-height: 300px
	}
	</style>
	
	<!-- modal scroll !-->
	<script type="text/javascript">
		$(document).ready(ajustamodal);
		$(window).resize(ajustamodal);
		function ajustamodal() {
		//var altura = $(window).height() - 155; //value corresponding to the modal heading + footer
		var altura = $(window).height() - 205; //value corresponding to the modal heading + footer
		$(".ativa-scroll").css({"height":altura,"overflow-y":"auto"});
	}
	</script>
	<!-- eof modal scroll !-->

  <div class="modal-header">
    <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
    <h4 class="modal-title" id="myModalLabel" align="center">Show/Hide <%= GetTerm("Prospecting") %> Columns</h4>
  </div>

	<form method="post" action="mainRecyclePoolCustomizeSaveShowHideColumnValues.asp" name="frmProspectingCustomizeColumnsRecyclePool">

      <!-- insert content in here !-->
      <div class="modal-body ativa-scroll">
 	      	
  	      	<!-- filtering !-->
	      	<div class="container-fluid">
		      	<div class="row">
 		      	
 		      	<!-- left column !-->
 		      	<div class="col-lg-2 col-md-3 col-sm-12 col-xs-12 left-column">
	 		      	<h4><br>Column Names</h4>
 		      	</div>
 		      	<!-- eof left column !-->
 		      	
 		      	<!-- right column !-->
 		      	<div class="col-lg-10 col-md-9 col-sm-12 col-xs-12 right-column">
	 		      	
		      	<!-- row !-->
		      	<div class="row">
     	
		      	<div class="col-lg-3 col-md-3 col-sm-12 col-xs-12">
			      	<strong>Click To Show/Hide</strong>
		      	</div>
				<%
					'************************
					'Read Settings_Reports
					'************************
					SQL = "SELECT * from Settings_Reports where ReportNumber = 1400 AND PoolForProspecting = 'Dead' AND UserNo = " & Session("userNo")
					Set cnn8 = Server.CreateObject("ADODB.Connection")
					cnn8.open (Session("ClientCnnString"))
					Set rs = Server.CreateObject("ADODB.Recordset")
					Set rs= cnn8.Execute(SQL)
					UseSettings_Reports = False
					If NOT rs.EOF Then
						UseSettings_Reports = True
						showHideColumns = rs("ReportSpecificData1")
					End If
					'****************************
					'End Read Settings_Reports
					'****************************
				%>
		      	<div class="col-lg-9 col-md-9 col-sm-12 col-xs-12">
					
					<div class="ck-button">
					<label><input type="checkbox" value="col_address" name="chkCol_address" <% If InStr(showHideColumns,"col_address") Then Response.Write("checked='checked'") %>><span>Street Address</span></label>
					</div>
					<div class="ck-button">
					<label><input type="checkbox" value="col_city" name="chkCol_city" <% If InStr(showHideColumns,"col_city") Then Response.Write("checked='checked'") %>><span>City</span></label>
					</div>
					<div class="ck-button">
					<label><input type="checkbox" value="col_state" name="chkCol_state" <% If InStr(showHideColumns,"col_state") Then Response.Write("checked='checked'") %>><span>State</span></label>
					</div>
					<div class="ck-button">
					<label><input type="checkbox" value="col_zip" name="chkCol_zip" <% If InStr(showHideColumns,"col_zip") Then Response.Write("checked='checked'") %>><span>Zip Code</span></label>
					</div>
					<div class="ck-button">
					<label><input type="checkbox" value="col_leadsource" name="chkCol_leadsource" <% If InStr(showHideColumns,"col_leadsource") Then Response.Write("checked='checked'") %>><span>Lead Source</span></label>
					</div>
					<div class="ck-button">
					<label><input type="checkbox" value="col_leadsource" name="chkCol_stage" <% If InStr(showHideColumns,"col_stage") OR showHideColumns = "" Then Response.Write("checked='checked'") %>><span>Stage</span></label>
					</div>
					<div class="ck-button">
					<label><input type="checkbox" value="col_industry" name="chkCol_industry" <% If InStr(showHideColumns,"col_industry") Then Response.Write("checked='checked'") %>><span>Industry</span></label>
					</div>
					<div class="ck-button">
					<label><input type="checkbox" value="col_numemployees" name="chkCol_numemployees" <% If InStr(showHideColumns,"col_numemployees") Then Response.Write("checked='checked'") %>><span>Number of Employees</span></label>
					</div>
					<div class="ck-button">
					<label><input type="checkbox" value="col_owner" name="chkCol_owner" <% If InStr(showHideColumns,"col_owner") Then Response.Write("checked='checked'") %>><span>Prospect Owner</span></label>
					</div>
					<div class="ck-button">
					<label><input type="checkbox" value="col_createddate" name="chkCol_createddate" <% If InStr(showHideColumns,"col_createddate") Then Response.Write("checked='checked'") %>><span>Prospect Created Date</span></label>
					</div>
					<div class="ck-button">
					<label><input type="checkbox" value="col_createdby" name="chkCol_createdby" <% If InStr(showHideColumns,"col_createdby") Then Response.Write("checked='checked'") %>><span>Prospect Created By</span></label>
					</div>
					<div class="ck-button">
					<label><input type="checkbox" value="col_telemarketer" name="chkCol_telemarketer" <% If InStr(showHideColumns,"col_telemarketer") Then Response.Write("checked='checked'") %>><span>Telemarketer</span></label>
					</div>
					<div class="ck-button">
					<label><input type="checkbox" value="col_numpantries" name="chkCol_numpantries" <% If InStr(showHideColumns,"col_numpantries") Then Response.Write("checked='checked'") %>><span>Number of Pantries</span></label>
					</div>
					<div class="ck-button">
					<label><input type="checkbox" value="col_prospectid" name="chkCol_prospectid" <% If InStr(showHideColumns,"col_prospectid") Then Response.Write("checked='checked'") %>><span>Prospect ID</span></label>
					</div>

		      	</div>
 		      	  		      			      	
		      	</div>
		      	<!-- eof row !-->
		      	
		      	
 		      	</div>
 		      	<!-- eof right column !-->
 		      	
		      	</div>
	      	</div>
       
       </div>
      <!-- eof content insertion !-->
      
      <div class="modal-footer">
        <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
         <a href="#" onClick="document.frmProspectingCustomizeColumnsRecyclePool.submit()"><button type="button" class="btn btn-primary">Save Show/Hide Columns</button></a>     
      </div>
      </form>
    </div>
  </div>
</div>
<!------------------------------------------------------------------------------>
<!-- modal for showing and hiding columns in the prospecting table/grid view !-->
<!------------------------------------------------------------------------------>	
	

	<!-- delete prospect modal starts here !-->
	<!-- modal starts here !-->
		<div class="modal fade" id="myProspectingDeleteModal" tabindex="-1" role="dialog" aria-labelledby="myProspectingDeleteLabel">
		  <div class="modal-dialog" role="document">
		    <div class="modal-content">
		      <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		        <h4 class="modal-title" id="myProspectingDeleteModalLabel">Delete Prospect(s)</h4>
		      </div>
		      <form name="frmDeleteProspects" id="frmDeleteProspects" method="post" action="deleteRecyclePoolProspectsFromModal.asp">
			      <div class="modal-body">
	
					<div class="col-lg-12" id="deleteProspectInfo">
					</div>
	
					<div class="clearfix"></div>
						  
			       </div>
			      <div class="modal-footer">
			        <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
			        <button type="submit" class="btn btn-danger" data-dismiss="modal" onclick="frmDeleteProspects.submit()">Delete Prospect(s)</button>
			      </div>
		      </form>
		    </div>
		  </div>
		</div>
	<!-- modal ends here !-->
	<!-- deleted prospect modal ends here !-->
	

	<!-- new prospect group modal starts here !-->
	<!-- modal starts here !-->
		<div class="modal fade" id="myProspectingWatchModal" tabindex="-1" role="dialog" aria-labelledby="myProspectingWatchLabel">
		  <div class="modal-dialog" role="document">
		    <div class="modal-content">
		      <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		        <h4 class="modal-title" id="myProspectingModalLabel">Watch This Prospect</h4>
		      </div>
		      <form name="frmCreateNewProspectWatch" id="frmCreateNewProspectWatch">
			      <div class="modal-body">
	
					<div class="col-lg-12">
						<div class="form-group">
							<label class="col-sm-12 control-label">Some Fields Here</label>
							<input type="text" class="form-control required" id="txtProspectingWatch" name="txtProspectingWatch">
						</div>
					</div>
	
					<div class="clearfix"></div>
						  
			       </div>
			      <div class="modal-footer">
			        <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
			        <button type="button" class="btn btn-primary" data-dismiss="modal">Watch Prospect</button>
			      </div>
		      </form>
		    </div>
		  </div>
		</div>
	<!-- modal ends here !-->
	<!-- new prospect group modal ends here !-->
	
	

	<!-- new prospect group modal starts here !-->
	<!-- modal starts here !-->
		<div class="modal fade" id="saveAsNewProspectFilterViewRecyclePool" tabindex="-1" role="dialog" aria-labelledby="saveAsNewProspectFilterViewRecyclePool">
		  <div class="modal-dialog" role="document">
		  
		 			  
				<script language="JavaScript">
				<!--
				
				   function validateViewNameRecPool()
				    {
								    				       
					   var viewNameInputField = $("#txtNewFilterReportViewName").val();
					   var viewNameSelectBox = $("#selExistingFilterViewNames option:selected").val();
					   		    
				       if (viewNameInputField == "" && viewNameSelectBox == "") {
				            swal("Please select a name or enter a new name to save this view.");
				            return false;
				       }
				       
				       if (viewNameInputField == "Default" || viewNameInputField == "DEFAULT" || viewNameInputField == "default" || viewNameInputField == "Current" || viewNameInputField == "current" || viewNameInputField == "CURRENT"|| viewNameInputField == "All Prospects" || viewNameInputField == "all prospects" || viewNameInputField == "ALL PROSPECTS") {
				            swal("A view cannot be named DEFAULT, CURRENT or ALL PROSPECTS, they are reserved names.");
				            return false;
				       }
				       return true;
				
				    }
				// -->
				</script>  

		    <div class="modal-content">
		      <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		        <h4 class="modal-title" id="myProspectingModalLabel">Save This View</h4>
		      </div>
		      <form name="frmCreateNewFilterReportViewRecyclePool" id="frmCreateNewFilterReportViewRecyclePool">
			      <div class="modal-body">
	
					<div class="col-lg-12">
						<div class="form-group">
							<label class="col-sm-12 control-label" style="margin-left:-15px;">Save This View As:</label>
							
					      	<%'Report View Name Dropdown
					      	
					      	CurrentViewName = MUV_READ("CRMVIEWSTATE")
					      	CurrentViewNameForSQL = Replace(MUV_READ("CRMVIEWSTATE"),"'","''")
					      	 
					  	  	SQL = "SELECT * FROM Settings_Reports WHERE ReportNumber = 1400 AND UserNo = " & Session("userNo") & " AND PoolForProspecting = 'Dead' "
					  	  	SQL = SQL & " AND UserReportName <> 'Current'  AND UserReportName <> 'Default' AND UserReportName <> 'All Prospects' "
					  	  	SQL = SQL & " ORDER BY UserReportName "
					
							Set cnn8 = Server.CreateObject("ADODB.Connection")
							cnn8.open (Session("ClientCnnString"))
							Set rs = Server.CreateObject("ADODB.Recordset")
							rs.CursorLocation = 3 
							Set rs = cnn8.Execute(SQL)
						
							If NOT rs.EOF Then
							%>
								<!-- Display Report View Names -->
								<select class="form-control when-line" style="width:100%;height:50px;display:inline;margin-left:0px;" name="selExistingFilterViewNames" id="selExistingFilterViewNames">
								<option value=""> -- Choose An Existing View Name To Overwrite -- </option>
								<%
									Do
										selReportName = Replace(rs("UserReportName"),"''","'")
										If MUV_READ("CRMVIEWSTATE") = selReportName Then
											%><option value="<%= selReportName %>" selected="selected"><%= selReportName %></option><%
										Else
											%><option value="<%= selReportName %>"><%= selReportName %></option><%
										End If
										rs.movenext
									Loop until rs.eof
								%>		
								</select>
								<!-- End Display Report View Names -->
							<%
							End If
							set rs = Nothing
							cnn8.close
							set cnn8 = Nothing
					      	%>
							
							<br><br>
							<label class="col-sm-12 control-label" style="margin-left:-15px;">Enter a New View Name:</label>
							<input type="text" class="form-control required"  style="width:100%;height:50px;display:inline;margin-left:0px;" id="txtNewFilterReportViewName" name="txtNewFilterReportViewName">
						</div>
					</div>
	
					<div class="clearfix"></div>
						  
			       </div>
			      <div class="modal-footer">
			        <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
			        <button type="button" class="btn btn-primary" onclick="if (validateViewNameRecPool()) saveAsNewProspectFilterViewRecyclePool();" id="saveFilterReportViewNameButton">Save Filter View</button>
			      </div>
		      </form>
		    </div>
		  </div>
		</div>
	<!-- modal ends here !-->
	<!-- new prospect group modal ends here !-->



	

	<!-- new prospect group modal starts here !-->
	<!-- modal starts here !-->
		<div class="modal fade" id="editFilterViewNameRecyclePool" tabindex="-1" role="dialog" aria-labelledby="editFilterViewNameRecyclePool">
		  <div class="modal-dialog" role="document">
		  
		 			  
				<script language="JavaScript">
				<!--
				
				   function validateEditViewNameRecPool()
				    {
								    				       
					   var viewName = $("#txtUpdatedFilterReportViewNameRecPool").val();
					   		    
				       if (viewName == "") {
				            swal("Please enter a name to save this view.");
				            return false;
				       }				       
				       	
				       if (viewName == "Default" || viewName == "DEFAULT" || viewName == "default" || viewName == "Current" || viewName == "current" || viewName == "CURRENT"|| viewName == "All Prospects" || viewName == "all prospects" || viewName == "ALL PROSPECTS") {
				            swal("A view cannot be named DEFAULT, CURRENT or ALL PROSPECTS, they are reserved names.");
				            return false;
				       }
				
				       return true;
				
				    }
				// -->
				</script>  

		    <div class="modal-content">
		      <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		        <h4 class="modal-title" id="myProspectingModalLabel">Save This View</h4>
		      </div>
		      <form name="frmEditFilterViewNameRecyclePool" id="frmEditFilterViewNameRecyclePool">
			      <div class="modal-body">
	
					<div class="col-lg-12">
						<div class="form-group">
							<label class="col-sm-12 control-label" style="margin-left:-15px;">Save This View As:</label>
							<% CurrentViewName = MUV_READ("CRMVIEWSTATERECPOOL") %>
							
							<input type="text" class="form-control required" id="txtUpdatedFilterReportViewNameRecPool" name="txtUpdatedFilterReportViewNameRecPool">
							
							<input type="hidden" name="originalViewNameRecPool" id="originalViewNameRecPool" value="<%= CurrentViewName %>">
							
						</div>
					</div>
	
					<div class="clearfix"></div>
						  
			       </div>
			      <div class="modal-footer">
			        <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
			        <button type="button" class="btn btn-primary" onclick="if (validateEditViewNameRecPool()) renameProspectFilterViewRecyclePool();" id="updateFilterReportViewNameButton">Update Filter Name</button>
			      </div>
		      </form>
		    </div>
		  </div>
		</div>
	<!-- modal ends here !-->
	<!-- new prospect group modal ends here !-->


				  
	<!-- delete prospect view modal starts here !-->	
		<div class="modal fade" id="deleteProspectViewRecyclePool" tabindex="-1" role="dialog" aria-labelledby="myProspectingRecycleLabel">
		  <div class="modal-dialog" role="document">
		    <div class="modal-content">
		      <div class="modal-header">
					<button type="button" class="close" data-dismiss="modal" aria-hidden="true"><span class="glyphicon glyphicon-remove" aria-hidden="true"></span></button>
					<h4 class="modal-title custom_align" id="Heading">Delete this View</h4>
				</div>
				
				<form name="frmDeleteProspectViewRecyclePool" id="frmDeleteProspectViewRecyclePool" method="post" action="deleteRecyclePoolProspectFilterViewFromModal.asp">

					<div class="modal-body">
						<h3><%= CurrentViewName %></h3>
						<div class="alert alert-danger"><span class="glyphicon glyphicon-warning-sign"></span> Are you sure you want to delete this view?</div>
						<input type="hidden" name="viewNameToDelete" id="viewNameToDelete" value="<%= CurrentViewName %>">
					</div>
					
					<div class="modal-footer">
						<button type="button" class="btn btn-default" data-dismiss="modal"><span class="glyphicon glyphicon-remove"></span> No</button>
						<button type="button" class="btn btn-success" onclick="frmDeleteProspectViewRecyclePool.submit()"><span class="glyphicon glyphicon-ok-sign"></span> Yes, Delete This View</button>
					</div>
				
				</form>
			</div>
			<!-- /.modal-content --> 
		</div>
		<!-- /.modal-dialog --> 
	</div>
 	<!-- delete prospect view modal ends here !-->   


	<!-- recycle prospect modal starts here !-->
	<!-- modal starts here !-->
		<div class="modal fade" id="myProspectingRecycleModal" tabindex="-1" role="dialog" aria-labelledby="myProspectingRecycleLabel">
		  <div class="modal-dialog" role="document">
		    <div class="modal-content">
		      <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
		        <h4 class="modal-title" id="myProspectingRecycleModalLabel">Recycle This Prospect</h4>
		      </div>
		      <form name="frmProspectRecycle" id="frmProspectRecycle">
			      <div class="modal-body">
	
					<div class="col-lg-12">
						<div class="form-group">
							<label class="col-sm-12 control-label">Some Fields Here</label>
							<input type="text" class="form-control required" id="txtProspectRecycle" name="txtProspectRecycle">
						</div>
					</div>
	
					<div class="clearfix"></div>
						  
			       </div>
			      <div class="modal-footer">
			        <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
			        <button type="button" class="btn btn-primary" data-dismiss="modal">Recycle Prospect</button>
			      </div>
		      </form>
		    </div>
		  </div>
		</div>
	<!-- modal ends here !-->
	<!-- recycle prospect modal ends here !-->
	
	


<!-- ******************************************************************************************************************************** -->
<!-- END MODAL WINDOW DESIGN AND DEFINITIONS -->
<!-- ******************************************************************************************************************************** -->
	

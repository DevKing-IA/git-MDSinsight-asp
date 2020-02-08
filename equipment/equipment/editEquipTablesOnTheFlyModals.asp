<!-- **************************************************************************************************************************** -->
<!-- MODAL FOR ADD/EDIT EQUIPMENT OPTIONS ON THE FLY BEGINS HERE !-->
<!-- **************************************************************************************************************************** -->



<div class="modal fade" id="equipAddNewStatusCodeModal" tabindex="-1" role="dialog" aria-labelledby="equipAddNewStatusCodeLabel">
	
	<style type="text/css">
	
		.select-line{
			margin-bottom: 15px;
		}
		
		.form-control{
			min-width: 100px;
		}
		
		.textarea-box{
			min-width: 260px;
		}
		
		.required{
			border-left:3px solid red;
		}
		
		.row-line{
			margin-bottom: 25px;
		}
		
		.control-label-modal-2 {
		    font-size: 12px;
		    font-weight: normal;
		    padding-top: 10px; 
		}		
		
		.control-label-modal {
		    font-size: 12px;
		    font-weight: normal;
		    padding-top: 1-0px; 
		}	
	</style>
	
	<script language="JavaScript">
	<!--
	    function validateStatusCodeForm()
	    {

	        if (document.frmAddNewStatusCode.txtBackendSystemCode.value == "") {
	            swal("Status/Backend Code cannot be blank.");
	            return false;
	        }

	
	        if (document.frmAddNewStatusCode.txtStatusCodeDesc.value == "") {
	            swal("Status Code name/desc cannot be blank.");
	            return false;
	        }
	
	        return true;
	
	    }
	// -->
	</script>   

	<div class="modal-dialog" role="document">
						
		<div class="modal-content">
		
			<form method="post" action="#" name="frmAddNewStatusCode" id="frmAddNewStatusCode">
		    
				<!-- modal header !-->
				<div class="modal-header">		
					<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
					<h4 class="modal-title" id="equipAddNewStatusCodeModalTitle">Add New Status Code</h4>
				</div>
				<!-- eof modal header !-->
		  
				<!-- modal body !-->
				<div class="modal-body">
				
					<input type="hidden" name="txtEquipIntRecID" id="txtEquipIntRecID" value="">
					
					<div id="equipAddNewStatusCodeModalContent">
						
						<div class="row row-line">
							<div class="form-group col-lg-12">
								<label for="txtBackendSystemCode" class="col-sm-4 control-label-modal-2">Status Code/Backend System Code</label>	
				    			<div class="col-sm-7">
				    				<input type="text" class="form-control required" id="txtBackendSystemCode" name="txtBackendSystemCode">
				    			</div>
							</div>
						</div>	
						
					
						<div class="row row-line">
							<div class="form-group col-lg-12">
								<label for="txtStatusCode" class="col-sm-4 control-label-modal-2">Status Code Desc</label>	
				    			<div class="col-sm-7">
				    				<input type="text" class="form-control required" id="txtStatusCodeDesc" name="txtStatusCodeDesc">
				    			</div>
							</div>
						</div>
						
						
						<div class="row row-line">
							<div class="form-group col-lg-6">
								<label for="chkAvailableForPlacement" class="col-sm-9 control-label-modal">Available for Placement?</label>	
				    			<div class="col-sm-1">
				    				<input type="checkbox" id="chkAvailableForPlacement" name="chkAvailableForPlacement">
				    			</div>
							</div>			
						</div>
						
						<div class="row row-line">
							<div class="form-group col-lg-6">
								<label for="chkAvailableForPlacement" class="col-sm-9 control-label-modal">Generates Rental Revenue?</label>	
				    			<div class="col-sm-1">
				    				<input type="checkbox" id="chkGeneratesRentalRevenue" name="chkGeneratesRentalRevenue">
				    			</div>
							</div>
						</div>
						
						
					</div>
				</div>
			
				<!-- modal footer !-->
			    <div class="modal-footer">
					      			      	      
					<!-- close / save !-->
					<div class="col-lg-12">
						<button type="button" class="btn btn-default" data-dismiss="modal">Close</button>
						<button type="submit" class="btn btn-primary" id="btnEquipAddNewStatusCodeSubmit" data-dismiss="modal">Add Status Code</button>
					</div>
					<!-- eof close / save !-->
			
				</div>
				<!-- eof modal footer !-->
				
			</form>

		</div>
		<!-- eof modal content !-->
</div>
<!-- eof modal dialog !-->
</div>
<!-- eof modal !-->




<div class="modal fade" id="equipAddNewConditionCodeModal" tabindex="-1" role="dialog" aria-labelledby="equipAddNewConditionCodeLabel">
	
	<style type="text/css">
	
		.select-line{
			margin-bottom: 15px;
		}
		
		.form-control{
			min-width: 100px;
		}
		
		.textarea-box{
			min-width: 260px;
		}
		
		.required{
			border-left:3px solid red;
		}
		
		.row-line{
			margin-bottom: 25px;
		}
		
		.control-label-modal-2 {
		    font-size: 12px;
		    font-weight: normal;
		    padding-top: 10px; 
		}		
		
		.control-label-modal {
		    font-size: 12px;
		    font-weight: normal;
		    padding-top: 1-0px; 
		}	
	</style>
	
	<script language="JavaScript">
	<!--
	    function validateConditionCodeForm()
	    {
	
	        if (document.frmAddNewConditionCode.txtCondition.value == "") {
	            swal("Condition cannot be blank.");
	            return false;
	        }
	
	        return true;
	
	    }
	// -->
	</script>   

	<div class="modal-dialog" role="document">
						
		<div class="modal-content">
		
			<form method="post" action="#" name="frmAddNewConditionCode" id="frmAddNewConditionCode">
		    
				<!-- modal header !-->
				<div class="modal-header">		
					<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
					<h4 class="modal-title" id="equipAddNewConditionCodeModalTitle">Add New Condition Code</h4>
				</div>
				<!-- eof modal header !-->
		  
				<!-- modal body !-->
				<div class="modal-body">
				
					<input type="hidden" name="txtEquipIntRecID" id="txtEquipIntRecID" value="">
					
					<div id="equipAddNewConditionCodeModalContent">
					

						<div class="row row-line">
							<div class="form-group col-lg-12">
								<label for="txtCondition" class="col-sm-3 control-label">Condition</label>	
				    			<div class="col-sm-6">
				    				<input type="text" class="form-control required" id="txtCondition" name="txtCondition" >
				    			</div>
							</div>
						</div>
				
				
						<div class="row row-line">
							<div class="form-group col-lg-12">
								<label for="txtCondition" class="col-sm-3 control-label">Description (up to 8000 chars)</label>	
				    			<div class="col-sm-6">
				    				<textarea class="form-control" id="txtConditionDescription" name="txtConditionDescription"></textarea>
				    			</div>
							</div>
						</div>
						
					</div>
				</div>
			
				<!-- modal footer !-->
			    <div class="modal-footer">
					      			      	      
					<!-- close / save !-->
					<div class="col-lg-12">
						<button type="button" class="btn btn-default" data-dismiss="modal">Close</button>
						<button type="submit" class="btn btn-primary" id="btnEquipAddNewConditionCodeSubmit" data-dismiss="modal">Add Condition Code</button>
					</div>
					<!-- eof close / save !-->
			
				</div>
				<!-- eof modal footer !-->
				
			</form>

		</div>
		<!-- eof modal content !-->
</div>
<!-- eof modal dialog !-->
</div>
<!-- eof modal !-->








<div class="modal fade" id="equipAddNewMovementCodeModal" tabindex="-1" role="dialog" aria-labelledby="equipAddNewMovementCodeLabel">
	
	<style type="text/css">
	
		.select-line{
			margin-bottom: 15px;
		}
		
		.form-control{
			min-width: 100px;
		}
		
		.textarea-box{
			min-width: 260px;
		}
		
		.required{
			border-left:3px solid red;
		}
		
		.row-line{
			margin-bottom: 25px;
		}
		
		.control-label-modal-2 {
		    font-size: 12px;
		    font-weight: normal;
		    padding-top: 10px; 
		}		
		
		.control-label-modal {
		    font-size: 12px;
		    font-weight: normal;
		    padding-top: 1-0px; 
		}	
	</style>
	
	<script language="JavaScript">
	<!--
	    function validateMovementCodeForm()
	    {
	
	        if (document.frmAddNewMovementCode.txtMovementCode.value == "") {
	            swal("Movement Code cannot be blank.");
	            return false;
	        }
	
	        if (document.frmAddNewMovementCode.txtMovementCodeDesc.value == "") {
	            swal("Movement Code description cannot be blank.");
	            return false;
	        }
	
	        return true;
	
	    }
	// -->
	</script>   

	<div class="modal-dialog" role="document">
						
		<div class="modal-content">
		
			<form method="post" action="#" name="frmAddNewMovementCode" id="frmAddNewMovementCode">
		    
				<!-- modal header !-->
				<div class="modal-header">		
					<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
					<h4 class="modal-title" id="equipAddNewMovementCodeModalTitle">Add New Movement Code</h4>
				</div>
				<!-- eof modal header !-->
		  
				<!-- modal body !-->
				<div class="modal-body">
				
					<input type="hidden" name="txtEquipIntRecID" id="txtEquipIntRecID" value="">
					
					<div id="equipAddNewMovementCodeModalContent">
									
						<div class="row row-line">
							<div class="form-group col-lg-12">
								<label for="txtBackendSystemCode" class="col-sm-3 control-label">Movement Code</label>	
				    			<div class="col-sm-6">
				    				<input type="text" class="form-control required" id="txtMovementCode" name="txtMovementCode">
				    			</div>
							</div>
						</div>	
				
						<div class="row row-line">
							<div class="form-group col-lg-12">
								<label for="txtMovementCode" class="col-sm-3 control-label">Movement Code Desc</label>	
				    			<div class="col-sm-6">
				    				<input type="text" class="form-control required" id="txtMovementCodeDesc" name="txtMovementCodeDesc">
				    			</div>
							</div>
						</div>
						
					</div>
				</div>
			
				<!-- modal footer !-->
			    <div class="modal-footer">
					      			      	      
					<!-- close / save !-->
					<div class="col-lg-12">
						<button type="button" class="btn btn-default" data-dismiss="modal">Close</button>
						<button type="submit" class="btn btn-primary" id="btnEquipAddNewMovementCodeSubmit" data-dismiss="modal">Add Movement Code</button>
					</div>
					<!-- eof close / save !-->
			
				</div>
				<!-- eof modal footer !-->
				
			</form>

		</div>
		<!-- eof modal content !-->
</div>
<!-- eof modal dialog !-->
</div>
<!-- eof modal !-->









<div class="modal fade" id="equipAddNewAcquisitionCodeModal" tabindex="-1" role="dialog" aria-labelledby="equipAddNewAcquisitionCodeLabel">
	
	<style type="text/css">
	
		.select-line{
			margin-bottom: 15px;
		}
		
		.form-control{
			min-width: 100px;
		}
		
		.textarea-box{
			min-width: 260px;
		}
		
		.required{
			border-left:3px solid red;
		}
		
		.row-line{
			margin-bottom: 25px;
		}
		
		.control-label-modal-2 {
		    font-size: 12px;
		    font-weight: normal;
		    padding-top: 10px; 
		}		
		
		.control-label-modal {
		    font-size: 12px;
		    font-weight: normal;
		    padding-top: 1-0px; 
		}	
	</style>
	
	<script language="JavaScript">
	<!--
	    function validateAcquisitionCodeForm()
	    {
	
	        if (document.frmAddNewAcquisitionCode.txtAcquisitionCode.value == "") {
	            swal("Acquisition Code cannot be blank.");
	            return false;
	        }
	
	        if (document.frmAddNewAcquisitionCode.txtAcquisitionCodeDesc.value == "") {
	            swal("Acquisition Code description cannot be blank.");
	            return false;
	        }
	
	        return true;
	
	    }
	// -->
	</script>   

	<div class="modal-dialog" role="document">
						
		<div class="modal-content">
		
			<form method="post" action="#" name="frmAddNewAcquisitionCode" id="frmAddNewAcquisitionCode">
		    
				<!-- modal header !-->
				<div class="modal-header">		
					<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
					<h4 class="modal-title" id="equipAddNewAcquisitionCodeModalTitle">Add New Acquisition Code</h4>
				</div>
				<!-- eof modal header !-->
		  
				<!-- modal body !-->
				<div class="modal-body">
				
					<input type="hidden" name="txtEquipIntRecID" id="txtEquipIntRecID" value="">
					
					<div id="equipAddNewAcquisitionCodeModalContent">
									
						<div class="row row-line">
							<div class="form-group col-lg-12">
								<label for="txtBackendSystemCode" class="col-sm-3 control-label">Acquisition Code</label>	
				    			<div class="col-sm-6">
				    				<input type="text" class="form-control required" id="txtAcquisitionCode" name="txtAcquisitionCode">
				    			</div>
							</div>
						</div>	
				
						<div class="row row-line">
							<div class="form-group col-lg-12">
								<label for="txtAcquisitionCode" class="col-sm-3 control-label">Acquisition Code Desc</label>	
				    			<div class="col-sm-6">
				    				<input type="text" class="form-control required" id="txtAcquisitionCodeDesc" name="txtAcquisitionCodeDesc">
				    			</div>
							</div>
						</div>
						
					</div>
				</div>
			
				<!-- modal footer !-->
			    <div class="modal-footer">
					      			      	      
					<!-- close / save !-->
					<div class="col-lg-12">
						<button type="button" class="btn btn-default" data-dismiss="modal">Close</button>
						<button type="submit" class="btn btn-primary" id="btnEquipAddNewAcquisitionCodeSubmit" data-dismiss="modal">Add Acquisition Code</button>
					</div>
					<!-- eof close / save !-->
			
				</div>
				<!-- eof modal footer !-->
				
			</form>

		</div>
		<!-- eof modal content !-->
</div>
<!-- eof modal dialog !-->
</div>
<!-- eof modal !-->




<!-- **************************************************************************************************************************** -->
<!-- MODAL FOR ADD/EDIT EQUIPMENT OPTIONS ON THE FLY ENDS HERE !-->
<!-- **************************************************************************************************************************** -->

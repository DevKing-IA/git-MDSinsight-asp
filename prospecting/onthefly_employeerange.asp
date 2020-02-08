<SCRIPT LANGUAGE="JavaScript">
<!--
	
	function isInt(value) {
	  return !isNaN(value) && (function(x) { return (x | 0) === x; })(parseFloat(value))
	}
	
    function validateEmployeeRangeForm()
    {

        if (document.frmAddEmployeeRange.txtEmployeeRange1.value == "") {
            swal("Beginning Employee Range can not be blank.");
            return false;
        }
        
        if (!isInt(document.frmAddEmployeeRange.txtEmployeeRange1.value)) {
             swal("Beginning Employee Range must be a whole number.");
            return false;
        }

         if (document.frmAddEmployeeRange.txtEmployeeRange2.value == "") {
            swal("Ending Employee Range can not be blank.");
            return false;
        }
        
        if (!isInt(document.frmAddEmployeeRange.txtEmployeeRange2.value)) {
             swal("Ending Employee Range must be a whole number.");
            return false;
        }
        
		if(parseInt(document.frmAddEmployeeRange.txtEmployeeRange1.value) > parseInt(document.frmAddEmployeeRange.txtEmployeeRange2.value))
		{
            swal("Ending employee range must be greater than beginning employee range.");
            return false;
		}
		 
		if(parseInt(document.frmAddEmployeeRange.txtEmployeeRange1.value)==parseInt(document.frmAddEmployeeRange.txtEmployeeRange2.value))
		{
	        swal("Beginning and ending employee ranges cannot be equal");
	        return false;
		
		}       
        
       return true;

    }
// -->
</SCRIPT>   


<style>

.form-control-modal {
    display: block;
    width: 100%;
    height: 34px;
    padding: 6px 12px;
    font-size: 14px;
    line-height: 1.42857143;
    color: #555;
    background-color: #fff;
    background-image: none;
    border: 1px solid #ccc;
    border-radius: 4px;
    -webkit-box-shadow: inset 0 1px 1px rgba(0,0,0,.075);
    box-shadow: inset 0 1px 1px rgba(0,0,0,.075);
    -webkit-transition: border-color ease-in-out .15s,-webkit-box-shadow ease-in-out .15s;
    -o-transition: border-color ease-in-out .15s,box-shadow ease-in-out .15s;
    transition: border-color ease-in-out .15s,box-shadow ease-in-out .15s;
}
</style> 
 
<!-- Employee Range modal starts here !-->
<div class="modal fade" id="ONTHEFLYmodalEmployeeRange" tabindex="-1" role="dialog" aria-labelledby="ONTHEFLYmodalEmployeeRangeLabel">
	<div class="modal-dialog" role="document">
		<div class="modal-content">
			<div class="modal-header">
				<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
				<h4 class="modal-title" id="ONTHEFLYmodalEmployeeRangeTitle">Add New Employee Range</h4>
			</div>
			<div class="modal-body">
			
				<form method="POST" action="#" name="frmAddEmployeeRange" id="frmAddEmployeeRange">
					<input type="hidden" id="txtpid" name="txtpid" value="<%=InternalRecordIdentifier %>">
					<div class="row row-line">
						<div class="form-group col-lg-12">
							<label for="txtEmployeeRange1" class="col-md-4 control-label">Beginning Employee Range</label>	
			    			<div class="col-sm-5">
			    				<input type="text" class="form-control-modal required" id="txtEmployeeRange1" name="txtEmployeeRange1">
			    			</div>
						</div>
						
						<div class="form-group col-lg-12">	
							<label for="txtEmployeeRange2" class="col-md-4 control-label">Ending Employee Range</label>	
			    			<div class="col-sm-5">
			    				<input type="text" class="form-control-modal required" id="txtEmployeeRange2" name="txtEmployeeRange2">
			    			</div>
						</div>
					</div>
					
					<div class="modal-footer">
				        <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
				        <button type="submit" class="btn btn-primary">Save</button>
					</div>
				</form>
			</div>
		</div>
	</div>
</div>
<!-- Employee Range modal ends here !-->

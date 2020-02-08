<SCRIPT LANGUAGE="JavaScript">
<!--
    function validateLeadSourceForm()
    {

        if (document.frmAddLeadSource.txtLeadSource.value == "") {
            swal("Lead source can not be blank.");
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
 
<!-- Lead Source modal starts here !-->
<div class="modal fade" id="ONTHEFLYmodalLeadSource" tabindex="-1" role="dialog" aria-labelledby="ONTHEFLYmodalLeadSourcelabel">
	<div class="modal-dialog" role="document">
		<div class="modal-content">
			<div class="modal-header">
				<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
				<h4 class="modal-title" id="ONTHEFLYmodalLeadSourceTitle">Add New Lead Source</h4>
			</div>
			<div class="modal-body">
			
				<form method="POST" action="#" name="frmAddLeadSource" id="frmAddLeadSource">
					<input type="hidden" id="txtpid" name="txtpid" value="<%=InternalRecordIdentifier %>">
					<div class="row row-line">
						<div class="form-group col-lg-12">
							<label for="txtLeadSource" class="col-sm-3 control-label">Lead Source</label>	
			    			<div class="col-sm-6">
			    				<input type="text" class="form-control-modal required" id="txtLeadSource" name="txtLeadSource">
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
<!-- Lead Source modal ends here !-->

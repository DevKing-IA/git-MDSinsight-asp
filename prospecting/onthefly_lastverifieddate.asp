

 


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
 
<!-- Verify Date modal starts here !-->
<div class="modal fade" id="myProspectingModalEditVerifyDate" tabindex="-1" role="dialog" aria-labelledby="myProspectingModalEditVerifyDatelabel">
	<div class="modal-dialog" role="document">
		<div class="modal-content">
			<div class="modal-header">
				<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
				<h4 class="modal-title" id="myProspectingModalEditVerifyDateTitle">Last Verified Date</h4>
			</div>
			<div class="modal-body">
			
					<form method="POST" action="#" name="frmUpdateVerifyDate" id="frmUpdateVerifyDate">

		<div class="row row-line">

									<div class="col-lg-12" style="margin-top:15px;">	
							<div class="form-group">

								<div class="col-lg-5" style="padding-left:0px;">
									<label class="control-label" style="padding-left:0px;">Date:</label>
								</div>
								<div class="col-lg-7">								  	
					                <div class="input-group date" id="datetimepickerVerifyDate">
					                    <input type="text" class="form-control" name="txtProspectEditVerifyDate" id="txtProspectEditVerifyDate">
					                    <span class="input-group-addon">
					                        <span class="glyphicon glyphicon-calendar"></span>
					                    </span>
					                </div>
					             </div>
							</div>
						</div>
			
           

            
		</div>
		
	    <!-- cancel / submit !-->
		<div class="modal-footer">
				        <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
				        <button type="submit" class="btn btn-primary">Save</button>
					</div>
                    
                    <input type="hidden" name="dateInternalRecordIdentifier" value="<%=InternalRecordIdentifier%>">
		
	</form>
			</div>
		</div>
	</div>
</div>
<!-- Verify Date modal ends here !-->

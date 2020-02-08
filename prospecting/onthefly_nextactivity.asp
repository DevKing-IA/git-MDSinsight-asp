

<link rel="stylesheet" type="text/css" href="<%= BaseURL %>css/tabs.css" />

<SCRIPT LANGUAGE="JavaScript">
<!--
    function validateActivityForm()
    {

        if (document.frmAddActivity.txtActivity.value == "") {
            swal("Activity can not be blank.");
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
 
<!-- Next activity modal starts here !-->
<div class="modal fade" id="ONTHEFLYmodalNextActivity" tabindex="-1" role="dialog" aria-labelledby="ONTHEFLYmodalNextActivitylabel">
	<div class="modal-dialog" role="document">
		<div class="modal-content">
			<div class="modal-header">
				<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
				<h4 class="modal-title" id="ONTHEFLYmodalNextActivityTitle">Add New Next Activity</h4>
			</div>
			<div class="modal-body">
			
					<form method="POST" action="#" name="frmAddNextActivity" id="frmAddNextActivity">

		<div class="row row-line">

			<div class="form-group col-lg-12">
				<label for="txtActivity" class="control-label"><strong>Activity</strong></label>
    			
    				<input type="text" class="form-control required" id="txtActivity" name="txtActivity" >
    			
			</div>
			
            <div class="form-group col-lg-12">
            	<input type="radio" name="optApptOrMeet" id="optApptOrMeet" value="Appointment" >
				<label for="optApptOrMeet" class="control-label"><strong>Create Appointment</strong></label>
			</div>
            
             <div class="form-group col-lg-12">
             	<input type="radio" name="optApptOrMeet" id="optApptOrMeet" value="Meeting">
				<label for="optApptOrMeet" class="control-label"><strong>Create Meeting</strong></label>
			</div>
            
            <div class="form-group col-lg-12">
				<input type="radio" name="optApptOrMeet" id="optApptOrMeet" checked value="Neither">
                <label for="Neither" class="control-label"><strong>Neither</strong></label>
			</div>
            

            
		</div>
		
	    <!-- cancel / submit !-->
		<div class="modal-footer">
				        <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
				        <button type="submit" class="btn btn-primary">Save</button>
					</div>
		
	</form>
			</div>
		</div>
	</div>
</div>
<!-- Next Activity modal ends here !-->

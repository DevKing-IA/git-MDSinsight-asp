<!--#include file="../../inc/header.asp"-->

<SCRIPT LANGUAGE="JavaScript">
<!--

	$(window).load(function()
	{
	   var phones = [{ "mask": "(###) ###-####"}];
	    $('#txtPhoneNumber').inputmask({ 
	        mask: phones, 
	        greedy: false, 
	        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
	    $('#txtFaxNumber').inputmask({ 
	        mask: phones, 
	        greedy: false, 
	        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
	        
	});
	
	function isValidPhone(p) {
	  var phoneRe = /^(1\s|1|)?((\(\d{3}\))|\d{3})(\-|\s)?(\d{3})(\-|\s)?(\d{4})$/;
	  var digits = p.replace(/\D/g, "");
	  return phoneRe.test(digits);
	}
	
	function isValidEmail(email) 
	{
	    var re = /\S+@\S+\.\S+/;
	    return re.test(email);
	}	
	

   function validateVendorForm()
    {
    
       if (document.frmAddVendor.txtVendorCompanyName.value == "") {
            swal("Vendor company name cannot be blank.");
            return false;
       }

       if ((document.frmAddVendor.txtPrimaryContactEmailAddress.value !== "") && (isValidEmail(document.frmAddVendor.txtPrimaryContactEmailAddress.value) == false)) {
           swal("The primary contact email address is invalid. Please enter any format like the following: anystring@anystring.anystring.");
           return false;
       }
       if ((document.frmAddVendor.txtTechnicalContactEmailAddress.value !== "") && (isValidEmail(document.frmAddVendor.txtTechnicalContactEmailAddress.value) == false)) {
           swal("The technical contact email address is invalid. Please enter any format like the following: anystring@anystring.anystring.");
           return false;
       }
       if ((document.frmAddVendor.txtPhoneNumber.value !== "") && (isValidPhone(document.frmAddVendor.txtPhoneNumber.value) == false)) {
           swal("The phone number is invalid. Please enter the number in the following format: (555) 555-5555.");
           return false;
       }     
       if ((document.frmAddVendor.txtFaxNumber.value !== "") && (isValid(document.frmAddVendor.txtFaxNumber.value) == false)) {
           swal("The fax number is invalid. Please enter the number in the following format: (555) 555-5555.");
           return false;
       }
       return true;
    }
// -->
</SCRIPT>   


<!-- password strength meter !-->

<style type="text/css">

.input-group {
	margin-top:10px;
}
.select-line{
	margin-bottom: 15px;
}

.enable-disable{
	margin-top:20px;
}

.row-line{
	margin-bottom: 25px;
}

.table th, tr, td{
	font-weight: normal;
}

.table>thead>tr>th{
	border: 0px;
}
.table thead>tr>th,.table tbody>tr>th,.table tfoot>tr>th,.table thead>tr>td,.table tbody>tr>td,.table tfoot>tr>td{
border:0px;
}

.when-col{
	width: 10%;
}

.reference-col{
	width: 45%;
}

.has-more-col{
	width: 12%;
}

.form-control{
	min-width: 100px;
}

.textarea-box{
	min-width: 260px;
}

.custom-container{
	max-width:800px;
	margin:0 auto;
}

.control-label{
	font-size:12px;
	font-weight:normal;
	padding-top:10px;
}
.control-label-last{
	padding-top:0px;
}

.required{
	border-left:3px solid red;
}
	</style>
<!-- eof password strength meter !-->

<h1 class="page-header"> Add New Vendor</h1>

<div class="custom-container">

	<form method="POST" action="addVendor_submit.asp" name="frmAddVendor" id="frmAddVendor" onsubmit="return validateVendorForm();">

		<div class="row row-line">			
			
              <div class="form-group">         	                
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-briefcase"></i></div>
	                    	<input type="text" class="form-control required" id="txtVendorCompanyName" placeholder="Company Name" name="txtVendorCompanyName">
	                   </div>
	                </div> 
               </div>
               
	          <div class="form-group">
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-globe"></i></div>
	                    	<input type="text" class="form-control required" id="txtVendorAPIKey" placeholder="Vendor API Key" name="txtVendorAPIKey">
	                   </div>
	                </div> 
               </div>


              <div class="form-group">   
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-user"></i></div>
	                    	<input type="text" class="form-control" id="txtPrimaryContactName" placeholder="Primary Contact Name" name="txtPrimaryContactName">
	                   </div>
	                </div> 	
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-envelope"></i></div>
	                    	<input type="text" class="form-control" id="txtPrimaryContactEmailAddress" placeholder="Primary Contact Email" name="txtPrimaryContactEmailAddress">
	                   </div>
	                </div>  
               </div>               
               
   
              <div class="form-group">   
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-user"></i></div>
	                    	<input type="text" class="form-control" id="txtTechnicalContactName" placeholder="Technical Contact Name" name="txtTechnicalContactName">
	                   </div>
	                </div> 	
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-envelope"></i></div>
	                    	<input type="text" class="form-control" id="txtTechnicalContactEmailAddress" placeholder="Technical Contact Email" name="txtTechnicalContactEmailAddress">
	                   </div>
	                </div> 
	                 
               </div>

              <div class="form-group">   
              
                    <div class="col-sm-4">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-address-card"></i></div>
	                    	<input type="text" class="form-control" id="txtAddressLine2" placeholder="Suite, Floor #, etc." name="txtAddressLine2">
	                   </div>
	                </div> 


	                <div class="col-sm-8">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-address-card"></i></div>
	                    	<input type="text" class="form-control" id="txtAddressLine1" placeholder="Street Address" name="txtAddressLine1">
	                   </div>
	                </div> 
	           </div>     
           

              <div class="form-group">          
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-address-book"></i></div>
	                    	<input type="text" class="form-control" id="txtCity" placeholder="City" name="txtCity">
	                   </div>
	                </div> 
	                <div class="col-sm-3">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-address-book"></i></div>
                    		<select data-placeholder="Choose State" class="C_Country_Modal form-control" id="txtState" name="txtState"> 
                    			<option value="">State</option>
								<!--#include file="statelist.asp"-->
							</select>				
	                   </div>
	                </div> 
	                <div class="col-sm-3">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-address-book"></i></div>
	                    	<input type="text" class="form-control" id="txtZipCode" placeholder="Zip" name="txtZipCode">
	                   </div>
	                </div> 
	                
	          </div>   
	            
              <div class="form-group">
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-phone"></i></div>
	                    	<input type="text" class="form-control" id="txtPhoneNumber" placeholder="Phone Number" name="txtPhoneNumber">
	                   </div>
	                </div> 
	 
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-fax"></i></div>
	                    	<input type="text" class="form-control" id="txtFaxNumber" placeholder="Fax Number" name="txtFaxNumber">
	                   </div>
	                </div> 
	 
               </div>

              <div class="form-group">
              
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-globe"></i></div>
	                    	<input type="text" class="form-control" id="txtWebsite" placeholder="Website" name="txtwebsite">
	                   </div>
	                </div> 
	 
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-envelope"></i></div>
	                    	<input type="text" class="form-control" id="txtAccountNumber" placeholder="Account Number" name="txtAccountNumber">
	                   </div>
	                </div> 
	 
               </div>
			
              <div class="form-group">
              
				<div class="col-lg-12">
					<div class="input-group">
						<div class="input-group-addon"><i class="fa fa-sticky-note"></i></div>	
	    				<textarea class="form-control" id="txtNotes" name="txtNotes" rows="4" placeholder="Notes"></textarea>

	    			</div>
				</div>
	 
               </div>



			
		</div>
		
	    <!-- cancel / submit !-->
		<div class="row row-line">
			<div class="col-lg-12 alertbutton">
				<div class="col-lg-12">
					<a href="<%= BaseURL %>filemaint/AP/Vendors/main.asp">
	    				<button type="button" class="btn btn-default">&lsaquo; Cancel &amp; Go Back To Vendors List</button>
					</a>
					<button type="submit" class="btn btn-primary"><i class="far fa-save"></i> Save</button>
				</div>
		    </div>
		</div>
		
	</form>
</div>

<!--#include file="../../inc/footer-main.asp"-->

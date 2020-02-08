<!--#include file="../../inc/header.asp"-->

<%
ClientKeyForFileNames = MUV_READ("ClientKeyForFileNames")
%>
<SCRIPT LANGUAGE="JavaScript">

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
	

   function validateManufacturerForm()
    {
    
       if (document.frmAddManufacturer.txtManufacturerName.value == "") {
            swal("Manufacturer name cannot be blank.");
            return false;
       }
       if (document.frmAddManufacturer.txtInsightAssetTagPrefix.value == "") {
            swal("The Insight asset tag prefix cannot be blank.");
            return false;
       }       
       if ((document.frmAddManufacturer.txtEmailAddress.value !== "") && (isValidEmail(document.frmAddManufacturer.txtEmailAddress.value) == false)) {
           swal("The manufacturer email address is invalid. Please enter any format like the following: anystring@anystring.anystring.");
           return false;
       }
       if ((document.frmAddManufacturer.txtPhoneNumber.value !== "") && (isValidPhone(document.frmAddManufacturer.txtPhoneNumber.value) == false)) {
           swal("The manufacturer phone number is invalid. Please enter the number in the following format: (555) 555-5555.");
           return false;
       }     
       if ((document.frmAddManufacturer.txtFaxNumber.value !== "") && (isValid(document.frmAddManufacturer.txtFaxNumber.value) == false)) {
           swal("The manufacturer fax number is invalid. Please enter the number in the following format: (555) 555-5555.");
           return false;
       }
       
       return true;
    }
</script>
       
 

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
		max-width:1000px;
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

<h1 class="page-header"> Add New <%= GetTerm("Equipment") %> Manufacturer</h1>

<div class="custom-container">

	<form method="POST" action="addManufacturer_Submit.asp" name="frmAddManufacturer" id="frmAddManufacturer" onsubmit="return validateManufacturerForm();">

		<div class="row row-line">			
		
		<h4>General Manufacturer Information</h4>
			
              <div class="form-group">         	                
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-briefcase"></i></div>
	                    	<input type="text" class="form-control required" id="txtManufacturerName" placeholder="Manufacturer Name" name="txtManufacturerName">
	                   </div>
	                </div> 
               </div>
 

              <div class="form-group">         	                
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-tag"></i></div>
	                    	<input type="text" class="form-control required" id="txtInsightAssetTagPrefix" placeholder="Insight Asset Tag Prefix" name="txtInsightAssetTagPrefix">
	                   </div>
	                </div> 
               </div>
                        

              <div class="form-group">   
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-address-card"></i></div>
	                    	<input type="text" class="form-control" id="txtAddressLine1" placeholder="Street Address" name="txtAddressLine1">
	                   </div>
	                </div> 
	           </div>     
           

	          <div class="form-group">
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-address-card-o"></i></div>
	                    	<input type="text" class="form-control" id="txtAddressLine2" placeholder="Suite, Floor #, etc." name="txtAddressLine2">
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
	                <div class="col-sm-8">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-envelope"></i></div>
	                    	<input type="text" class="form-control" id="txtEmailAddress" placeholder="Email Address" name="txtEmailAddress">
	                   </div>
	                </div> 
	                 
               </div>
               
		</div>
				
		<hr>
				
		
	    <!-- cancel / submit !-->
		<div class="row row-line">
			<div class="col-lg-12 alertbutton">
				<div class="col-lg-12">
					<a href="<%= BaseURL %>equipment/manufacturers/main.asp">
	    				<button type="button" class="btn btn-default">&lsaquo; Cancel &amp; Go Back To Manufacturers List</button>
					</a>
					<button type="submit" class="btn btn-primary"><i class="far fa-save"></i> Save</button>
				</div>
		    </div>
		</div>
		
	</form>
</div>

<!--#include file="../../inc/footer-main.asp"-->

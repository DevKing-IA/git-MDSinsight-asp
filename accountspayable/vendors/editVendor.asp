<!--#include file="../../inc/header.asp"-->

<% InternalRecordIdentifier = Request.QueryString("i") 
If InternalRecordIdentifier = "" Then Response.Redirect("main.asp")
%>

<link rel="stylesheet" type="text/css" href="<%= BaseURL %>css/tabs.css" />


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
	

   function validateEditVendorForm()
    {
    
       if (document.frmEditVendor.txtVendorCompanyName.value == "") {
            swal("Vendor company name cannot be blank.");
            return false;
       }
       if ((document.frmEditVendor.txtPrimaryContactEmailAddress.value !== "") && (isValidEmail(document.frmEditVendor.txtPrimaryContactEmailAddress.value) == false)) {
           swal("The primary contact email address is invalid. Please enter any format like the following: anystring@anystring.anystring.");
           return false;
       }
       if ((document.frmEditVendor.txtTechnicalContactEmailAddress.value !== "") && (isValidEmail(document.frmEditVendor.txtTechnicalContactEmailAddress.value) == false)) {
           swal("The technical contact email address is invalid. Please enter any format like the following: anystring@anystring.anystring.");
           return false;
       }
       if ((document.frmEditVendor.txtPhoneNumber.value !== "") && (isValidPhone(document.frmEditVendor.txtPhoneNumber.value) == false)) {
           swal("The phone number is invalid. Please enter the number in the following format: (555) 555-5555.");
           return false;
       }     
       if ((document.frmEditVendor.txtFaxNumber.value !== "") && (isValid(document.frmEditVendor.txtFaxNumber.value) == false)) {
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



<%
SQL = "SELECT * FROM AP_Vendor where InternalRecordIdentifier = " & InternalRecordIdentifier 

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If NOT rs.EOF Then
	VendorAPIKey = rs.Fields("VendorAPIKey")
	VendorCompanyName = rs.Fields("VendorCompanyName")
	VendorPrimaryContactName = rs.Fields("VendorPrimaryContactName")
	VendorPrimaryContactEmail = rs.Fields("VendorPrimaryContactEmail")
	VendorTechnicalContactName = rs.Fields("VendorTechnicalContactName")
	VendorTechnicalContactEmail = rs.Fields("VendorTechnicalContactEmail")
	VendorAddress = rs.Fields("VendorAddress")
	VendorAddress2 = rs.Fields("VendorAddress2")
	VendorCity = rs.Fields("VendorCity")
	VendorState = rs.Fields("VendorState")
	VendorZip = rs.Fields("VendorZip")
	VendorPhone = rs.Fields("VendorPhone")
	VendorFax = rs.Fields("VendorFax")
	VendorWebsite = rs.Fields("Website")
	VendorAccountNumber = rs.Fields("AccountNumber")
	VendorNotes = rs.Fields("Notes")	
End If
set rs = Nothing
cnn8.close
set cnn8 = Nothing

%>


<h1 class="page-header"> Edit Vendor</h1>

<div class="custom-container">

	<form method="POST" action="editVendor_submit.asp" name="frmEditVendor " id="frmEditVendor " onsubmit="return validateEditVendorForm();">

			
	<div class="row row-line">		
	
			  <input type="hidden" id="txtInternalRecordIdentifier" name="txtInternalRecordIdentifier" value="<%= InternalRecordIdentifier %>">	
			
              <div class="form-group">         	                
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-briefcase"></i></div>
	                    	<input type="text" class="form-control required" id="txtVendorCompanyName" placeholder="Company Name" name="txtVendorCompanyName" value="<%= VendorCompanyName %>">
	                   </div>
	                </div> 
               </div>
               
	          <div class="form-group">
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-globe"></i></div>
	                    	<input type="text" class="form-control required" id="txtVendorAPIKey" placeholder="Vendor API Key" name="txtVendorAPIKey" value="<%= VendorAPIKey %>">
	                   </div>
	                </div> 
               </div>


              <div class="form-group">   
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-user"></i></div>
	                    	<input type="text" class="form-control" id="txtPrimaryContactName" placeholder="Primary Contact Name" name="txtPrimaryContactName" value="<%= VendorPrimaryContactName %>">
	                   </div>
	                </div> 	
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-envelope"></i></div>
	                    	<input type="text" class="form-control" id="txtPrimaryContactEmailAddress" placeholder="Primary Contact Email" name="txtPrimaryContactEmailAddress" value="<%= VendorPrimaryContactEmail %>">
	                   </div>
	                </div>  
               </div>               
               
   
              <div class="form-group">   
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-user"></i></div>
	                    	<input type="text" class="form-control" id="txtTechnicalContactName" placeholder="Technical Contact Name" name="txtTechnicalContactName" value="<%= VendorTechnicalContactName %>">
	                   </div>
	                </div> 	
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-envelope"></i></div>
	                    	<input type="text" class="form-control" id="txtTechnicalContactEmailAddress" placeholder="Technical Contact Email" name="txtTechnicalContactEmailAddress" value="<%= VendorTechnicalContactEmail %>">
	                   </div>
	                </div> 
	                 
               </div>

              <div class="form-group">   
              
                    <div class="col-sm-4">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-address-card"></i></div>
	                    	<input type="text" class="form-control" id="txtAddressLine2" placeholder="Suite, Floor #, etc." name="txtAddressLine2" value="<%= VendorAddress2 %>">
	                   </div>
	                </div> 


	                <div class="col-sm-8">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-address-card"></i></div>
	                    	<input type="text" class="form-control" id="txtAddressLine1" placeholder="Street Address" name="txtAddressLine1" value="<%= VendorAddress %>">
	                   </div>
	                </div> 
	           </div> 
              <div class="form-group">          
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-address-book"></i></div>
	                    	<input type="text" class="form-control" id="txtCity" placeholder="City" name="txtCity" value="<%= VendorCity %>">
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
	                    	<input type="text" class="form-control" id="txtZipCode" placeholder="Zip" name="txtZipCode" value="<%= VendorZip %>">
	                   </div>
	                </div> 
	                
	          </div>   
	            
              <div class="form-group">
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-phone"></i></div>
	                    	<input type="text" class="form-control" id="txtPhoneNumber" placeholder="Phone Number" name="txtPhoneNumber" value="<%= VendorPhone %>">
	                   </div>
	                </div> 
	 
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-fax"></i></div>
	                    	<input type="text" class="form-control" id="txtFaxNumber" placeholder="Fax Number" name="txtFaxNumber" value="<%= VendorFax %>">
	                   </div>
	                </div> 
	 
               </div>


              <div class="form-group">
              
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-globe"></i></div>
	                    	<input type="text" class="form-control" id="txtWebsite" placeholder="Website" name="txtwebsite" value="<%= VendorWebsite %>">
	                   </div>
	                </div> 
	 
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-envelope"></i></div>
	                    	<input type="text" class="form-control" id="txtAccountNumber" placeholder="Account Number" name="txtAccountNumber" value="<%= VendorAccountNumber %>">
	                   </div>
	                </div> 
	 
               </div>
			
              <div class="form-group">
              
				<div class="col-lg-12">
					<div class="input-group">
						<div class="input-group-addon"><i class="fa fa-sticky-note"></i></div>	
	    				<textarea class="form-control" id="txtNotes" name="txtNotes" rows="4" placeholder="Notes" value="<%= VendorNotes %>"></textarea>

	    			</div>
				</div>
	 
               </div>
			
			
		</div>	    <!-- cancel / submit !-->
		<div class="row row-line">
			<div class="col-lg-12 alertbutton">
				<div class="col-lg-12">
					<a href="<%= BaseURL %>filemaint/AP/Vendors/main.asp">
	    				<button type="button" class="btn btn-default">&lsaquo; Cancel &amp; Go Back To Vendors List</button>
					</a>
					<button type="submit" class="btn btn-primary"><i class="far fa-save"></i> Save Changes</button>
				</div>
		    </div>
		</div>
		
	</form>
</div>

<!--#include file="../../inc/footer-main.asp"-->

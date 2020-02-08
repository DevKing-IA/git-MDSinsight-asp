<!--#include file="../../inc/header.asp"-->

<% 

InternalRecordIdentifier = Request.QueryString("i") 

If InternalRecordIdentifier = "" Then Response.Redirect("main.asp")

ClientKeyForFileNames = MUV_READ("ClientKeyForFileNames")

%>

<link rel="stylesheet" type="text/css" href="<%= BaseURL %>css/tabs.css" />


<script type="text/javascript">


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
    
       if (document.frmEditManufacturer.txtManufacturerName.value == "") {
            swal("Manufacturer name cannot be blank.");
            return false;
       }
       if (document.frmEditManufacturer.txtInsightAssetTagPrefix.value == "") {
            swal("The Insight asset tag prefix cannot be blank.");
            return false;
       }           
       if ((document.frmEditManufacturer.txtEmailAddress.value !== "") && (isValidEmail(document.frmEditManufacturer.txtEmailAddress.value) == false)) {
           swal("The manufacturer email address is invalid. Please enter any format like the following: anystring@anystring.anystring.");
           return false;
       }
       if ((document.frmEditManufacturer.txtPhoneNumber.value !== "") && (isValidPhone(document.frmEditManufacturer.txtPhoneNumber.value) == false)) {
           swal("The manufacturer phone number is invalid. Please enter the number in the following format: (555) 555-5555.");
           return false;
       }     
       if ((document.frmEditManufacturer.txtFaxNumber.value !== "") && (isValid(document.frmEditManufacturer.txtFaxNumber.value) == false)) {
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


<%
SQL = "SELECT * FROM EQ_Manufacturers where InternalRecordIdentifier = " & InternalRecordIdentifier 

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If NOT rs.EOF Then
	manufacturerName = rs.Fields("ManufacturerName")
	InsightAssetTagPrefix = rs.Fields("InsightAssetTagPrefix")
	manufacturerAddress = rs.Fields("Address1")
	manufacturerAddress2 = rs.Fields("Address2")
	manufacturerCity = rs.Fields("City")
	manufacturertate = rs.Fields("State")
	manufacturerZip = rs.Fields("Zip")
	manufacturerPhone = rs.Fields("Phone")
	manufacturerFax = rs.Fields("Fax")
	manufacturerEmail = rs.Fields("Email")
End If
set rs = Nothing
cnn8.close
set cnn8 = Nothing

%>


<h1 class="page-header"> Edit <%= GetTerm("Equipment") %> Manufacturer</h1>

<div class="custom-container">

	<form method="POST" action="editManufacturer_submit.asp" name="frmEditManufacturer" id="frmEditManufacturer" onsubmit="return validateManufacturerForm();">
	
	<div class="row row-line">	
	
		<h4>General Manufacturer Information</h4>	
	
			  <input type="hidden" id="txtInternalRecordIdentifier" name="txtInternalRecordIdentifier" value="<%= InternalRecordIdentifier %>">	
			
              <div class="form-group">         	                
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-briefcase"></i></div>
	                    	<input type="text" class="form-control required" id="txtManufacturerName" placeholder="Manufacturer Name" name="txtManufacturerName" value="<%= manufacturerName %>">
	                   </div>
	                </div> 
               </div>
               
              <div class="form-group">         	                
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-tag"></i></div>
	                    	<input type="text" class="form-control required" id="txtInsightAssetTagPrefix" placeholder="Insight Asset Tag Prefix" name="txtInsightAssetTagPrefix"  value="<%= InsightAssetTagPrefix %>">
	                   </div>
	                </div> 
               </div>

              <div class="form-group">   
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-address-card"></i></div>
	                    	<input type="text" class="form-control" id="txtAddressLine1" placeholder="Street Address" name="txtAddressLine1" value="<%= manufacturerAddress %>">
	                   </div>
	                </div> 
	           </div>     
           

	          <div class="form-group">
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-address-card-o"></i></div>
	                    	<input type="text" class="form-control" id="txtAddressLine2" placeholder="Suite, Floor #, etc." name="txtAddressLine2" value="<%= manufacturerAddress2 %>">
	                   </div>
	                </div> 
               </div>

              <div class="form-group">          
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-address-book"></i></div>
	                    	<input type="text" class="form-control" id="txtCity" placeholder="City" name="txtCity" value="<%= manufacturerCity %>">
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
	                    	<input type="text" class="form-control" id="txtZipCode" placeholder="Zip" name="txtZipCode" value="<%= manufacturerZip %>">
	                   </div>
	                </div> 
	                
	          </div>   
	            
              <div class="form-group">
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-phone"></i></div>
	                    	<input type="text" class="form-control" id="txtPhoneNumber" placeholder="Phone Number" name="txtPhoneNumber" value="<%= manufacturerPhone %>">
	                   </div>
	                </div> 
	 
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-fax"></i></div>
	                    	<input type="text" class="form-control" id="txtFaxNumber" placeholder="Fax Number" name="txtFaxNumber" value="<%= manufacturerFax %>">
	                   </div>
	                </div> 
	 
               </div>

	            <div class="form-group">   
	                <div class="col-sm-8">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-envelope"></i></div>
	                    	<input type="text" class="form-control" id="txtEmailAddress" placeholder="Email Address" name="txtEmailAddress" value="<%= manufacturerEmail %>">
	                   </div>
	                </div> 
	                 
               </div>
		
			
		</div>	    <!-- cancel / submit !-->
		
		<hr>
			<div class="row row-line">
			<div class="col-lg-12 alertbutton">
				<div class="col-lg-12">
					<a href="<%= BaseURL %>equipment/manufacturers/main.asp"><button type="button" class="btn btn-default">&lsaquo; Cancel &amp; Go Back To Manufacturers List</button></a>
					<button type="submit" class="btn btn-primary"><i class="far fa-save"></i> Save Changes</button>

				</div>
		    </div>
		</div>
		
	</form>
</div>

<!--#include file="../../inc/footer-main.asp"-->

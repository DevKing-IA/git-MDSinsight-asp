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
	

   function validateEditPartnerForm()
    {
    
       if (document.frmEditPartner.txtPartnerCompanyName.value == "") {
            swal("Partner company name cannot be blank.");
            return false;
       }
       if (document.frmEditPartner.txtPartnerAPIKey.value == "") {
            swal("Partner API Key cannot be blank.");
            return false;
       }
       if ((document.frmEditPartner.txtPrimaryContactEmailAddress.value !== "") && (isValidEmail(document.frmEditPartner.txtPrimaryContactEmailAddress.value) == false)) {
           swal("The primary contact email address is invalid. Please enter any format like the following: anystring@anystring.anystring.");
           return false;
       }
       if ((document.frmEditPartner.txtTechnicalContactEmailAddress.value !== "") && (isValidEmail(document.frmEditPartner.txtTechnicalContactEmailAddress.value) == false)) {
           swal("The technical contact email address is invalid. Please enter any format like the following: anystring@anystring.anystring.");
           return false;
       }
       if ((document.frmEditPartner.txtPhoneNumber.value !== "") && (isValidPhone(document.frmEditPartner.txtPhoneNumber.value) == false)) {
           swal("The phone number is invalid. Please enter the number in the following format: (555) 555-5555.");
           return false;
       }     
       if ((document.frmEditPartner.txtFaxNumber.value !== "") && (isValid(document.frmEditPartner.txtFaxNumber.value) == false)) {
           swal("The fax number is invalid. Please enter the number in the following format: (555) 555-5555.");
           return false;
       }
		    
		if ($("input[name=optMappedOrPassedTaxable][value='DefinedCode']:checked").length > 0) {
			if (document.frmEditPartner.txtUnmappedTaxableSKUToPass.value == "") {
				swal("If you want to specify an unmapped taxable product, make sure to select a product code and unit of measure.");
				return false;
			}
		}
		
		if ($("input[name=optMappedOrPassedNonTaxable][value='DefinedCode']:checked").length > 0) {
			if (document.frmEditPartner.txtUnmappedNonTaxableSKUToPass.value == "") {
				swal("If you want to specify an unmapped non taxable product, make sure to select a product code and unit of measure.");
				return false;
			}
		}

		
		if ($("input[name=optMappedOrPassedCustAccount][value='DefinedAccount']:checked").length > 0) {
			if (document.frmEditPartner.txtUnmappedCustomerIDToPass.value == "") {
				swal("If you want to specify an unmapped customer number, make sure to select a customer account number to map to.");
				return false;
			}
		}

       return true;
    }



	$(function () {
	
		var taxableProductsRadioButtonvalue = $('input[optMappedOrPassedTaxable]:checked').val();
		var nonTaxableProductsRadioButtonvalue = $('input[optMappedOrPassedNonTaxable]:checked').val();
		
		if (taxableProductsRadioButtonvalue == 'PassedCode' || $("#txtUnmappedTaxableSKUToPass").val() !== "") {
		
			$("#taxableUM").show();
			$("input[name=optMappedOrPassedTaxable][value='DefinedCode']").prop("checked","checked");
			
            var UnmappedTaxableSKU = $("#txtUnmappedTaxableSKUToPass").val();
            var UnmappedTaxableUM = $("#txtUnmappedTaxableUMToPass").val();
            
            

       		$.ajax({
				type:"POST",
				url: "../../inc/InSightFuncs_AjaxForOrderAPI.asp",
				cache: false,
				data: "action=returnsUMSForUnmappedTaxableProductCode&prodSKU=" + encodeURIComponent(UnmappedTaxableSKU) + "&UM=" + encodeURIComponent(UnmappedTaxableUM),
				success: function(response)
				 {
				 	$("#taxableUM").html(response); 
				 	$("#taxableUM").show();
				 }
			});	//end ajax post
			
		}
		else {
			$("#taxableUM").hide();
		}
		
		

		if (nonTaxableProductsRadioButtonvalue == 'PassedCode' || $("#txtUnmappedNonTaxableSKUToPass").val() !== "") {
		
			$("#nonTaxableUM").show();
			$("input[name=optMappedOrPassedNonTaxable][value='DefinedCode']").prop("checked","checked");
			
            var UnmappedNonTaxableSKU = $("#txtUnmappedNonTaxableSKUToPass").val();
            var UnmappedNonTaxableUM = $("#txtUnmappedNonTaxableUMToPass").val();
        
       		$.ajax({
				type:"POST",
				url: "../../inc/InSightFuncs_AjaxForOrderAPI.asp",
				cache: false,
				data: "action=returnsUMSForUnmappedNonTaxableProductCode&prodSKU=" + encodeURIComponent(UnmappedNonTaxableSKU) + "&UM=" + encodeURIComponent(UnmappedNonTaxableUM),
				success: function(response)
				 {
				 	$("#nonTaxableUM").html(response); 
				 	$("#nonTaxableUM").show();
				 }
			});	//end ajax post
		}
		else {
			$("#nonTaxableUM").hide();
		}
	
		var autocompleteJSONFileURLProducts = "../../clientfiles/<%= ClientKeyForFileNames %>/autocomplete/product_list_<%= ClientKeyForFileNames %>.json";
		var autocompleteJSONFileURLAccount = "../../clientfiles/<%= ClientKeyForFileNames %>/autocomplete/customer_account_list_CSZ_<%= ClientKeyForFileNames %>.json";
		
		var optionsTaxableProducts = {
		  url: autocompleteJSONFileURLProducts,
		  placeholder: "Search for a product by SKU or description",
		  getValue: "name",
		  list: {	
	        onChooseEvent: function() {
	            var UnmappedTaxableSKU = $("#txtUnmappedTaxableSKU").getSelectedItemData().code;
	            $("input[name=optMappedOrPassedTaxable][value='DefinedCode']").prop("checked","checked");
	            $("#txtUnmappedTaxableSKUToPass").html("");
	       		$.ajax({
					type:"POST",
					url: "../../inc/InSightFuncs_AjaxForOrderAPI.asp",
					cache: false,
					data: "action=returnsUMSForUnmappedTaxableProductCode&prodSKU=" + encodeURIComponent(UnmappedTaxableSKU) + "&UM=",
					success: function(response)
					 {
					 	$("#taxableUM").html(response); 
					 	$("#taxableUM").show();
					 }
				});	//end ajax post
	            $("#txtUnmappedTaxableSKUToPass").val(UnmappedTaxableSKU);
        	},		  
		    match: {
		      enabled: true
			},
			maxNumberOfElements: 20		
		  },
		  theme: "round"
		};
		$("#txtUnmappedTaxableSKU").easyAutocomplete(optionsTaxableProducts);


		var optionsNonTaxableProducts = {
		  url: autocompleteJSONFileURLProducts,
		  placeholder: "Search for a product by SKU or description",
		  getValue: "name",
		  list: {	
	        onChooseEvent: function() {
	            var UnmappedNonTaxableSKU = $("#txtUnmappedNonTaxableSKU").getSelectedItemData().code;
	            $("input[name=optMappedOrPassedNonTaxable][value='DefinedCode']").prop("checked","checked");
	            $("#nonTaxableUM").html("");
	       		$.ajax({
					type:"POST",
					url: "../../inc/InSightFuncs_AjaxForOrderAPI.asp",
					cache: false,
					data: "action=returnsUMSForUnmappedNonTaxableProductCode&prodSKU=" + encodeURIComponent(UnmappedNonTaxableSKU) + "&UM=",
					success: function(response)
					 {
					 	$("#nonTaxableUM").html(response); 
					 	$("#nonTaxableUM").show();
					 }
				});	//end ajax post
	            $("#txtUnmappedNonTaxableSKUToPass").val(UnmappedNonTaxableSKU);
        	},		  
		    match: {
		      enabled: true
			},
			maxNumberOfElements: 20		
		  },
		  theme: "round"
		};
		$("#txtUnmappedNonTaxableSKU").easyAutocomplete(optionsNonTaxableProducts);
				

		

		var optionsAccount = {
		  url: autocompleteJSONFileURLAccount,
		  placeholder: "Search for a customer by id, name, city, state, zip",
		  getValue: "name",
		  list: {	
	        onChooseEvent: function() {
	            var custID = $("#txtUnmappedCustomerID").getSelectedItemData().code;
	            $("input[name=optMappedOrPassedCustAccount][value='DefinedAccount']").prop("checked","checked");
	            $("#txtUnmappedCustomerIDToPass").val(custID);
        	},		  
		    match: {
		      enabled: true
			},
			maxNumberOfElements: 20		
		  },
		  theme: "round"
		};
		$("#txtUnmappedCustomerID").easyAutocomplete(optionsAccount);

	})
</script>
       


<style type="text/css">
	#searchIcon {
	    position: relative;
	    z-index: 2;
	    color: #888;
	    top: -31px;
	    left: 410px;
	}
	#searchIcon2 {
	    position: relative;
	    z-index: 2;
	    color: #888;
	    top: -31px;
	    left: 410px;
	}
	
	.easy-autocomplete.eac-round {
	  font-family: "Open Sans", "Helvetica Neue",Helvetica,Arial,sans-serif;
	}
	.easy-autocomplete.eac-round input {
	  border: 1px solid #888;
	  border-radius: 4px;
	  color: #888;
	  font-family: inherit;
	  font-size: 16px;
	  font-weight: 400;
	  margin: 0;
	  min-width: 250px !important;
	  max-width: 400px !important;
	  padding: 10px;
	}
	.easy-autocomplete.eac-round input:hover, .easy-autocomplete.eac-round input:focus {
	  border-color: #3079ed;
	}
	.easy-autocomplete.eac-round ul {
	  background: none;
	  border-color: #888;
	  border-width: 0;
	  box-shadow: none;
	  min-width: 300px;
	  top: 20px;
	}
	.easy-autocomplete.eac-round ul li, .easy-autocomplete.eac-round ul .eac-category {
	  background: #fff;
	  border-color: #3079ed;
	  border-width: 0 2px;
	  font-size: 14px;
	  padding: 8px 12px;
	  transition: all 0.4s ease 0s;
	}
	.easy-autocomplete.eac-round ul li.selected, .easy-autocomplete.eac-round ul .eac-category.selected {
	  background-color: #d4e3fb;
	}
	.easy-autocomplete.eac-round ul li:first-child, .easy-autocomplete.eac-round ul .eac-category:first-child {
	  border-radius: 10px 10px 0 0;
	  border-width: 2px 2px 0;
	}
	.easy-autocomplete.eac-round ul li:last-child, .easy-autocomplete.eac-round ul .eac-category:last-child {
	  border-radius: 0 0 10px 10px;
	  border-width: 0 2px 2px;
	}
	.easy-autocomplete.eac-round ul li b, .easy-autocomplete.eac-round ul .eac-category b {
	  font-weight: 700;
	}

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
SQL = "SELECT * FROM IC_Partners where InternalRecordIdentifier = " & InternalRecordIdentifier 

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If NOT rs.EOF Then
	partnerAPIKey = rs.Fields("partnerAPIKey")
	partnerCompanyName = rs.Fields("partnerCompanyName")
	partnerPrimaryContactName = rs.Fields("partnerPrimaryContactName")
	partnerPrimaryContactEmail = rs.Fields("partnerPrimaryContactEmail")
	partnerTechnicalContactName = rs.Fields("partnerTechnicalContactName")
	partnerTechnicalContactEmail = rs.Fields("partnerTechnicalContactEmail")
	partnerAddress = rs.Fields("partnerAddress")
	partnerAddress2 = rs.Fields("partnerAddress2")
	partnerCity = rs.Fields("partnerCity")
	partnerState = rs.Fields("partnerState")
	partnerZip = rs.Fields("partnerZip")
	partnerPhone = rs.Fields("partnerPhone")
	partnerFax = rs.Fields("partnerFax")
	partnerUnmappedTaxableSKU = rs.Fields("partnerUnmappedTaxableSKU")
	partnerUnmappedTaxableUM = rs.Fields("partnerUnmappedTaxableUM")
	partnerUnmappedTaxablePassOriginalSKU = rs.Fields("partnerUnmappedTaxablePassOriginalSKU")
	partnerUnmappedNonTaxableSKU = rs.Fields("partnerUnmappedNonTaxableSKU")
	partnerUnmappedNonTaxableUM = rs.Fields("partnerUnmappedNonTaxableUM")
	partnerUnmappedNonTaxablePassOriginalSKU = rs.Fields("partnerUnmappedNonTaxablePassOriginalSKU")
	partnerUnmappedCustomerID = rs.Fields("partnerUnmappedCustomerID")
	partnerUnmappedPassOriginalCustomerID = rs.Fields("partnerUnmappedPassOriginalCustomerID")
	partnerRejectsBlankProdDescs = rs.Fields("partnerRejectsBlankProdDescs")
	partnerRejectsBlankProdUOMS = rs.Fields("partnerRejectsBlankProdUOMS")	
End If
set rs = Nothing
cnn8.close
set cnn8 = Nothing

%>


<h1 class="page-header"> Edit Partner</h1>

<div class="custom-container">

	<form method="POST" action="editPartner_submit.asp" name="frmEditPartner" id="frmEditPartner" onsubmit="return validateEditPartnerForm();">
	
	<div class="row row-line">	
	
		<h4>General Partner Information, API Key</h4>	
	
			  <input type="hidden" id="txtInternalRecordIdentifier" name="txtInternalRecordIdentifier" value="<%= InternalRecordIdentifier %>">	
			
              <div class="form-group">         	                
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-briefcase"></i></div>
	                    	<input type="text" class="form-control required" id="txtPartnerCompanyName" placeholder="Company Name" name="txtPartnerCompanyName" value="<%= partnerCompanyName %>">
	                   </div>
	                </div> 
               </div>
               
	          <div class="form-group">
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-globe"></i></div>
	                    	<input type="text" class="form-control required" id="txtPartnerAPIKey" placeholder="Partner API Key" name="txtPartnerAPIKey" value="<%= partnerAPIKey %>">
	                   </div>
	                </div> 
               </div>


              <div class="form-group">   
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-user"></i></div>
	                    	<input type="text" class="form-control" id="txtPrimaryContactName" placeholder="Primary Contact Name" name="txtPrimaryContactName" value="<%= partnerPrimaryContactName %>">
	                   </div>
	                </div> 	
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-envelope"></i></div>
	                    	<input type="text" class="form-control" id="txtPrimaryContactEmailAddress" placeholder="Primary Contact Email" name="txtPrimaryContactEmailAddress" value="<%= partnerPrimaryContactEmail %>">
	                   </div>
	                </div>  
               </div>               
               
   
              <div class="form-group">   
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-user"></i></div>
	                    	<input type="text" class="form-control" id="txtTechnicalContactName" placeholder="Technical Contact Name" name="txtTechnicalContactName" value="<%= partnerTechnicalContactName %>">
	                   </div>
	                </div> 	
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-envelope"></i></div>
	                    	<input type="text" class="form-control" id="txtTechnicalContactEmailAddress" placeholder="Technical Contact Email" name="txtTechnicalContactEmailAddress" value="<%= partnerTechnicalContactEmail %>">
	                   </div>
	                </div> 
	                 
               </div>

              <div class="form-group">   
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-address-card"></i></div>
	                    	<input type="text" class="form-control" id="txtAddressLine1" placeholder="Street Address" name="txtAddressLine1" value="<%= partnerAddress %>">
	                   </div>
	                </div> 
	           </div>     
           

	          <div class="form-group">
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-address-card-o"></i></div>
	                    	<input type="text" class="form-control" id="txtAddressLine2" placeholder="Suite, Floor #, etc." name="txtAddressLine2" value="<%= partnerAddress2 %>">
	                   </div>
	                </div> 
               </div>

              <div class="form-group">          
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-address-book"></i></div>
	                    	<input type="text" class="form-control" id="txtCity" placeholder="City" name="txtCity" value="<%= partnerCity %>">
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
	                    	<input type="text" class="form-control" id="txtZipCode" placeholder="Zip" name="txtZipCode" value="<%= partnerZip %>">
	                   </div>
	                </div> 
	                
	          </div>   
	            
              <div class="form-group">
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-phone"></i></div>
	                    	<input type="text" class="form-control" id="txtPhoneNumber" placeholder="Phone Number" name="txtPhoneNumber" value="<%= partnerPhone %>">
	                   </div>
	                </div> 
	 
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-fax"></i></div>
	                    	<input type="text" class="form-control" id="txtFaxNumber" placeholder="Fax Number" name="txtFaxNumber" value="<%= partnerFax %>">
	                   </div>
	                </div> 
	 
               </div>

			
			
		</div>	    <!-- cancel / submit !-->
		
		<hr>
		
		
		<div class="row row-line">		
		
			<h4>Specify the action to take based on blank product fields</h4>

              <div class="form-group">         	                
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<h5><input type="checkbox" id="chkRejectsBlankProdDescs" name="chkRejectsBlankProdDescs" <% If partnerRejectsBlankProdDescs = 1 OR partnerRejectsBlankProdDescs = true Then Response.Write("CHECKED") %>> 
	                    	API Rejects Products with Blank Descriptions</h5>
	                   </div>
	                </div> 
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<h5><input type="checkbox" id="chkRejectsBlankProdUOMS" name="chkRejectsBlankProdUOMS" <% If partnerRejectsBlankProdUOMS = 1 OR partnerRejectsBlankProdUOMS = true Then Response.Write("CHECKED") %>> 
	                    	API Rejects Products with Blank UOMs</h5>
	                   </div>
	                </div> 
               </div>
		</div>
		
		<hr>
		
			
		<div class="row row-line">		
		
			<h4>Specify the action to take when a taxable product id is not found in the product mapping table</h4>

              <div class="form-group">         	                
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<h5><input type="radio" id="optMappedOrPassedTaxable" name="optMappedOrPassedTaxable" <% If partnerUnmappedTaxablePassOriginalSKU = 1 OR partnerUnmappedTaxablePassOriginalSKU = true Then Response.Write("CHECKED") %> value="PassedCode"> 
	                    	Pass Through Using Original Product Code</h5>
	                   </div>
	                </div> 
	                <div class="col-sm-6">
	                  <div>
	                    	<h5><input type="radio" id="optMappedOrPassedTaxable" name="optMappedOrPassedTaxable" <% If partnerUnmappedTaxablePassOriginalSKU = 0 OR partnerUnmappedTaxablePassOriginalSKU = false Then Response.Write("CHECKED") %> value="DefinedCode"> 
	                    	Map to The Following Product Code/UM:</h5>
			        		<!-- select taxable product code !-->
								<input id="txtUnmappedTaxableSKU" name="txtUnmappedTaxableSKU" class="form-control" value="<%= partnerUnmappedTaxableSKU %>">
								<input type="hidden" id="txtUnmappedTaxableSKUToPass" name="txtUnmappedTaxableSKUToPass" value="<%= partnerUnmappedTaxableSKU %>">
								<input type="hidden" id="txtUnmappedTaxableUMToPass" name="txtUnmappedTaxableUMToPass" value="<%= partnerUnmappedTaxableUM %>">

								<i id="searchIcon" class="fa fa-search fa-2x"></i>
							<!-- eof select taxable product code !-->
	                   </div>
	                   <div id="taxableUM">   	                                   		
					   </div>
	                </div> 
               </div>
		</div>
		
		<hr>
		
		
		<div class="row row-line">		
		
			<h4>Specify the action to take when a non taxable product id is not found in the product mapping table</h4>

              <div class="form-group">         	                
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<h5><input type="radio" id="optMappedOrPassedNonTaxable" name="optMappedOrPassedNonTaxable" <% If partnerUnmappedNonTaxablePassOriginalSKU = 1 OR partnerUnmappedNonTaxablePassOriginalSKU = true Then Response.Write("CHECKED") %> value="PassedCode"> 
	                    	Pass Through Using Original Product Code</h5>
	                   </div>
	                </div> 
	                <div class="col-sm-6">
	                  <div>
	                    	<h5><input type="radio" id="optMappedOrPassedNonTaxable" name="optMappedOrPassedNonTaxable" <% If partnerUnmappedNonTaxablePassOriginalSKU = 0 OR partnerUnmappedNonTaxablePassOriginalSKU = false Then Response.Write("CHECKED") %> value="DefinedCode"> 
	                    	Map to The Following Product Code/UM:</h5>
			        		<!-- select non taxable product code !-->
								<input id="txtUnmappedNonTaxableSKU" name="txtUnmappedNonTaxableSKU" class="form-control" value="<%= partnerUnmappedNonTaxableSKU %>">
								<input type="hidden" id="txtUnmappedNonTaxableSKUToPass" name="txtUnmappedNonTaxableSKUToPass" value="<%= partnerUnmappedNonTaxableSKU %>">
								<input type="hidden" id="txtUnmappedNonTaxableUMToPass" name="txtUnmappedNonTaxableUMToPass" value="<%= partnerUnmappedNonTaxableUM %>">
								<i id="searchIcon" class="fa fa-search fa-2x"></i>
							<!-- eof select non taxable product code !-->
	                   </div>
	                   <div id="nonTaxableUM">               		
					   </div>
	                   
	                </div> 
               </div>
		</div>
		
		<hr>
		
		
		<div class="row row-line">		
		
			<h4>Specify the action to take when a customer id is not found in the customer mapping table</h4>

              <div class="form-group">         	                
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<h5><input type="radio" id="optMappedOrPassedCustAccount" name="optMappedOrPassedCustAccount" <% If partnerUnmappedPassOriginalCustomerID = 1 OR partnerUnmappedPassOriginalCustomerID = true Then Response.Write("CHECKED") %> value="PassedAccount"> 
	                    	Pass Through Using Original Customer Account Number</h5>
	                   </div>
	                </div> 
	                <div class="col-sm-6">
	                  <div>
	                    	<h5><input type="radio" id="optMappedOrPassedCustAccount" name="optMappedOrPassedCustAccount" <% If partnerUnmappedPassOriginalCustomerID = 0 OR partnerUnmappedPassOriginalCustomerID = false Then Response.Write("CHECKED") %> value="DefinedAccount"> 
	                    	Map to The Following Customer Account Number:</h5>
			        		<!-- select customer account code !-->
								<input id="txtUnmappedCustomerID" name="txtUnmappedCustomerID" class="form-control" value="<%= partnerUnmappedCustomerID %>">
								<input type="hidden" id="txtUnmappedCustomerIDToPass" name="txtUnmappedCustomerIDToPass" value="<%= partnerUnmappedCustomerID %>">
								<i id="searchIcon2" class="fa fa-search fa-2x"></i>
							<!-- eof select customer account code !-->
	                   </div>
	                </div> 
               </div>
		</div>
		
		<hr>
				
		
		<div class="row row-line">
			<div class="col-lg-12 alertbutton">
				<div class="col-lg-12">
					<a href="<%= BaseURL %>inventorycontrol/partners/main.asp"><button type="button" class="btn btn-default">&lsaquo; Cancel &amp; Go Back To Partners List</button></a>
					<button type="submit" class="btn btn-primary"><i class="far fa-save"></i> Save Changes</button>

				</div>
		    </div>
		</div>
		
	</form>
</div>

<!--#include file="../../inc/footer-main.asp"-->

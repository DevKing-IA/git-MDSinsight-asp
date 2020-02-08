<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->


<script src="<%= baseURL %>js/moment.min.js" type="text/javascript"></script>
<script type="text/javascript" src="<%= BaseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.min.js"></script>
<link href="<%= BaseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.css" rel="stylesheet">

<script>

	function highlightBlankFields(){
	  $(".input-group .form-control").each(function() {
	     var val = $(this).val();
	     if(val == "" || val == 0) {
	       $(this).css({ backgroundColor:'#ffff99' });
	     }
	     else {
	     	$(this).css({ backgroundColor:'#fff' });
	     }
	  });
	}

	$(window).load(function()
	{
	   var phones = [{ "mask": "(###) ###-####"}];
	    $('#txtPhoneNumber').inputmask({ 
	        mask: phones, 
	        greedy: false, 
	        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
	    $('#txtCellPhoneNumber').inputmask({ 
	        mask: phones, 
	        greedy: false, 
	        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
	    $('#txtFaxNumber').inputmask({ 
	        mask: phones, 
	        greedy: false, 
	        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
	        
	});

	$(document).ready(function() {
	
		highlightBlankFields();
		
		$(".input-group .form-control").blur(function() {
		  highlightBlankFields()
		});	
		
		
		$('input, select, textarea').each(
		    function(index){  
		        var input = $(this);
		        //console.log(input.attr('name'));
		    }
		); 
		
			
		$("#txtAccountNumber").focusout(function() {
						
			var passedNewCustID = $("#txtAccountNumber").val();
			
	    	$.ajax({
				type:"POST",
				url: "../../inc/InSightFuncs_AjaxForARAP.asp",
				cache: false,
				data: "action=CheckIfCustomerIDAlreadyExists&passedNewCustID=" + encodeURIComponent(passedNewCustID) + "&passedCurrCustID=''",
				success: function(response)
				 {
	               	 if (response == "CUSTIDALREADYEXISTS") {
	               	 	swal("That Account Number Already Exists for Another Customer.");
	               	 	$("#txtAccountNumber").val('');
	               	 }
				 }		
			});			
		
		        
		});
		
		
     
	});
</script>


<script language="JavaScript">
<!--
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

   function validateAddCustomerForm()
    {

       if (document.frmAddCustomer.txtAccountNumber.value == "") {
            swal("Account Number cannot be blank.");
            return false;
       }
       if (document.frmAddCustomer.txtCompanyName.value == "") {
            swal("Company name cannot be blank.");
            return false;
       }
       

       if (document.frmAddCustomer.txtBillToContactFirstName.value == "") {
            swal("Billing contact first name cannot be blank.");
            return false;
       }
       if (document.frmAddCustomer.txtBillToContactLastName.value == "") {
            swal("Billing contact last name cannot be blank.");
            return false;
       }
       if (document.frmAddCustomer.txtBillToCompanyName.value == "") {
            swal("Bill to company name cannot be blank.");
            return false;
       }
       if (document.frmAddCustomer.txtBillToAddressLine1.value == "") {
            swal("Bill to address cannot be blank.");
            return false;
       }
       if (document.frmAddCustomer.txtBillToCity.value == "") {
            swal("Bill to city cannot be blank.");
            return false;
       }
       if (document.frmAddCustomer.txtBillToState.value == "") {
            swal("Bill to state cannot be blank.");
            return false;
       }
        if (document.frmAddCustomer.txtBillToZipCode.value == "") {
            swal("Bill to zip code cannot be blank.");
            return false;
       }
       if (document.frmAddCustomer.txtBillToCountry.value == "") {
            swal("Bill to country cannot be blank.");
            return false;
       }
       if ((document.frmAddCustomer.txtBillToPhoneNumber.value !== "") && (isValidPhone(document.frmAddCustomer.txtBillToPhoneNumber.value) == false)) {
           swal("The Billing phone number is invalid. Please enter the number in the following format: (555) 555-5555.");
           return false;
       }
       if ((document.frmAddCustomer.txtBillToEmailAddress.value !== "") && (isValidEmail(document.frmAddCustomer.txtBillToEmailAddress.value) == false)) {
           swal("The Billing contact email address is invalid. Please enter any format like the following: anystring@anystring.anystring.");
           return false;
       }


 
       if (document.frmAddCustomer.txtShipToContactFirstName.value == "") {
            swal("Shipping contact first name cannot be blank.");
            return false;
       }
       if (document.frmAddCustomer.txtShipToContactLastName.value == "") {
            swal("Shipping contact last name cannot be blank.");
            return false;
       }
       if (document.frmAddCustomer.txtShipToCompanyName.value == "") {
            swal("Ship to company name cannot be blank.");
            return false;
       }
       if (document.frmAddCustomer.txtShipToAddressLine1.value == "") {
            swal("Ship to address cannot be blank.");
            return false;
       }
       if (document.frmAddCustomer.txtShipToCity.value == "") {
            swal("Ship to city cannot be blank.");
            return false;
       }
       if (document.frmAddCustomer.txtShipToState.value == "") {
            swal("Ship to state cannot be blank.");
            return false;
       }
        if (document.frmAddCustomer.txtShipToZipCode.value == "") {
            swal("Ship to zip code cannot be blank.");
            return false;
       }
       if (document.frmAddCustomer.txtShipToCountry.value == "") {
            swal("Ship to country cannot be blank.");
            return false;
       }
       if ((document.frmAddCustomer.txtShipToPhoneNumber.value !== "") && (isValidPhone(document.frmAddCustomer.txtShipToPhoneNumber.value) == false)) {
           swal("The shipping phone number is invalid. Please enter the number in the following format: (555) 555-5555.");
           return false;
       }
       if ((document.frmAddCustomer.txtShipToEmailAddress.value !== "") && (isValidEmail(document.frmAddCustomer.txtShipToEmailAddress.value) == false)) {
           swal("The shipping contact email address is invalid. Please enter any format like the following: anystring@anystring.anystring.");
           return false;
       }

	   
      
       return true;

    }
    
    
	function SetShipping(checked) {  
	
	      if (checked) {  
	                document.getElementById('txtShipToContactFirstName').value = document.getElementById('txtBillToContactFirstName').value;   
	                document.getElementById('txtShipToContactLastName').value = document.getElementById('txtBillToContactLastName').value;   
	                document.getElementById('txtShipToCompanyName').value = document.getElementById('txtBillToCompanyName').value;   
	                document.getElementById('txtShipToAddressLine1').value = document.getElementById('txtBillToAddressLine1').value;   
	                document.getElementById('txtShipToAddressLine2').value = document.getElementById('txtBillToAddressLine2').value;
	                document.getElementById('txtShipToCity').value = document.getElementById('txtBillToCity').value;
	                document.getElementById('txtShipToState').value = document.getElementById('txtBillToState').value;
	                document.getElementById('txtShipToZipCode').value = document.getElementById('txtBillToZipCode').value;
	                document.getElementById('txtShipToCountry').value = document.getElementById('txtBillToCountry').value;
	                document.getElementById('txtShipToPhoneNumber').value = document.getElementById('txtBillToPhoneNumber').value;
	                document.getElementById('txtShipToEmailAddress').value = document.getElementById('txtBillToEmailAddress').value;  
	      } else {  
	                document.getElementById('txtShipToContactFirstName').value = '';   
	                document.getElementById('txtShipToContactLastName').value = '';   
	                document.getElementById('txtShipToCompanyName').value = '';   
	                document.getElementById('txtShipToAddressLine1').value = '';   
	                document.getElementById('txtShipToAddressLine2').value = '';  
	                document.getElementById('txtShipToCity').value = '';
	                document.getElementById('txtShipToState').value = '';
	                document.getElementById('txtShipToZipCode').value = '';
	                document.getElementById('txtShipToCountry').value = ''; 
	                document.getElementById('txtShipToPhoneNumber').value = '';
	                document.getElementById('txtShipToEmailAddress').value = '';
	      }  
	}  
  
    
// -->
</script>   

<style type="text/css">

/*Colored Content Boxes
------------------------------------*/

	.container{
		width: 100%;
	}
	
	.quick-info-block {
	  padding: 3px 20px;
	  text-align: center;
	  margin-bottom: 20px;
	  border-radius: 7px;
	}
	
	.quick-info-block p{
	  color: #fff;
	  font-size:16px;
	}
	.quick-info-block h2 {
	  color: #fff;
	  font-size:20px;
	  margin-bottom:25px;
	}

	.quick-info-block h2.black {
	  color: #000;
	  font-size:20px;
	  margin-bottom:25px;
	}
	
	.quick-info-block h2 a:hover{
	  text-decoration: none;
	}
	
	.quick-info-block-light,
	.quick-info-block-default {
	  background: #fafafa;
	  border: solid 1px #eee; 
	}
	
	.quick-info-block-default:hover {
	  box-shadow: 0 0 8px #eee;
	}
	
	.quick-info-block-light p,
	.quick-info-block-light h2,
	.quick-info-block-default p,
	.quick-info-block-default h2 {
	  color: #555;
	}

	.quick-info-block-u {
	  background: #72c02c;
	}
	.quick-info-block-blue {
	  background: #80B8FF;
	}
	.quick-info-block-red {
	  background: #e74c3c;
	}
	.quick-info-block-sea {
	  background: #1abc9c;
	}
	.quick-info-block-grey {
	  background: #f8f8f8;
	}
	.quick-info-block-yellow {
	  background: #f1c40f;
	}
	.quick-info-block-orange {
	  background: #e67e22;
	}
	.quick-info-block-green {
	  background: #2ecc71;
	}
	.quick-info-block-purple {
	  background: #9b6bcc;
	}
	.quick-info-block-aqua {
	  background: #27d7e7;
	}
	.quick-info-block-brown {
	  background: #9c8061;
	}
	.quick-info-block-dark-blue {
	  background: #4765a0;
	}
	.quick-info-block-light-green {
	  background: #79d5b3;
	}
	.quick-info-block-dark {
	  background: #555;
	}
	.quick-info-block-light {
	  background: #ecf0f1;
	}
	
	textarea.form-control {
    	height: 100px; !important;
    	width:385px !important;
    	border-radius:3px !important;
	}
		
	hr.tile {
	    border: 0;
	    height: 3px;
	    background-image: linear-gradient(to right, rgba(0, 0, 0, 0), rgba(255, 255, 255, 0.95), rgba(0, 0, 0, 0));
	}

	.red-line{
		border-left:3px solid red;
	}   
	
	input[type=checkbox] {
	  transform:scale(1.5, 1.5);
	}	  
	

	.sameascheckbox{
	    /*display: block;*/
	    width: 100%;
	    height: 20px;
	    padding: 6px 12px !important;
	    font-size: 14px;
	    line-height: 1.42857143 !important;
	    color: #555;
	    background-color: #fff;
	    border: 1px solid #ccc; 
	    border-radius: 4px;
	    margin-right:0px;
	    vertical-align: middle;
	}

</style>
<!-- eof css !-->

<h1 class="page-header"><i class="fas fa-users-medical"></i>&nbsp;Add Customer&nbsp;&nbsp;(All Fields w/Red Line Required)
	<!-- customize !-->
	<div class="col pull-right">
	</div>
	<!-- eof customize !-->
</h1>

		
<form autocomplete="off" action="<%= BaseURL %>accountsreceivable/customermininfo/addCustomerSubmit.asp" method="POST" name="frmAddCustomer" id="frmAddCustomer" onsubmit="return validateAddCustomerForm();" class="form-horizontal track-event-form bv-form">

<input autocomplete="false" name="hidden" type="text" style="display:none;">
<div class="container pull-left">
<div class="row">


	  <div class="col-md-12">
	  
			<div class="quick-info-block quick-info-block-blue">
			<h2 class="heading-md"><i class="fas fa-file-user"></i>&nbsp;Main Account Information</h2>
	
	              <div class="form-group">         
		                <div class="col-sm-4">
		                  <div class="input-group">
		                    	<div class="input-group-addon"><i class="fas fw fa-id-card-alt"></i></div>
		                    	<input type="text" class="form-control red-line" id="txtAccountNumber" placeholder="Account Number" name="txtAccountNumber">
		                   </div>
		                </div> 
	
		                <div class="col-sm-4">
		                  <div class="input-group">
		                    	<div class="input-group-addon"><i class="fas fw fa-building"></i></div>
		                    	<input type="text" class="form-control red-line" id="txtCompanyName" placeholder="Company Name" name="txtCompanyName">
		                   </div>
		                </div> 
		                
		                <div class="col-sm-4">
			               <div class="form-group">
			               		<div class="col-sm-5"><p>Last Price Change Date</p></div>
								<div class="col-sm-7">
							        <div class="input-group date" id="datetimepickerLastPriceChangeDate">
							            <input type="text" class="form-control" id="txtLastPriceChangeDate" name="txtLastPriceChangeDate" value="">
							            <span class="input-group-addon">
							                <span class="glyphicon glyphicon-calendar"></span>
							            </span>
							        </div>
							    </div>
							</div>   
							<script type="text/javascript">
					            $(function () {
					                $('#datetimepickerLastPriceChangeDate').datetimepicker({
					                   maxDate: moment(),
					                   useCurrent: false,
					                   format: 'MM/DD/YYYY',
					                   ignoreReadonly: true
					                });  
					            });
					        </script> 	
					    </div>	                
	               </div>
			</div>
	  </div>

      <div class="col-md-6">

		<div class="quick-info-block quick-info-block-green">
		<h2 class="heading-md"><i class="fas fa-file-invoice-dollar"></i>&nbsp;Customer Account/Billing Information</h2> 

              <div class="form-group">
	                            
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fas fw fa-user-circle"></i></div>
	                    	<input type="text" class="form-control red-line" id="txtBillToContactFirstName" placeholder="Billing Contact First Name" name="txtBillToContactFirstName">
	                   </div>
	                </div> 
	                
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="far fw fa-user-circle"></i></div>
	                    	<input type="text" class="form-control red-line" id="txtBillToContactLastName" placeholder="Billing Contact Last Name" name="txtBillToContactLastName">
	                   </div>
	                </div> 
	 
               </div>
               

              <div class="form-group">          	                
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fw fa-briefcase"></i></div>
	                    	<input type="text" class="form-control red-line" id="txtBillToCompanyName" placeholder="Bill To Company Name" name="txtBillToCompanyName">
	                   </div>
	                </div> 
               </div>
               

              <div class="form-group">         
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fw fa-address-card"></i></div>
	                    	<input type="text" class="form-control red-line" id="txtBillToAddressLine1" placeholder="Billing Street Address" name="txtBillToAddressLine1">
	                   </div>
	                </div> 
	           </div>     
	                
	                
	          <div class="form-group">
	                
	                <div class="col-sm-8">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fad fw fa-building"></i></div>
	                    	<input type="text" class="form-control" id="txtBillToAddressLine2" placeholder="Suite, Floor #, etc." name="txtBillToAddressLine2">
	                   </div>
	                </div> 
	 
               </div>



              <div class="form-group">
	                            
	                <div class="col-sm-9">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fas fw fa-city"></i></div>
	                    	<input type="text" class="form-control red-line" id="txtBillToCity" placeholder="Bill To City" name="txtBillToCity">
	                   </div>
	                </div> 
	                
	          </div>     
	          <div class="form-group">

	                <div class="col-sm-7">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fas fw fa-map-pin"></i></div>
                    		<select data-placeholder="Choose State" class="C_Country_Modal form-control red-line" id="txtBillToState" name="txtBillToState"> 
                    			<option value="">Bill To State</option>
								<!--#include file="statelist.asp"-->
							</select>				
		
	                   </div>
	                </div> 
	                <div class="col-sm-5">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fas fw fa-mailbox"></i></div>
	                    	<input type="text" class="form-control red-line" id="txtBillToZipCode" placeholder="Bill To Zip Code" name="txtBillToZipCode">
	                   </div>
	                </div> 
	 
               </div>
               
                
              <div class="form-group">          	                
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fw fa-globe"></i></div>
                    		<select data-placeholder="Choose Country" class="C_Country_Modal form-control red-line" id="txtBillToCountry" name="txtBillToCountry"> 
								<option value="">Bill To Country</option>
								<!--#include file="countrylist.asp"-->
							</select>
	                   </div>
	                </div>  	 
               </div>
              
              <div class="form-group">  
              
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fw fa-phone"></i></div>
	                    	<input type="text" class="form-control red-line" id="txtBillToPhoneNumber" placeholder="Billing Phone Number" name="txtBillToPhoneNumber">
	                   </div>
	                </div> 
        
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fw fa-envelope"></i></div>
	                    	<input type="text" class="form-control red-line" id="txtBillToEmailAddress" placeholder="Billing Email Address" name="txtBillToEmailAddress">
	                   </div>
	                </div> 
	          </div>    
               
              
                    
			</div>
        <!-- END QUICK INFO BOX -->
        
        
      </div><!-- end col-md-6 -->


      
     <div class="col-md-6">
		
		<div class="quick-info-block quick-info-block-grey">
		<h2 class="heading-md black"><i class="fas fa-truck"></i>&nbsp;Customer Shipping Information</h2>

				<div class="form-group">
				  <div class="col-md-1">
					<div style="white-space:nowrap;display:inline;">
						<input type="checkbox" class="sameascheckbox" onclick="SetShipping(this.checked);">&nbsp;<h4 style="display:inline;">Same as Billing Info</h4>
					</div>
				  </div>
			    </div>

              <div class="form-group">
	                            
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fas fw fa-user-circle"></i></div>
	                    	<input type="text" class="form-control red-line" id="txtShipToContactFirstName" placeholder="Shipping Contact First Name" name="txtShipToContactFirstName">
	                   </div>
	                </div> 
	                
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="far fw fa-user-circle"></i></div>
	                    	<input type="text" class="form-control red-line" id="txtShipToContactLastName" placeholder="Shipping Contact Last Name" name="txtShipToContactLastName">
	                   </div>
	                </div> 
	 
               </div>
               

              <div class="form-group">          	                
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fw fa-briefcase"></i></div>
	                    	<input type="text" class="form-control red-line" id="txtShipToCompanyName" placeholder="Ship To Company Name" name="txtShipToCompanyName">
	                   </div>
	                </div> 
               </div>
               

              <div class="form-group">         
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fw fa-address-card"></i></div>
	                    	<input type="text" class="form-control red-line" id="txtShipToAddressLine1" placeholder="Shipping Street Address" name="txtShipToAddressLine1">
	                   </div>
	                </div> 
	           </div>     
	                
	                
	          <div class="form-group">
	                
	                <div class="col-sm-8">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fad fw fa-building"></i></div>
	                    	<input type="text" class="form-control" id="txtShipToAddressLine2" placeholder="Suite, Floor #, etc." name="txtShipToAddressLine2">
	                   </div>
	                </div> 
	 
               </div>



              <div class="form-group">
	                            
	                <div class="col-sm-9">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fas fw fa-city"></i></div>
	                    	<input type="text" class="form-control red-line" id="txtShipToCity" placeholder="Ship To City" name="txtShipToCity">
	                   </div>
	                </div> 
	                
	          </div>     
	          <div class="form-group">

	                <div class="col-sm-7">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fas fw fa-map-pin"></i></div>
                    		<select data-placeholder="Choose State" class="C_Country_Modal form-control red-line" id="txtShipToState" name="txtShipToState"> 
                    			<option value="">Ship To State</option>
								<!--#include file="statelist.asp"-->
							</select>				
		
	                   </div>
	                </div> 
	                <div class="col-sm-5">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fas fw fa-mailbox"></i></div>
	                    	<input type="text" class="form-control red-line" id="txtShipToZipCode" placeholder="Ship To Zip Code" name="txtShipToZipCode">
	                   </div>
	                </div> 
	 
               </div>
               
                
              <div class="form-group">          	                
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fw fa-globe"></i></div>
                    		<select data-placeholder="Choose Country" class="C_Country_Modal form-control red-line" id="txtShipToCountry" name="txtShipToCountry"> 
								<option value="">Ship To Country</option>
								<!--#include file="countrylist.asp"-->
							</select>
	                   </div>
	                </div>  	 
               </div>
              
              <div class="form-group">  
              
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fw fa-phone"></i></div>
	                    	<input type="text" class="form-control red-line" id="txtShipToPhoneNumber" placeholder="Shipping Phone Number" name="txtShipToPhoneNumber">
	                   </div>
	                </div> 
        
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fw fa-envelope"></i></div>
	                    	<input type="text" class="form-control red-line" id="txtShipToEmailAddress" placeholder="Shipping Email Address" name="txtShipToEmailAddress">
	                   </div>
	                </div> 
	          </div>    
               
              
                    
			</div>
        <!-- END QUICK INFO BOX -->
        
      </div><!-- end col-md-6 -->


 </div> <!-- end row -->
        
        
<div class="form-group pull-right">
	<div class="col-lg-12">
		<button class="btn btn-primary btn-lg btn-block" href="<%= BaseURL %>accountsreceivable/customermininfo/addCustomerSubmit.asp" role="button" type="submit"><i class="fas fa-save"></i>&nbsp;SAVE THIS CUSTOMER</button>
	</div>
</div>
        
</div> <!-- end container -->

</form>
 
 <!-- tabs js !-->
 <script type="text/javascript">
 $('a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
  e.target // newly activated tab
  e.relatedTarget // previous active tab
})
 </script>
 <!-- eof tabs js !-->

<!--#include file="../../inc/footer-main.asp"-->
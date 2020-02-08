<!--#include file="../inc/header.asp"-->

<%
ClientKeyForFileNames = MUV_READ("ClientKeyForFileNames")
%>

<script type="text/javascript">

	$(function () {
		var autocompleteJSONFileURL = "../clientfiles/<%= ClientKeyForFileNames %>/autocomplete/customer_account_list_CSZ_<%= ClientKeyForFileNames %>.json";
		//var autocompleteJSONFileURL = "../clientfiles/1106d/autocomplete/customer_account_list_CSZ_POS_1106d.json";
		var options = {
		  url: autocompleteJSONFileURL,
		  placeholder: "Search for a customer by name, account, city, state, zip",
		  getValue: "name",
		  list: {	
	        onChooseEvent: function() {
	            var custID = $("#txtCustID").getSelectedItemData().code;
	            
				 if (custID!=""){
				 	$.ajax({
						type:"POST",
						url: "../inc/InSightAjaxFuncs.asp",
						data: "action=selectAccount&custID="+encodeURIComponent(custID),
							success: function(msg){
								window.location = "addServiceMemo.asp";
							}
					}) 
				
				  }

        	},		  
		    match: {
		      enabled: true
			},
			maxNumberOfElements: 15		
		  },
		  theme: "round"
		};
		$("#txtCustID").easyAutocomplete(options);

	})
</script>


<link rel="stylesheet" type="text/css" href="<%= BaseURL %>css/tabs.css" />

<style type="text/css">
	  
	.table-info .table>tbody>tr>td, .table>tbody>tr>th, .table>tfoot>tr>td, .table>tfoot>tr>th, .table>thead>tr>td, .table>thead>tr>th{
		border: 0px;
		font-weight: bold;
		line-height: 0.8;
	}

</style>

          
<SCRIPT LANGUAGE="JavaScript">
<!--
    function validateServiceMemoform()
    {

        if (document.frmAddServiceMemo.txtContactName.value == "") {
	        swal("Contact name can not be blank.")
            return false;
        }
        if (document.frmAddServiceMemo.txtLocation.value == "") {
            swal("Problem Location can not be blank.");
            return false;
        }
        if (document.frmAddServiceMemo.txtDescription.value == "") {
            swal("Problem description can not be blank.");
            return false;
        }


        return true;

    }
    
   
// -->
</SCRIPT>     

<style type="text/css">
	.alert{
 		padding: 6px 12px;
	}
	
	.form-control{
		margin-bottom: 20px;
	}
	
	a:hover{
		text-decoration: none;
	}
</style>

<% SelectedCustomer = Session("ServiceCustID") %>
<h1 class="page-header"><i class="fa fa-wrench"></i> Submit A New Service Ticket</h1>

	
	<form method="POST" action="addservicememo_submit.asp" name="frmAddServiceMemo" onsubmit="return validateServiceMemoform();">		    
      
		<!-- row !-->		
		<div class="row ">
			<div class="col-lg-8">
				<div class="row">
					<!-- select company !-->
					<div class="col-lg-4 col-md-4 col-sm-12 col-xs-12">
						<input id="txtCustID" name="txtCustID">
						<i id="searchIcon" class="fa fa-search fa-2x"></i>
					</div>
					<!-- eof select company !-->
				</div>
			</div>
		</div>

		<!-- row !-->		
		<div class="row ">
			<div class="col-lg-8">
				<div class="row">
					<!--account number !-->
					<div class="col-lg-8 col-md-8 col-sm-12 col-xs-12">
						<% If SelectedCustomer  <> "" Then %>
							<!--#include file="../inc/commonCustomerDisplay.asp"-->
						<% End If %>				        
					</div>					
					<!-- eof account number !-->
				</div>
			</div>
			<!-- eof row !-->
		</div>
		
		<% If SelectedCustomer  <> "" Then 'Only show all this stuff if a customer has been selected %>
		 <!-- main row !-->

		 <div class="row">
 
			 <!-- left col !-->
			 <div class="col-lg-7 col-md-7 col-sm-12 col-xs-12">
 
 		        <!-- row !-->			
			    <div class="row">

					<!-- Contact Name !-->
					<div class="col-lg-4">
						<strong>Contact Name</strong>
						<input type="text" id="txtContactName" name="txtContactName" class="form-control" >
					</div>
					<!-- Contact Name !-->
					
					<!-- Contact Phone !-->
					<div class="col-lg-4">
						<strong>Contact Phone</strong>
						<input type="text" id="txtContactPhone" name="txtContactPhone"   class="form-control">
					</div>
					<!-- Contact Phone !-->
		    	
				</div>
				<!-- eof row !-->
		    	
				<!-- row !-->			
				<div class="row">

					<div class="col-lg-4">
						<strong>Contact Email</strong>
						<input type="text" id="txtContactEmail" name="txtContactEmail" class="form-control">
					</div>

					<!-- Problem Location !-->
					<div class="col-lg-4">
						<strong>Problem Location</strong> <small>(Please include floor # if applicable)</small>
						<input type="text" id="txtLocation" name="txtLocation"   class="form-control">
					</div>
					<!-- Problem Location !-->

				<!-- end row !-->		
				</div>
			
			</div>
			
		</div>
		<!-- eof main row !-->

	
		<!-- right col !-->
		<div class="row">
			<div class="col-lg-5 col-md-5 col-sm-12 col-xs-12">	  
				<!-- Description of problem !-->
				<strong>Please enter a description as completely as possible.</strong>
				<textarea name="txtDescription" id="txtDescription" rows="5" spellcheck="True" class="form-control"></textarea>
				<!-- Description of problem !-->
			</div>
		</div>
		<!-- eof right col !-->
		<% End If%>
		
		
		<div class="row">
		
		<div class="col-lg-12">	<br>
		    <a href="<%= BaseURL %>service/main.asp">
		    	<button type="button" class="btn btn-default">&lsaquo; Cancel &amp; Go Back</button>
			</a>
			<% If SelectedCustomer <> "" Then 'Only show all this stuff if a customer has been selected %>
				<button type="submit" class="btn btn-primary"><i class="fa fa-upload"></i> Submit</button>
			<% End If %>
		</div>
	</div>
	<!-- eof row !-->    
</form>

<!--#include file="../inc/footer-service.asp"-->

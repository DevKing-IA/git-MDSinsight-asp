 <%'**********************
' **** Contacts Tab *****
'************************
%>
<%

SQLBillToStates = "SELECT StateCode,StateName FROM PR_States ORDER BY StateName ASC"
SQLBillToCountries = "SELECT CountryCode,CountryName FROM PR_Countries ORDER BY CountryOrder ASC, CountryName ASC"

Set cnnBillToStates = Server.CreateObject("ADODB.Connection")
cnnBillToStates.open (Session("ClientCnnString"))

Set rsBillToStates = Server.CreateObject("ADODB.Recordset")
Set rsBillToStates = cnnBillToStates.Execute(SQLBillToStates)

BillToStates = ("[{""id"":"""",""title"":""Select""},")

If not rsBillToStates.EOF Then
	sep = ""
	Do While Not rsBillToStates.EOF
			BillToStates = BillToStates & (sep)
			sep = ","
			BillToStates = BillToStates & ("{")
			BillToStates = BillToStates & ("""id"":""" & Replace(rsBillToStates("StateCode"), """", "\""") & """")
			BillToStates = BillToStates & (",""title"":""" & Replace(rsBillToStates("StateName"), """", "\""") & """")
			BillToStates = BillToStates & ("}")
		rsBillToStates.MoveNext						
	Loop
End If
BillToStates = BillToStates & ("]")
Set rsBillToStates = Nothing


Set rsBillToCountries = Server.CreateObject("ADODB.Recordset")
Set rsBillToCountries = cnnBillToStates.Execute(SQLBillToCountries)

BillToCountries = ("[")
If not rsBillToCountries.EOF Then
	sep = ""
	Do While Not rsBillToCountries.EOF
			BillToCountries = BillToCountries & (sep)
			sep = ","
			BillToCountries = BillToCountries & ("{")
			BillToCountries = BillToCountries & ("""id"":""" & Replace(rsBillToCountries("CountryCode"), """", "\""") & """")
			BillToCountries = BillToCountries & (",""title"":""" & Replace(rsBillToCountries("CountryName"), """", "\""") & """")
			BillToCountries = BillToCountries & ("}")
		rsBillToCountries.MoveNext						
	Loop
End If
BillToCountries = BillToCountries & ("]")
Set rsBillToCountries = Nothing


%> 


<%'********************
' **** Contacts Tab****
'**********************
%>


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


   function validateAddEditBillToLocationFields(updateActionId)
    {

       var txtBillToCompany = $("#txtCompanyTab" + updateActionId).val();
       if (txtBillToCompany == "") {
            swal("Bill to company name cannot be blank.");
            return false;
       }
       
       var txtBillToContactFirstName = $("#txtContactFirstNameTab" + updateActionId).val();
       if (txtBillToContactFirstName == "") {
            swal("Billing contact first name cannot be blank.");
            return false;
       }
       
       var txtBillToContactLastName = $("#txtContactLastNameTab" + updateActionId).val();
       if (txtBillToContactLastName == "") {
            swal("Billing contact last name cannot be blank.");
            return false;
       }
 
       var txtBillToAddressLine1 = $("#txtBillToAddress1" + updateActionId).val();
       if (txtBillToAddressLine1 == "") {
            swal("Bill to address cannot be blank.");
            return false;
       }
      
       var txtBillToCity = $("#txtCityTab" + updateActionId).val();
       if (txtBillToCity == "") {
            swal("Bill to city cannot be blank.");
            return false;
       }
       
       var txtBillToState = $("#txtStateTab" + updateActionId).val();
       if (txtBillToState == "") {
            swal("Bill to state cannot be blank.");
            return false;
       }
       
       var txtBillToZip = $("#txtPostalTab" + updateActionId).val();
        if (txtBillToZip == "") {
            swal("Bill to zip code cannot be blank.");
            return false;
       }
       
       var txtBillToCountry = $("#txtCountryTab" + updateActionId).val();
       if (txtBillToCountry == "") {
            swal("Bill to country cannot be blank.");
            return false;
       }
 
       //var txtBillToPhone = $("#txtPhoneTab" + updateActionId).val();
       //if ((txtBillToPhone == "") || (typeof txtBillToPhone == 'undefined')) {
           	//swal("The Billing phone number is invalid. Please enter the number in the following format: (555) 555-5555.");
           	//return false;
       //}
       
       var txtBillToEmail = $("#txtEmailTab" + updateActionId).val(); 
       if ((txtBillToEmail !== "") && (isValidEmail(txtBillToEmail) == false)) {
           swal("The Billing contact email address is invalid. Please enter any format like the following: anystring@anystring.anystring.");
           return false;
       }
      
       return true;

    }
    
    
    
// -->
</script>   


<div role="tabpanel" class="tab-pane fade active in" id="billtos">

	<div class="input-group narrow-results"><span class="input-group-addon">Narrow Results</span>
		<input id="filter-billtolocations" type="text" class="form-control filter-search-width" placeholder="Type here...">
	</div>
	  
	<p><button type="button" class="btn btn-success" onclick="ajaxRowNewBillingLocation();">Create New Bill To Location</button></p>
			  
	  <div class="table-responsive">
            <table class="table table-striped table-condensed table-hover table-bordered sortable standard-font">
              <thead>
                <tr>
                  <th width="2%">Default</th>
                  <th width="5%">CustID</th>
				  <th width="15%">Company</th>
				  <th width="6%">Contact First Name</th>
				  <th width="6%">Contact Last Name</th>			  
				  <th width="12%">Address 1, Address 2</th>
				  <th width="10%">City, State Zip</th>
                  <th width="6%">Country</th>
				  <th width="12%">Email</th>
				  <th width="6%">Phone</th>
				  <th width="6%">Fax</th>                  
                  <th class="sorttable_nosort" width="7%">Actions</th>
                </tr>
              </thead>
             
              <tbody id="ajaxContainerBillingLocations" class='searchable-billtos ajax-loading'></tbody>

		</table>
		

	</div>
</div>
<%'************************
' **** eof Contacts Tab****
'**************************
%>

<script>
	var BillToStates = <%= BillToStates %>;
	var BillToCountries = <%= BillToCountries %>;
	
	var curcontactid =0;
	var value_default = {};
	var DefaultBillToFound = false;
	
	$(document).ready(function () { 

		ajaxLoadBillingLocations();
		
		//contact title modal window submit
		$('#frmAddContactTitleTab').submit(function(e) {
			
			if ($('#frmAddContactTitleTab #txtContactTitleTab').val()==''){
				 swal("Contact title can not be blank.");
				return false;
			}
			
	        return false;
	    });			
					
	});
			
   function checkDefaultBillTo(el) {
   
		if ($('#chkDefaultBillTo' + el).is(':checked')) {
		
			//Make sure the user is not un-checking the default bill to location
		
			$.ajax({
				type:"POST",
				url: "../../inc/InSightFuncs_AjaxForARAP.asp",
				cache: false,
				data: "action=CheckIfDefaultBillingLocation&i=" + encodeURIComponent(el),
				success: function(response)
				 {
				 	if (response == "OKTODELETE") {
						$('#ajaxContainerBillingLocations #chkDefaultBillTo' + el).prop('checked', false);			 	
					}
				 	else {
				 	 	$('#ajaxContainerBillingLocations #chkDefaultBillTo' + el).prop('checked', true);
				 	 	swal("You must have one default billing location. Please set or create another default billing location and then un-check.");
					}       	 
	             }	
			});	
		
			
		}  

    }		
        
    
	function ajaxRowNewBillingLocation() {
		var value = {};
		
		value.id = 0;				//id
		value.DefaultBillTo = 0;	//primary
		value.BillToCustNum = "";	//customer id
		value.BillToCompany = "";	//company
		value.BillToContactFirstName = "";		//firstname
		value.BillToContactLastName = "";		//lastname
		value.BillToAddress1 = "";	//Address1
		value.BillToAddress2 = "";	//Address2
		value.BillToCity = "";		//City
		value.BillToState = "";		//State
		value.BillToZip = "";		//Postal Code
		value.BillToCountry = "";	//Country
		value.BillToEmail = "";		//email
		value.BillToPhone = "";		//phone
		value.BillToFax = "";		//fax
		
		$('#ajaxRowBillingLocations-' + 0 + '').remove();
		
		if (DefaultBillToFound){
			$("#ajaxContainerBillingLocations").prepend(ajaxRowhtmlBillingLocations(value_default));
		} else {
			$("#ajaxContainerBillingLocations").prepend(ajaxRowhtmlBillingLocations(value));
		}
		
	   var phonesTab = [{ "mask": "(###) ###-####"}];
	    $('#txtPhoneTab0').inputmask({ 
	        mask: phonesTab, 
	        greedy: false, 
	        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
	    $('#txtFaxTab0').inputmask({ 
	        mask: phonesTab, 
	        greedy: false, 
	        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
	
	}
	
	
	function ajaxRowhtmlBillingLocations(value) {
	
		var BillToStateSelect = '<select class="form-control" data-type="BillToState" name="txtBillToState" id="txtStateTab' + value.id + '">';		
		$.each(BillToStates, function (key, BillToState) {
			BillToStateSelect +='<option value="'+BillToState.id+'" ' + (value.BillToState+""==BillToState.id+""?'selected':'') + '>'+BillToState.title+'</option>';
		});		
		BillToStateSelect +='</select>';
		
		var BillToCountrySelect = '<select class="form-control" data-type="BillToCountry" name="txtBillToCountry" id="txtCountryTab' + value.id + '">';		
		$.each(BillToCountries, function (key, BillToCountry) {
			BillToCountrySelect +='<option value="'+BillToCountry.id+'" ' + (value.BillToCountry+""==BillToCountry.id+""?'selected':'') + '>'+BillToCountry.title+'</option>';
		});		
		BillToCountrySelect +='</select>';
		
		
		if (value.DefaultBillTo==1){
			DefaultBillToFound= true;
			value_default.id = 0;				//id
			value_default.DefaultBillTo = 0;	//primary
			value_default.BillToCustNum = value.BillToCustNum;	//customer id
			value_default.BillToCompany = "";	//company
			value_default.BillToContactFirstName = "";		//firstname
			value_default.BillToContactLastName = "";		//lastname
			value_default.BillToAddress1 = "";//Address1
			value_default.BillToAddress2 = "";//Address2
			value_default.BillToCity = "";		//City
			value_default.BillToState = "";		//State
			value_default.BillToZip = "";		//Postal Code
			value_default.BillToCountry = "";	//Country		
			value_default.BillToEmail = "";		//email
			value_default.BillToPhone = "";		//phone
			value_default.BillToFax = "";		//fax
								
		}
		
		var btns = '\
					<div class="visibleRowView btn-group btn-group-sm"><a class="btn btn-primary" onclick="ajaxRowMode(\'BillingLocations\', ' + value.id + ', \'Edit\');"><i class="fa fa-edit"></i></a><a class="btn btn-danger" onclick="ajaxDeleteBillingLocations(\'delete\', ' + value.id + ');"><i class="fas fa-trash-alt"></i></a></div>\
					<div class="visibleRowEdit btn-group btn-group-sm"><a class="btn btn-success" onclick="ajaxSaveInsertBillingLocations(\'save\', ' + value.id + ');"><i class="fa fa-save"></i></a><a class="btn btn-default" id="btnEdit" onclick="ajaxRowMode(\'BillingLocations\', ' + value.id + ', \'View\');"><i class="fa fa-times"></i></a></div>\
				';
		if(value.id==0)
			btns = '\
					<div class="visibleRowEdit btn-group btn-group-sm"><a class="btn btn-success" onclick="ajaxSaveInsertBillingLocations(\'insert\', ' + value.id + ');"><i class="fa fa-save"></i></a><a class="btn btn-default" onclick="ajaxRowMode(\'BillingLocations\', ' + value.id + ', \'View\');"><i class="fa fa-times"></i></a></div>\
				';	
		var htmlBillingLocations = '\
			<tr id="ajaxRowBillingLocations-' + value.id + '" class="' + (value.id==0?'ajaxRowEdit':'ajaxRowView') + '">\
				<td>\
					<div class="visibleRowView"><input type="checkbox" disabled="true" data-type="DefaultBillTo" ' + (value.DefaultBillTo==1?'checked':'') + ' id="chkDefaultBillTo' + (value.id) + '" value="' + (value.id) + '"></div>\
					<div class="visibleRowEdit"><input type="checkbox" onClick="checkDefaultBillTo(' + value.id + ');" data-type="DefaultBillTo" ' + (value.DefaultBillTo==1?'checked':'') + ' id="chkDefaultBillTo' + (value.id) + '" value="' + (value.id) + '"></div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.BillToCustNum + '</div>\
					<div class="visibleRowEdit">' + value.BillToCustNum + '</div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.BillToCompany + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="BillToCompany" id="txtCompanyTab' + value.id + '" name="txtBillToCompany" value="' + value.BillToCompany + '" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.BillToContactFirstName + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="BillToContactFirstName" id="txtContactFirstNameTab' + value.id + '" name="txtBillToContactFirstName" value="' + value.BillToContactFirstName + '" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.BillToContactLastName + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="BillToContactLastName" id="txtContactLastNameTab' + value.id + '" name="txtBillToContactLastName" value="' + value.BillToContactLastName + '" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.BillToAddress1 + ', ' + value.BillToAddress2 + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="BillToAddress1" name="txtBillToAddress1" id="txtAddress1Tab' + value.id + '" value="' + value.BillToAddress1 + '" placeholder="Address1" />\
					<input class="form-control" data-type="BillToAddress2" name="txtBillToAddress2" id="txtAddress2Tab' + value.id + '" value="' + value.BillToAddress2 + '"  placeholder="Address2" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.BillToCity + ', '+value.BillToState  + ' '+value.BillToZip + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="BillToCity" name="txtBillToCity" id="txtCityTab' + value.id + '" value="' + value.BillToCity + '"  placeholder="City" />'+ BillToStateSelect +'<input class="form-control" data-type="BillToZip" name="txtBillToZip" id="txtPostalTab' + value.id + '" value="' + value.BillToZip + '"  placeholder="Zip" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.BillToCountry + '</div>\
					<div class="visibleRowEdit">' + BillToCountrySelect + '</div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.BillToEmail + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="BillToEmail" name="txtBillToEmail" id="txtEmailTab' + value.id + '" value="' + value.BillToEmail + '" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.BillToPhone + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="BillToPhone" name="txtBillToPhone" id="txtPhoneTab' + value.id + '" value="' + value.BillToPhone + '" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.BillToFax + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="BillToFax" name="txtBillToFax" id="txtFaxTab' + value.id + '" value="' + value.BillToFax + '" /></div>\
				</td>\
				<td class="text-center">'+btns+'</td>\
		   </tr>\
			';
		return htmlBillingLocations;
		//<div class="visibleRowEdit"><input class="form-control" data-type="Country" id="txtCountryTab' + value.id + '" value="' + value.Country + '"  placeholder="Country" /></div>\
	}
	
	
	
	
	
	function ajaxLoadBillingLocations(updateAction, updateActionId) {

		
		$("#ajaxContainerBillingLocations").addClass("ajax-loading");
		var url = "ajax/ar_billto.asp?i=<%= InternalRecordIdentifier %>&cid=<%= customerID %>";
		var jsondata = {};
		jsondata.updateAction = updateAction;
		jsondata.updateActionId = updateActionId;
		

		if(updateAction=="save" || updateAction=="insert"){
			jsondata.DefaultBillTo	= $('#ajaxRowBillingLocations-' + updateActionId + ' [data-type="DefaultBillTo"]').is(':checked')?1:0;
			jsondata.BillToCustNum	= $('#ajaxRowBillingLocations-' + updateActionId + ' [data-type="BillToCustNum"]').val();
			jsondata.BillToCompany	= $('#ajaxRowBillingLocations-' + updateActionId + ' [data-type="BillToCompany"]').val();
			jsondata.BillToContactFirstName	= $('#ajaxRowBillingLocations-' + updateActionId + ' [data-type="BillToContactFirstName"]').val();
			jsondata.BillToContactLastName = $('#ajaxRowBillingLocations-' + updateActionId + ' [data-type="BillToContactLastName"]').val();
			jsondata.BillToAddress1		= $('#ajaxRowBillingLocations-' + updateActionId + ' [data-type="BillToAddress1"]').val();
			jsondata.BillToAddress2		= $('#ajaxRowBillingLocations-' + updateActionId + ' [data-type="BillToAddress2"]').val();
			jsondata.BillToCity 		= $('#ajaxRowBillingLocations-' + updateActionId + ' [data-type="BillToCity"]').val();
			jsondata.BillToState 		= $('#ajaxRowBillingLocations-' + updateActionId + ' [data-type="BillToState"]').val();
			jsondata.BillToZip			= $('#ajaxRowBillingLocations-' + updateActionId + ' [data-type="BillToZip"]').val();
			jsondata.BillToCountry 		= $('#ajaxRowBillingLocations-' + updateActionId + ' [data-type="BillToCountry"]').val();
			jsondata.BillToEmail		= $('#ajaxRowBillingLocations-' + updateActionId + ' [data-type="BillToEmail"]').val();
			jsondata.BillToPhone		= $('#ajaxRowBillingLocations-' + updateActionId + ' [data-type="BillToPhone"]').val();
			jsondata.BillToFax 			= $('#ajaxRowBillingLocations-' + updateActionId + ' [data-type="BillToFax"]').val();
			
		}
			
		$.ajax({
			type: "POST",
			url: url,
			dataType: "json",
			data: jsondata,
			success: function (data) {
				//if (!data || data+""=="") { window.location.href = window.location.href + ""; return; }				
				var htmlBillingLocations = "";
				$.each(data, function (key, value) {
					htmlBillingLocations += ajaxRowhtmlBillingLocations(value);
				});
				$("#ajaxContainerBillingLocations").html(htmlBillingLocations);
				
				setTimeout(function(){
					$("#ajaxContainerBillingLocations").removeClass("ajax-loading");
				}, 0);
				
			},
			failure: function (data) {
				$("#ajaxContainerBillingLocations").html("Failed To Load Billing Locations");
				setTimeout(function(){
					$("#ajaxContainerBillingLocations").removeClass("ajax-loading");
				}, 0);
				
			}
			
		});
			

		
	}
	
	
	
	
	function ajaxSaveInsertBillingLocations(updateAction, updateActionId) {

		
		$("#ajaxContainerBillingLocations").addClass("ajax-loading");
		var url = "ajax/ar_billto.asp?i=<%= InternalRecordIdentifier %>&cid=<%= customerID %>";
		var jsondata = {};
		jsondata.updateAction = updateAction;
		jsondata.updateActionId = updateActionId;
		

		if(updateAction=="save" || updateAction=="insert"){
			jsondata.DefaultBillTo	= $('#ajaxRowBillingLocations-' + updateActionId + ' [data-type="DefaultBillTo"]').is(':checked')?1:0;
			jsondata.BillToCustNum	= $('#ajaxRowBillingLocations-' + updateActionId + ' [data-type="BillToCustNum"]').val();
			jsondata.BillToCompany	= $('#ajaxRowBillingLocations-' + updateActionId + ' [data-type="BillToCompany"]').val();
			jsondata.BillToContactFirstName	= $('#ajaxRowBillingLocations-' + updateActionId + ' [data-type="BillToContactFirstName"]').val();
			jsondata.BillToContactLastName = $('#ajaxRowBillingLocations-' + updateActionId + ' [data-type="BillToContactLastName"]').val();
			jsondata.BillToAddress1		= $('#ajaxRowBillingLocations-' + updateActionId + ' [data-type="BillToAddress1"]').val();
			jsondata.BillToAddress2		= $('#ajaxRowBillingLocations-' + updateActionId + ' [data-type="BillToAddress2"]').val();
			jsondata.BillToCity 		= $('#ajaxRowBillingLocations-' + updateActionId + ' [data-type="BillToCity"]').val();
			jsondata.BillToState 		= $('#ajaxRowBillingLocations-' + updateActionId + ' [data-type="BillToState"]').val();
			jsondata.BillToZip			= $('#ajaxRowBillingLocations-' + updateActionId + ' [data-type="BillToZip"]').val();
			jsondata.BillToCountry 		= $('#ajaxRowBillingLocations-' + updateActionId + ' [data-type="BillToCountry"]').val();
			jsondata.BillToEmail		= $('#ajaxRowBillingLocations-' + updateActionId + ' [data-type="BillToEmail"]').val();
			jsondata.BillToPhone		= $('#ajaxRowBillingLocations-' + updateActionId + ' [data-type="BillToPhone"]').val();
			jsondata.BillToFax 			= $('#ajaxRowBillingLocations-' + updateActionId + ' [data-type="BillToFax"]').val();
			
			if (validateAddEditBillToLocationFields(updateActionId)) {
			
				$.ajax({
					type: "POST",
					url: url,
					dataType: "json",
					data: jsondata,
					success: function (data) {
						//if (!data || data+""=="") { window.location.href = window.location.href + ""; return; }				
						var htmlBillingLocations = "";
						$.each(data, function (key, value) {
							htmlBillingLocations += ajaxRowhtmlBillingLocations(value);
						});
						$("#ajaxContainerBillingLocations").html(htmlBillingLocations);
						
						setTimeout(function(){
							$("#ajaxContainerBillingLocations").removeClass("ajax-loading");
						}, 0);
						
					},
					failure: function (data) {
						$("#ajaxContainerBillingLocations").html("Failed To Load Billing Locations");
						setTimeout(function(){
							$("#ajaxContainerBillingLocations").removeClass("ajax-loading");
						}, 0);
						
					}
					
				});
				
			}
			else {
				setTimeout(function(){
					$("#ajaxContainerBillingLocations").removeClass("ajax-loading");
				}, 0);
						
			}
		}
		
	}
	
	
	function ajaxDeleteBillingLocations(updateAction, updateActionId) {


		//Make sure the user is not trying to delete the default billing location
		if (updateAction == "delete") {
			$.ajax({
				type:"POST",
				url: "../../inc/InSightFuncs_AjaxForARAP.asp",
				cache: false,
				data: "action=CheckIfDefaultBillingLocation&i=" + encodeURIComponent(updateActionId),
				success: function(response)
				 {
				 	if (response == "OKTODELETE") {

						var r = confirm("Are your sure you want to delete this billing location?");
						
						if (r == true) {
	
							$("#ajaxContainerBillingLocations").addClass("ajax-loading");
							var url = "ajax/ar_billto.asp?i=<%= InternalRecordIdentifier %>&cid=<%= customerID %>";
							var jsondata = {};
							jsondata.updateAction = updateAction;
							jsondata.updateActionId = updateActionId;
							
							$.ajax({
								type: "POST",
								url: url,
								dataType: "json",
								data: jsondata,
								success: function (data) {				
									var htmlBillingLocations = "";
									$.each(data, function (key, value) {
										htmlBillingLocations += ajaxRowhtmlBillingLocations(value);
									});
									$("#ajaxContainerBillingLocations").html(htmlBillingLocations);
									
									setTimeout(function(){
										$("#ajaxContainerBillingLocations").removeClass("ajax-loading");
									}, 0);
									
								},
								failure: function (data) {
									$("#ajaxContainerBillingLocations").html("Failed To Load Billing Locations");
									setTimeout(function(){
										$("#ajaxContainerBillingLocations").removeClass("ajax-loading");
									}, 0);
									
								}
								
							});

						}				 	
					}
				 	else {
				 	 	swal("You cannot delete the default billing location. Please set or create another default billing location and then delete.");
					}       	 
	             }	
			});	
		}
		
	}
	
</script>


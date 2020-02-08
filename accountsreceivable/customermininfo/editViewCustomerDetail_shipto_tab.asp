 <%'**********************
' **** Contacts Tab *****
'************************
%>
<%

SQLShipToStates = "SELECT StateCode,StateName FROM PR_States ORDER BY StateName ASC"
SQLShipToCountries = "SELECT CountryCode,CountryName FROM PR_Countries ORDER BY CountryOrder ASC, CountryName ASC"

Set cnnShipToStates = Server.CreateObject("ADODB.Connection")
cnnShipToStates.open (Session("ClientCnnString"))

Set rsShipToStates = Server.CreateObject("ADODB.Recordset")
Set rsShipToStates = cnnShipToStates.Execute(SQLShipToStates)

ShipToStates = ("[{""id"":"""",""title"":""Select""},")

If not rsShipToStates.EOF Then
	sep = ""
	Do While Not rsShipToStates.EOF
			ShipToStates = ShipToStates & (sep)
			sep = ","
			ShipToStates = ShipToStates & ("{")
			ShipToStates = ShipToStates & ("""id"":""" & Replace(rsShipToStates("StateCode"), """", "\""") & """")
			ShipToStates = ShipToStates & (",""title"":""" & Replace(rsShipToStates("StateName"), """", "\""") & """")
			ShipToStates = ShipToStates & ("}")
		rsShipToStates.MoveNext						
	Loop
End If
ShipToStates = ShipToStates & ("]")
Set rsShipToStates = Nothing


Set rsShipToCountries = Server.CreateObject("ADODB.Recordset")
Set rsShipToCountries = cnnShipToStates.Execute(SQLShipToCountries)

ShipToCountries = ("[")
If not rsShipToCountries.EOF Then
	sep = ""
	Do While Not rsShipToCountries.EOF
			ShipToCountries = ShipToCountries & (sep)
			sep = ","
			ShipToCountries = ShipToCountries & ("{")
			ShipToCountries = ShipToCountries & ("""id"":""" & Replace(rsShipToCountries("CountryCode"), """", "\""") & """")
			ShipToCountries = ShipToCountries & (",""title"":""" & Replace(rsShipToCountries("CountryName"), """", "\""") & """")
			ShipToCountries = ShipToCountries & ("}")
		rsShipToCountries.MoveNext						
	Loop
End If
ShipToCountries = ShipToCountries & ("]")
Set rsShipToCountries = Nothing


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


   function validateAddEditShipToLocationFields(updateActionId)
    {

       var txtShipToCompany = $("#txtCompanyTab" + updateActionId).val();
       if (txtShipToCompany == "") {
            swal("Ship to company name cannot be blank.");
            return false;
       }
       
       var txtShipToContactFirstName = $("#txtContactFirstNameTab" + updateActionId).val();
       if (txtShipToContactFirstName == "") {
            swal("Shipping contact first name cannot be blank.");
            return false;
       }
       
       var txtShipToContactLastName = $("#txtContactLastNameTab" + updateActionId).val();
       if (txtShipToContactLastName == "") {
            swal("Shipping contact last name cannot be blank.");
            return false;
       }
 
       var txtShipToAddressLine1 = $("#txtShipToAddress1" + updateActionId).val();
       if (txtShipToAddressLine1 == "") {
            swal("Ship to address cannot be blank.");
            return false;
       }
      
       var txtShipToCity = $("#txtCityTab" + updateActionId).val();
       if (txtShipToCity == "") {
            swal("Ship to city cannot be blank.");
            return false;
       }
       
       var txtShipToState = $("#txtStateTab" + updateActionId).val();
       if (txtShipToState == "") {
            swal("Ship to state cannot be blank.");
            return false;
       }
       
       var txtShipToZip = $("#txtPostalTab" + updateActionId).val();
        if (txtShipToZip == "") {
            swal("Ship to zip code cannot be blank.");
            return false;
       }
       
       var txtShipToCountry = $("#txtCountryTab" + updateActionId).val();
       if (txtShipToCountry == "") {
            swal("Ship to country cannot be blank.");
            return false;
       }
 
       //var txtShipToPhone = $("#txtPhoneTab" + updateActionId).val();
       //if ((txtShipToPhone == "") || (typeof txtShipToPhone == 'undefined')) {
           	//swal("The Shipping phone number is invalid. Please enter the number in the following format: (555) 555-5555.");
           	//return false;
       //}
       
       var txtShipToEmail = $("#txtEmailTab" + updateActionId).val(); 
       if ((txtShipToEmail !== "") && (isValidEmail(txtShipToEmail) == false)) {
           swal("The Shipping contact email address is invalid. Please enter any format like the following: anystring@anystring.anystring.");
           return false;
       }
      
       return true;

    }
    
    
    
// -->
</script>   


<div role="tabpanel" class="tab-pane fade" id="shiptos">

	<div class="input-group narrow-results"><span class="input-group-addon">Narrow Results</span>
		<input id="filter-ShipTolocations" type="text" class="form-control filter-search-width" placeholder="Type here...">
	</div>
	  
	<p><button type="button" class="btn btn-success" onclick="ajaxRowNewShippingLocation();">Create New Ship to Location</button></p>
			  
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
             
              <tbody id="ajaxContainerShippingLocations" class='searchable-shiptos ajax-loading'></tbody>

		</table>
		

	</div>
</div>
<%'************************
' **** eof Contacts Tab****
'**************************
%>

<script>
	var ShipToStates = <%= ShipToStates %>;
	var ShipToCountries = <%= ShipToCountries %>;
	
	var curcontactid =0;
	var value_default = {};
	var DefaultShipToFound = false;
	
	$(document).ready(function () { 

		ajaxLoadShippingLocations();
		
		//contact title modal window submit
		$('#frmAddContactTitleTab').submit(function(e) {
			
			if ($('#frmAddContactTitleTab #txtContactTitleTab').val()==''){
				 swal("Contact title can not be blank.");
				return false;
			}
			
	        return false;
	    });			
					
	});
			
   function checkDefaultShipTo(el) {
   
		if ($('#chkDefaultShipTo' + el).is(':checked')) {
		
			//Make sure the user is not un-checking the default Ship to location
		
			$.ajax({
				type:"POST",
				url: "../../inc/InSightFuncs_AjaxForARAP.asp",
				cache: false,
				data: "action=CheckIfDefaultShippingLocation&i=" + encodeURIComponent(el),
				success: function(response)
				 {
				 	if (response == "OKTODELETE") {
						$('#ajaxContainerShippingLocations #chkDefaultShipTo' + el).prop('checked', false);			 	
					}
				 	else {
				 	 	$('#ajaxContainerShippingLocations #chkDefaultShipTo' + el).prop('checked', true);
				 	 	swal("You must have one default Shipping location. Please set or create another default Shipping location and then un-check.");
					}       	 
	             }	
			});	
		
			
		}  

    }		
        
    
	function ajaxRowNewShippingLocation() {
		var value = {};
		
		value.id = 0;				//id
		value.DefaultShipTo = 0;	//primary
		value.ShipToCustNum = "";	//customer id
		value.ShipToCompany = "";	//company
		value.ShipToContactFirstName = "";		//firstname
		value.ShipToContactLastName = "";		//lastname
		value.ShipToAddress1 = "";	//Address1
		value.ShipToAddress2 = "";	//Address2
		value.ShipToCity = "";		//City
		value.ShipToState = "";		//State
		value.ShipToZip = "";		//Postal Code
		value.ShipToCountry = "";	//Country
		value.ShipToEmail = "";		//email
		value.ShipToPhone = "";		//phone
		value.ShipToFax = "";		//fax
		
		$('#ajaxRowShippingLocations-' + 0 + '').remove();
		
		if (DefaultShipToFound){
			$("#ajaxContainerShippingLocations").prepend(ajaxRowhtmlShippingLocations(value_default));
		} else {
			$("#ajaxContainerShippingLocations").prepend(ajaxRowhtmlShippingLocations(value));
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
	
	
	function ajaxRowhtmlShippingLocations(value) {
	
		var ShipToStateSelect = '<select class="form-control" data-type="ShipToState" name="txtShipToState" id="txtStateTab' + value.id + '">';		
		$.each(ShipToStates, function (key, ShipToState) {
			ShipToStateSelect +='<option value="'+ShipToState.id+'" ' + (value.ShipToState+""==ShipToState.id+""?'selected':'') + '>'+ShipToState.title+'</option>';
		});		
		ShipToStateSelect +='</select>';
		
		var ShipToCountrySelect = '<select class="form-control" data-type="ShipToCountry" name="txtShipToCountry" id="txtCountryTab' + value.id + '">';		
		$.each(ShipToCountries, function (key, ShipToCountry) {
			ShipToCountrySelect +='<option value="'+ShipToCountry.id+'" ' + (value.ShipToCountry+""==ShipToCountry.id+""?'selected':'') + '>'+ShipToCountry.title+'</option>';
		});		
		ShipToCountrySelect +='</select>';
		
		
		if (value.DefaultShipTo==1){
			DefaultShipToFound= true;
			value_default.id = 0;				//id
			value_default.DefaultShipTo = 0;	//primary
			value_default.ShipToCustNum = value.ShipToCustNum;	//customer id
			value_default.ShipToCompany = "";	//company
			value_default.ShipToContactFirstName = "";		//firstname
			value_default.ShipToContactLastName = "";		//lastname
			value_default.ShipToAddress1 = "";//Address1
			value_default.ShipToAddress2 = "";//Address2
			value_default.ShipToCity = "";		//City
			value_default.ShipToState = "";		//State
			value_default.ShipToZip = "";		//Postal Code
			value_default.ShipToCountry = "";	//Country		
			value_default.ShipToEmail = "";		//email
			value_default.ShipToPhone = "";		//phone
			value_default.ShipToFax = "";		//fax
								
		}
		
		var btns = '\
					<div class="visibleRowView btn-group btn-group-sm"><a class="btn btn-primary" onclick="ajaxRowMode(\'ShippingLocations\', ' + value.id + ', \'Edit\');"><i class="fa fa-edit"></i></a><a class="btn btn-danger" onclick="ajaxDeleteShippingLocations(\'delete\', ' + value.id + ');"><i class="fas fa-trash-alt"></i></a></div>\
					<div class="visibleRowEdit btn-group btn-group-sm"><a class="btn btn-success" onclick="ajaxSaveInsertShippingLocations(\'save\', ' + value.id + ');"><i class="fa fa-save"></i></a><a class="btn btn-default" id="btnEdit" onclick="ajaxRowMode(\'ShippingLocations\', ' + value.id + ', \'View\');"><i class="fa fa-times"></i></a></div>\
				';
		if(value.id==0)
			btns = '\
					<div class="visibleRowEdit btn-group btn-group-sm"><a class="btn btn-success" onclick="ajaxSaveInsertShippingLocations(\'insert\', ' + value.id + ');"><i class="fa fa-save"></i></a><a class="btn btn-default" onclick="ajaxRowMode(\'ShippingLocations\', ' + value.id + ', \'View\');"><i class="fa fa-times"></i></a></div>\
				';	
		var htmlShippingLocations = '\
			<tr id="ajaxRowShippingLocations-' + value.id + '" class="' + (value.id==0?'ajaxRowEdit':'ajaxRowView') + '">\
				<td>\
					<div class="visibleRowView"><input type="checkbox" disabled="true" data-type="DefaultShipTo" ' + (value.DefaultShipTo==1?'checked':'') + ' id="chkDefaultShipTo' + (value.id) + '" value="' + (value.id) + '"></div>\
					<div class="visibleRowEdit"><input type="checkbox" onClick="checkDefaultShipTo(' + value.id + ');" data-type="DefaultShipTo" ' + (value.DefaultShipTo==1?'checked':'') + ' id="chkDefaultShipTo' + (value.id) + '" value="' + (value.id) + '"></div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.ShipToCustNum + '</div>\
					<div class="visibleRowEdit">' + value.ShipToCustNum + '</div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.ShipToCompany + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="ShipToCompany" id="txtCompanyTab' + value.id + '" name="txtShipToCompany" value="' + value.ShipToCompany + '" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.ShipToContactFirstName + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="ShipToContactFirstName" id="txtContactFirstNameTab' + value.id + '" name="txtShipToContactFirstName" value="' + value.ShipToContactFirstName + '" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.ShipToContactLastName + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="ShipToContactLastName" id="txtContactLastNameTab' + value.id + '" name="txtShipToContactLastName" value="' + value.ShipToContactLastName + '" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.ShipToAddress1 + ', ' + value.ShipToAddress2 + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="ShipToAddress1" name="txtShipToAddress1" id="txtAddress1Tab' + value.id + '" value="' + value.ShipToAddress1 + '" placeholder="Address1" />\
					<input class="form-control" data-type="ShipToAddress2" name="txtShipToAddress2" id="txtAddress2Tab' + value.id + '" value="' + value.ShipToAddress2 + '"  placeholder="Address2" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.ShipToCity + ', '+value.ShipToState  + ' '+value.ShipToZip + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="ShipToCity" name="txtShipToCity" id="txtCityTab' + value.id + '" value="' + value.ShipToCity + '"  placeholder="City" />'+ ShipToStateSelect +'<input class="form-control" data-type="ShipToZip" name="txtShipToZip" id="txtPostalTab' + value.id + '" value="' + value.ShipToZip + '"  placeholder="Zip" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.ShipToCountry + '</div>\
					<div class="visibleRowEdit">' + ShipToCountrySelect + '</div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.ShipToEmail + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="ShipToEmail" name="txtShipToEmail" id="txtEmailTab' + value.id + '" value="' + value.ShipToEmail + '" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.ShipToPhone + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="ShipToPhone" name="txtShipToPhone" id="txtPhoneTab' + value.id + '" value="' + value.ShipToPhone + '" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.ShipToFax + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="ShipToFax" name="txtShipToFax" id="txtFaxTab' + value.id + '" value="' + value.ShipToFax + '" /></div>\
				</td>\
				<td class="text-center">'+btns+'</td>\
		   </tr>\
			';
		return htmlShippingLocations;
		//<div class="visibleRowEdit"><input class="form-control" data-type="Country" id="txtCountryTab' + value.id + '" value="' + value.Country + '"  placeholder="Country" /></div>\
	}
	
	
	
	
	
	function ajaxLoadShippingLocations(updateAction, updateActionId) {

		
		$("#ajaxContainerShippingLocations").addClass("ajax-loading");
		var url = "ajax/ar_ShipTo.asp?i=<%= InternalRecordIdentifier %>&cid=<%= customerID %>";
		var jsondata = {};
		jsondata.updateAction = updateAction;
		jsondata.updateActionId = updateActionId;
		

		if(updateAction=="save" || updateAction=="insert"){
			jsondata.DefaultShipTo	= $('#ajaxRowShippingLocations-' + updateActionId + ' [data-type="DefaultShipTo"]').is(':checked')?1:0;
			jsondata.ShipToCustNum	= $('#ajaxRowShippingLocations-' + updateActionId + ' [data-type="ShipToCustNum"]').val();
			jsondata.ShipToCompany	= $('#ajaxRowShippingLocations-' + updateActionId + ' [data-type="ShipToCompany"]').val();
			jsondata.ShipToContactFirstName	= $('#ajaxRowShippingLocations-' + updateActionId + ' [data-type="ShipToContactFirstName"]').val();
			jsondata.ShipToContactLastName = $('#ajaxRowShippingLocations-' + updateActionId + ' [data-type="ShipToContactLastName"]').val();
			jsondata.ShipToAddress1		= $('#ajaxRowShippingLocations-' + updateActionId + ' [data-type="ShipToAddress1"]').val();
			jsondata.ShipToAddress2		= $('#ajaxRowShippingLocations-' + updateActionId + ' [data-type="ShipToAddress2"]').val();
			jsondata.ShipToCity 		= $('#ajaxRowShippingLocations-' + updateActionId + ' [data-type="ShipToCity"]').val();
			jsondata.ShipToState 		= $('#ajaxRowShippingLocations-' + updateActionId + ' [data-type="ShipToState"]').val();
			jsondata.ShipToZip			= $('#ajaxRowShippingLocations-' + updateActionId + ' [data-type="ShipToZip"]').val();
			jsondata.ShipToCountry 		= $('#ajaxRowShippingLocations-' + updateActionId + ' [data-type="ShipToCountry"]').val();
			jsondata.ShipToEmail		= $('#ajaxRowShippingLocations-' + updateActionId + ' [data-type="ShipToEmail"]').val();
			jsondata.ShipToPhone		= $('#ajaxRowShippingLocations-' + updateActionId + ' [data-type="ShipToPhone"]').val();
			jsondata.ShipToFax 			= $('#ajaxRowShippingLocations-' + updateActionId + ' [data-type="ShipToFax"]').val();
			
		}
			
		$.ajax({
			type: "POST",
			url: url,
			dataType: "json",
			data: jsondata,
			success: function (data) {
				//if (!data || data+""=="") { window.location.href = window.location.href + ""; return; }				
				var htmlShippingLocations = "";
				$.each(data, function (key, value) {
					htmlShippingLocations += ajaxRowhtmlShippingLocations(value);
				});
				$("#ajaxContainerShippingLocations").html(htmlShippingLocations);
				
				setTimeout(function(){
					$("#ajaxContainerShippingLocations").removeClass("ajax-loading");
				}, 0);
				
			},
			failure: function (data) {
				$("#ajaxContainerShippingLocations").html("Failed To Load Shipping Locations");
				setTimeout(function(){
					$("#ajaxContainerShippingLocations").removeClass("ajax-loading");
				}, 0);
				
			}
			
		});
			

		
	}
	
	
	
	
	function ajaxSaveInsertShippingLocations(updateAction, updateActionId) {

		
		$("#ajaxContainerShippingLocations").addClass("ajax-loading");
		var url = "ajax/ar_ShipTo.asp?i=<%= InternalRecordIdentifier %>&cid=<%= customerID %>";
		var jsondata = {};
		jsondata.updateAction = updateAction;
		jsondata.updateActionId = updateActionId;
		

		if(updateAction=="save" || updateAction=="insert"){
			jsondata.DefaultShipTo	= $('#ajaxRowShippingLocations-' + updateActionId + ' [data-type="DefaultShipTo"]').is(':checked')?1:0;
			jsondata.ShipToCustNum	= $('#ajaxRowShippingLocations-' + updateActionId + ' [data-type="ShipToCustNum"]').val();
			jsondata.ShipToCompany	= $('#ajaxRowShippingLocations-' + updateActionId + ' [data-type="ShipToCompany"]').val();
			jsondata.ShipToContactFirstName	= $('#ajaxRowShippingLocations-' + updateActionId + ' [data-type="ShipToContactFirstName"]').val();
			jsondata.ShipToContactLastName = $('#ajaxRowShippingLocations-' + updateActionId + ' [data-type="ShipToContactLastName"]').val();
			jsondata.ShipToAddress1		= $('#ajaxRowShippingLocations-' + updateActionId + ' [data-type="ShipToAddress1"]').val();
			jsondata.ShipToAddress2		= $('#ajaxRowShippingLocations-' + updateActionId + ' [data-type="ShipToAddress2"]').val();
			jsondata.ShipToCity 		= $('#ajaxRowShippingLocations-' + updateActionId + ' [data-type="ShipToCity"]').val();
			jsondata.ShipToState 		= $('#ajaxRowShippingLocations-' + updateActionId + ' [data-type="ShipToState"]').val();
			jsondata.ShipToZip			= $('#ajaxRowShippingLocations-' + updateActionId + ' [data-type="ShipToZip"]').val();
			jsondata.ShipToCountry 		= $('#ajaxRowShippingLocations-' + updateActionId + ' [data-type="ShipToCountry"]').val();
			jsondata.ShipToEmail		= $('#ajaxRowShippingLocations-' + updateActionId + ' [data-type="ShipToEmail"]').val();
			jsondata.ShipToPhone		= $('#ajaxRowShippingLocations-' + updateActionId + ' [data-type="ShipToPhone"]').val();
			jsondata.ShipToFax 			= $('#ajaxRowShippingLocations-' + updateActionId + ' [data-type="ShipToFax"]').val();
			
			if (validateAddEditShipToLocationFields(updateActionId)) {
			
				$.ajax({
					type: "POST",
					url: url,
					dataType: "json",
					data: jsondata,
					success: function (data) {
						//if (!data || data+""=="") { window.location.href = window.location.href + ""; return; }				
						var htmlShippingLocations = "";
						$.each(data, function (key, value) {
							htmlShippingLocations += ajaxRowhtmlShippingLocations(value);
						});
						$("#ajaxContainerShippingLocations").html(htmlShippingLocations);
						
						setTimeout(function(){
							$("#ajaxContainerShippingLocations").removeClass("ajax-loading");
						}, 0);
						
					},
					failure: function (data) {
						$("#ajaxContainerShippingLocations").html("Failed To Load Shipping Locations");
						setTimeout(function(){
							$("#ajaxContainerShippingLocations").removeClass("ajax-loading");
						}, 0);
						
					}
					
				});
				
			}
			else {
				setTimeout(function(){
					$("#ajaxContainerShippingLocations").removeClass("ajax-loading");
				}, 0);
						
			}
		}
		
	}
	
	
	function ajaxDeleteShippingLocations(updateAction, updateActionId) {


		//Make sure the user is not trying to delete the default Shipping location
		if (updateAction == "delete") {
			$.ajax({
				type:"POST",
				url: "../../inc/InSightFuncs_AjaxForARAP.asp",
				cache: false,
				data: "action=CheckIfDefaultShippingLocation&i=" + encodeURIComponent(updateActionId),
				success: function(response)
				 {
				 	if (response == "OKTODELETE") {

						var r = confirm("Are your sure you want to delete this Shipping location?");
						
						if (r == true) {
	
							$("#ajaxContainerShippingLocations").addClass("ajax-loading");
							var url = "ajax/ar_ShipTo.asp?i=<%= InternalRecordIdentifier %>&cid=<%= customerID %>";
							var jsondata = {};
							jsondata.updateAction = updateAction;
							jsondata.updateActionId = updateActionId;
							
							$.ajax({
								type: "POST",
								url: url,
								dataType: "json",
								data: jsondata,
								success: function (data) {				
									var htmlShippingLocations = "";
									$.each(data, function (key, value) {
										htmlShippingLocations += ajaxRowhtmlShippingLocations(value);
									});
									$("#ajaxContainerShippingLocations").html(htmlShippingLocations);
									
									setTimeout(function(){
										$("#ajaxContainerShippingLocations").removeClass("ajax-loading");
									}, 0);
									
								},
								failure: function (data) {
									$("#ajaxContainerShippingLocations").html("Failed To Load Shipping Locations");
									setTimeout(function(){
										$("#ajaxContainerShippingLocations").removeClass("ajax-loading");
									}, 0);
									
								}
								
							});

						}				 	
					}
				 	else {
				 	 	swal("You cannot delete the default Shipping location. Please set or create another default Shipping location and then delete.");
					}       	 
	             }	
			});	
		}
		
	}
	
</script>



<style>
 .decisionMaker {
	font-size:22px;
	font-weight:bold;
	font-color:#00FF00;
}

 .decisionMaker2 {
}

</style>
<div role="tabpanel" class="tab-pane fade in" id="contacts">
	  
		<div class="input-group narrow-results"> <span class="input-group-addon">Narrow Results</span>
			<input id="filter-contacts" type="text" class="form-control filter-search-width" placeholder="Type here...">
		</div>
			  
	  <div class="table-responsive">
            <table class="table table-striped table-condensed table-hover table-bordered sortable standard-font">
              <thead>
                <tr>
                  <th width="2%">Primary<br>Contact</th>
                  <th width="2%">Decision<br>Maker</th>
				  <th width="3%">Suffix</th>
				  <th width="5%">First Name</th>
				  <th width="5%">Last Name</th>
				  <th width="10%">Contact Notes (Not Sales Notes)</th>				  
				  <th width="9%">Title</th>
				  <th width="8%">Email</th>
				  <th width="8%">Phone</th>
				  <th width="6%">Ext</th>
				  <th width="6%">Cell</th>
				  <th width="2%">Do Not Email</th>
				  <th width="8%">Address</th>
                  <th width="8%">City<br>State<br>Postal Code</th>
                  <th width="8%">Country</th>
                </tr>
              </thead>
             
              <tbody id="ajaxContainerContacts" class='searchable-contacts ajax-loading'></tbody>

		</table>
		

	</div>
</div>
<%'************************
' **** eof Contacts Tab****
'**************************
%>

<script>
	$(document).ready(function () { 
	
		ajaxLoadContacts(); 
						
	});
			
	function ajaxRowNewContacts() {
		var value = {};
		
		value.id = 0;				//id
		value.PrimaryContact = 0;	//primary
		value.DecisionMaker = 0;	//decisionmaker
		value.Suffix = "";			//suffix
		value.FirstName = "";		//firstname
		value.LastName = "";		//lastname
		value.Notes = "";			//notes
		value.ContactTitle = ""; 	//title DROPDOWN
		value.ContactTitleNumber = 0; 	//title DROPDOWN
		value.Email = "";			//email
		value.Phone = "";			//phone
		value.PhoneExt = "";		//extension
		value.Cell = "";			//cell
		value.DoNotEmail = 0;		//donotemail
		value.Fax = "";				//fax
		value.Address1 = "";		//Address1
		value.Address2 = "";		//Address2
		value.City = "";			//City
		value.State = "";			//State
		value.PostalCode = "";		//Postal
		value.Country = "";			//Country
		
		
		$('#ajaxRowContacts-' + 0 + '').remove();		
		$("#ajaxContainerContacts").prepend(ajaxRowHtmlContacts(value));
	}
	
	
	function ajaxRowHtmlContacts(value) {
		var htmlContacts = '\
			<tr id="ajaxRowContacts-' + value.id + '" class="' + (value.id==0?'ajaxRowEdit':'ajaxRowView') + '">\
				<td>\
					<div class="visibleRowView"><input type="checkbox" disabled="true" data-type="PrimaryContact" ' + (value.PrimaryContact==1?'checked':'') + ' id="chkPrimaryContact' + (value.id) + '" value="' + (value.id) + '"></div>\
					<div class="visibleRowEdit"><input type="checkbox" disabled="true" data-type="PrimaryContact" ' + (value.PrimaryContact==1?'checked':'') + ' id="chkPrimaryContact' + (value.id) + '" value="' + (value.id) + '"></div>\
				</td>\
				<td>\
					<div class="visibleRowView"><input type="checkbox" disabled="true" data-type="DecisionMaker" ' + (value.DecisionMaker==1?'checked':'') + ' id="chkDecisionMaker' + (value.id) + '" value="' + (value.id) + '"></div>\
					<div class="visibleRowEdit"><input type="checkbox" disabled="true" data-type="DecisionMaker" ' + (value.DecisionMaker==1?'checked':'') + ' id="chkDecisionMaker' + (value.id) + '" value="' + (value.id) + '"></div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.Suffix + '</div>\
					<div class="visibleRowEdit">' + value.Suffix + '</div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.FirstName + '</div>\
					<div class="visibleRowEdit">' + value.FirstName + '</div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.LastName + '</div>\
					<div class="visibleRowEdit">' + value.LastName + '</div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.Notes + '</div>\
					<div class="visibleRowEdit">' + value.Notes + '</div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.ContactTitle + '</div>\
					<div class="visibleRowEdit">' + value.ContactTitle + '</div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.Email + '</div>\
					<div class="visibleRowEdit">' + value.Email + '</div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.Phone + '</div>\
					<div class="visibleRowEdit">' + value.Phone + '</div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.PhoneExt + '</div>\
					<div class="visibleRowEdit">' + value.PhoneExt + '</div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.Cell + '</div>\
					<div class="visibleRowEdit">' + value.Cell + '</div>\
				</td>\
				<td>\
					<div class="visibleRowView"><input type="checkbox" disabled="true" data-type="DoNotEmail" ' + (value.DoNotEmail==1?'checked':'') + ' id="chkDoNotEmail' + (value.id) + '" value="' + (value.id) + '"></div>\
					<div class="visibleRowEdit"><input type="checkbox" disabled="true" data-type="DoNotEmail" ' + (value.DoNotEmail==1?'checked':'') + ' id="chkDoNotEmail' + (value.id) + '" value="' + (value.id) + '"></div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.Address1 + ', '+value.Address2 + '</div>\
					<div class="visibleRowEdit">' + value.Address1 + ', '+value.Address2 + '</div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.City + ', '+value.State  + ' '+value.PostalCode + '</div>\
					<div class="visibleRowEdit">' + value.City + ', '+value.State  + ' '+value.PostalCode + '</div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.Country + '</div>\
					<div class="visibleRowEdit">' + value.Country + '</div>\
				</td>\
		   </tr>\
			';
		return htmlContacts;
	}
	
	function ajaxLoadContacts(updateAction, updateActionId) {
		$("#ajaxContainerContacts").addClass("ajax-loading");
		var url = "ajax/pr_contacts.asp?i=<%= InternalRecordIdentifier %>";
		var jsondata = {};
		jsondata.updateAction = updateAction;
		jsondata.updateActionId = updateActionId;
				
		$.ajax({
			type: "POST",
			url: url,
			dataType: "json",
			data: jsondata,
			success: function (data) {				
				var htmlContacts = "";
				$.each(data, function (key, value) {
					htmlContacts += ajaxRowHtmlContacts(value);
				});
				$("#ajaxContainerContacts").html(htmlContacts);
				
				setTimeout(function(){
					$("#ajaxContainerContacts").removeClass("ajax-loading");
				}, 0);
				
			},
			failure: function (data) {
				$("#ajaxContainerContacts").html("Failed To Load Contacts");
				setTimeout(function(){
					$("#ajaxContainerContacts").removeClass("ajax-loading");
				}, 0);
				
			}
			
		});
	}
</script>

<%'**********************
' **** Contacts Tab *****
'************************
%>
<%

SQLContactStates = "SELECT StateCode,StateName FROM PR_States ORDER BY StateName ASC"
SQLContactCountries = "SELECT CountryCode,CountryName FROM PR_Countries ORDER BY CountryOrder ASC, CountryName ASC"

Set cnnContactStates = Server.CreateObject("ADODB.Connection")
cnnContactStates.open (Session("ClientCnnString"))

Set rsContactStates = Server.CreateObject("ADODB.Recordset")
Set rsContactStates = cnnContactStates.Execute(SQLContactStates)

ContactStates = ("[{""id"":"""",""title"":""Select""},")

If not rsContactStates.EOF Then
	sep = ""
	Do While Not rsContactStates.EOF
			ContactStates = ContactStates & (sep)
			sep = ","
			ContactStates = ContactStates & ("{")
			ContactStates = ContactStates & ("""id"":""" & Replace(rsContactStates("StateCode"), """", "\""") & """")
			ContactStates = ContactStates & (",""title"":""" & Replace(rsContactStates("StateName"), """", "\""") & """")
			ContactStates = ContactStates & ("}")
		rsContactStates.MoveNext						
	Loop
End If
ContactStates = ContactStates & ("]")
Set rsContactStates = Nothing


Set rsContactCountries = Server.CreateObject("ADODB.Recordset")
Set rsContactCountries = cnnContactStates.Execute(SQLContactCountries)

ContactCountries = ("[")
If not rsContactCountries.EOF Then
	sep = ""
	Do While Not rsContactCountries.EOF
			ContactCountries = ContactCountries & (sep)
			sep = ","
			ContactCountries = ContactCountries & ("{")
			ContactCountries = ContactCountries & ("""id"":""" & Replace(rsContactCountries("CountryCode"), """", "\""") & """")
			ContactCountries = ContactCountries & (",""title"":""" & Replace(rsContactCountries("CountryName"), """", "\""") & """")
			ContactCountries = ContactCountries & ("}")
		rsContactCountries.MoveNext						
	Loop
End If
ContactCountries = ContactCountries & ("]")
Set rsContactCountries = Nothing




SQLContactTitles = "SELECT *, InternalRecordIdentifier as id FROM PR_ContactTitles ORDER BY ContactTitle"

Set rsContactTitles = Server.CreateObject("ADODB.Recordset")
rsContactTitles.CursorLocation = 3 
Set rsContactTitles = cnnContactStates.Execute(SQLContactTitles)

'ContactTitles = ("[")
ContactTitles = ("[{""id"":""0"",""title"":""Select""},")
If userCanEditCRMOnTheFly(Session("UserNO")) = True AND GetCRMAddEditMenuPermissionLevel(Session("UserNO")) = vbTrue Then
ContactTitles  = ContactTitles & ("{""id"":""-1"",""title"":""Add a new Job Title""},")
End If
If not rsContactTitles.EOF Then
	sep = ""
	Do While Not rsContactTitles.EOF
			ContactTitles = ContactTitles & (sep)
			sep = ","
			ContactTitles = ContactTitles & ("{")
			ContactTitles = ContactTitles & ("""id"":""" & Replace(rsContactTitles("id"), """", "\""") & """")
			ContactTitles = ContactTitles & (",""title"":""" & Replace(rsContactTitles("ContactTitle"), """", "\""") & """")
			ContactTitles = ContactTitles & ("}")
		rsContactTitles.MoveNext						
	Loop
End If
ContactTitles = ContactTitles & ("]")
Set rsContactTitles = Nothing

cnnContactStates.Close
Set cnnContactStates = Nothing



%> 


<%'********************
' **** Contacts Tab****
'**********************
%>

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
	  
	<p> <button type="button" class="btn btn-success" onclick="ajaxRowNewContacts();">New Contact</button> </p>
			  
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
                  <th class="sorttable_nosort" width="7%">Actions</th>
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
	var ContactStates = <%= ContactStates %>;
	var ContactCountries = <%= ContactCountries %>;
	var ContactTitles = <%= ContactTitles %>;
	var SuffixTitles = [{"id":"","title":"Select"},{"id":"Mr.","title":"Mr."},{"id":"Mrs.","title":"Mrs."},{"id":"Miss","title":"Miss"},{"id":"Dr.","title":"Dr."},{"id":"Ms.","title":"Ms."}];
	
	
	var curcontactid =0;
	var value_default = {};
	var primarycontactfound = false;
	
	$(document).ready(function () { 

		ajaxLoadContacts();
		
	//contact title modal window submit
	$('#frmAddContactTitleTab').submit(function(e) {
		
		if ($('#frmAddContactTitleTab #txtContactTitleTab').val()==''){
			 swal("Contact title can not be blank.");
			return false;
		}
		
		$("#ONTHEFLYmodalContactTitleTab .btn-primary").html("Saving...");
        $.ajax({
            type: "POST",
            url: "onthefly_contacttitle_submit.asp",
            data: $('#frmAddContactTitleTab').serialize(),
            success: function(response) {				

				$.ajax({
					url: "onthefly_selectboxes.asp",
					data: {section: "txtTitleforTab"},
					dataType: "json",
					success: function(response2) {

						
						ContactTitles = response2;
						
						$('#ajaxRowContacts-' + curcontactid + ' [data-type="ContactTitleNumber"]').empty();
						
						var titlerow='';
												
						
						$.each(ContactTitles, function (key, ContactTitle) {
							if (ContactTitle.id== -1){
								titlerow ='<option value="'+ContactTitle.id+'"  style="font-weight:bold">'+ContactTitle.title+'</option>';
							} else {
								titlerow ='<option value="'+ContactTitle.id+'">'+ContactTitle.title+'</option>';
							}
							$('#ajaxRowContacts-' + curcontactid + ' [data-type="ContactTitleNumber"]').append(titlerow);
						});
						
						
						//ajaxLoadContacts();
						$("#ONTHEFLYmodalContactTitleTab .modal-body").hide();
						$("#ONTHEFLYmodalContactTitleTab .modal-footer").hide();
						$("#ONTHEFLYmodalContactTitleTab .modal-body2").html('Contact title added successfully<br><br><button type="button" class="btn btn-default" data-dismiss="modal" aria-label="Close">Close</button>');
						$("#ONTHEFLYmodalContactTitleTab .modal-body2").show();
						
					},
					error: function() {
						$("#ONTHEFLYmodalContactTitleTab .btn-primary").html("Save");
						//alert('Error add industry');
					}
				});
						
				
				//ajaxLoadContacts();
				//$("#ONTHEFLYmodalContactTitleTab .modal-body").html('Contact title added successfully<br><br><button type="button" class="btn btn-default" data-dismiss="modal" aria-label="Close">Close</button>');
				
            },
            error: function() {
				$("#ONTHEFLYmodalContactTitleTab .btn-primary").html("Save");
                //alert('Error add industry');
            }
        });
        return false;
    });			
					
	});
			
   function checkPrimary(el) {
		if ($('#chkPrimaryContact' + el).is(':checked')) {
			$('#chkPrimaryContact' + el).prop('checked', false);
		}  
    }		
   function checkDecisionMaker(el) {
		if ($('#chkDecisionMaker' + el).is(':checked')) {
			$('#chkDecisionMaker' + el).prop('checked', false);
		}  
    }	
    
   function checkDoNotEmail(el) {
		if ($('#chkDoNotEmail' + el).is(':checked')) {
			$('#chkDoNotEmail' + el).prop('checked', false);
		}  
    }	
    
    
	function ajaxRowNewContacts() {
		var value = {};
		
		value.id = 0;				//id
		value.PrimaryContact = 0;	//primary
		value.DecisionMaker = 0;	//decisionmaker
		value.Suffix = "";			//suffix
		value.SuffixTitleID = "";  // suffix id
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
		value.Address1 = "<%=Street%> <%=Suit%>";		//Address1
		value.Address2 = "";		//Address2
		value.City = "<%=City%>";			//City
		value.State = "<%=State%>";			//State
		value.PostalCode = "<%=PostalCode%>";		//Postal
		value.Country = "<%=Country%>";			//Country
		
		$('#ajaxRowContacts-' + 0 + '').remove();
		
		if (primarycontactfound){
			$("#ajaxContainerContacts").prepend(ajaxRowHtmlContacts(value_default));
		} else {
			$("#ajaxContainerContacts").prepend(ajaxRowHtmlContacts(value));
		}
		
	   var phonesTab = [{ "mask": "(###) ###-####"}];
	    $('#txtPhoneTab0').inputmask({ 
	        mask: phonesTab, 
	        greedy: false, 
	        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
	    $('#txtCellTab0').inputmask({ 
	        mask: phonesTab, 
	        greedy: false, 
	        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
	    $('#txtFaxTab0').inputmask({ 
	        mask: phonesTab, 
	        greedy: false, 
	        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
	
		
	}
	
	
	function ajaxRowHtmlContacts(value) {
		var ContactStateSelect = '<select class="form-control" data-type="State">';		
		$.each(ContactStates, function (key, ContactState) {
			ContactStateSelect +='<option value="'+ContactState.id+'" ' + (value.State+""==ContactState.id+""?'selected':'') + '>'+ContactState.title+'</option>';
		});		
		ContactStateSelect +='</select>';
		
		var ContactCountrySelect = '<select class="form-control" data-type="Country">';		
		$.each(ContactCountries, function (key, ContactCountry) {
			ContactCountrySelect +='<option value="'+ContactCountry.id+'" ' + (value.Country+""==ContactCountry.id+""?'selected':'') + '>'+ContactCountry.title+'</option>';
		});		
		ContactCountrySelect +='</select>';
		
		var ContactTitlesSelect = '<select class="form-control" data-type="ContactTitleNumber" onchange="contacttitlechanged(this.value,'+value.id+');">';
		$.each(ContactTitles, function (key, ContactTitle) {
			if (ContactTitle.id== -1){
				ContactTitlesSelect +='<option value="'+ContactTitle.id+'"  style="font-weight:bold" ' + (value.ContactTitleNumber+""==ContactTitle.id+""?'selected':'') + '>'+ContactTitle.title+'</option>';
			} else {
				ContactTitlesSelect +='<option value="'+ContactTitle.id+'" ' + (value.ContactTitleNumber+""==ContactTitle.id+""?'selected':'') + '>'+ContactTitle.title+'</option>';
			}
		});
		ContactTitlesSelect +='</select>';
		
		var SuffixTitlesSelect = '<select class="form-control" data-type="Suffix" >';
		$.each(SuffixTitles, function (key, SuffixTitle) {
			if (SuffixTitle.id== ""){
				SuffixTitlesSelect +='<option value="'+SuffixTitle.id+'"  style="font-weight:bold" ' + (value.Suffix+""==SuffixTitle.id+""?'selected':'') + '>'+SuffixTitle.title+'</option>';
			} else {
				SuffixTitlesSelect +='<option value="'+SuffixTitle.id+'" ' + (value.Suffix+""==SuffixTitle.id+""?'selected':'') + '>'+SuffixTitle.title+'</option>';
			}
		});
		SuffixTitlesSelect +='</select>';
		
		if (value.PrimaryContact==1){
			primarycontactfound = true;
			value_default.id = 0;				//id
			value_default.PrimaryContact = 0;	//primary
			value_default.DecisionMaker = 0;	//decisionmaker
			value_default.Suffix = "";			//suffix
			value_default.SuffixTitleID = "";  // suffix id
			value_default.FirstName = "";		//firstname
			value_default.LastName = "";		//lastname
			value_default.Notes = "";			//notes
			value_default.ContactTitle = ""; 	//title DROPDOWN
			value_default.ContactTitleNumber = 0; 	//title DROPDOWN
			value_default.Email = "";			//email
			value_default.Phone = "";			//phone
			value_default.PhoneExt = "";		//extension
			value_default.Cell = "";			//cell
			value_default.DoNotEmail = 0;		//donotemail
			value_default.Fax = "";				//fax
			value_default.Address1 = value.Address1;		//Address1
			value_default.Address2 = value.Address2;		//Address2
			value_default.City = value.City;			//City
			value_default.State = value.State;			//State
			value_default.PostalCode = value.PostalCode;		//Postal
			value_default.Country = value.Country;			//Country							
		}
		
		var btns = '\
					<div class="visibleRowView btn-group btn-group-sm"><a class="btn btn-primary" onclick="ajaxRowMode(\'Contacts\', ' + value.id + ', \'Edit\');"><i class="fa fa-edit"></i></a><a class="btn btn-danger" onclick="ajaxLoadContacts(\'delete\', ' + value.id + ');"><i class="fas fa-trash-alt"></i></a></div>\
					<div class="visibleRowEdit btn-group btn-group-sm"><a class="btn btn-success" onclick="ajaxLoadContacts(\'save\', ' + value.id + ');"><i class="fa fa-save"></i></a><a class="btn btn-default" id="btnEdit" onclick="ajaxRowMode(\'Contacts\', ' + value.id + ', \'View\');"><i class="fa fa-times"></i></a></div>\
				';
		if(value.id==0)
			btns = '\
					<div class="visibleRowEdit btn-group btn-group-sm"><a class="btn btn-success" onclick="ajaxLoadContacts(\'insert\', ' + value.id + ');"><i class="fa fa-save"></i></a><a class="btn btn-default" onclick="ajaxRowMode(\'Contacts\', ' + value.id + ', \'View\');"><i class="fa fa-times"></i></a></div>\
				';	
		var htmlContacts = '\
			<tr id="ajaxRowContacts-' + value.id + '" class="' + (value.id==0?'ajaxRowEdit':'ajaxRowView') + '">\
				<td>\
					<div class="visibleRowView"><input type="checkbox" disabled="true" data-type="PrimaryContact" ' + (value.PrimaryContact==1?'checked':'') + ' id="chkPrimaryContact' + (value.id) + '" value="' + (value.id) + '"></div>\
					<div class="visibleRowEdit"><input type="checkbox" onClick="checkPrimary(' + value.id + ');" data-type="PrimaryContact" ' + (value.PrimaryContact==1?'checked':'') + ' id="chkPrimaryContact' + (value.id) + '" value="' + (value.id) + '"></div>\
				</td>\
				<td>\
					<div class="visibleRowView"><input type="checkbox" disabled="true" data-type="DecisionMaker" ' + (value.DecisionMaker==1?'checked':'') + ' id="chkDecisionMaker' + (value.id) + '" value="' + (value.id) + '"></div>\
					<div class="visibleRowEdit"><input type="checkbox" onClick="checkDecisionMaker(' + value.id + ');" data-type="DecisionMaker" ' + (value.DecisionMaker==1?'checked':'') + ' id="chkDecisionMaker' + (value.id) + '" value="' + (value.id) + '"></div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.Suffix + '</div>\
					<div class="visibleRowEdit">'+ SuffixTitlesSelect +'</div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.FirstName + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="FirstName" value="' + value.FirstName + '" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.LastName + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="LastName" value="' + value.LastName + '" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.Notes + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="Notes" value="' + value.Notes.replace(/"/g, '&quot;') + '" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.ContactTitle + '</div>\
					<div class="visibleRowEdit">'+ ContactTitlesSelect +'</div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.Email + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="Email" value="' + value.Email + '" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.Phone + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="Phone" id="txtPhoneTab' + value.id + '" value="' + value.Phone + '" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.PhoneExt + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="PhoneExt" value="' + value.PhoneExt + '" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.Cell + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="Cell" id="txtCellTab' + value.id + '" value="' + value.Cell + '" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView"><input type="checkbox" disabled="true" data-type="DoNotEmail" ' + (value.DoNotEmail==1?'checked':'') + ' id="chkDoNotEmail' + (value.id) + '" value="' + (value.id) + '"></div>\
					<div class="visibleRowEdit"><input type="checkbox" onClick="checkDoNotEmail(' + value.id + ');" data-type="DoNotEmail" ' + (value.DoNotEmail==1?'checked':'') + ' id="chkDoNotEmail' + (value.id) + '" value="' + (value.id) + '"></div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.Address1 + ', '+value.Address2 + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="Address1" id="txtAddress1Tab' + value.id + '" value="' + value.Address1 + '" placeholder="Address1" />\
					<input class="form-control" data-type="Address2" id="txtAddress2Tab' + value.id + '" value="' + value.Address2 + '"  placeholder="Address2" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.City + ', '+value.State  + ' '+value.PostalCode + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="City" id="txtCityTab' + value.id + '" value="' + value.City + '"  placeholder="City" />'+ContactStateSelect+'<input class="form-control" data-type="PostalCode" id="txtPostalTab' + value.id + '" value="' + value.PostalCode + '"  placeholder="Postal" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.Country + '</div>\
					<div class="visibleRowEdit">' + ContactCountrySelect + '</div>\
				</td>\
				<td class="text-center">'+btns+'</td>\
		   </tr>\
			';
		return htmlContacts;
		//<div class="visibleRowEdit"><input class="form-control" data-type="Country" id="txtCountryTab' + value.id + '" value="' + value.Country + '"  placeholder="Country" /></div>\
	}
	function contacttitlechanged(val,val_id){
		if (val== -1){
			
			curcontactid = val_id;
			//deselect add new row			
			$('#ajaxRowContacts-' + val_id + ' [data-type="ContactTitleNumber"] option[selected="selected"]').each(
				function() {
					$(this).removeAttr('selected');
				}
			);

			// mark the first option as selected
			$('#ajaxRowContacts-' + val_id + ' [data-type="ContactTitleNumber"] option:first').attr('selected','selected');
			
			//show modal
			$('#frmAddContactTitleTab #txtContactTitleTab').val('');
			$("#ONTHEFLYmodalContactTitleTab .btn-primary").html("Save");
			$("#ONTHEFLYmodalContactTitleTab .modal-footer").show();
			$("#ONTHEFLYmodalContactTitleTab .modal-body2").hide();
			$("#ONTHEFLYmodalContactTitleTab .modal-body").show();
			$('#ONTHEFLYmodalContactTitleTab').modal('show');
			
		}
	}
	
	function ajaxLoadContacts(updateAction, updateActionId) {
		if (updateAction=="save" || updateAction=="insert"){
			var contitle = $('#ajaxRowContacts-' + updateActionId + ' [data-type="ContactTitleNumber"]').val();
			if (contitle==0 || contitle== -1){
				swal("Please select contact title");
				return false;	
			}
		}
		if (updateAction == "delete" && !confirm("Are your sure you want to delete this contact?")) return;
		$("#ajaxContainerContacts").addClass("ajax-loading");
		var url = "ajax/pr_contacts.asp?i=<%= InternalRecordIdentifier %>";
		var jsondata = {};
		jsondata.updateAction = updateAction;
		jsondata.updateActionId = updateActionId;
		

		if(updateAction=="save" || updateAction=="insert"){
			jsondata.PrimaryContact	= $('#ajaxRowContacts-' + updateActionId + ' [data-type="PrimaryContact"]').is(':checked')?1:0;
			jsondata.DecisionMaker	= $('#ajaxRowContacts-' + updateActionId + ' [data-type="DecisionMaker"]').is(':checked')?1:0;
			jsondata.Suffix			= $('#ajaxRowContacts-' + updateActionId + ' [data-type="Suffix"]').val();
			jsondata.FirstName		= $('#ajaxRowContacts-' + updateActionId + ' [data-type="FirstName"]').val();
			jsondata.LastName		= $('#ajaxRowContacts-' + updateActionId + ' [data-type="LastName"]').val();
			jsondata.Notes			= $('#ajaxRowContacts-' + updateActionId + ' [data-type="Notes"]').val();
			jsondata.ContactTitleNumber = $('#ajaxRowContacts-' + updateActionId + ' [data-type="ContactTitleNumber"]').val();
			jsondata.Email			= $('#ajaxRowContacts-' + updateActionId + ' [data-type="Email"]').val();
			jsondata.Phone			= $('#ajaxRowContacts-' + updateActionId + ' [data-type="Phone"]').val();
			jsondata.PhoneExt		= $('#ajaxRowContacts-' + updateActionId + ' [data-type="PhoneExt"]').val();
			jsondata.Cell			= $('#ajaxRowContacts-' + updateActionId + ' [data-type="Cell"]').val();
			jsondata.DoNotEmail		= $('#ajaxRowContacts-' + updateActionId + ' [data-type="DoNotEmail"]').is(':checked')?1:0;
			jsondata.Fax 			= $('#ajaxRowContacts-' + updateActionId + ' [data-type="Fax"]').val();
			jsondata.Address1 		= $('#ajaxRowContacts-' + updateActionId + ' [data-type="Address1"]').val();
			jsondata.Address2 		= $('#ajaxRowContacts-' + updateActionId + ' [data-type="Address2"]').val();
			jsondata.City 			= $('#ajaxRowContacts-' + updateActionId + ' [data-type="City"]').val();
			jsondata.State 			= $('#ajaxRowContacts-' + updateActionId + ' [data-type="State"]').val();
			jsondata.PostalCode		= $('#ajaxRowContacts-' + updateActionId + ' [data-type="PostalCode"]').val();
			jsondata.Country 		= $('#ajaxRowContacts-' + updateActionId + ' [data-type="Country"]').val();
		}
		
		$.ajax({
			type: "POST",
			url: url,
			dataType: "json",
			data: jsondata,
			success: function (data) {
				//if (!data || data+""=="") { window.location.href = window.location.href + ""; return; }				
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


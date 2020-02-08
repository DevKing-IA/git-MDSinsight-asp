<%'*************************
' **** Competitors Tab *****
'***************************
%>
<%
'SQLCompetitorNames = "SELECT * FROM PR_Competitors WHERE InternalRecordIdentifier NOT IN "
'SQLCompetitorNames = SQLCompetitorNames & "(SELECT CompetitorRecID FROM PR_ProspectCompetitors WHERE ProspectRecID = " & InternalRecordIdentifier  & ") "

SQLCompetitorNames = "SELECT * FROM PR_Competitors ORDER BY CompetitorName"
Set cnnCompetitorNames = Server.CreateObject("ADODB.Connection")
cnnCompetitorNames.open (Session("ClientCnnString"))
Set rsCompetitorNames = Server.CreateObject("ADODB.Recordset")
rsCompetitorNames.CursorLocation = 3 
Set rsCompetitorNames = cnnCompetitorNames.Execute(SQLCompetitorNames)

CompetitorNames = ("[")
If not rsCompetitorNames.EOF Then
	sep = ""
	Do While Not rsCompetitorNames.EOF
			CompetitorNames = CompetitorNames & (sep)
			sep = ","
			CompetitorNames = CompetitorNames & ("{")
			CompetitorNames = CompetitorNames & ("""id"":""" & Replace(rsCompetitorNames("InternalRecordIdentifier"), """", "\""") & """")
			CompetitorNames = CompetitorNames & (",""name"":""" & Replace(rsCompetitorNames("CompetitorName"), """", "\""") & """")
			CompetitorNames = CompetitorNames & ("}")
		rsCompetitorNames.MoveNext						
	Loop
End If
CompetitorNames = CompetitorNames & ("]")
Set rsCompetitorNames = Nothing
cnnCompetitorNames.Close
Set cnnCompetitorNames = Nothing


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
<div role="tabpanel" class="tab-pane fade in" id="competitors">
	  
	<p> <button type="button" class="btn btn-success" onclick="ajaxRowNewCompetitors();">Add Another Competitior</button> </p>
	
	<div id="ajaxContainerCompetitorsNoPrimary"></div>

		<div class="input-group narrow-results"> <span class="input-group-addon">Narrow Results</span>
			<input id="filter-competitors" type="text" class="form-control filter-search-width" placeholder="Type here...">
		</div>
	  
	  <div class="table-responsive">
            <table  class="table table-striped table-condensed table-hover table-bordered sortable standard-font">
              <thead>
                <tr>
                  <th width="5%">Primary</th>
                  <th width="10%">Competitor Name</th>
                  <th>Address Information</th>
                  <th width="30%">Notes</th>
                  <th width="5%">Bottled Water</th>
                  <th width="5%">Filtered Water</th>
				  <th width="5%">OCS</th>
				  <th width="5%">OCS<br>Supply</th>
				  <th width="5%">Office<br>Supplies</th>
				  <th width="5%">Vending</th>
				  <th width="5%">Micro<br>Market</th>				  
				  <th width="5%">Pantry</th>
                  <th class="sorttable_nosort" width="7%">Actions</th>
                </tr>
              </thead>
             
              <tbody id="ajaxContainerCompetitors" class='searchable-competitors ajax-loading'></tbody>

		</table>
		

	</div>
</div>

<script>

	var CompetitorNames = <%= CompetitorNames %>;

	$(document).ready(function () { 
	
			ajaxLoadCompetitors(); 
			
	});
	
	
	
   function checkPrimaryCompetitor(el) {
		if ($('#chkPrimaryCompetitor' + el).is(':checked')) {
			$('#chkPrimaryCompetitor' + el).prop('checked', false);
		}  
    }	
   function checkBottledWater(el) {
		if ($('#chkBottledWater' + el).is(':checked')) {
			$('#chkBottledWater' + el).prop('checked', false);
		}  
    }		
   function checkFilteredWater(el) {
		if ($('#chkFilteredWater' + el).is(':checked')) {
			$('#chkFilteredWater' + el).prop('checked', false);
		}  
    }	
   function checkOCS(el) {
		if ($('#chkOCS' + el).is(':checked')) {
			$('#chkOCS' + el).prop('checked', false);
		}  
    }	
   function checkOCS_Supply(el) {
		if ($('#chkOCS_Supply' + el).is(':checked')) {
			$('#chkOCS_Supply' + el).prop('checked', false);
		}  
    }	
   function checkOfficeSupplies(el) {
		if ($('#chkOfficeSupplies' + el).is(':checked')) {
			$('#chkOfficeSupplies' + el).prop('checked', false);
		}  
    }	
   function checkVending(el) {
		if ($('#chkVending' + el).is(':checked')) {
			$('#chkVending' + el).prop('checked', false);
		}  
    }	
   function checkMicroMarket(el) {
		if ($('#chkMicroMarket' + el).is(':checked')) {
			$('#chkMicroMarket' + el).prop('checked', false);
		}  
    }	
   function checkPantry(el) {
		if ($('#chkPantry' + el).is(':checked')) {
			$('#chkPantry' + el).prop('checked', false);
		}  
    }	
	
 
	function ajaxRowNewCompetitors() {
		var value = {};
		value.CompetitorRecID = 0;
		value.PrimaryCompetitor = 0;
		value.Notes = "";
		value.BottledWater = 0;
		value.FilteredWater = 0;
		value.OCS = 0;
		value.OCS_Supply = 0;
		value.OfficeSupplies = 0;
		value.Vending = 0;
		value.MicroMarket = 0;
		value.Pantry = 0;
		value.CompInternalRecordIdentifier = 0;
		value.CompetitorName = "";
		value.AddressInformation = "";		
		$('#ajaxRowCompetitors-' + 0 + '').remove();		
		$("#ajaxContainerCompetitors").prepend(ajaxRowHtmlCompetitors(value));
	}
	
	
	function ajaxRowHtmlCompetitors(value) {
		var CompetitorNamesSelect = '<select class="form-control" data-type="CompetitorRecID">';
		$.each(CompetitorNames, function (key, CompetitorName) {
			CompetitorNamesSelect+='<option value="'+CompetitorName.id+'" ' + (value.CompetitorRecID +""==CompetitorName.id+""?'selected':'') + '>'+CompetitorName.name+'</option>';
		});
		CompetitorNamesSelect+='</select>';
		var btns = '\
					<div class="visibleRowView btn-group btn-group-sm"><a class="btn btn-primary" onclick="ajaxRowMode(\'Competitors\', ' + value.CompetitorRecID + ', \'Edit\');"><i class="fa fa-edit"></i></a><a class="btn btn-danger" onclick="ajaxLoadCompetitors(\'delete\', ' + value.CompetitorRecID + ');"><i class="fas fa-trash-alt"></i></a></div>\
					<div class="visibleRowEdit btn-group btn-group-sm"><a class="btn btn-success" onclick="ajaxLoadCompetitors(\'save\', ' + value.CompetitorRecID + ');"><i class="fa fa-save"></i></a><a class="btn btn-default" id="btnEdit" onclick="ajaxRowMode(\'Competitors\', ' + value.CompetitorRecID + ', \'View\');"><i class="fa fa-times"></i></a></div>\
				';
		if(value.CompetitorRecID==0)
			btns = '\
					<div class="visibleRowEdit btn-group btn-group-sm"><a class="btn btn-success" onclick="ajaxLoadCompetitors(\'insert\', ' + value.CompetitorRecID + ');"><i class="fa fa-save"></i></a><a class="btn btn-default" onclick="ajaxRowMode(\'Competitors\', ' + value.CompetitorRecID + ', \'View\');"><i class="fa fa-times"></i></a></div>\
				';			
		var htmlCompetitors = '\
			<tr id="ajaxRowCompetitors-' + value.CompetitorRecID + '" class="' + (value.CompetitorRecID==0?'ajaxRowEdit':'ajaxRowView') + '">\
				<td>\
					<div class="visibleRowView"><input type="checkbox" disabled="true" data-type="PrimaryCompetitor" ' + (value.PrimaryCompetitor==1?'checked':'') + ' id="chkPrimaryCompetitor' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
					<div class="visibleRowEdit"><input type="checkbox" onClick="checkPrimaryCompetitor(' + value.CompetitorRecID + ');" data-type="PrimaryCompetitor" ' + (value.PrimaryCompetitor==1?'checked':'') + ' id="chkPrimaryCompetitor' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.CompetitorName + '</div>\
					<div class="visibleRowEdit">'+ CompetitorNamesSelect +'</div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.AddressInformation + '</div>\
					<div class="visibleRowEdit">---</div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.Notes + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="Notes" value="' + value.Notes.replace(/"/g, '&quot;') + '" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView"><input type="checkbox" disabled="true" data-type="BottledWater" ' + (value.BottledWater==1?'checked':'') + ' id="chkBottledWater' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
					<div class="visibleRowEdit"><input type="checkbox" onClick="checkBottledWater(' + value.CompetitorRecID + ');" data-type="BottledWater" ' + (value.BottledWater==1?'checked':'') + ' id="chkBottledWater' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
				</td>\
				<td>\
					<div class="visibleRowView"><input type="checkbox" disabled="true" data-type="FilteredWater" ' + (value.FilteredWater==1?'checked':'') + ' id="chkFilteredWater' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
					<div class="visibleRowEdit"><input type="checkbox" onClick="checkFilteredWater(' + value.CompetitorRecID + ');" data-type="FilteredWater" ' + (value.FilteredWater==1?'checked':'') + ' id="chkFilteredWater' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
				</td>\
				<td>\
					<div class="visibleRowView"><input type="checkbox" disabled="true" data-type="OCS" ' + (value.OCS==1?'checked':'') + ' id="chkOCS' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
					<div class="visibleRowEdit"><input type="checkbox" onClick="checkOCS(' + value.CompetitorRecID + ');" data-type="OCS" ' + (value.OCS==1?'checked':'') + ' id="chkOCS' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
				</td>\
				<td>\
					<div class="visibleRowView"><input type="checkbox" disabled="true" data-type="OCS_Supply" ' + (value.OCS_Supply==1?'checked':'') + ' id="chkOCS_Supply' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
					<div class="visibleRowEdit"><input type="checkbox" onClick="checkOCS(' + value.CompetitorRecID + ');" data-type="OCS_Supply" ' + (value.OCS_Supply==1?'checked':'') + ' id="chkOCS_Supply' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
				</td>\
				<td>\
					<div class="visibleRowView"><input type="checkbox" disabled="true" data-type="OfficeSupplies" ' + (value.OfficeSupplies==1?'checked':'') + ' id="chkOfficeSupplies' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
					<div class="visibleRowEdit"><input type="checkbox" onClick="checkOfficeSupplies(' + value.CompetitorRecID + ');" data-type="OfficeSupplies" ' + (value.OfficeSupplies==1?'checked':'') + ' id="chkOfficeSupplies' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
				</td>\
				<td>\
					<div class="visibleRowView"><input type="checkbox" disabled="true" data-type="Vending" ' + (value.Vending==1?'checked':'') + ' id="chkVending' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
					<div class="visibleRowEdit"><input type="checkbox" onClick="checkVending(' + value.CompetitorRecID + ');" data-type="Vending" ' + (value.Vending==1?'checked':'') + ' id="chkOfficeSupplies' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
				</td>\
				<td>\
					<div class="visibleRowView"><input type="checkbox" disabled="true" data-type="MicroMarket" ' + (value.MicroMarket==1?'checked':'') + ' id="chkMicroMarket' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
					<div class="visibleRowEdit"><input type="checkbox" onClick="checkMicroMarket(' + value.CompetitorRecID + ');" data-type="MicroMarket" ' + (value.MicroMarket==1?'checked':'') + ' id="chkMicroMarket' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
				</td>\
				<td>\
					<div class="visibleRowView"><input type="checkbox" disabled="true" data-type="Pantry" ' + (value.Pantry==1?'checked':'') + ' id="chkPantry' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
					<div class="visibleRowEdit"><input type="checkbox" onClick="checkPantry(' + value.CompetitorRecID + ');" data-type="Pantry" ' + (value.Pantry==1?'checked':'') + ' id="chkPantry' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
				</td>\
				<td class="text-center">'+btns+'</td>\
		   </tr>\
			';
		return htmlCompetitors;
	}
	function ajaxLoadCompetitors(updateAction, updateActionId) {
		if (updateAction == "delete" && !confirm("Are your sure you want to delete this competitor?")) return;
		$("#ajaxContainerCompetitors").addClass("ajax-loading");
		var url = "ajax/pr_competitors.asp?i=<%= InternalRecordIdentifier %>";
		var jsondata = {};
		jsondata.updateAction = updateAction;
		jsondata.updateActionId = updateActionId;

		if(updateAction=="save" || updateAction=="insert"){
			jsondata.CompetitorRecID	= $('#ajaxRowCompetitors-' + updateActionId + ' [data-type="CompetitorRecID"]').val();
			jsondata.PrimaryCompetitor	= $('#ajaxRowCompetitors-' + updateActionId + ' [data-type="PrimaryCompetitor"]').is(':checked')?1:0;
			jsondata.Notes				= $('#ajaxRowCompetitors-' + updateActionId + ' [data-type="Notes"]').val();
			jsondata.BottledWater		= $('#ajaxRowCompetitors-' + updateActionId + ' [data-type="BottledWater"]').is(':checked')?1:0;
			jsondata.FilteredWater		= $('#ajaxRowCompetitors-' + updateActionId + ' [data-type="FilteredWater"]').is(':checked')?1:0;
			jsondata.OCS				= $('#ajaxRowCompetitors-' + updateActionId + ' [data-type="OCS"]').is(':checked')?1:0;
			jsondata.OCS_Supply			= $('#ajaxRowCompetitors-' + updateActionId + ' [data-type="OCS_Supply"]').is(':checked')?1:0;
			jsondata.OfficeSupplies		= $('#ajaxRowCompetitors-' + updateActionId + ' [data-type="OfficeSupplies"]').is(':checked')?1:0;
			jsondata.Vending			= $('#ajaxRowCompetitors-' + updateActionId + ' [data-type="Vending"]').is(':checked')?1:0;
			jsondata.MicroMarket		= $('#ajaxRowCompetitors-' + updateActionId + ' [data-type="MicroMarket"]').is(':checked')?1:0;
			jsondata.Pantry				= $('#ajaxRowCompetitors-' + updateActionId + ' [data-type="Pantry"]').is(':checked')?1:0;
		}
		
		$.ajax({
			type: "POST",
			url: url,
			dataType: "json",
			data: jsondata,
			success: function (data) {
				//if (!data || data+""=="") { window.location.href = window.location.href + ""; return; }				
				var htmlCompetitors = "";
				$.each(data, function (key, value) {
				
					//alert(key + " : " + value.PrimaryCompetitor);
					
					if ((key == 0) && (value.PrimaryCompetitor !== 1)) 
					{
						$("#ajaxContainerCompetitorsNoPrimary").html("<h3><font color='#ff0000'>You must select a primary competitor.</font></h3>");
					}
					
					if ((key == 0) && (value.PrimaryCompetitor == 1)) 
					{
						$("#ajaxContainerCompetitorsNoPrimary").html("");
					}
					
					htmlCompetitors += ajaxRowHtmlCompetitors(value);
				});
				$("#ajaxContainerCompetitors").html(htmlCompetitors);
				
				
				
				setTimeout(function(){
					$("#ajaxContainerCompetitors").removeClass("ajax-loading");
				}, 0);
				
			},
			failure: function (data) {
				$("#ajaxContainerCompetitors").html("Failed To Load Competitors");
				setTimeout(function(){
					$("#ajaxContainerCompetitors").removeClass("ajax-loading");
				}, 0);
				
			}
			
		});
	}
</script>


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
                </tr>
              </thead>
             
              <tbody id="ajaxContainerCompetitors" class='searchable-competitors ajax-loading'></tbody>

		</table>
		

	</div>
</div>

<script>

	$(document).ready(function () { 
	
			ajaxLoadCompetitors(); 
			
	});
	
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
		var htmlCompetitors = '\
			<tr id="ajaxRowCompetitors-' + value.CompetitorRecID + '" class="' + (value.CompetitorRecID==0?'ajaxRowEdit':'ajaxRowView') + '">\
				<td>\
					<div class="visibleRowView"><input type="checkbox" disabled="true" data-type="PrimaryCompetitor" ' + (value.PrimaryCompetitor==1?'checked':'') + ' id="chkPrimaryCompetitor' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
					<div class="visibleRowEdit"><input type="checkbox" disabled="true" data-type="PrimaryCompetitor" ' + (value.PrimaryCompetitor==1?'checked':'') + ' id="chkPrimaryCompetitor' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.CompetitorName + '</div>\
					<div class="visibleRowEdit">' + value.CompetitorName + '</div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.AddressInformation + '</div>\
					<div class="visibleRowEdit">' + value.AddressInformation + '</div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.Notes + '</div>\
					<div class="visibleRowEdit">' + value.Notes + '</div>\
				</td>\
				<td>\
					<div class="visibleRowView"><input type="checkbox" disabled="true" data-type="BottledWater" ' + (value.BottledWater==1?'checked':'') + ' id="chkBottledWater' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
					<div class="visibleRowEdit"><input type="checkbox" disabled="true" data-type="BottledWater" ' + (value.BottledWater==1?'checked':'') + ' id="chkBottledWater' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
				</td>\
				<td>\
					<div class="visibleRowView"><input type="checkbox" disabled="true" data-type="FilteredWater" ' + (value.FilteredWater==1?'checked':'') + ' id="chkFilteredWater' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
					<div class="visibleRowEdit"><input type="checkbox" disabled="true" data-type="FilteredWater" ' + (value.FilteredWater==1?'checked':'') + ' id="chkFilteredWater' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
				</td>\
				<td>\
					<div class="visibleRowView"><input type="checkbox" disabled="true" data-type="OCS" ' + (value.OCS==1?'checked':'') + ' id="chkOCS' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
					<div class="visibleRowEdit"><input type="checkbox" disabled="true" data-type="OCS" ' + (value.OCS==1?'checked':'') + ' id="chkOCS' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
				</td>\
				<td>\
					<div class="visibleRowView"><input type="checkbox" disabled="true" data-type="OCS_Supply" ' + (value.OCS_Supply==1?'checked':'') + ' id="chkOCS_Supply' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
					<div class="visibleRowEdit"><input type="checkbox" disabled="true" data-type="OCS_Supply" ' + (value.OCS_Supply==1?'checked':'') + ' id="chkOCS_Supply' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
				</td>\
				<td>\
					<div class="visibleRowView"><input type="checkbox" disabled="true" data-type="OfficeSupplies" ' + (value.OfficeSupplies==1?'checked':'') + ' id="chkOfficeSupplies' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
					<div class="visibleRowEdit"><input type="checkbox" disabled="true" data-type="OfficeSupplies" ' + (value.OfficeSupplies==1?'checked':'') + ' id="chkOfficeSupplies' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
				</td>\
				<td>\
					<div class="visibleRowView"><input type="checkbox" disabled="true" data-type="Vending" ' + (value.Vending==1?'checked':'') + ' id="chkVending' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
					<div class="visibleRowEdit"><input type="checkbox" disabled="true" data-type="Vending" ' + (value.Vending==1?'checked':'') + ' id="chkVending' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
				</td>\
				<td>\
					<div class="visibleRowView"><input type="checkbox" disabled="true" data-type="MicroMarket" ' + (value.MicroMarket==1?'checked':'') + ' id="chkMicroMarket' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
					<div class="visibleRowEdit"><input type="checkbox" disabled="true" data-type="MicroMarket" ' + (value.MicroMarket==1?'checked':'') + ' id="chkMicroMarket' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
				</td>\
				<td>\
					<div class="visibleRowView"><input type="checkbox" disabled="true" data-type="Pantry" ' + (value.Pantry==1?'checked':'') + ' id="chkPantry' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
					<div class="visibleRowEdit"><input type="checkbox" disabled="true" data-type="Pantry" ' + (value.Pantry==1?'checked':'') + ' id="chkPantry' + (value.CompetitorRecID) + '" value="' + (value.CompetitorRecID) + '"></div>\
				</td>\
		   </tr>\
			';
		return htmlCompetitors;
	}
	function ajaxLoadCompetitors(updateAction, updateActionId) {

		$("#ajaxContainerCompetitors").addClass("ajax-loading");
		var url = "ajax/pr_competitors.asp?i=<%= InternalRecordIdentifier %>";
		var jsondata = {};
		jsondata.updateAction = updateAction;
		jsondata.updateActionId = updateActionId;
		
		$.ajax({
			type: "POST",
			url: url,
			dataType: "json",
			data: jsondata,
			success: function (data) {			
				var htmlCompetitors = "";
				$.each(data, function (key, value) {
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

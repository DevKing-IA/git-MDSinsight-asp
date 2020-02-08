<script LANGUAGE="JavaScript">
<!--
	function showHide(){ 
		//create an object reference to the div containing images 
		var oimageDiv = document.getElementById('searchingimageDiv');
		//set display to inline if currently none, otherwise to none 
		oimageDiv.style.display=(oimageDiv.style.display=='none')?'inline':'none'; 
		
		var cancelFiltersBtn = document.getElementById('cancelFiltersBtn');
		var saveFiltersBtn = document.getElementById('saveFiltersBtn');
		cancelFiltersBtn.style.display = 'none';
		saveFiltersBtn.style.display = 'none';
		
	} 

	function getCheckedRadioValue(radio_group) {
	    for (var i = 0; i < radio_group.length; i++) {
	        var button = radio_group[i];
	        if (button.checked) {
	            return button;
	        }
	    }
	    return undefined;
	}
	    	
	
	function isOneStageUnqualifiedChecked() {
	    return ($('[name="chkStageUnqualified"]:checked').length > 0);
	}
	
	function isOneStageLostChecked() {
	    return ($('[name="chkStageLost"]:checked').length > 0);
	}

	
	function isInteger(str) {
    	var r = /^-?[0-9]*[1-9][0-9]*$/;
    	return r.test(str);
	}
	
	

	function validateCustomizeForm()	{
		
		//***************************************
		//These are the checks for the stages
		//***************************************
    	//If they chose the option button for stage date filters
		    	
    	 

		if ($('#optStageUnqualifiedDateRange').is(':checked'))
		{
			if ($("#selUnqualifiedStageDateRangeCustom").find(":selected").val() == "") 
			{
				//Date Range Checking Here
				if (document.frmProspectingCustomizeDataFilters.txtStageUnqualifiedRangeStartDate.value == "" || 
					document.frmProspectingCustomizeDataFilters.txtStageUnqualifiedRangeEndDate.value == "") 
					{		
						swal("Please make sure both UNQUALIFIED filter dates are filled in or quick pick range has been selected.")
						return false;
					}		
			}
			else
			{
				//swal("Please select at least one stage for stage date filtering.")
				//return false;
			}
   		}



		if ($('#optStageLostDateRange').is(':checked'))
		{
			if ($("#selLostStageDateRangeCustom").find(":selected").val() == "") 
			{
				//Date Range Checking Here
				if (document.frmProspectingCustomizeDataFilters.txtStageLostDateRangeStartDate.value == "" || 
					document.frmProspectingCustomizeDataFilters.txtStageLostDateRangeEndDate.value == "") 
					{		
						swal("Please make sure both LOST filter dates are filled in or quick pick range has been selected.")
						return false;
					}		
			}
			else
			{
				//swal("Please select at least one stage for stage date filtering.")
				//return false;
			}
   		}




		if ($('#optProspectCreatedDateRange').is(':checked'))
		{
			if ($("#selProspectCreatedDateRangeCustom").find(":selected").val() == "") 
			{
				//Date Range Checking Here
				if (document.frmProspectingCustomizeDataFilters.txtProspectCreatedRangeStartDate.value == "" || 
					document.frmProspectingCustomizeDataFilters.txtProspectCreatedRangeEndDate.value == "") 
					{		
						swal("Please make sure both CREATED DATE filter dates are filled in or quick pick range has been selected.")
						return false;
					}		
			}
			else
			{
				//swal("Please select at least one stage for stage date filtering.")
				//return false;
			}
   		}

		//***************************************
		//These are the checks for the employee ranges
		//***************************************
    	//If they chose the option button for employee number filters
    	
	    var radiosNumEmployees = document.getElementsByName("optNumEmployeesRangeCompare");
	    var empRadioSelected = false;
	
	    var i = 0;
	    while (!empRadioSelected && i < radiosNumEmployees.length) {
	        if (radiosNumEmployees[i].checked) empRadioSelected = true;
	        i++;        
	    }
	
    	
    	if (empRadioSelected)
		{
		
			var checkedEmployeeRadioButtonValue = getCheckedRadioValue(document.frmProspectingCustomizeDataFilters.optNumEmployeesRangeCompare);
		
			if (checkedEmployeeRadioButtonValue.value == "ByCustomNumber") 
			{
		
				if (document.frmProspectingCustomizeDataFilters.txtEmployeeRangeComparisonNumberSingle.value == "") 
					{		
						swal("Please specify number of employees to filter by.")
						return false;
					}
				if (!isInteger(document.frmProspectingCustomizeDataFilters.txtEmployeeRangeComparisonNumberSingle.value))
					{		
						swal("Please enter a whole number for the number of employees.")
						return false;
					}

			}
			
			if (checkedEmployeeRadioButtonValue.value == "ByCustomRange") 
			
			{
		
				if (document.frmProspectingCustomizeDataFilters.txtEmployeeCustomRangeNumber1.value == "") 
					{		
						swal("Please specify a starting range number for number of employees to filter by.")
						return false;
					}
				if (!isInteger(document.frmProspectingCustomizeDataFilters.txtEmployeeCustomRangeNumber1.value))
					{		
						swal("Please enter a whole number for the starting number of employees.")
						return false;
					}
				
				
				if (document.frmProspectingCustomizeDataFilters.txtEmployeeCustomRangeNumber2.value == "") 
					{		
						swal("Please specify an ending range number for number of employees to filter by.")
						return false;
					}
				if (!isInteger(document.frmProspectingCustomizeDataFilters.txtEmployeeCustomRangeNumber2.value))
					{		
						swal("Please enter a whole number for the ending number of employees.")
						return false;
					}				
				
			}
			
			
			
   		}





		//***************************************
		//These are the checks for the pantry ranges
		//***************************************
    	//If they chose the option button for pantry number filters
 

	    var radiosNumPantries = document.getElementsByName("optNumPantriesCompare");
	    var pantryRadioSelected = false;
	
	    var x = 0;
	    while (!pantryRadioSelected && x < radiosNumPantries.length) {
	        if (radiosNumPantries[x].checked) pantryRadioSelected = true;
	        x++;        
	    }
	
    	if (pantryRadioSelected)
		{
		
	
			var checkedPantryRadioButtonValue = getCheckedRadioValue(document.frmProspectingCustomizeDataFilters.optNumPantriesCompare);
		
			if (checkedPantryRadioButtonValue.value == "ByCustomNumber") 
			
			{
		
				if (document.frmProspectingCustomizeDataFilters.txtNumPantriesComparisonNumberSingle.value == "") 
					{		
						swal("Please specify number of pantries to filter by.")
						return false;
					}
				if (!isInteger(document.frmProspectingCustomizeDataFilters.txtNumPantriesComparisonNumberSingle.value))
					{		
						swal("Please enter a whole number for the number of pantries.")
						return false;
					}
			}
			
			if (checkedPantryRadioButtonValue.value == "ByCustomRange") 
			
			{
		
				if (document.frmProspectingCustomizeDataFilters.txtNumPantriesCustomRangeNumber1.value == "") 
					{		
						swal("Please specify a starting range number for number of pantries to filter by.")
						return false;
					}
				if (!isInteger(document.frmProspectingCustomizeDataFilters.txtNumPantriesCustomRangeNumber1.value))
					{		
						swal("Please enter a whole number for the starting number of pantries.")
						return false;
					}
				
				if (document.frmProspectingCustomizeDataFilters.txtNumPantriesCustomRangeNumber2.value == "") 
					{		
						swal("Please specify an ending range number for number of pantries to filter by.")
						return false;
					}
				if (!isInteger(document.frmProspectingCustomizeDataFilters.txtNumPantriesCustomRangeNumber2.value))
					{		
						swal("Please enter a whole number for the ending number of pantries.")
						return false;
					}
				
				
			}
			
   		}

		showHide();
        return true; 
        
    } //end validateCustomizeForm() function
// -->
</SCRIPT>


<style type="text/css">
	.ativa-scroll{
		max-height: 300px
	}
	
	.sellperiodform{
		margin-top: 20px;
	}
		
	.categories-checkboxes{
		font-size: 12px;
	}
	
	.categories-checkboxes input{
		margin-right: 6px;
	}
	
	.container-modal{
		border-bottom: 1px solid #e5e5e5;
		margin-bottom: 10px;
	}
	
	.stagerangedatepicker {
		position: absolute;
		bottom: 10px;
		right: 24px;
		top: auto;
		cursor: pointer;
	}
	
	.calendar-table tbody{
	 	height:200px !important;
	}	
	
	select[multiple], select[size] {
	    height: 100px;
	}
	</style>
<%
	If Right(MUV_READ("CLIENTID"),1) = "d" Then
		ClientKeyForFileName = LEFT(MUV_READ("CLIENTID"), (LEN(MUV_READ("CLIENTID"))-1))
	Else
		ClientKeyForFileName = (MUV_READ("CLIENTID"))
	End If		
%>
<!-- modal scroll !-->
<script type="text/javascript">
  $(document).ready(ajustamodal);
  $(window).resize(ajustamodal);
  function ajustamodal() {
    var altura = $(window).height(); //value corresponding to the modal heading + footer
    $(".ativa-scroll").css({"height":"500px","overflow-y":"auto"});
  }
  
  $(document).ready(function() {

		var autocompleteCityJSONFileURL = "../clientfiles/<%= ClientKeyForFileName %>/autocomplete/prospect_city_list.json";
		var autocompleteStateJSONFileURL = "../clientfiles/<%= ClientKeyForFileName %>/autocomplete/prospect_state_list.json";
		var autocompleteZipJSONFileURL = "../clientfiles/<%= ClientKeyForFileName %>/autocomplete/prospect_zip_list.json";
		
		var optionsCity = {
		  url: autocompleteCityJSONFileURL,
		  placeholder: "Start Typing ...",
		  getValue: "city",
		  list: {	
	        onChooseEvent: function() {
	            var city = $("#txtCityFilter").getSelectedItemData().city;
        	},		  
		    match: {
		      enabled: true
			},
			maxNumberOfElements: 10		
		  },
		  theme: "plate-dark"
		};
		$("#txtCityFilter").easyAutocomplete(optionsCity);
		$(".easy-autocomplete").removeAttr("style");
		
		
		var optionsState = {
		  url: autocompleteStateJSONFileURL,
		  placeholder: "Start Typing ...",
		  getValue: "state",
		  list: {	
	        onChooseEvent: function() {
	            var state = $("#txtStateFilter").getSelectedItemData().state;
        	},		  
		    match: {
		      enabled: true
			},
			maxNumberOfElements: 10		
		  },
		  theme: "plate-dark"
		};
		$("#txtStateFilter").easyAutocomplete(optionsState);
		
		
		
		
		var optionsZip = {
		  url: autocompleteZipJSONFileURL,
		  placeholder: "Start Typing ...",
		  getValue: "zip",
		  list: {	
	        onChooseEvent: function() {
	            var zipCode = $("#txtZipFilter").getSelectedItemData().zip;
        	},		  
		    match: {
		      enabled: true
			},
			maxNumberOfElements: 10		
		  },
		  theme: "plate-dark"
		};
		$("#txtZipFilter").easyAutocomplete(optionsZip);
		

		$("#resetModalFormFields").on('click',function(){
		   $('#frmProspectingCustomizeDataFilters').trigger("reset");
		   $('option').attr('selected', false);
		   $(':radio').prop('checked', false);
		   $('input:checkbox').removeAttr('checked');
		   $(':input').val('');
		   $('.left-column').removeClass("activefilter");
		});		 
		   
	    	
		$("#btnClearLocation").click(function() {
			$("#txtCityFilter").val('');
			$("#txtStateFilter").val('');
			$("#txtZipFilter").val('');
		});		
	
	
		$("#btnClearStagesLost").click(function() {
		
			$(".checkLost").removeAttr('checked');
			$('input[name="optStageLostDateRange"]').prop('checked', false);
			$("#txtStageLostDateRangeStartDate").val('');
			$("#txtStageLostDateRangeEndDate").val('');
			$('#selLostStageDateRangeCustom').children().removeProp('selected');
			$("#stageRangeLost").val('');
		});	
		
		$("#btnClearStagesUnqualified").click(function() {
		
			$(".checkUnqualified").removeAttr('checked');
			$('input[name="optStageUnqualifiedDateRange"]').prop('checked', false);
			$("#txtStageUnqualifiedRangeStartDate").val('');
			$("#txtStageUnqualifiedRangeEndDate").val('');
			$('#selUnqualifiedStageDateRangeCustom').children().removeProp('selected');
			$("#stageRangeUnqualified").val('');
		
		});	
		
		$("#btnClearLeadSource").click(function() { 
			$('#selLeadSourceNumber').children().removeProp('selected');   
		});	
		$("#btnSelectAllLeadSource").click(function() {
			$('#selLeadSourceNumber option').prop('selected', true);
		});	

	
		$("#btnClearIndustry").click(function() {
			$('#selIndustryNumber').children().removeProp('selected');
		});	
		$("#btnSelectAllIndustry").click(function() {
			$('#selIndustryNumber option').prop('selected', true);
		});	

		
		$("#btnClearTelemarketers").click(function() {
			$('#selTelemarketerUserNo').children().removeProp('selected');
		});	
		$("#btnSelectAllTelemarketers").click(function() {
			$('#selTelemarketerUserNo option').prop('selected', true);
		});	
		
		
		
		$("#btnClearOwners").click(function() {
			$('#selProspectOwnerUserNo').children().removeProp('selected');
		});	
		$("#btnSelectAllOwners").click(function() {
			$('#selProspectOwnerUserNo option').prop('selected', true);
		});	

		
		$("#btnClearCreatedBy").click(function() {
			$('#selProspectCreatedByUserNo').children().removeProp('selected');
		});	
		$("#btnSelectAllCreatedBy").click(function() {
			$('#selProspectCreatedByUserNo option').prop('selected', true);
		});	

		
		$("#btnClearCreatedDate").click(function() {
			$("#optProspectCreatedDateRange").removeAttr("checked");
			$("#txtProspectCreatedRangeStartDate").val('');
			$("#txtProspectCreatedRangeEndDate").val('');
			$("#prospectCreatedRange").val('');
			$('#selProspectCreatedDateRangeCustom').children().removeProp('selected');
			
		});	

		$("#btnClearNumEmployees").click(function() {
		
			$('input[name="optNumEmployeesRangeCompare"]').prop('checked', false);
			$('#selEmployeeRangeNo').prop('selectedIndex', -1);
			$('#selEmployeeRangeComparisonOperator').prop('selectedIndex', -1);
			$("#txtEmployeeRangeComparisonNumberSingle").val('');
			$("#txtEmployeeCustomRangeNumber1").val('');
			$("#txtEmployeeCustomRangeNumber2").val('');

		});	
		
   
		$("#btnClearNumPantries").click(function() {

			$('input[name="optNumPantriesCompare"]').prop('checked', false);
			$('#selNumPantriesComparisonOperator').prop('selectedIndex', -1);
			$("#txtNumPantriesComparisonNumberSingle").val('');
			$("#txtNumPantriesCustomRangeNumber1").val('');
			$("#txtNumPantriesCustomRangeNumber2").val('');

		});	

		$("#selLostStageDateRangeCustom").change(function() {
			
			selectedValue = $(this).find(":selected").val();
			
		    if (selectedValue != "") 
		    {
		    	$("#stageRangeLost").val('');
		    	$("#txtStageLostDateRangeStartDate").val('');
		    	$("#txtStageLostDateRangeEndDate").val('');
		    	$("#optStageLostDateRange").prop("checked","checked");
		    }
		    else
		    {
		    	$("#optStageLostDateRange").prop("checked","");
		    }
		});
		
		$("#selUnqualifiedStageDateRangeCustom").change(function() {
		
			selectedValue = $(this).find(":selected").val();
			
		    if (selectedValue != "") 
		    {
		    	$("#stageRangeUnqualified").val('');
		    	$("#txtStageUnqualifiedRangeStartDate").val('');
		    	$("#txtStageUnqualifiedRangeEndDate").val('');
		    	$("#optStageUnqualifiedDateRange").prop("checked","checked");
		    }
		    else
		    {
		    	$("#optStageUnqualifiedDateRange").prop("checked","");
		    }

		});


		$("#selProspectCreatedDateRangeCustom").change(function() {
		
			selectedValue = $(this).find(":selected").val();
			
		    if (selectedValue != "") 
		    {
		    	$("#prospectCreatedRange").val('');
		    	$("#txtProspectCreatedRangeStartDate").val('');
		    	$("#txtProspectCreatedRangeEndDate").val('');
		    	$("#optProspectCreatedDateRange").prop("checked","checked");
		    }
		    else
		    {
		    	$("#optProspectCreatedDateRange").prop("checked","");
		    }
		});
		
		
	});
		
</script>
<!-- eof modal scroll !-->

<%

function mmddyy(input)
    dim m: m = month(input)
    dim d: d = day(input)
    if (m < 10) then m = "0" & m
    if (d < 10) then d = "0" & d

    mmddyy = m & "/" & d & "/" & right(year(input), 2)
end function

Function dateCustomFormat(date)
	x = FormatDateTime(date, 2) 
	d = split(x, "/")
	dateCustomFormat = d(2) & "-" & d(0) & "-" & d(1)
End Function


%>
	
<!-- modal box !-->
<div class="modal fade bs-modal-filter-prospecting-data-recycle-pool" tabindex="-1" role="dialog" aria-labelledby="myLargeModalLabel" aria-hidden="true">
	<div class="modal-dialog modal-lg modal-height">
		<div class="modal-content">
			<div class="modal-header">
				<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
				<h4 class="modal-title" id="myModalLabel" align="center">Customize and Filter <%= GetTerm("Recycle Pool") %> Screen</h4>
			</div>
			<%
			'************************
			'Read Settings_Reports
			'************************
			
			SQLReportName = Replace(MUV_READ("CRMVIEWSTATERECPOOL"),"'","''")
			
			SQL = "SELECT * FROM Settings_Reports WHERE ReportNumber = 1400 AND UserNo = " & Session("userNo") & " AND PoolForProspecting = 'Dead' AND UserReportName = '" & SQLReportName & "'"
			
			Set cnn8 = Server.CreateObject("ADODB.Connection")
			cnn8.open (Session("ClientCnnString"))
			Set rs = Server.CreateObject("ADODB.Recordset")
			Set rs= cnn8.Execute(SQL)
			UseSettings_Reports = False
			If NOT rs.EOF Then
				UseSettings_Reports = True
				ReportSpecificData2  = rs("ReportSpecificData2")
				ReportSpecificData3  = rs("ReportSpecificData3")
				ReportSpecificData4  = rs("ReportSpecificData4")
				ReportSpecificData5  = rs("ReportSpecificData5")
				ReportSpecificData6  = rs("ReportSpecificData6")
				ReportSpecificData7  = rs("ReportSpecificData7")
				ReportSpecificData8  = rs("ReportSpecificData8")
				ReportSpecificData9  = rs("ReportSpecificData9")
				ReportSpecificData10  = rs("ReportSpecificData10")
				ReportSpecificData11  = rs("ReportSpecificData11")
				ReportSpecificData12  = rs("ReportSpecificData12")
				ReportSpecificData13  = rs("ReportSpecificData13")
				ReportSpecificData14  = rs("ReportSpecificData14")
				ReportSpecificData15  = rs("ReportSpecificData15")
				ReportSpecificData16  = rs("ReportSpecificData16")
				ReportSpecificData17  = rs("ReportSpecificData17")
				ReportSpecificData18  = rs("ReportSpecificData18")
				ReportSpecificData19  = rs("ReportSpecificData19")
				ReportSpecificData20  = rs("ReportSpecificData20")
				ReportSpecificData21  = rs("ReportSpecificData21")
				ReportSpecificData22  = rs("ReportSpecificData22")
				ReportSpecificData23  = rs("ReportSpecificData23")
				ReportSpecificData24  = rs("ReportSpecificData24")
				ReportSpecificData25  = rs("ReportSpecificData25")
				ReportSpecificData26  = rs("ReportSpecificData26")
				ReportSpecificData27  = rs("ReportSpecificData27")
				ReportSpecificData28  = rs("ReportSpecificData28")
			End If
			'****************************
			'End Read Settings_Reports
			'****************************
			%>
  
			<!-- eof modal scroll !-->
			
	<form method="post" action="mainRecyclePoolCustomizeSaveDataFilterValues.asp" id="frmProspectingCustomizeDataFilters" name="frmProspectingCustomizeDataFilters" onsubmit="return validateCustomizeForm();">

	      <!-- insert content in here !-->
	      <div class="modal-body ativa-scroll">


 	      	<!-- date ranges !-->
	      	<div class="container-fluid container-modal">
	      	
		      	<div class="row">
      	
 		      	<!-- left column !-->
 		      	<div class="col-lg-2 col-md-2 col-sm-12 col-xs-12 left-column">
	 		      	<h4><br>Clear</h4>
 		      	</div>
 		      	<!-- eof left column !-->
      	
 		      	<!-- right column !-->
 		      	<div class="col-lg-10 col-md-10 col-sm-12 col-xs-12 right-column">
	      	
		      	<!-- row !-->
			      	<div class="row">
				        <!-- First Date !-->
				    	<div class="col-xs-8 col-sm-1 col-md-12 col-lg-8">
							<button type="button" class="btn btn-warning btn-lg btn-block" id="resetModalFormFields">
								<i class="fa fa-refresh" aria-hidden="true"></i>&nbsp;Clear All Selected Options/Filters
							</button>
				        </div>
    			  	</div>
    			  	<!-- eof row !-->
      	
 		      	</div>
 		      	<!-- eof right column !-->
	      	</div>
      	</div>
      	<!-- eof date ranges !-->    
      	
 	      	<!-- date ranges !-->
	      	<div class="container-fluid container-modal">
	      	
		      	<div class="row">
      	
 		      	<!-- left column !-->
 		      	<div class="col-lg-2 col-md-2 col-sm-12 col-xs-12 left-column <% If ReportSpecificData15 <> "" OR ReportSpecificData16 <> "" OR ReportSpecificData17 <> "" Then Response.Write("activefilter") %>">
	 		      	<h4><br>Location<br><br><button type="button" class="btn btn-info btn-xs btn-block" id="btnClearLocation">clear locations</button></h4>
 		      	</div>
 		      	<!-- eof left column !-->
      	
 		      	<!-- right column !-->
 		      	<div class="col-lg-3 col-md-3 col-sm-3 col-xs-3 right-column">
		      		<!-- row !-->
			      	<div class="row">
				        <!-- First Date !-->
				    	<div class="col-xs-12 col-sm-12 col-md-112 col-lg-12">
				    		<% If ReportSpecificData15 <> "" Then %>
					    		City: <input type="text" class="form-control" name="txtCityFilter" id="txtCityFilter" style="width: 160px; !important;" value="<%= ReportSpecificData15 %>">
					    	<% Else %>
					    		City: <input type="text" class="form-control" name="txtCityFilter" id="txtCityFilter" style="width: 160px; !important;">
					    	<% End If %>
				        </div>
				  	</div>
				  	<!-- eof row !-->

 		      	</div>
 		      	<!-- eof right column !-->
 		      	
 		      	<!-- right column !-->
 		      	<div class="col-lg-3 col-md-3 col-sm-3 col-xs-3 right-column">
		      		<!-- row !-->
			      	<div class="row">
				        <!-- First Date !-->
				    	<div class="col-xs-12 col-sm-12 col-md-112 col-lg-12">
				    		<% If ReportSpecificData16 <> "" Then %>
					    		State: <input type="text" class="form-control" name="txtStateFilter" id="txtStateFilter" style="width: 160px; !important;" value="<%= ReportSpecificData16 %>">
					    	<% Else %>
					    		State: <input type="text" class="form-control" name="txtStateFilter" id="txtStateFilter" style="width: 160px; !important;">
					    	<% End If %>
				        </div>
				  	</div>
				  	<!-- eof row !-->

 		      	</div>
 		      	<!-- eof right column !-->


 		      	<!-- right column !-->
 		      	<div class="col-lg-3 col-md-3 col-sm-3 col-xs-3 right-column">
		      		<!-- row !-->
			      	<div class="row">
				        <!-- First Date !-->
				    	<div class="col-xs-12 col-sm-12 col-md-112 col-lg-12">
				    		<% If ReportSpecificData17 <> "" Then %>
					    		Zip: <input type="text" class="form-control" name="txtZipFilter" id="txtZipFilter" style="width: 160px; !important;" value="<%= ReportSpecificData17 %>">
					    	<% Else %>
					    		Zip: <input type="text" class="form-control" name="txtZipFilter" id="txtZipFilter" style="width: 160px; !important;">
					    	<% End If %>
				        </div>
				  	</div>
				  	<!-- eof row !-->

 		      	</div>
 		      	<!-- eof right column !-->
 		      	
	      	</div>
      	</div>
      	<!-- eof date ranges !-->    
      	
      	
      	
   	
   	
   	
   	
   	
   	

      


	
  	 	      	
  	<!-- categories !-->
  	<div class="container-fluid container-modal">
      	<div class="row">
  	
	      	<!-- left column !-->
	      	<div class="col-lg-2 col-md-3 col-sm-12 col-xs-12 left-column <% If ReportSpecificData23 <> "" OR ReportSpecificData24 <> "" Then Response.Write("activefilter") %>">
 		      	<h4><br>Stage - Lost<br><br><button type="button" class="btn btn-info btn-xs btn-block" id="btnClearStagesLost">clear selections</button></h4>
	      	</div>
	      	<!-- eof left column !-->
  	
	      	<!-- right column !-->
	      	<div class="col-lg-10 col-md-9 col-sm-12 col-xs-12 right-column check-row">
      	
	      	<!-- row !-->
		      	<div class="row categories-checkboxes">
      	
			      	<div class="col-lg-12">
					  <div class="checkbox">
						<label>
						  <input type="checkbox" class="checkLost" id="checkAllLost"> <strong>CHECK / UNCHECK ALL LOST REASONS</strong>
						</label>
					  </div>
			      	</div>

					<div class="checkbox">
  
		     			<%
		     				If ReportSpecificData23 <> "" Then
			     				selectedLostReasonsToFilterArray = Split(ReportSpecificData23,",")
			     				upperBoundLost = ubound(selectedLostReasonsToFilterArray)
			     			Else
			     				upperBoundLost = -1
			     				selectedLostReasonsToFilterArray = ""
			     			End If
							 
				      		'Get all stages
				      	  	SQLStagesLost = "SELECT * FROM PR_Reasons WHERE ReasonType='Lost' OR ReasonType='Unqualifying and Lost' ORDER BY InternalRecordIdentifier"
		
							Set cnnStagesLost = Server.CreateObject("ADODB.Connection")
							cnnStagesLost.open (Session("ClientCnnString"))
							Set rsStagesLost = Server.CreateObject("ADODB.Recordset")
							rsStagesLost.CursorLocation = 3 
							Set rsStagesLost = cnnStagesLost.Execute(SQLStagesLost)
								
							If not rsStagesLost.EOF Then
								Do
									%>
							      	<div class="col-lg-6 col-md-6 col-sm-12 col-xs-12 ">
								      	<label>
								      	<%
								      	stageIsSelected = "False"
								      	
								      	For i=0 to upperBoundLost
										   If cInt(selectedLostReasonsToFilterArray(i)) = cInt(rsStagesLost("InternalRecordIdentifier")) Then
										   		stageIsSelected = "True"
										   End If
										Next

								      	%>
								      	<input type="checkbox" class="checkLost" <% If stageIsSelected = "True" Then Response.Write("checked='checked'") %> id="chkStageLost<%= rsStagesLost("InternalRecordIdentifier") %>" name="chkStageLost" value="<%= rsStagesLost("InternalRecordIdentifier") %>"><%= rsStagesLost("Reason") %><br>
								      	</label>
							      	</div>   
									<%
									rsStagesLost.movenext
								Loop until rsStagesLost.eof
							End If
		     			%>
    			
		     			<!-- select all / deselect all checkboxes !-->
						<script type="text/javascript">
						$("#checkAllLost").click(function () {
					    $(".checkLost").prop('checked', $(this).prop('checked'));
						});	</script>

					</div> 		      			      	
		      	</div>
		      	<!-- eof row !-->
		      	
		      	<div class="row" style="margin-top:35px;">
		      	
			        
				    <div class="col-xs-12 col-sm-12 col-md-12 col-lg-12">

				    	<div class="col-lg-12" style="margin-left:-15px;">
					    	<input type="radio" id="optStageLostDateRange" name="optStageLostDateRange" value="WasLostInDateRange" <% If ReportSpecificData24 <> "" AND ReportSpecificData25 <> "" Then Response.write("checked") %>>
							<%
							If ReportSpecificData24 <> "" AND ReportSpecificData25 <> "" Then
								startDateStageLostRange = dateCustomFormat(ReportSpecificData24)
								endDateStageLostRange = dateCustomFormat(ReportSpecificData25) 
							Else
								thirtyDaysAgo = DateAdd("d",-30, DateSerial(Year(Now()), Month(Now()), Day(Now())))
								startDateStageLostRange = dateCustomFormat(thirtyDaysAgo)
								endDateStageLostRange = dateCustomFormat(Now())
							End If
	
							%>
							Where prospect was "lost" in the date range:
						</div> 	
				    </div>	
				 </div>
				 
		      	<div class="row" style="margin-top:15px;">
				    <div class="col-xs-12 col-sm-12 col-md-12 col-lg-12" style="margin-left:15px;">

				    	<div class="col-lg-5" style="margin-left:-15px;">
							Choose a "Quick Pick" Date Range<br>
							<select class="form-control" name="selLostStageDateRangeCustom" id="selLostStageDateRangeCustom">
								<option value="">Select A Date Range</option>
								<option value="Today">Today</option>
								<option value="This Week">This Week</option>
								<option value="Last Week">Last Week</option>
								<option value="Past 10 Days">Past 10 Days</option>
								<option value="Past 15 Days">Past 15 Days</option>
								<option value="Past 30 Days">Past 30 Days</option>
								<option value="This Month">This Month</option>
								<option value="Last Month">Last Month</option>
								<option value="Year To Date">Year To Date</option>
								<option value="Last Year">Last Year</option>
								<option value="All Dates">All Dates</option>
							</select>
						</div> 
						<div class="col-lg-1"><strong>OR</strong></div>
						<div class="col-lg-6">
							Use Calendar To Select Dates<br>					
                            <input type="text" id="stageRangeLost" class="form-control">
                            <i class="glyphicon glyphicon-calendar fa fa-calendar stagerangedatepicker"></i>	
                        </div>	
						<input type="hidden" id="txtStageLostDateRangeStartDate" name="txtStageLostDateRangeStartDate" value="<%= startDateStageLostRange %>">
						<input type="hidden" id="txtStageLostDateRangeEndDate" name="txtStageLostDateRangeEndDate" value="<%= endDateStageLostRange %>">
				    </div>	
				 </div>
				 
	      	</div>
	      	<!-- eof right column !-->
      	</div>
  	</div>
  	<!-- eof categories !-->      	

  
   	
   	
   	
   	
   	
   	
   	
   	
   	
   	
   	
   	

	
  	 	      	
  	<!-- categories !-->
  	<div class="container-fluid container-modal">
      	<div class="row">
  	
	      	<!-- left column !-->
	      	<div class="col-lg-2 col-md-3 col-sm-12 col-xs-12 left-column <% If ReportSpecificData26 <> "" OR ReportSpecificData27 <> "" Then Response.Write("activefilter") %>">
 		      	<h4><br>Stage - Unqualified<br><br><button type="button" class="btn btn-info btn-xs btn-block" id="btnClearStagesUnqualified">clear selections</button></h4>
	      	</div>
	      	<!-- eof left column !-->
  	
	      	<!-- right column !-->
	      	<div class="col-lg-10 col-md-9 col-sm-12 col-xs-12 right-column check-row">
      	
	      	<!-- row !-->
		      	<div class="row categories-checkboxes">
      	
			      	<div class="col-lg-12">
					  <div class="checkbox">
						<label>
						  <input type="checkbox" class="checkUnqualified" id="checkAllUnqualified"> <strong>CHECK / UNCHECK ALL UNQUALIFIED REASONS</strong>
						</label>
					  </div>
			      	</div>

					<div class="checkbox">
  
		     			<%
		     				If ReportSpecificData26 <> "" Then
			     				selectedUnqualifiedReasonsToFilterArray = Split(ReportSpecificData26,",")
			     				upperBoundUnqualified = ubound(selectedUnqualifiedReasonsToFilterArray)
			     			Else
			     				upperBoundUnqualified = -1
			     				selectedUnqualifiedReasonsToFilterArray = ""
			     			End If
							 
				      		'Get all stages
				      	  	SQLStagesUnqualified = "SELECT * FROM PR_Reasons WHERE ReasonType='Unqualifying' OR ReasonType='Unqualifying and Lost' ORDER BY InternalRecordIdentifier"
		
							Set cnnStagesUnqualified = Server.CreateObject("ADODB.Connection")
							cnnStagesUnqualified.open (Session("ClientCnnString"))
							Set rsStagesUnqualified = Server.CreateObject("ADODB.Recordset")
							rsStagesUnqualified.CursorLocation = 3 
							Set rsStagesUnqualified = cnnStagesUnqualified.Execute(SQLStagesUnqualified)
								
							If not rsStagesUnqualified.EOF Then
								Do
									%>
							      	<div class="col-lg-6 col-md-6 col-sm-12 col-xs-12 ">
								      	<label>
								      	<%
								      	stageUnqualifiedIsSelected = "False"
								      	
								      	For i=0 to upperBoundUnqualified
										   If cInt(selectedUnqualifiedReasonsToFilterArray(i)) = cInt(rsStagesUnqualified("InternalRecordIdentifier")) Then
										   		stageUnqualifiedIsSelected = "True"
										   End If
										Next

								      	%>
								      	<input type="checkbox" class="checkUnqualified" <% If stageUnqualifiedIsSelected = "True" Then Response.Write("checked='checked'") %> id="chkStageUnqualified<%= rsStagesUnqualified("InternalRecordIdentifier") %>" name="chkStageUnqualified" value="<%= rsStagesUnqualified("InternalRecordIdentifier") %>"><%= rsStagesUnqualified("Reason") %><br>
								      	</label>
							      	</div>   
									<%
									rsStagesUnqualified.movenext
								Loop until rsStagesUnqualified.eof
							End If
		     			%>
    			
		     			<!-- select all / deselect all checkboxes !-->
						<script type="text/javascript">
						$("#checkAllUnqualified").click(function () {
					    $(".checkUnqualified").prop('checked', $(this).prop('checked'));
						});	</script>

					</div> 		      			      	
		      	</div>
		      	<!-- eof row !-->
		      	
		      	<div class="row" style="margin-top:35px;">
		      	
				    
				    <div class="col-xs-12 col-sm-12 col-md-12 col-lg-12" style="margin-bottom:10px;">
				    	<div class="col-lg-12" style="margin-left:-15px;">
					    	<input type="radio" id="optStageUnqualifiedDateRange" name="optStageUnqualifiedDateRange" value="WasUnqualifiedInDateRange" <% If ReportSpecificData27 <> "" AND ReportSpecificData28 <> "" Then Response.write("checked") %>>
						
							<%
							If ReportSpecificData27 <> "" AND ReportSpecificData28 <> "" Then
								startDateStageNotChangedRange = dateCustomFormat(ReportSpecificData27)
								endDateStageNotChangedRange = dateCustomFormat(ReportSpecificData28) 
							Else
								thirtyDaysAgo = DateAdd("d",-30, DateSerial(Year(Now()), Month(Now()), Day(Now())))
								startDateStageNotChangedRange = dateCustomFormat(thirtyDaysAgo)
								endDateStageNotChangedRange = dateCustomFormat(Now())
							End If
	
							%>
							Where prospect was "unqualified" in the date range:
						</div> 
			        
					</div>
				 </div>
				 
		      	<div class="row" style="margin-top:15px;">
				    <div class="col-xs-12 col-sm-12 col-md-12 col-lg-12" style="margin-left:15px;">

				    	<div class="col-lg-5" style="margin-left:-15px;">
							Choose a "Quick Pick" Date Range<br>
							<select class="form-control" name="selUnqualifiedStageDateRangeCustom" id="selUnqualifiedStageDateRangeCustom">
								<option value="">Select A Date Range</option>
								<option value="Today">Today</option>
								<option value="This Week">This Week</option>
								<option value="Last Week">Last Week</option>
								<option value="Past 10 Days">Past 10 Days</option>
								<option value="Past 15 Days">Past 15 Days</option>
								<option value="Past 30 Days">Past 30 Days</option>
								<option value="This Month">This Month</option>
								<option value="Last Month">Last Month</option>
								<option value="Year To Date">Year To Date</option>
								<option value="Last Year">Last Year</option>
								<option value="All Dates">All Dates</option>
							</select>
						</div> 
						<div class="col-lg-1"><strong>OR</strong></div>
						<div class="col-lg-6">
							Use Calendar To Select Dates<br>					
                            <input type="text" id="stageRangeUnqualified" class="form-control">
                            <i class="glyphicon glyphicon-calendar fa fa-calendar stagerangedatepicker"></i>	
                        </div>
						<input type="hidden" id="txtStageUnqualifiedRangeStartDate" name="txtStageUnqualifiedRangeStartDate" value="<%= startDateStageNotChangedRange %>">
						<input type="hidden" id="txtStageUnqualifiedRangeEndDate" name="txtStageUnqualifiedRangeEndDate" value="<%= endDateStageNotChangedRange %>">
				    </div>	
				 </div>
				 
	      	</div>
	      	<!-- eof right column !-->
      	</div>
  	</div>
  	<!-- eof categories !-->      	

  
   	
   	
   	
   	   	
   	
   	
   	
   	
   	
   	
   	
   	
   	
   	
   	
   	
   	
	<!-- exclusions !-->
  	<div class="container-fluid container-modal">
      	<div class="row">
  	
      	<!-- left column !-->
      	<div class="col-lg-2 col-md-3 col-sm-12 col-xs-12 left-column leadsource <% If ReportSpecificData6 <> "" Then Response.Write("activefilter") %>">
	      	<h4><br>Lead Source<br><br><button type="button" class="btn btn-info btn-xs btn-block" id="btnClearLeadSource">clear selections</button>
	      	<button type="button" class="btn btn-success btn-xs btn-block" id="btnSelectAllLeadSource">select all</button></h4>
      	</div>
      	<!-- eof left column !-->
  	
      	<!-- right column !-->
      	<div class="col-lg-10 col-md-9 col-sm-12 col-xs-12 right-column">
  	
  	
      	<!-- row !-->
      	<div class="row">

	      	<div class="col-lg-7 col-md-7 col-sm-12 col-xs-12">
		      	<select class="form-control" name="selLeadSourceNumber" id="selLeadSourceNumber" multiple="multiple">
	  	  			<%
	     				If ReportSpecificData6 <> "" Then
		     				selectedLeadSourcesToFilterArray = Split(ReportSpecificData6,",")
		     				upperBoundLeadSource = ubound(selectedLeadSourcesToFilterArray)
		     			Else
		     				upperBoundLeadSource = -1
		     				selectedLeadSourcesToFilterArray = ""
		     			End If
  	  			
						SQLLeadSource = "SELECT * FROM PR_LeadSources ORDER BY LeadSource"

						Set cnnLeadSource = Server.CreateObject("ADODB.Connection")
						cnnLeadSource.open (Session("ClientCnnString"))
						Set rsLeadSource = Server.CreateObject("ADODB.Recordset")
						rsLeadSource.CursorLocation = 3 
						Set rsLeadSource = cnnLeadSource.Execute(SQLLeadSource)
							
						If not rsLeadSource.EOF Then
							Do

						      	leadSourceIsSelected = "False"
						      	
						      	For i=0 to upperBoundLeadSource
								   If cInt(selectedLeadSourcesToFilterArray(i)) = cInt(rsLeadSource("InternalRecordIdentifier")) Then
								   		leadSourceIsSelected = "True"
								   End If
								Next

						      	%>

								<option value="<%= rsLeadSource("InternalRecordIdentifier") %>" <% If leadSourceIsSelected = "True" Then Response.Write("selected") %>><%= rsLeadSource("LeadSource") %></option><%
								rsLeadSource.movenext
							Loop until rsLeadSource.eof
						End If
						set rsLeadSource = Nothing
						cnnLeadSource.close
						set cnnLeadSource = Nothing
					%>
			    </select>
	      	</div>
	      	<div class="col-lg-5 col-md-5 col-sm-12 col-xs-12">
	      		SHIFT+CLICK To Select Multiple Values
	      	</div>	      	
      	</div>
      	<!-- eof row !-->
	  	</div>
	  	<!-- eof right column !-->
	</div>
	</div>
	   	
   	
    	
   	
	<!-- exclusions !-->
  	<div class="container-fluid container-modal">
      	<div class="row">
  	
      	<!-- left column !-->
      	<div class="col-lg-2 col-md-3 col-sm-12 col-xs-12 left-column <% If ReportSpecificData7 <> "" Then Response.Write("activefilter") %>">
	      	<h4><br>Industry<br><br><button type="button" class="btn btn-info btn-xs btn-block" id="btnClearIndustry">clear selections</button>
	      	<button type="button" class="btn btn-success btn-xs btn-block" id="btnSelectAllIndustry">select all</button></h4>
      	</div>
      	<!-- eof left column !-->
  	
      	<!-- right column !-->
      	<div class="col-lg-10 col-md-9 col-sm-12 col-xs-12 right-column">
  	
  	
      	<!-- row !-->
      	<div class="row">

	      	<div class="col-lg-7 col-md-7 col-sm-12 col-xs-12">
		      	<select class="form-control" name="selIndustryNumber" id="selIndustryNumber" multiple="multiple">
	  	  			<%
	  	  			
	     				If ReportSpecificData7 <> "" Then
		     				selectedIndustriesToFilterArray = Split(ReportSpecificData7,",")
		     				upperBoundIndustries = ubound(selectedIndustriesToFilterArray)
		     			Else
		     				upperBoundIndustries = -1
		     				selectedIndustriesToFilterArray = ""
		     			End If
	  	  			
			      	  	SQLIndustry = "SELECT * FROM PR_Industries ORDER BY Industry "
	
						Set cnnIndustry = Server.CreateObject("ADODB.Connection")
						cnnIndustry.open (Session("ClientCnnString"))
						Set rsIndustry = Server.CreateObject("ADODB.Recordset")
						rsIndustry.CursorLocation = 3 
						Set rsIndustry = cnnIndustry.Execute(SQLIndustry)
							
						If not rsIndustry.EOF Then
							Do
						      	industryIsSelected = "False"
						      	
						      	For i=0 to upperBoundIndustries
								   If cInt(selectedIndustriesToFilterArray(i)) = cInt(rsIndustry("InternalRecordIdentifier")) Then
								   		industryIsSelected = "True"
								   End If
								Next
	
									%><option value="<%= rsIndustry("InternalRecordIdentifier") %>" <% If industryIsSelected = "True" Then Response.Write("selected") %>><%= rsIndustry("Industry") %></option>
									
									<%
								rsIndustry.movenext
							Loop until rsIndustry.eof
						End If
						set rsIndustry = Nothing
						cnnIndustry.close
						set cnnIndustry = Nothing
					%>
			    </select>
	      	</div>
	      	<div class="col-lg-5 col-md-5 col-sm-12 col-xs-12">
	      		SHIFT+CLICK To Select Multiple Values
	      	</div>
      	</div>
      	<!-- eof row !-->
	  	</div>
	  	<!-- eof right column !-->
	</div>
	</div>
	   	


   	
   	
	<!-- exclusions !-->
  	<div class="container-fluid container-modal">
      	<div class="row">
  	
      	<!-- left column !-->
      	<div class="col-lg-2 col-md-3 col-sm-12 col-xs-12 left-column <% If ReportSpecificData8 <> "" Then Response.Write("activefilter") %>">
	      	<h4><br>Telemarketer<br><br><button type="button" class="btn btn-info btn-xs btn-block" id="btnClearTelemarketers">clear selections</button>
	      	<button type="button" class="btn btn-success btn-xs btn-block" id="btnSelectAllTelemarketers">select all</button></h4>
      	</div>
      	<!-- eof left column !-->
  	
      	<!-- right column !-->
      	<div class="col-lg-10 col-md-9 col-sm-12 col-xs-12 right-column">
  	
  	
      	<!-- row !-->
      	<div class="row">

	      	<div class="col-lg-7 col-md-7 col-sm-12 col-xs-12">
		      	<select class="form-control" name="selTelemarketerUserNo" id="selTelemarketerUserNo" multiple="multiple">
	  	  			<%

	     				If ReportSpecificData8 <> "" Then
		     				selectedTelemarketersToFilterArray = Split(ReportSpecificData8,",")
		     				upperBoundTelemarketers = ubound(selectedTelemarketersToFilterArray )
		     			Else
		     				upperBoundTelemarketers = -1
		     				selectedIndustriesToFilterArray = ""
		     			End If
	  	  			
			      	  	SQLTelemarketer = "SELECT UserNo, userFirstName, userLastName, userType FROM " & MUV_Read("SQL_Owner") & ".tblUsers "
			      	  	SQLTelemarketer = SQLTelemarketer & "WHERE userArchived <> 1 AND userEnabled = 1"
			      	  	SQLTelemarketer = SQLTelemarketer & " AND userType = 'Telemarketing' "
			      	  	SQLTelemarketer = SQLTelemarketer & "ORDER BY userFirstName, userLastName"
			
						Set cnnTelemarketer = Server.CreateObject("ADODB.Connection")
						cnnTelemarketer.open (Session("ClientCnnString"))
						Set rsTelemarketer = Server.CreateObject("ADODB.Recordset")
						rsTelemarketer.CursorLocation = 3 
						Set rsTelemarketer = cnnTelemarketer.Execute(SQLTelemarketer)
					
						If not rsTelemarketer.EOF Then
							Do
								FullName = rsTelemarketer("userFirstName") & " " & rsTelemarketer("userLastName")
								
						      	telemarketerIsSelected = "False"
						      	
						      	For i=0 to upperBoundTelemarketers
								   If cInt(selectedTelemarketersToFilterArray(i)) = cInt(rsTelemarketer("UserNo")) Then
								   		telemarketerIsSelected = "True"
								   End If
								Next
								
								%>
								<option value="<%= rsTelemarketer("UserNo") %>" <% If telemarketerIsSelected = "True" Then Response.Write("selected") %>><%= FullName %></option>
								<%
								rsTelemarketer.movenext
							Loop until rsTelemarketer.eof
						End If
						set rsTelemarketer = Nothing
						cnnTelemarketer.close
						set cnnTelemarketer = Nothing
	  	  			
					%>
			    </select>
	      	</div>
	      	<div class="col-lg-5 col-md-5 col-sm-12 col-xs-12">
	      		SHIFT+CLICK To Select Multiple Values
	      	</div>	      	
      	</div>
      	<!-- eof row !-->
	  	</div>
	  	<!-- eof right column !-->
	</div>
	</div>



   	
   	
	<!-- exclusions !-->
  	<div class="container-fluid container-modal">
      	<div class="row">
  	
      	<!-- left column !-->
      	<div class="col-lg-2 col-md-3 col-sm-12 col-xs-12 left-column <% If ReportSpecificData9 <> "" Then Response.Write("activefilter") %>">
	      	<h4><br>Owner<br><br><button type="button" class="btn btn-info btn-xs btn-block" id="btnClearOwners">clear selections</button>
	      	<button type="button" class="btn btn-success btn-xs btn-block" id="btnSelectAllOwners">select all</button></h4>
      	</div>
      	<!-- eof left column !-->
  	
      	<!-- right column !-->
      	<div class="col-lg-10 col-md-9 col-sm-12 col-xs-12 right-column">
  	
  	
      	<!-- row !-->
      	<div class="row">

	      	<div class="col-lg-7 col-md-7 col-sm-12 col-xs-12">
		      	<select class="form-control" name="selProspectOwnerUserNo" id="selProspectOwnerUserNo" multiple="multiple">
	  	  			<%

	     				If ReportSpecificData9 <> "" Then
		     				selectedOwnersToFilterArray = Split(ReportSpecificData9,",")
		     				upperBoundOwners = ubound(selectedOwnersToFilterArray)
		     			Else
		     				upperBoundOwners = -1
		     				selectedOwnersToFilterArray = ""
		     			End If
	  	  			
			      	  	SQLOwner = "SELECT UserNo, userFirstName, userLastName, userType FROM " & MUV_Read("SQL_Owner") & ".tblUsers "
			      	  	SQLOwner = SQLOwner & "WHERE UserNo IN (SELECT DISTINCT OwnerUserNo FROM PR_Prospects WHERE Pool='Dead') "
			      	  	SQLOwner = SQLOwner & "ORDER BY userFirstName, userLastName"
			
						Set cnnOwner = Server.CreateObject("ADODB.Connection")
						cnnOwner.open (Session("ClientCnnString"))
						Set rsOwner = Server.CreateObject("ADODB.Recordset")
						rsOwner.CursorLocation = 3 
						Set rsOwner = cnnOwner.Execute(SQLOwner)
					
						If not rsOwner.EOF Then
							Do
								FullName = rsOwner("userFirstName") & " " & rsOwner("userLastName")

						      	ownerIsSelected = "False"
						      	
						      	For i=0 to upperBoundOwners
								   If cInt(selectedOwnersToFilterArray(i)) = cInt(rsOwner("UserNo")) Then
								   		ownerIsSelected = "True"
								   End If
								Next
								
								%>
								<option value="<%= rsOwner("UserNo") %>" <% If ownerIsSelected = "True" Then Response.Write("selected") %>><%= FullName %></option>
								<%
								
								rsOwner.movenext
							Loop until rsOwner.eof
						End If
						set rsOwner = Nothing
						cnnOwner.close
						set cnnOwner = Nothing
					%>
			    </select>
	      	</div>
	      	<div class="col-lg-5 col-md-5 col-sm-12 col-xs-12">
	      		SHIFT+CLICK To Select Multiple Values
	      	</div>	      	
      	</div>
      	<!-- eof row !-->
	  	</div>
	  	<!-- eof right column !-->
	</div>
	</div>
	   	
  




   	
	
	
	<!-- exclusions !-->
  	<div class="container-fluid container-modal">
      	<div class="row">

      	<!-- left column !-->
      	<div class="col-lg-2 col-md-3 col-sm-12 col-xs-12 left-column <% If ReportSpecificData10 <> "" Then Response.Write("activefilter") %>">
	      	<h4><br>Created By<br><br><button type="button" class="btn btn-info btn-xs btn-block" id="btnClearCreatedBy">clear selections</button>
	      	<button type="button" class="btn btn-success btn-xs btn-block" id="btnSelectAllCreatedBy">select all</button></h4>
      	</div>
      	<!-- eof left column !-->

      	<!-- right column !-->
      	<div class="col-lg-10 col-md-9 col-sm-12 col-xs-12 right-column">
  	
      	<!-- row !-->
	      	<div class="row">

	      	<div class="col-lg-7 col-md-7 col-sm-12 col-xs-12">
		      	<select class="form-control" name="selProspectCreatedByUserNo" id="selProspectCreatedByUserNo" multiple="multiple">
	  	  			<%

	     				If ReportSpecificData10 <> "" Then
		     				selectedOwnersToFilterArray = Split(ReportSpecificData10,",")
		     				upperBoundCreatedByUsers = ubound(selectedCreatedByUsersToFilterArray)
		     			Else
		     				upperBoundCreatedByUsers = -1
		     				selectedCreatedByUsersToFilterArray = ""
		     			End If
	  	  				  	  			
			      	  	SQLCreatedBy = "SELECT DISTINCT(CreatedByUserNo) FROM PR_Prospects WHERE Pool='Dead' ORDER BY CreatedByUserNo DESC"
			
						Set cnnCreatedBy = Server.CreateObject("ADODB.Connection")
						cnnCreatedBy.open (Session("ClientCnnString"))
						Set rsCreatedBy = Server.CreateObject("ADODB.Recordset")
						rsCreatedBy.CursorLocation = 3 
						Set rsCreatedBy = cnnCreatedBy.Execute(SQLCreatedBy)
					
						If not rsCreatedBy.EOF Then
							Do
						      	createdByUserIsSelected = "False"
						      	
						      	For i=0 to upperBoundCreatedByUsers
								   If cInt(selectedCreatedByUsersToFilterArray(i)) = cInt(rsCreatedBy("CreatedByUserNo")) Then
								   		createdByUserIsSelected = "True"
								   End If
								Next
								
								%>
								<option value="<%= rsCreatedBy("CreatedByUserNo") %>" <% If createdByUserIsSelected = "True" Then Response.Write("selected") %>><%= GetUserFirstAndLastNameByUserNo(rsCreatedBy("CreatedByUserNo")) %></option>
								<%
								
								rsCreatedBy.movenext
							Loop until rsCreatedBy.eof
						End If
						set rsCreatedBy = Nothing
						cnnCreatedBy.close
						set cnnCreatedBy = Nothing
					%>
			    </select>
	      	</div>
	      	<div class="col-lg-5 col-md-5 col-sm-12 col-xs-12">
	      		SHIFT+CLICK To Select Multiple Values
	      	</div>	      	
      	</div>
      	<!-- eof row !-->
	  	</div>
	  	<!-- eof right column !-->
	</div>
	</div>
	




   	
      	<!-- filtering !-->
      	<div class="container-fluid container-modal">
	      	<div class="row">
      	
	      	<!-- left column !-->
	      	<div class="col-lg-2 col-md-3 col-sm-12 col-xs-12 left-column <% If ReportSpecificData11 <> "" Then Response.Write("activefilter") %>">
 		      	<h4><br>Created Date<br><br><button type="button" class="btn btn-info btn-xs btn-block" id="btnClearCreatedDate">clear selections</button></h4>
	      	</div>
	      	<!-- eof left column !-->
      	
	      	<!-- right column !-->
	      	<div class="col-lg-10 col-md-9 col-sm-12 col-xs-12 right-column">
	      	<!-- row !-->
		      	<div class="row">
				    <div class="col-xs-12 col-sm-12 col-md-12 col-lg-12" style="margin-bottom:10px;">
				    	<div class="col-lg-12" style="margin-left:-15px;">
					    	<input type="radio" id="optProspectCreatedDateRange" name="optProspectCreatedDateRange" value="setrange" <% If ReportSpecificData11 <> "" AND ReportSpecificData12 <> "" Then Response.write("checked='checked'")%>>
					    	
							<%
							If ReportSpecificData11 <> "" AND ReportSpecificData12 <> "" Then
								startDateProspectCreatedRange = dateCustomFormat(ReportSpecificData11)
								endDateProspectCreatedRange = dateCustomFormat(ReportSpecificData12) 
							Else
								thirtyDaysAgo = DateAdd("d",-30, DateSerial(Year(Now()), Month(Now()), Day(Now())))
								startDateProspectCreatedRange = dateCustomFormat(thirtyDaysAgo)
								endDateProspectCreatedRange = dateCustomFormat(Now())
							End If
	
							%>
							Where prospect was created in the past date range:
						</div> 
	
		              </div>     
		  		</div>
		  		
				 
		      	<div class="row" style="margin-top:15px;">
				    <div class="col-xs-12 col-sm-12 col-md-12 col-lg-12" style="margin-left:15px;">

				    	<div class="col-lg-5" style="margin-left:-15px;">
							Choose a "Quick Pick" Date Range<br>
							<select class="form-control" name="selProspectCreatedDateRangeCustom" id="selProspectCreatedDateRangeCustom">
								<option value="">Select A Date Range</option>
								<option value="Today">Today</option>
								<option value="This Week">This Week</option>
								<option value="Last Week">Last Week</option>
								<option value="Past 10 Days">Past 10 Days</option>
								<option value="Past 15 Days">Past 15 Days</option>
								<option value="Past 30 Days">Past 30 Days</option>
								<option value="This Month">This Month</option>
								<option value="Last Month">Last Month</option>
								<option value="Year To Date">Year To Date</option>
								<option value="Last Year">Last Year</option>
								<option value="All Dates">All Dates</option>
							</select>
						</div> 
						<div class="col-lg-1"><strong>OR</strong></div>
						<div class="col-lg-6">
							Use Calendar To Select Dates<br>					
                            <input type="text" id="prospectCreatedRange" class="form-control">
                            <i class="glyphicon glyphicon-calendar fa fa-calendar stagerangedatepicker"></i>	
                        </div>	
						<input type="hidden" id="txtProspectCreatedRangeStartDate" name="txtProspectCreatedRangeStartDate" value="<%= startDateProspectCreatedRange %>">
						<input type="hidden" id="txtProspectCreatedRangeEndDate" name="txtProspectCreatedRangeEndDate" value="<%= endDateProspectCreatedRange %>">                        
				    </div>	
				 </div>
		  		
		  	<!-- eof row !-->
  	</div>
  	<!-- eof right column !-->
  		</div>
  	<!-- eof row !-->
  	</div>
  	<!-- eof right column !-->
	








  	<!-- filtering !-->
  	<div class="container-fluid container-modal">
      	<div class="row">
  	
      	<!-- left column !-->
      	<div class="col-lg-2 col-md-3 col-sm-12 col-xs-12 left-column <% If ReportSpecificData13 <> "" Then Response.Write("activefilter") %>">
	      	<h4><br># Employees<br><br><button type="button" class="btn btn-info btn-xs btn-block" id="btnClearNumEmployees">clear selections</button></h4>
      	</div>
      	<!-- eof left column !-->
      	
      	<%
      	
     		If ReportSpecificData13 <> "" Then
 				employeeFilterArray = Split(ReportSpecificData13,",")
 				uBoundEmployees = uBound(employeeFilterArray)
  			Else
 				employeeFilterArray = ""
 				uBoundEmployees = -1
 			End If
     		
     		If uBoundEmployees > 0 Then
     		
	     		selectedEmployeeFilterType = employeeFilterArray(0)
	     		
	     		If selectedEmployeeFilterType = "ByPredefinedRange" Then
	     			selectedEmployeeRange = employeeFilterArray(1)
	     			selectedEmployeeCompOperator = ""
	     			selectedEmployeeCompNumber = ""
	     			selectedEmployeeRangeNumber1 = ""
	     			selectedEmployeeRangeNumber2 = ""
	     		End If
	     		
	     		If selectedEmployeeFilterType = "ByCustomNumber" Then
	     			
		 			selectedEmployeeCompOperator = employeeFilterArray(1)
					selectedEmployeeCompNumber = employeeFilterArray(2)
					selectedEmployeeRange = ""
	     			selectedEmployeeRangeNumber1 = ""
	     			selectedEmployeeRangeNumber2 = ""				
	     		End If
	     		
	     		If selectedEmployeeFilterType = "ByCustomRange" Then
		   			selectedEmployeeRangeNumber1 = employeeFilterArray(1)
					selectedEmployeeRangeNumber2 = employeeFilterArray(2)
					selectedEmployeeRange = ""
	     			selectedEmployeeCompOperator = ""
	     			selectedEmployeeCompNumber = ""				
	     		End If
	     	Else
	   			selectedEmployeeFilterType = ""
	   			selectedEmployeeRangeNumber1 = ""
				selectedEmployeeRangeNumber2 = ""
				selectedEmployeeRange = ""
     			selectedEmployeeCompOperator = ""
     			selectedEmployeeCompNumber = ""			     	
	     	End If

      	%>
  	
      	<!-- right column !-->
      	<div class="col-lg-10 col-md-9 col-sm-12 col-xs-12 right-column">
      	<!-- row !-->
	      	<div class="row">
		      	<div class="col-lg-3 col-md-3 col-sm-12 col-xs-12">
			      	<input type="radio" id="optNumEmployeesRangeCompare" name="optNumEmployeesRangeCompare" value="ByPredefinedRange" <% If selectedEmployeeFilterType = "ByPredefinedRange" Then Response.write("checked") %>>&nbsp;&nbsp;<strong>By Range</strong>
		      	</div>
		      	<div class="col-lg-3 col-md-3 col-sm-12 col-xs-12">
			      	<select class="form-control" name="selEmployeeRangeNo" id="selEmployeeRangeNo">
			      	<%
						SQLEmployees = "SELECT *, Cast(LEFT(Range,CHARINDEX('-',Range)-1) as int) as Expr1 FROM PR_EmployeeRangeTable ORDER BY Expr1"
	
						Set cnnEmployees = Server.CreateObject("ADODB.Connection")
						cnnEmployees.open (Session("ClientCnnString"))
						Set rsEmployees = Server.CreateObject("ADODB.Recordset")
						rsEmployees.CursorLocation = 3 
						Set rsEmployees = cnnEmployees.Execute(SQLEmployees)
							
						If not rsEmployees.EOF Then
							Do
								If selectedEmployeeRange <> "" Then
									%><option value="<%= rsEmployees("InternalRecordIdentifier") %>" <% If cInt(selectedEmployeeRange) = cInt(rsEmployees("InternalRecordIdentifier")) Then Response.Write("selected") %>><%= rsEmployees("Range") %></option><%
								Else
									%><option value="<%= rsEmployees("InternalRecordIdentifier") %>"><%= rsEmployees("Range") %></option><%
								End If
								rsEmployees.movenext
							Loop until rsEmployees.eof
						End If
						set rsEmployees = Nothing
						cnnEmployees.close
						set cnnEmployees = Nothing
			      	%>
					</select>
		      	</div>
	      	</div>
	      	<!-- eof row !-->
		      	
	      	<!-- row !-->
	      	<div class="row">
 	
		      	<div class="col-lg-3 col-md-3 col-sm-12 col-xs-12">
			      	<input type="radio" id="optNumEmployeesRangeCompare" name="optNumEmployeesRangeCompare" value="ByCustomNumber" <% If selectedEmployeeFilterType = "ByCustomNumber" Then Response.write("checked") %>>&nbsp;&nbsp;<strong>By Number</strong>
		      	</div>
		      	
			      	<div class="col-lg-7 col-md-7 col-sm-12 col-xs-12">
				      	<select class="form-control" name="selEmployeeRangeComparisonOperator" id="selEmployeeRangeComparisonOperator" style="display:inline-block; width:20%">
				      		<option value="greater than" <% If selectedEmployeeCompOperator = "greater than" Then Response.Write("selected") %>>&#707;</option>
				      		<option value="greater than or equal to" <% If selectedEmployeeCompOperator = "greater than or equal to" Then Response.Write("selected") %>>&#707;&#61;</option>
				      		<option value="less than" <% If selectedEmployeeCompOperator = "less than" Then Response.Write("selected") %>>&#706;</option>
				      		<option value="less than or equal to" <% If selectedEmployeeCompOperator = "less than or equal to" Then Response.Write("selected") %>>&#706;&#61;</option>
						</select>
						
						<input type="text" class="form-control" name="txtEmployeeRangeComparisonNumberSingle" value="<%= selectedEmployeeCompNumber %>" id="txtEmployeeRangeComparisonNumberSingle" style="width:20%; display:inline-block !important;"> employees.

			      	</div>
		      	</div>
		      	<!-- eof row !-->
		     
	      	
	      	
	      	<!-- row !-->
	      	<div class="row">
 	
		      	<div class="col-lg-3 col-md-4 col-sm-12 col-xs-12">
			      	<input type="radio" id="optNumEmployeesRangeCompare" name="optNumEmployeesRangeCompare" value="ByCustomRange" <% If selectedEmployeeFilterType = "ByCustomRange" Then Response.write("checked") %>>&nbsp;&nbsp;<strong>By Custom Range</strong>
		      	</div>
		      	
			      	<div class="col-lg-7 col-md-6 col-sm-12 col-xs-12">
				      	Between
						<input type="text" class="form-control" name="txtEmployeeCustomRangeNumber1" value="<%= selectedEmployeeRangeNumber1 %>" id="txtEmployeeCustomRangeNumber1" style="width:20%; display:inline-block !important;">
						and 
						<input type="text" class="form-control" name="txtEmployeeCustomRangeNumber2" value="<%= selectedEmployeeRangeNumber2 %>" id="txtEmployeeCustomRangeNumber2" style="width:20%; display:inline-block !important;">
						employees.
			      	</div>
		      	</div>
		      	<!-- eof row !-->
		     
	      	</div>
	      	<!-- eof right column !-->
	      	
      	</div>
   	</div>
   	








  	<!-- filtering !-->
  	<div class="container-fluid container-modal">
      	<div class="row">
  	
      	<!-- left column !-->
      	<div class="col-lg-2 col-md-3 col-sm-12 col-xs-12 left-column <% If ReportSpecificData14 <> "" Then Response.Write("activefilter") %>">
	      	<h4><br># Pantries<br><br><button type="button" class="btn btn-info btn-xs btn-block" id="btnClearNumPantries">clear selections</button></h4>
      	</div>
      	<!-- eof left column !-->
      	
 		<%
     		If ReportSpecificData14 <> "" Then
 				PantryFilterArray = Split(ReportSpecificData14,",")
 				uBoundPantries = uBound(PantryFilterArray)
  			Else
 				PantryFilterArray = ""
 				uBoundPantries = -1
 			End If

     		If uBoundPantries > 0 Then
     		
	     		selectedPantryFilterType = PantryFilterArray(0)
	     		     		
	     		If selectedPantryFilterType = "ByCustomNumber" Then
		 			selectedPantryCompOperator = PantryFilterArray(1)
					selectedPantryCompNumber = PantryFilterArray(2)
	     			selectedPantryRangeNumber1 = ""
	     			selectedPantryRangeNumber2 = ""				
	     		End If
	     		
	     		If selectedPantryFilterType = "ByCustomRange" Then
		   			selectedPantryRangeNumber1 = PantryFilterArray(1)
					selectedPantryRangeNumber2 = PantryFilterArray(2)
	     			selectedPantryCompOperator = ""
	     			selectedPantryCompNumber = ""				
	     		End If
	     	Else
	     		selectedPantryFilterType = ""
	   			selectedPantryRangeNumber1 = ""
				selectedPantryRangeNumber2 = ""
     			selectedPantryCompOperator = ""
     			selectedPantryCompNumber = ""		     	
	     	End If

 		%>
  	
      	<!-- right column !-->
      	<div class="col-lg-10 col-md-9 col-sm-12 col-xs-12 right-column">
		      	
	      	<!-- row !-->
	      	<div class="row">
 	
		      	<div class="col-lg-3 col-md-3 col-sm-12 col-xs-12">
			      	<input type="radio" id="optNumPantriesCompare" name="optNumPantriesCompare" value="ByCustomNumber" <% If selectedPantryFilterType = "ByCustomNumber" Then Response.write("checked") %>>&nbsp;&nbsp;<strong>By Number</strong>
		      	</div>
		      	
			      	<div class="col-lg-7 col-md-7 col-sm-12 col-xs-12">
				      	<select class="form-control" name="selNumPantriesComparisonOperator" id="selNumPantriesComparisonOperator" style="display:inline-block; width:20%">
				      		<option value="equal to" <% If selectedPantryCompOperator = "equal to" Then Response.Write("selected") %>>&#61;</option>
				      		<option value="greater than" <% If selectedPantryCompOperator = "greater than" Then Response.Write("selected") %>>&#707;</option>
				      		<option value="greater than or equal to" <% If selectedPantryCompOperator = "greater than or equal to" Then Response.Write("selected") %>>&#707;&#61;</option>
				      		<option value="less than" <% If selectedPantryCompOperator = "less than" Then Response.Write("selected") %>>&#706;</option>
				      		<option value="less than or equal to" <% If selectedPantryCompOperator = "less than or equal to" Then Response.Write("selected") %>>&#706;&#61;</option>
						</select>
						
						<input type="text" class="form-control" value="<%= selectedPantryCompNumber %>" name="txtNumPantriesComparisonNumberSingle" id="txtNumPantriesComparisonNumberSingle" style="width:20%; display:inline-block !important;"> pantries.

			      	</div>
		      	</div>
		      	<!-- eof row !-->
		     
	      	
	      	
	      	<!-- row !-->
	      	<div class="row">
 	
		      	<div class="col-lg-3 col-md-4 col-sm-12 col-xs-12">
			      	<input type="radio" id="optNumPantriesCompare" name="optNumPantriesCompare" value="ByCustomRange" <% If selectedPantryFilterType = "ByCustomRange" Then Response.write("checked") %>>&nbsp;&nbsp;<strong>By Custom Range</strong>
		      	</div>
		      	
			      	<div class="col-lg-7 col-md-6 col-sm-12 col-xs-12">
				      	Between
						<input type="text" class="form-control" value="<%= selectedPantryRangeNumber1 %>" name="txtNumPantriesCustomRangeNumber1" id="txtNumPantriesCustomRangeNumber1" style="width:20%; display:inline-block !important;">
						and 
						<input type="text" class="form-control" value="<%= selectedPantryRangeNumber2 %>" name="txtNumPantriesCustomRangeNumber2" id="txtNumPantriesCustomRangeNumber2" style="width:20%; display:inline-block !important;">
						pantries.
			      	</div>
		      	</div>
		      	<!-- eof row !-->
		     
	      	</div>
	      	<!-- eof right column !-->
	      	
      	</div>
   	</div>
   	
   	
   	
  	

	   	
</div>
<style type="text/css">
	.datepicker.dropdown-menu {right: auto;}
</style>


<script type="text/javascript" src="http://cdn.jsdelivr.net/momentjs/latest/moment.min.js"></script>
<script type="text/javascript" src="http://cdn.jsdelivr.net/bootstrap.daterangepicker/2/daterangepicker.js"></script>
<link rel="stylesheet" type="text/css" href="http://cdn.jsdelivr.net/bootstrap.daterangepicker/2/daterangepicker.css">

<script type="text/javascript">

		startDate = moment();
		endDate = moment();

			
		startDateStageNotChangedRange = moment(Date.parse($('#txtStageUnqualifiedRangeStartDate').val()));
		endDateStageNotChangedRange = moment(Date.parse($('#txtStageUnqualifiedRangeEndDate').val()));
			
		startDateStageLostRange = moment(Date.parse($('#txtStageLostDateRangeStartDate').val()));
		endDateStageLostRange = moment(Date.parse($('#txtStageLostDateRangeEndDate').val()));


		startDateProspectCreatedRange = moment(Date.parse($('#txtProspectCreatedRangeStartDate').val()));
		endDateProspectCreatedRange = moment(Date.parse($('#txtProspectCreatedRangeEndDate').val()));
		

		$('#stageRangeLost').daterangepicker({
                opens: 'right',
                startDate: startDateStageNotChangedRange,
                endDate: endDateStageNotChangedRange,
                alwaysShowCalendars: true,
                timePicker: false,
                linkedCalendars: false,
                autoUpdateInput:true,
                autoApply:true,
                showWeekNumbers: true,
                showClear: true,
                buttonClasses: ['btn'],
                applyClass: 'green',
                cancelClass: 'default',
                format: 'MM/DD/YYYY',
                separator: ' to ',
                locale: {
                    applyLabel: 'Apply',
                    fromLabel: 'From',
                    toLabel: 'To',
                    daysOfWeek: ['Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa'],
                    monthNames: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
                    firstDay: 1
                }
            },
            function (start, end) {
                $('#txtStageLostDateRangeStartDate').val(start.format('MM/DD/YYYY'));
                $('#txtStageLostDateRangeEndDate').val(end.format('MM/DD/YYYY'));
				$("#optStageLostDateRange").prop("checked","checked");
				$("#selLostStageDateRangeCustom").val('');
            }
        );


		$('#stageRangeUnqualified').daterangepicker({
                opens: 'right',
                startDate: startDateStageLostRange,
                endDate: endDateStageLostRange,
                alwaysShowCalendars: true,
                timePicker: false,
                linkedCalendars: false,
                autoUpdateInput:true,
                autoApply:true,
                showWeekNumbers: true,
                showClear: true,
                ranges: {
                    'Today': [moment(), moment()],
                    'Next 30 Days': [moment(), moment().add(29, 'days')],
                    'This Month': [moment().startOf('month'), moment().endOf('month')],
                    'Next Month': [moment().add(1, 'month').startOf('month'), moment().add(1, 'month').endOf('month')],
                    'All Dates': [moment().set({'year': 2014, 'month': 0}), moment()],
                },
                buttonClasses: ['btn'],
                applyClass: 'green',
                cancelClass: 'default',
                format: 'MM/DD/YYYY',
                separator: ' to ',
                locale: {
                    applyLabel: 'Apply',
                    fromLabel: 'From',
                    toLabel: 'To',
                    customRangeLabel: 'Custom Range',
                    daysOfWeek: ['Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa'],
                    monthNames: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
                    firstDay: 1
                }
            },
            function (start, end) {				
                $('#txtStageUnqualifiedRangeStartDate').val(start.format('MM/DD/YYYY'));
                $('#txtStageUnqualifiedRangeEndDate').val(end.format('MM/DD/YYYY'));
				$("#optStageUnqualifiedDateRange").prop("checked","checked");
				$("#selUnqualifiedStageDateRangeCustom").val('');			
            }
        );



		$('#prospectCreatedRange').daterangepicker({
                opens: 'right',
                startDate: startDateProspectCreatedRange,
                endDate: endDateProspectCreatedRange,
                alwaysShowCalendars: true,
                timePicker: false,
                linkedCalendars: false,
                autoUpdateInput:true,
                autoApply:true,
                showWeekNumbers: true,
                showClear: true,
                ranges: {
                    'Today': [moment(), moment()],
                    'Next 30 Days': [moment(), moment().add(29, 'days')],
                    'This Month': [moment().startOf('month'), moment().endOf('month')],
                    'Next Month': [moment().add(1, 'month').startOf('month'), moment().add(1, 'month').endOf('month')],
                    'All Dates': [moment().set({'year': 2014, 'month': 0}), moment()],
                },
                buttonClasses: ['btn'],
                applyClass: 'green',
                cancelClass: 'default',
                format: 'MM/DD/YYYY',
                separator: ' to ',
                locale: {
                    applyLabel: 'Apply',
                    fromLabel: 'From',
                    toLabel: 'To',
                    customRangeLabel: 'Custom Range',
                    daysOfWeek: ['Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa'],
                    monthNames: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
                    firstDay: 1
                }
            },
            function (start, end) {
                $('#txtProspectCreatedRangeStartDate').val(start.format('MM/DD/YYYY'));
                $('#txtProspectCreatedRangeEndDate').val(end.format('MM/DD/YYYY'));
				$("#optProspectCreatedDateRange").prop("checked","checked");
				$("#selProspectCreatedDateRangeCustom").val('');
            }
        );

 
        //Set the initial state of the picker label
		$('#txtStageUnqualifiedRangeStartDate').val(startDateStageNotChangedRange.format('MM/DD/YYYY'));
		$('#txtStageUnqualifiedRangeEndDate').val(endDateStageNotChangedRange.format('MM/DD/YYYY'));
		$('#txtStageLostDateRangeStartDate').val(startDateStageLostRange.format('MM/DD/YYYY'));
		$('#txtStageLostDateRangeEndDate').val(endDateStageLostRange.format('MM/DD/YYYY'));
		$('#txtProspectCreatedRangeStartDate').val(startDateProspectCreatedRange.format('MM/DD/YYYY'));
		$('#txtProspectCreatedRangeEndDate').val(endDateProspectCreatedRange.format('MM/DD/YYYY'));

		
</script>


      <!-- eof content insertion !-->
      
     <div class="modal-footer">
	     <button type="button" id="cancelFiltersBtn" class="btn btn-default" data-dismiss="modal">Cancel</button>
	     <button type="submit" id="saveFiltersBtn" class="btn btn-primary" >Apply Filters</button>
	     <div id="searchingimageDiv" style="display:none; margin-top:-15px;" class="pull-right">Applying Prospect Filters...<img id="searchingimage1" src="../img/preloader.gif" alt="" /></div>
     </div>
			</form>
		</div>
	</div>
</div>
	<!-- eof modal box !-->
	



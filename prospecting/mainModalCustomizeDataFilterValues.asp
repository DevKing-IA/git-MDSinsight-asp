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
	    
	    
	function checkStageCheckboxes()
	{
	    var stageCheckboxes =document.getElementsByName("chkStage");
	    var atLeastOneStageChecked=false;
	    for(var i=0,l=stageCheckboxes.length;i<l;i++)
	    {
	        if(stageCheckboxes[i].checked)
	        {
	            atLeastOneStageChecked = true;
	            break;
	        }
	    }
	    if (atLeastOneStageChecked) 
	    	{return true;}
	    else 
	    	{return false;}
	}	
	
	
	
	function isOneChecked() {
	    return ($('[name="chkStage"]:checked').length > 0);
	}
	
	function isOneActivityChecked() {
	    return ($('[name="chkNextActivity"]:checked').length > 0);
	}
	
	
	function isInteger(str) {
    	var r = /^-?[0-9]*[1-9][0-9]*$/;
    	return r.test(str);
	}
	
	

	function validateCustomizeForm()	{
		
   		

		if ($('#optNextActivityScheduledDateRange').is(':checked'))
		{
			if ($("#selNextActivityScheduledDateRangeCustom").find(":selected").val() == "") 
			{
				//Date Range Checking Here
				if (document.frmProspectingCustomizeDataFilters.txtNextActivityScheduledDateRangeStartDate.value == "" || 
					document.frmProspectingCustomizeDataFilters.txtNextActivityScheduledDateRangeEndDate.value == "") 
					{		
						swal("Please make sure both next activity filter dates are selected or quick pick range has been selected.")
						return false;
					}		
			}
   		}


		
		if ($('#optStageChangeDatesNotChangedDays').is(':checked'))
		{
			if (document.frmProspectingCustomizeDataFilters.txtStageNotChangedDays.value == "") 
				{		
					swal("Please specify days that stage HAS NOT changed in.")
					return false;
				}
			else 
			{
				if (!isInteger(document.frmProspectingCustomizeDataFilters.txtStageNotChangedDays.value))
				{		
					swal("Please enter a whole number for days that stage HAS NOT changed in.")
					return false;
				}
			}
   		}
    	
    	
 		if ($('#optStageChangeDatesChangedDays').is(':checked'))
		{
			if (document.frmProspectingCustomizeDataFilters.txtStageChangedDays.value == "") 
				{		
					swal("Please specify days that stage HAS changed in.")
					return false;
				}
			else 
			{
				if (!isInteger(document.frmProspectingCustomizeDataFilters.txtStageChangedDays.value))
				{		
					swal("Please enter a whole number for days that stage HAS changed in.")
					return false;
				}
			}
   		}
   


		if ($('#optStageNotChangedDateRange').is(':checked'))
		{
			if ($("#selStageNotChangedDateRangeCustom").find(":selected").val() == "") 
			{
				//Date Range Checking Here
				if (document.frmProspectingCustomizeDataFilters.txtStageNotChangedDateRangeStartDate.value == "" || 
					document.frmProspectingCustomizeDataFilters.txtStageNotChangedDateRangeEndDate.value == "") 
					{		
						swal("Please make sure both stage filter dates are filled in or quick pick range has been selected.")
						return false;
					}		
			}
			else
			{
				//swal("Please select at least one stage for stage date filtering.")
				//return false;
			}
   		}



		if ($('#optStageChangeDateRange').is(':checked'))
		{
			if ($("#selStageChangedDateRangeCustom").find(":selected").val() == "") 
			{
				//Date Range Checking Here
				if (document.frmProspectingCustomizeDataFilters.txtStageChangedDateRangeStartDate.value == "" || 
					document.frmProspectingCustomizeDataFilters.txtStageChangedDateRangeEndDate.value == "") 
					{		
						swal("Please make sure both stage filter dates are filled in or quick pick range has been selected.")
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
		   
		$('input:radio[name="optStageChangeDatesNotChangedDays"]').change(
		    function(){
		        if (this.checked) {
		            // note that, as per comments, the 'changed' <input> will *always* be checked, as the change
		            // event only fires on checking an <input>, not on un-checking it.
		            $('input:radio[name="optStageChangeDatesChangedDays"]').removeAttr("checked");
		            $('input:radio[name="optStageNotChangedDateRange"]').removeAttr("checked");
		            $('input:radio[name="optStageChangeDateRange"]').removeAttr("checked");
		        }
		    });		    
		$('input:radio[name="optStageChangeDatesChangedDays"]').change(
		    function(){
		        if (this.checked) {
		            // note that, as per comments, the 'changed' <input> will *always* be checked, as the change
		            // event only fires on checking an <input>, not on un-checking it.
		            $('input:radio[name="optStageChangeDatesNotChangedDays"]').removeAttr("checked");
		            $('input:radio[name="optStageNotChangedDateRange"]').removeAttr("checked");
		            $('input:radio[name="optStageChangeDateRange"]').removeAttr("checked");
		        }
		    });		    
		$('input:radio[name="optStageNotChangedDateRange"]').change(
		    function(){
		        if (this.checked) {
		            // note that, as per comments, the 'changed' <input> will *always* be checked, as the change
		            // event only fires on checking an <input>, not on un-checking it.
		            $('input:radio[name="optStageChangeDatesNotChangedDays"]').removeAttr("checked");
		            $('input:radio[name="optStageChangeDatesChangedDays"]').removeAttr("checked");
		            $('input:radio[name="optStageChangeDateRange"]').removeAttr("checked");
		        }
		    });		    
		$('input:radio[name="optStageChangeDateRange"]').change(
		    function(){
		        if (this.checked) {
		            // note that, as per comments, the 'changed' <input> will *always* be checked, as the change
		            // event only fires on checking an <input>, not on un-checking it.
		            $('input:radio[name="optStageChangeDatesNotChangedDays"]').removeAttr("checked");
		            $('input:radio[name="optStageChangeDatesChangedDays"]').removeAttr("checked");
		            $('input:radio[name="optStageNotChangedDateRange"]').removeAttr("checked");
		        }
		    });	
		    

		$("#selStageNotChangedDateRangeCustom").change(function() {
		
			selectedValue = $(this).find(":selected").val();
			
		    if (selectedValue != "") 
		    {
		    	$("#stageRange1").val('');
		    	$("#txtStageNotChangedDateRangeStartDate").val('');
		    	$("#txtStageNotChangedDateRangeEndDate").val('');
		    	$("#optStageNotChangedDateRange").prop("checked","checked");
		    }
		    else
		    {
		    	$("#optStageNotChangedDateRange").prop("checked","");
		    }

		});



		$("#selStageChangedDateRangeCustom").change(function() {
		
			selectedValue = $(this).find(":selected").val();
			
		    if (selectedValue != "") 
		    {
		    	$("#stageRange2").val('');
		    	$("#txtStageChangedDateRangeStartDate").val('');
		    	$("#txtStageChangedDateRangeEndDate").val('');
		    	$("#optStageChangeDateRange").prop("checked","checked");
		    }
		    else
		    {
		    	$("#optStageChangeDateRange").prop("checked","");
		    }

		});


		    
		$("#selNextActivityScheduledDateRangeCustom").change(function() {
		
			selectedValue = $(this).find(":selected").val();
			
		    if (selectedValue != "") 
		    {
		    	$("#nextactivityRange1").val('');
		    	$("#txtNextActivityScheduledDateRangeStartDate").val('');
		    	$("#txtNextActivityScheduledDateRangeEndDate").val('');
		    	$("#optNextActivityScheduledDateRange").prop("checked","checked");
		    }
		    else
		    {
		    	$("#optNextActivityScheduledDateRange").prop("checked","");
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
	
		$("#btnClearLocation").click(function() {
			$("#txtCityFilter").val('');
			$("#txtStateFilter").val('');
			$("#txtZipFilter").val('');
		});		

		$("#btnClearNextActivity").click(function() {
		
			$(".checkNextActivity").removeAttr('checked');
			$('input[name="optNextActivityScheduledDateRange"]').prop('checked', false);
			$("#nextactivityRange1").val('');
			$("#txtNextActivityScheduledDateRangeStartDate").val('');
			$("#txtNextActivityScheduledDateRangeEndDate").val('');
			$('#selNextActivityScheduledDateRangeCustom').children().removeProp('selected');
		});	
	
		$("#btnClearStages").click(function() {
		
			$(".check").removeAttr('checked');
			$('input[name="optStageChangeDatesNotChangedDays"]').prop('checked', false);
			$("#txtStageNotChangedDays").val('');
			$('input[name="optStageChangeDatesChangedDays"]').prop('checked', false);
			$("#txtStageChangedDays").val('');
			$('input[name="optStageNotChangedDateRange"]').prop('checked', false);
			$("#txtStageNotChangedDateRangeStartDate").val('');
			$("#txtStageNotChangedDateRangeEndDate").val('');
			$('input[name="optStageChangeDateRange"]').prop('checked', false);
			$("#txtStageChangedDateRangeStartDate").val('');
			$("#txtStageChangedDateRangeEndDate").val('');
			$('#selStageNotChangedDateRangeCustom').children().removeProp('selected');
			$('#selStageChangedDateRangeCustom').children().removeProp('selected');
			$("#stageRange1").val('');
			$("#stageRange2").val('');
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
			$("#prospectCreatedRange").val('');
			$("#txtProspectCreatedRangeStartDate").val('');
			$("#txtProspectCreatedRangeEndDate").val('');
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



	});
</script>
<!-- eof modal scroll !-->

<%

Function dateCustomFormat(passeddate)
	x = FormatDateTime(passeddate, 2) 
	d = split(x, "/")
	dateCustomFormat = d(2) & "-" & d(0) & "-" & d(1)
End Function


%>
	
<!-- modal box !-->
<div class="modal fade bs-modal-filter-prospecting-data" tabindex="-1" role="dialog" aria-labelledby="myLargeModalLabel" aria-hidden="true">
	<div class="modal-dialog modal-lg modal-height">
		<div class="modal-content">
			<div class="modal-header">
				<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
				<h4 class="modal-title" id="myModalLabel" align="center">Customize and Filter <%= GetTerm("Prospecting") %> Screen</h4>
			</div>
			<%
			'************************
			'Read Settings_Reports
			'************************
			
			SQLReportName = Replace(MUV_READ("CRMVIEWSTATE"),"'","''")
			
			SQL = "SELECT * FROM Settings_Reports WHERE ReportNumber = 1400 AND UserNo = " & Session("userNo") & " AND PoolForProspecting = 'Live' AND UserReportName = '" & SQLReportName & "'"
			
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
			End If
			'****************************
			'End Read Settings_Reports
			'****************************
			%>
  
			<!-- eof modal scroll !-->
			
	<form method="post" action="mainCustomizeSaveDataFilterValues.asp" id="frmProspectingCustomizeDataFilters" name="frmProspectingCustomizeDataFilters" onsubmit="return validateCustomizeForm();">

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
	      	<div class="col-lg-2 col-md-3 col-sm-12 col-xs-12 left-column <% If ReportSpecificData19 <> "" OR ReportSpecificData20 <> "" Then Response.Write("activefilter") %>">
 		      	<h4><br>Next Activity<br><br><button type="button" class="btn btn-info btn-xs btn-block" id="btnClearNextActivity">clear selections</button></h4>
	      	</div>
	      	<!-- eof left column !-->
  	
	      	<!-- right column !-->
	      	<div class="col-lg-10 col-md-9 col-sm-12 col-xs-12 right-column check-row">
      	
	      	<!-- row !-->
		      	<div class="row categories-checkboxes">
      	
			      	<div class="col-lg-12">
					  <div class="checkbox">
						<label>
						  <input type="checkbox" class="checkNextActivity" id="checkAllNextActivity"> <strong>CHECK / UNCHECK ALL NEXT ACTIVITIES</strong>
						</label>
					  </div>
			      	</div>

					<div class="checkbox">
  
		     			<%
		     				If ReportSpecificData19 <> "" Then
			     				selectedNextActivityToFilterArray = Split(ReportSpecificData19,",")
			     				upperBoundNextActivity = ubound(selectedNextActivityToFilterArray)
			     			Else
			     				upperBoundNextActivity = -1
			     				selectedNextActivityToFilterArray = ""
			     			End If
							 
				      	  	SQLNextActivity = "SELECT * FROM PR_Activities ORDER BY Activity"

							Set cnnNextActivity = Server.CreateObject("ADODB.Connection")
							cnnNextActivity.open (Session("ClientCnnString"))
							Set rsNextActivity = Server.CreateObject("ADODB.Recordset")
							rsNextActivity.CursorLocation = 3 
							Set rsNextActivity = cnnNextActivity.Execute(SQLNextActivity)
							If not rsNextActivity.EOF Then
								Do
									%>
							      	<div class="col-lg-6 col-md-6 col-sm-12 col-xs-12 ">
								      	<label>
								      	<%
								      	activityIsSelected = "False"
								      	
								      	For i=0 to upperBoundNextActivity
										   If cInt(selectedNextActivityToFilterArray(i)) = cInt(rsNextActivity("InternalRecordIdentifier")) Then
										   		activityIsSelected = "True"
										   End If
										Next

								      	%>
								      	<input type="checkbox" class="checkNextActivity" <% If activityIsSelected = "True" Then Response.Write("checked='checked'") %> id="chkNextActivity<%= rsNextActivity("InternalRecordIdentifier") %>" name="chkNextActivity" value="<%= rsNextActivity("InternalRecordIdentifier") %>"><%= rsNextActivity("Activity") %><br>
								      	</label>
							      	</div>   
									<%
									rsNextActivity.movenext
								Loop until rsNextActivity.eof
							End If
							set rsNextActivity = Nothing
							cnnNextActivity.close
							set cnnNextActivity = Nothing
		     			%>
    			
		     			<!-- select all / deselect all checkboxes !-->
						<script type="text/javascript">
						$("#checkAllNextActivity").click(function () {
					    $(".checkNextActivity").prop('checked', $(this).prop('checked'));
						});	</script>

					</div> 		      			      	
		      	</div>
		      	<!-- eof row !-->
		      	
		      	<div class="row" style="margin-top:35px;">
		      	
				    
				    
				    <div class="col-xs-12 col-sm-12 col-md-12 col-lg-12" style="margin-bottom:10px;">
				    	<div class="col-lg-12" style="margin-left:-15px;">
					    	<input type="radio" id="optNextActivityScheduledDateRange" name="optNextActivityScheduledDateRange" value="NextActivityScheduledDateRange" <% If ReportSpecificData20="NextActivityScheduledDateRange" Then Response.write("checked") %>>
						
							<%
							If ReportSpecificData21 <> "" AND ReportSpecificData22 <> "" Then
								startDateNextActivityScheduledRange = dateCustomFormat(ReportSpecificData21)
								endDateNextActivityScheduledRange = dateCustomFormat(ReportSpecificData22) 
							Else
								thirtyDaysAgo = DateAdd("d",-30, DateSerial(Year(Now()), Month(Now()), Day(Now())))
								startDateNextActivityScheduledRange = dateCustomFormat(thirtyDaysAgo)
								endDateNextActivityScheduledRange = dateCustomFormat(Now())
							End If
	
							%>
							Where next activity is scheduled for:
						</div> 			        
					</div>
			      	<div class="row" style="margin-top:15px;">
					    <div class="col-xs-12 col-sm-12 col-md-12 col-lg-12" style="margin-left:25px;">
	
					    	<div class="col-lg-5" style="margin-left:-15px;">
								Choose a "Quick Pick" Date Range<br>
								<select class="form-control" name="selNextActivityScheduledDateRangeCustom" id="selNextActivityScheduledDateRangeCustom">
									<option value="">Select A Date Range</option>
									<option value="Today">Today</option>
									<option value="Tomorrow">Tomorrow</option>
									<option value="This Week">This Week</option>
									<option value="Next Week">Next Week</option>
									<option value="Next Week">Next Week</option>
									<option value="Next 10 Days">Next 10 Days</option>
									<option value="Next 15 Days">Next 15 Days</option>
									<option value="Next 30 Days">Next 30 Days</option>
									<option value="Past 10 Days">Past 10 Days</option>
									<option value="Past 15 Days">Past 15 Days</option>
									<option value="Past 30 Days">Past 30 Days</option>
									<option value="This Month">This Month</option>
									<option value="Last Month">Last Month</option>
									<option value="Next Month">Next Month</option>
									<option value="Year To Date">Year To Date</option>
									<option value="Last Year">Last Year</option>
									<option value="All Dates">All Dates</option>
								</select>
							</div> 
							<div class="col-lg-1"><strong>OR</strong></div>
							<div class="col-lg-6">
								Use Calendar To Select Dates<br>					
	                            <input type="text" id="nextactivityRange1" class="form-control">
	                            <i class="glyphicon glyphicon-calendar fa fa-calendar stagerangedatepicker"></i>	
	                        </div>	
							<input type="hidden" id="txtNextActivityScheduledDateRangeStartDate" name="txtNextActivityScheduledDateRangeStartDate" value="<%= startDateNextActivityScheduledRange %>">
							<input type="hidden" id="txtNextActivityScheduledDateRangeEndDate" name="txtNextActivityScheduledDateRangeEndDate" value="<%= endDateNextActivityScheduledRange %>">
							
					    </div>	
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
	      	<div class="col-lg-2 col-md-3 col-sm-12 col-xs-12 left-column <% If ReportSpecificData2 <> "" OR ReportSpecificData18 <> "" Then Response.Write("activefilter") %>">
 		      	<h4><br>Stages<br><br><button type="button" class="btn btn-info btn-xs btn-block" id="btnClearStages">clear selections</button></h4>
	      	</div>
	      	<!-- eof left column !-->
  	
	      	<!-- right column !-->
	      	<div class="col-lg-10 col-md-9 col-sm-12 col-xs-12 right-column check-row">
      	
	      	<!-- row !-->
		      	<div class="row categories-checkboxes">
      	
			      	<div class="col-lg-12">
					  <div class="checkbox">
						<label>
						  <input type="checkbox" class="check" id="checkAll"> <strong>CHECK / UNCHECK ALL STAGES</strong>
						</label>
					  </div>
			      	</div>

					<div class="checkbox">
  
		     			<%
		     				If ReportSpecificData18 <> "" Then
			     				selectedStagesToFilterArray = Split(ReportSpecificData18,",")
			     				upperBound = ubound(selectedStagesToFilterArray)
			     			Else
			     				upperBound = -1
			     				selectedStagesToFilterArray = ""
			     			End If
							 
				      		'Get all stages
				      	  	SQLStages = "SELECT * FROM PR_Stages WHERE StageType = 'Primary' OR StageType = 'Secondary' AND InternalRecordIdentifier <>0 ORDER BY SortOrder"
		
							Set cnnStages = Server.CreateObject("ADODB.Connection")
							cnnStages.open (Session("ClientCnnString"))
							Set rsStages = Server.CreateObject("ADODB.Recordset")
							rsStages.CursorLocation = 3 
							Set rsStages = cnnStages.Execute(SQLStages)
								
							If not rsStages.EOF Then
								Do
									%>
							      	<div class="col-lg-6 col-md-6 col-sm-12 col-xs-12 ">
								      	<label>
								      	<%
								      	stageIsSelected = "False"
								      	
								      	For i=0 to upperBound
										   If cInt(selectedStagesToFilterArray(i)) = cInt(rsStages("InternalRecordIdentifier")) Then
										   		stageIsSelected = "True"
										   End If
										Next

								      	%>
								      	<input type="checkbox" class="check" <% If stageIsSelected = "True" Then Response.Write("checked='checked'") %> id="chkStage<%= rsStages("InternalRecordIdentifier") %>" name="chkStage" value="<%= rsStages("InternalRecordIdentifier") %>"><%= rsStages("Stage") %><br>
								      	</label>
							      	</div>   
									<%
									rsStages.movenext
								Loop until rsStages.eof
							End If
		     			%>
    			
		     			<!-- select all / deselect all checkboxes !-->
						<script type="text/javascript">
						$("#checkAll").click(function () {
					    $(".check").prop('checked', $(this).prop('checked'));
						});	</script>

					</div> 		      			      	
		      	</div>
		      	<!-- eof row !-->
		      	
		      	<div class="row" style="margin-top:35px;">
		      	
			    	<div class="col-xs-12 col-sm-12 col-md-12 col-lg-12" style="margin-bottom:10px;">
				    	<input type="radio" id="optStageChangeDatesNotChangedDays" name="optStageChangeDatesNotChangedDays" value="HasNotChangedInXDays" <% If ReportSpecificData2="HasNotChangedInXDays" Then Response.write("checked") %>>
				    	Where stage <strong>HAS NOT</strong> changed in 
				    	
			    		<% If ReportSpecificData2 = "HasNotChangedInXDays" Then %>
								<% If ReportSpecificData3 <> "" Then %>
									<input type="text" class="form-control" name="txtStageNotChangedDays" id="txtStageNotChangedDays" style="width:10%; display:inline-block !important; height:30px;" value="<%= ReportSpecificData3 %>"> days.
								<% Else %>
									<input type="text" class="form-control" name="txtStageNotChangedDays" id="txtStageNotChangedDays" style="width:10%; display:inline-block !important; height:30px;"> days.
								<% End If %>
						<% Else %>
							<input type="text" class="form-control" name="txtStageNotChangedDays" id="txtStageNotChangedDays" style="width:10%; display:inline-block !important; height:30px;"> days.
						<% End If %>
				    </div>

			    	<div class="col-xs-12 col-sm-12 col-md-12 col-lg-12" style="margin-bottom:35px;">
				    	<input type="radio" id="optStageChangeDatesChangedDays" name="optStageChangeDatesChangedDays" value="HasChangedInXDays" <% If ReportSpecificData2="HasChangedInXDays" Then Response.write("checked") %>>
				    	Where stage <strong>HAS</strong> changed in the past 
				    	
			    		<% If ReportSpecificData2 = "HasChangedInXDays" Then %>
								<% If ReportSpecificData3 <> "" Then %>
									<input type="text" class="form-control" name="txtStageChangedDays" id="txtStageChangedDays" style="width:10%; display:inline-block !important; height:30px;" value="<%= ReportSpecificData3 %>"> days.
								<% Else %>
									<input type="text" class="form-control" name="txtStageChangedDays" id="txtStageChangedDays" style="width:10%; display:inline-block !important; height:30px;"> days.
								<% End If %>
						<% Else %>
							<input type="text" class="form-control" name="txtStageChangedDays" id="txtStageChangedDays" style="width:10%; display:inline-block !important; height:30px;"> days.
						<% End If %>
				    </div>
				    
				    
				    <div class="col-xs-12 col-sm-12 col-md-12 col-lg-12" style="margin-bottom:10px;">
				    	<div class="col-lg-12" style="margin-left:-15px;">
					    	<input type="radio" id="optStageNotChangedDateRange" name="optStageNotChangedDateRange" value="HasNotChangedInDateRange" <% If ReportSpecificData2="HasNotChangedInDateRange" Then Response.write("checked") %>>
							<%
							If ReportSpecificData4 <> "" AND ReportSpecificData5 <> "" AND ReportSpecificData2 = "HasNotChangedInDateRange" Then
								startDateStageNotChangedRange = dateCustomFormat(ReportSpecificData4)
								endDateStageNotChangedRange = dateCustomFormat(ReportSpecificData5) 
							Else
								thirtyDaysAgo = DateAdd("d",-30, DateSerial(Year(Now()), Month(Now()), Day(Now())))
								startDateStageNotChangedRange = dateCustomFormat(thirtyDaysAgo)
								endDateStageNotChangedRange = dateCustomFormat(Now())
							End If
	
							%>
							Where stage <strong>HAS NOT</strong> changed in the past date range:
						</div> 			        
					</div>
			      	<div class="row" style="margin-top:15px;">
					    <div class="col-xs-12 col-sm-12 col-md-12 col-lg-12" style="margin-left:25px;">
					    	<div class="col-lg-5" style="margin-left:-15px;">
								Choose a "Quick Pick" Date Range<br>
								<select class="form-control" name="selStageNotChangedDateRangeCustom" id="selStageNotChangedDateRangeCustom">
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
	                            <input type="text" id="stageRange1" class="form-control">
	                            <i class="glyphicon glyphicon-calendar fa fa-calendar stagerangedatepicker"></i>	
	                        </div>	
							<input type="hidden" id="txtStageNotChangedDateRangeStartDate" name="txtStageNotChangedDateRangeStartDate" value="<%= startDateStageNotChangedRange %>">
							<input type="hidden" id="txtStageNotChangedDateRangeEndDate" name="txtStageNotChangedDateRangeEndDate" value="<%= endDateStageNotChangedRange %>">
					    </div>	
					 </div>
					
					
					<br><br>
			        
				    <div class="col-xs-12 col-sm-12 col-md-12 col-lg-12">

				    	<div class="col-lg-12" style="margin-left:-15px;">
					    	<input type="radio" id="optStageChangeDateRange" name="optStageChangeDateRange" value="HasChangedInDateRange" <% If ReportSpecificData2="HasChangedInDateRange" Then Response.write("checked") %>>

							<%
							If ReportSpecificData4 <> "" AND ReportSpecificData5 <> "" AND ReportSpecificData2 = "HasChangedInDateRange" Then
								startDateStageChangedRange = dateCustomFormat(ReportSpecificData4)
								endDateStageChangedRange = dateCustomFormat(ReportSpecificData5) 
							Else
								thirtyDaysAgo = DateAdd("d",-30, DateSerial(Year(Now()), Month(Now()), Day(Now())))
								startDateStageChangedRange = dateCustomFormat(thirtyDaysAgo)
								endDateStageChangedRange = dateCustomFormat(Now())
							End If
	
							%>
							Where stage <strong>HAS</strong> changed in the past date range:
						</div> 
				    </div>
			      	<div class="row" style="margin-top:15px;">
					    <div class="col-xs-12 col-sm-12 col-md-12 col-lg-12" style="margin-left:25px;">
					    	<div class="col-lg-5" style="margin-left:-15px;">
								Choose a "Quick Pick" Date Range<br>
								<select class="form-control" name="selStageChangedDateRangeCustom" id="selStageChangedDateRangeCustom">
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
	                            <input type="text" id="stageRange2" class="form-control">
	                            <i class="glyphicon glyphicon-calendar fa fa-calendar stagerangedatepicker"></i>	
	                        </div>	
							<input type="hidden" id="txtStageChangedDateRangeStartDate" name="txtStageChangedDateRangeStartDate" value="<%= startDateStageChangedRange %>">
							<input type="hidden" id="txtStageChangedDateRangeEndDate" name="txtStageChangedDateRangeEndDate" value="<%= endDateStageChangedRange %>">
					    </div>	
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
			      	  	SQLOwner = SQLOwner & "WHERE UserNo IN (SELECT DISTINCT OwnerUserNo FROM PR_Prospects WHERE Pool='Live') "
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
	  	  				  	  			
			      	  	SQLCreatedBy = "SELECT DISTINCT(CreatedByUserNo) FROM PR_Prospects WHERE Pool='Live' ORDER BY CreatedByUserNo DESC"
			
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
		      	<div class="row" style="margin-top:15px;">
				    <div class="col-xs-12 col-sm-12 col-md-12 col-lg-12" style="margin-left:25px;">
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

			
		startDateStageNotChangedRange = moment(Date.parse($('#txtStageNotChangedDateRangeStartDate').val()));
		endDateStageNotChangedRange = moment(Date.parse($('#txtStageNotChangedDateRangeEndDate').val()));
			
		startDateStageChangedRange = moment(Date.parse($('#txtStageChangedDateRangeStartDate').val()));
		endDateStageChangedRange = moment(Date.parse($('#txtStageChangedDateRangeEndDate').val()));


		startDateProspectCreatedRange = moment(Date.parse($('#txtProspectCreatedRangeStartDate').val()));
		endDateProspectCreatedRange = moment(Date.parse($('#txtProspectCreatedRangeEndDate').val()));

		startDateNextActivityRange = moment(Date.parse($('#txtNextActivityScheduledDateRangeStartDate').val()));
		endDateNextActivityRange = moment(Date.parse($('#txtNextActivityScheduledDateRangeEndDate').val()));
		

		$('#stageRange1').daterangepicker({
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
                $('#txtStageNotChangedDateRangeStartDate').val(start.format('MM/DD/YYYY'));
                $('#txtStageNotChangedDateRangeEndDate').val(end.format('MM/DD/YYYY'));
				$("#optStageNotChangedDateRange").prop("checked","checked");
				$("#optStageChangeDateRange").prop("checked","");
				$("#optStageChangeDatesChangedDays").prop("checked","");
				$("#optStageChangeDatesNotChangedDays").prop("checked","");
				$("#selStageNotChangedDateRangeCustom").val('');
				
            }
        );


		$('#stageRange2').daterangepicker({
                opens: 'right',
                startDate: startDateStageChangedRange,
                endDate: endDateStageChangedRange,
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
                $('#txtStageChangedDateRangeStartDate').val(start.format('MM/DD/YYYY'));
                $('#txtStageChangedDateRangeEndDate').val(end.format('MM/DD/YYYY'));
				$("#optStageChangeDateRange").prop("checked","checked");
				$("#optStageNotChangedDateRange").prop("checked","");
				$("#optStageChangeDatesChangedDays").prop("checked","");
				$("#optStageChangeDatesNotChangedDays").prop("checked","");
				$("#selStageChangedDateRangeCustom").val('');
				
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



	$('#nextactivityRange1').daterangepicker({
                opens: 'right',
                startDate: startDateNextActivityRange,
                endDate: endDateNextActivityRange,
                alwaysShowCalendars: true,
                timePicker: false,
                linkedCalendars: false,
                autoUpdateInput:true,
                autoApply:true,
                showWeekNumbers: true,
                showClear: true,
                ranges: {
                    'Today': [moment(), moment()],
                    'Today And Older': [moment().subtract(365, 'days'), moment()],
                    '<%=ProspectActivityDefaultDaysToShow %> Days and Older': [moment().subtract(365, 'days'), moment().add(<%=ProspectActivityDefaultDaysToShow %>, 'days'),],
                    'Next 5 Days': [moment(), moment().add(4, 'days')],
                    'Next 10 Days': [moment(), moment().add(9, 'days')],
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
                $('#txtNextActivityScheduledDateRangeStartDate').val(start.format('MM/DD/YYYY'));
                $('#txtNextActivityScheduledDateRangeEndDate').val(end.format('MM/DD/YYYY'));
				$("#optNextActivityScheduledDateRange").prop("checked","checked");
            }
        );
       
        //Set the initial state of the picker label
		$('#txtStageNotChangedDateRangeStartDate').val(startDateStageNotChangedRange.format('MM/DD/YYYY'));
		$('#txtStageNotChangedDateRangeEndDate').val(endDateStageNotChangedRange.format('MM/DD/YYYY'));
		$('#txtStageChangedDateRangeStartDate').val(startDateStageChangedRange.format('MM/DD/YYYY'));
		$('#txtStageChangedDateRangeEndDate').val(endDateStageChangedRange.format('MM/DD/YYYY'));
		$('#txtProspectCreatedRangeStartDate').val(startDateProspectCreatedRange.format('MM/DD/YYYY'));
		$('#txtProspectCreatedRangeEndDate').val(endDateProspectCreatedRange.format('MM/DD/YYYY'));
		$('#txtNextActivityScheduledDateRangeStartDate').val(startDateNextActivityRange.format('MM/DD/YYYY'));
		$('#txtNextActivityScheduledDateRangeEndDate').val(endDateNextActivityRange.format('MM/DD/YYYY'));
		$("#selNextActivityScheduledDateRangeCustom").val('');

		
</script>


      <!-- eof content insertion !-->
      
     <div class="modal-footer">
	     <button type="button" id="cancelFiltersBtn" class="btn btn-default" data-dismiss="modal">Cancel</button>
	     <button type="submit" id="saveFiltersBtn" class="btn btn-primary">Apply Filters</button>
	     <div id="searchingimageDiv" style="display:none; margin-top:-15px;" class="pull-right">Applying Prospect Filters...<img id="searchingimage1" src="../img/preloader.gif" alt="" /></div>
     </div>
			</form>
		</div>
	</div>
</div>
	<!-- eof modal box !-->
	



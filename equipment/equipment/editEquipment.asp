<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/InSightFuncs_Equipment.asp"-->


<!-- datetime picker !-->
<script src="<%= baseURL %>js/moment.min.js" type="text/javascript"></script>
<script type="text/javascript" src="<%= BaseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.min.js"></script>
<link href="<%= BaseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.css" rel="stylesheet">
<!-- end datetime picker !-->

<% 

InternalRecordIdentifier = Request.QueryString("i") 
EquipIntRecID = Request.QueryString("i")
If InternalRecordIdentifier = "" Then Response.Redirect("findEquipment.asp")


SQLEquipment = "SELECT * FROM EQ_Equipment where InternalRecordIdentifier = " & InternalRecordIdentifier 

Set cnnEquipment = Server.CreateObject("ADODB.Connection")
cnnEquipment.open (Session("ClientCnnString"))
Set rsEquipment = Server.CreateObject("ADODB.Recordset")
rsEquipment.CursorLocation = 3 
Set rsEquipment = cnnEquipment.Execute(SQLEquipment)
	
If not rsEquipment.EOF Then

	ModelIntRecID = rsEquipment("ModelIntRecID")
	StatusCodeIntRecID = rsEquipment("StatusCodeIntRecID")
	SerialNumber = rsEquipment("SerialNumber")
	AssetTag1 = rsEquipment("AssetTag1")
	AssetTag2 = rsEquipment("AssetTag2")
	AssetTag3 = rsEquipment("AssetTag3")
	AssetTag4 = rsEquipment("AssetTag4")
	AcquisitionCodeIntRecID = rsEquipment("AcquisitionCodeIntRecID")
	PurchasedFromVendorID = rsEquipment("PurchasedFromVendorID")
	PurchasedViaPONumber = rsEquipment("PurchasedViaPONumber")
	PurchaseDate = rsEquipment("PurchaseDate")
	PurchaseCost = rsEquipment("PurchaseCost")
	ReplacementCost = rsEquipment("ReplacementCost")
	AcquiredConditionIntRecID = rsEquipment("AquiredConditionIntRecID")
	CurrentConditionIntRecID = rsEquipment("CurrentConditionIntRecID")
	WarrentyStartDate = rsEquipment("WarrentyStartDate")
	WarrentyEndDate = rsEquipment("WarrentyEndDate")
	Comments = rsEquipment("Comments")
	Color = rsEquipment("Color")
	Size= rsEquipment("Size")
	
	If PurchaseDate <> "" Then
		PurchaseDate = FormatDateTime(PurchaseDate,2)
	End If
	
	If WarrentyStartDate <> "" Then
		WarrentyStartDate = FormatDateTime(WarrentyStartDate,2)
	End If
	
	If WarrentyEndDate <> "" Then
		WarrentyEndDate = FormatDateTime(WarrentyEndDate,2)
	End If
	
End If
set rsEquipment = Nothing
cnnEquipment.close
set cnnEquipment = Nothing

%>


<link rel="stylesheet" type="text/css" href="<%= BaseURL %>css/tabs.css" />

<SCRIPT LANGUAGE="JavaScript">
<!--
	$(document).ready(function() {
	
	
		$('#equipAddNewStatusCodeModal').on('show.bs.modal', function(e) {
				    	
		    //get data-id attribute of the clicked prospect
		    var EquipIntRecID = $("#txtInternalRecordIdentifier").val();
	
		    //populate the textbox with the id of the clicked prospect
		    $(e.currentTarget).find('input[name="txtEquipIntRecID"]').val(EquipIntRecID);
		    	    
		    var $modal = $(this);
	

		});

	
	    $('#selStatusCodeIntRecID').change(function() { //jQuery Change Function
	        var opval = $(this).val(); //Get value from select element
	        if(opval=="addnewstatuscode"){ //Compare it and if true
	            $('#equipAddNewStatusCodeModal').modal("show"); //Open Modal
	        }
	    });	
	    
	    	
	    $('#selModelIntRecID').change(function() { //jQuery Change Function
	    
	    	var EquipIntRecID = $("#txtInternalRecordIdentifier").val();
	    	var ModelIntRecID = $(this).val(); //Get value from selected model element

		 	$.ajax({
				type:"POST",
				url: "../../inc/InSightFuncs_AjaxForEquipment.asp",
				data: "action=GetInsightAssetTagByEquipAndModelIntRecID&equipIntRecID=" + encodeURIComponent(EquipIntRecID) + "&modelIntRecID=" + encodeURIComponent(ModelIntRecID),
				success: function(newInsightAssetTag)
				 {
				 	$("#InsightAssetTag").html("");
				 	$("#InsightAssetTag").html(newInsightAssetTag);
	             }
			});
	    });	



		$("#btnEquipAddNewStatusCodeSubmit").on('click', function(e) {
		
		  	event.preventDefault();
		  	
		  	var InternalRecordIdentifier = $("#txtInternalRecordIdentifier").val();
		  	var StatusCodeDesc = $("#txtStatusCodeDesc").val();
		  	var BackendSystemCode = $("#txtBackendSystemCode").val();
		  	var AvailableForPlacement = $("#chkAvailableForPlacement").val();
		  	var GeneratesRevenue = $("#chkGeneratesRentalRevenue").val();

		  	
		  	if (validateStatusCodeForm()){
		  	
			  	if ((StatusCodeDesc) && (BackendSystemCode)) {
			    	$.ajax({
						type:"POST",
						url: "../../inc/InSightFuncs_AjaxForEquipment.asp",
						data: "action=StatusCodeExistsInDB&scd=" + encodeURIComponent(StatusCodeDesc) + "&bsc=" + encodeURIComponent(BackendSystemCode),
						success: function(response)
						 {
							 if (response == 'Both') {
								swal("The status code and description already exists. Please enter new information or cancel.");
								return false;
							 }
							 else if (response == 'STATUSCODEDESC') {
								swal("The status code description already exists. Please enter a new description.");
								return false;
							 }
							 else if (response == 'BACKENDSYSTEMCODE') {
								swal("The status code/backend system code already exists. Please enter a new status code.");
								return false;
							 }
							 
							 else {
							 	$.ajax({
									type:"POST",
									url: "../../inc/InSightFuncs_AjaxForEquipment.asp",
									data: "action=InsertStatusCodeIntoDB&scd=" + encodeURIComponent(StatusCodeDesc) + "&bsc=" + encodeURIComponent(BackendSystemCode) + "&afp=" + encodeURIComponent(AvailableForPlacement) + "&gr=" + encodeURIComponent(GeneratesRevenue),
									success: function(newIntRecID)
									 {
									 	if (newIntRecID != "Error"){
										 	optionText = StatusCodeDesc + " (" + BackendSystemCode + ")"
										 	$('#selStatusCodeIntRecID').append($("<option/>").val(newIntRecID).text(optionText));
										 	$('#selStatusCodeIntRecID option[value="' + newIntRecID + '"]').prop('selected', true);
										 	$('#equipAddNewStatusCodeModal').hide();
										 	swal("The status code was saved successfully and is now in your list.");
										 	return false;
										 }
										 else {
										 	swal("Error inserting new status code.");
										 	return false;
										 }
						             }
								});
							 }
							 
			             }
					});
		  		}
		  	}
		});	








		$('#equipAddNewMovementCodeModal').on('show.bs.modal', function(e) {
				    	
		    //get data-id attribute of the clicked prospect
		    var EquipIntRecID = $("#txtInternalRecordIdentifier").val();
	
		    //populate the textbox with the id of the clicked prospect
		    $(e.currentTarget).find('input[name="txtEquipIntRecID"]').val(EquipIntRecID);
		    	    
		    var $modal = $(this);
	

		});

	
	    $('#selMovementCodeIntRecID').change(function() { //jQuery Change Function
	        var opval = $(this).val(); //Get value from select element
	        if(opval=="addnewmovementcode"){ //Compare it and if true
	            $('#equipAddNewMovementCodeModal').modal("show"); //Open Modal
	        }
	    });	
	    	

		$("#btnEquipAddNewMovementCodeSubmit").on('click', function(e) {
		
		  	event.preventDefault();
		  	
		  	var InternalRecordIdentifier = $("#txtInternalRecordIdentifier").val();
		  	var MovementCode = $("#txtMovementCode").val();
		  	var MovementCodeDesc = $("#txtMovementCodeDesc").val();

		  	
		  	if (validateMovementCodeForm()){
		  	
			  	if ((MovementCode) && (MovementCodeDesc)) {
			    	$.ajax({
						type:"POST",
						url: "../../inc/InSightFuncs_AjaxForEquipment.asp",
						data: "action=MovementCodeExistsInDB&mc=" + encodeURIComponent(MovementCode) + "&mcd=" + encodeURIComponent(MovementCodeDesc),
						success: function(response)
						 {
							 if (response == 'BOTH') {
								swal("The movement code and description already exists. Please enter new information or cancel.");
								return false;
							 }
							 else if (response == 'MOVEMENTCODE') {
								swal("The movement code already exists. Please enter a new movement code.");
								return false;
							 }
							 else if (response == 'MOVEMENTCODEDESC') {
								swal("The movement code description already exists. Please enter a new description.");
								return false;
							 }
							 
							 else {
							 	$.ajax({
									type:"POST",
									url: "../../inc/InSightFuncs_AjaxForEquipment.asp",
									data: "action=InsertMovementCodeIntoDB&mc=" + encodeURIComponent(MovementCode) + "&mcd=" + encodeURIComponent(MovementCodeDesc),
									success: function(newIntRecID)
									 {
									 	if (newIntRecID != "Error"){
										 	optionText = MovementCode + " (" + MovementCodeDesc + ")"
										 	$('#selMovementCodeIntRecID').append($("<option/>").val(newIntRecID).text(optionText));
										 	$('#selMovementCodeIntRecID option[value="' + newIntRecID + '"]').prop('selected', true);
										 	$('#equipAddNewMovementCodeModal').hide();
										 	swal("The movement code was saved successfully and is now in your list.");
										 	return false;
										 }
										 else {
										 	swal("Error inserting new movement code.");
										 	return false;
										 }
						             }
								});
							 }
							 
			             }
					});
		  		}
		  	}
		});	










		$('#equipAddNewAcquisitionCodeModal').on('show.bs.modal', function(e) {
				    	
		    //get data-id attribute of the clicked prospect
		    var EquipIntRecID = $("#txtInternalRecordIdentifier").val();
	
		    //populate the textbox with the id of the clicked prospect
		    $(e.currentTarget).find('input[name="txtEquipIntRecID"]').val(EquipIntRecID);
		    	    
		    var $modal = $(this);
	

		});

	
	    $('#selAcquisitionCodeIntRecID').change(function() { //jQuery Change Function
	        var opval = $(this).val(); //Get value from select element
	        if(opval=="addnewacquisitioncode"){ //Compare it and if true
	            $('#equipAddNewAcquisitionCodeModal').modal("show"); //Open Modal
	        }
	    });	
	    	

		$("#btnEquipAddNewAcquisitionCodeSubmit").on('click', function(e) {
		
		  	event.preventDefault();
		  	
		  	var InternalRecordIdentifier = $("#txtInternalRecordIdentifier").val();
		  	var AcquisitionCode = $("#txtAcquisitionCode").val();
		  	var AcquisitionCodeDesc = $("#txtAcquisitionCodeDesc").val();

		  	
		  	if (validateAcquisitionCodeForm()){
		  	
			  	if ((AcquisitionCode) && (AcquisitionCodeDesc)) {
			    	$.ajax({
						type:"POST",
						url: "../../inc/InSightFuncs_AjaxForEquipment.asp",
						data: "action=AcquisitionCodeExistsInDB&ac=" + encodeURIComponent(AcquisitionCode) + "&acd=" + encodeURIComponent(AcquisitionCodeDesc),
						success: function(response)
						 {
							 if (response == 'BOTH') {
								swal("The Acquisition code and description already exists. Please enter new information or cancel.");
								return false;
							 }
							 else if (response == 'ACQUISITIONCODE') {
								swal("The Acquisition code already exists. Please enter a new Acquisition code.");
								return false;
							 }
							 else if (response == 'ACQUISITIONCODEDESC') {
								swal("The Acquisition code description already exists. Please enter a new description.");
								return false;
							 }
							 
							 else {
							 	$.ajax({
									type:"POST",
									url: "../../inc/InSightFuncs_AjaxForEquipment.asp",
									data: "action=InsertAcquisitionCodeIntoDB&ac=" + encodeURIComponent(AcquisitionCode) + "&acd=" + encodeURIComponent(AcquisitionCodeDesc),
									success: function(newIntRecID)
									 {
									 	if (newIntRecID != "Error"){
										 	optionText = AcquisitionCodeDesc + " (" + AcquisitionCode + ")"
										 	$('#selAcquisitionCodeIntRecID').append($("<option/>").val(newIntRecID).text(optionText));
										 	$('#selAcquisitionCodeIntRecID option[value="' + newIntRecID + '"]').prop('selected', true);
										 	$('#equipAddNewAcquisitionCodeModal').hide();
										 	swal("The Acquisition code was saved successfully and is now in your list.");
										 	return false;
										 }
										 else {
										 	swal("Error inserting new Acquisition code.");
										 	return false;
										 }
						             }
								});
							 }
							 
			             }
					});
		  		}
		  	}
		});	








		$('#equipAddNewConditionCodeModal').on('show.bs.modal', function(e) {
				    	
		    //get data-id attribute of the clicked prospect
		    var EquipIntRecID = $("#txtInternalRecordIdentifier").val();
	
		    //populate the textbox with the id of the clicked prospect
		    $(e.currentTarget).find('input[name="txtEquipIntRecID"]').val(EquipIntRecID);
		    	    
		    var $modal = $(this);
	

		});

	
	    $('#selAcquiredConditionIntRecID').change(function() { //jQuery Change Function
	        var opval = $(this).val(); //Get value from select element
	        if(opval=="addnewconditioncode"){ //Compare it and if true
	            $('#equipAddNewConditionCodeModal').modal("show"); //Open Modal
	        }
	    });	
	    	

	    $('#selCurrentConditionIntRecID').change(function() { //jQuery Change Function
	        var opval = $(this).val(); //Get value from select element
	        if(opval=="addnewconditioncode"){ //Compare it and if true
	            $('#equipAddNewConditionCodeModal').modal("show"); //Open Modal
	        }
	    });	
	    

		$("#btnEquipAddNewConditionCodeSubmit").on('click', function(e) {
		
		  	event.preventDefault();
		  	
		  	var InternalRecordIdentifier = $("#txtInternalRecordIdentifier").val();
		  	var ConditionCode = $("#txtCondition").val();
		  	var ConditionCodeDesc = $("#txtConditionDescription").val();

		  	
		  	if (validateConditionCodeForm()){
		  	
			  	if (ConditionCode) {
			    	$.ajax({
						type:"POST",
						url: "../../inc/InSightFuncs_AjaxForEquipment.asp",
						data: "action=ConditionCodeExistsInDB&cc=" + encodeURIComponent(ConditionCode) + "&ccd=" + encodeURIComponent(ConditionCodeDesc),
						success: function(response)
						 {
							 if (response == 'CONDITIONCODE') {
								swal("The condition code already exists. Please enter new information or cancel.");
								return false;
							 }
							 
							 else {
							 	$.ajax({
									type:"POST",
									url: "../../inc/InSightFuncs_AjaxForEquipment.asp",
									data: "action=InsertConditionCodeIntoDB&cc=" + encodeURIComponent(ConditionCode) + "&ccd=" + encodeURIComponent(ConditionCodeDesc),
									success: function(newIntRecID)
									 {
									 	if (newIntRecID != "Error"){
										 	optionText = ConditionCode;
										 	$('#selConditionCodeIntRecID').append($("<option/>").val(newIntRecID).text(optionText));
										 	$('#selConditionCodeIntRecID option[value="' + newIntRecID + '"]').prop('selected', true);
										 	$('#equipAddNewConditionCodeModal').hide();
										 	swal("The condition code was saved successfully and is now in your list.");
										 }
										 else {
										 	swal("Error inserting new condition code.");
										 }
						             }
								});
							 }
							 
			             }
					});
		  		}
		  	}
		});	









		
		$("#btnSubmitEditEquipmentForm").on('click', function(e) {
		
		  	event.preventDefault();
		  	
		  	var InternalRecordIdentifier = $("#txtInternalRecordIdentifier").val();
		  	var AssetTag1 = $("#txtAssetTag1").val();
		  	
		  	if (validateEditEquipmentForm()){
		  	
			  	if (AssetTag1) {
			    	$.ajax({
						type:"POST",
						url: "../../inc/InSightFuncs_AjaxForEquipment.asp",
						data: "action=AssetTag1ExistsInDB&intRecID=" + encodeURIComponent(InternalRecordIdentifier) + "&assetTag1=" + encodeURIComponent(AssetTag1),
						success: function(response)
						 {
							 if (response == 'True') {
								swal("Asset tag 1 already exists as asset tag 1 for another piece of equipment. Please enter a new asset tag 1.");
								return false;
							 }
							 else {
							 	$("#frmEditEquipment").submit();
							 }
			             }
					});
		  		}
		  	}
		});	
		
	
		$('#btnGenerateAssetTag1').on('click', function(e) {

		    var InternalRecordIdentifier = $("#txtInternalRecordIdentifier").val();
		    		    		    		
	    	$.ajax({
				type:"POST",
				url: "../../inc/InSightFuncs_AjaxForEquipment.asp",
				data: "action=GenerateAssetTagForEquipment&intRecID=" + encodeURIComponent(InternalRecordIdentifier) + "&tagNum=1",
				success: function(response)
				 {
					 if (response !== '') {
					   	$("#txtAssetTag1").val(response);
					 }
				 	
	             }
			});
    	});	
	

		$('#btnGenerateAssetTag2').on('click', function(e) {

		    var InternalRecordIdentifier = $("#txtInternalRecordIdentifier").val();
		    		    		    		
	    	$.ajax({
				type:"POST",
				url: "../../inc/InSightFuncs_AjaxForEquipment.asp",
				data: "action=GenerateAssetTagForEquipment&intRecID=" + encodeURIComponent(InternalRecordIdentifier) + "&tagNum=2",
				success: function(response)
				 {
					 if (response !== '') {
					   	$("#txtAssetTag2").val(response);
					 }
				 	
	             }
			});
    	});	
    	
		$('#btnGenerateAssetTag3').on('click', function(e) {

		    var InternalRecordIdentifier = $("#txtInternalRecordIdentifier").val();
		    		    		    		
	    	$.ajax({
				type:"POST",
				url: "../../inc/InSightFuncs_AjaxForEquipment.asp",
				data: "action=GenerateAssetTagForEquipment&intRecID=" + encodeURIComponent(InternalRecordIdentifier) + "&tagNum=3",
				success: function(response)
				 {
					 if (response !== '') {
					   	$("#txtAssetTag3").val(response);
					 }
				 	
	             }
			});
    	});	
    	
		$('#btnGenerateAssetTag4').on('click', function(e) {

		    var InternalRecordIdentifier = $("#txtInternalRecordIdentifier").val();
		    		    		    		
	    	$.ajax({
				type:"POST",
				url: "../../inc/InSightFuncs_AjaxForEquipment.asp",
				data: "action=GenerateAssetTagForEquipment&intRecID=" + encodeURIComponent(InternalRecordIdentifier) + "&tagNum=4",
				success: function(response)
				 {
					 if (response !== '') {
					   	$("#txtAssetTag4").val(response);
					 }
				 	
	             }
			});
    	});	
    	
    	
		var today = new Date();
		var dd = today.getDate();
		var mm = today.getMonth()+1; //January is 0!
		var yyyy = today.getFullYear();
		
		if(dd<10) {
		    dd = '0'+dd
		} 
		
		if(mm<10) {
		    mm = '0'+mm
		} 
		
		today = mm + '/' + dd + '/' + yyyy;
	 
		
		if ('<%= PurchaseDate %>') {
		    $('#datetimepickerPurchaseDate').datetimepicker({
		    	useCurrent: false,
		        format: 'MM/DD/YYYY',
		        ignoreReadonly: true,
		        sideBySide: true, 
		        defaultDate: moment(new Date('<%= PurchaseDate %>'))
		    	
			});	
		}
		else {
		    $('#datetimepickerPurchaseDate').datetimepicker({
		    	useCurrent: false,
		        format: 'MM/DD/YYYY',
		        ignoreReadonly: true,
		        sideBySide: true, 
		        defaultDate: moment(new Date(today))
		    	
			});	
		}
   
		
		if ('<%= WarrentyStartDate %>') {
		    $('#datetimepickerWarrentyStartDate').datetimepicker({
		    	useCurrent: false,

		        format: 'MM/DD/YYYY',
		        ignoreReadonly: true,
		        sideBySide: true, 
		        defaultDate: moment(new Date('<%= WarrentyStartDate %>'))
		    	
			});	
		}
		else {
		    $('#datetimepickerWarrentyStartDate').datetimepicker({
		    	useCurrent: false,
		        format: 'MM/DD/YYYY',
		        ignoreReadonly: true,
		        sideBySide: true, 
		        //defaultDate: moment(new Date(today))
		    	
			});	
		}
		
		if ('<%= WarrentyEndDate %>') {
		    $('#datetimepickerWarrentyEndDate').datetimepicker({
		    	useCurrent: false,
		        format: 'MM/DD/YYYY',
		        ignoreReadonly: true,
		        sideBySide: true, 
		        defaultDate: moment(new Date('<%= WarrentyEndDate %>'))
		    	
			});	
		}
		else {
		    $('#datetimepickerWarrentyEndDate').datetimepicker({
		    	useCurrent: false,
		        format: 'MM/DD/YYYY',
		        ignoreReadonly: true,
		        sideBySide: true, 
		        //defaultDate: moment(new Date(today))
		    	
			});	
		}
		
					
	});	
	

	
    function validateEditEquipmentForm()
    {

		var ddlModel = document.getElementById("selModelIntRecID");
		var selectedValueModel = ddlModel.options[ddlModel.selectedIndex].value;
		
		if (selectedValueModel == "")
		{
			swal("You must choose a model for a piece of equipment.");
			return false;
		}

		var ddlStatusCode = document.getElementById("selStatusCodeIntRecID");
		var selectedValueStatusCode = ddlStatusCode.options[ddlStatusCode.selectedIndex].value;
		
		if ((selectedValueStatusCode == "") || (selectedValueStatusCode == "addnewstatuscode"))
		{
			swal("You must choose a status code for a piece of equipment.");
			return false;
		}
		

        if (document.frmEditEquipment.txtAssetTag1.value == "") {
            swal("You must enter at least one asset tag.");
            return false;
        }

        if (document.frmEditEquipment.txtPurchaseCost.value == "") {
            swal("You must enter the purchase cost.");
            return false;
        }
		
        if (document.frmEditEquipment.txtPurchaseCost.value != "") {
        
        	if (isNaN(document.frmEditEquipment.txtPurchaseCost.value)) {
            	swal("Please enter numbers only for the purchase cost.");
            	return false;
           	}
        }
        
        if (document.frmEditEquipment.txtReplacementCost.value == "") {
            swal("You must enter the replacement cost.");
            return false;
        }
        
        if (document.frmEditEquipment.txtReplacementCost.value != "") {
        
        	if (isNaN(document.frmEditEquipment.txtReplacementCost.value)) {
            	swal("Please enter numbers only for the replacement cost.");
            	return false;
           	}
        }


        if ((document.frmEditEquipment.txtWarrentyStartDate.value != "") && (document.frmEditEquipment.txtWarrentyEndDate.value != "")) {
   
			var d1 = new Date(document.frmEditEquipment.txtWarrentyStartDate.value);
			var d2 = new Date(document.frmEditEquipment.txtWarrentyEndDate.value);
			
			//var same = d1.getTime() === d2.getTime();
			//var notSame = d1.getTime() !== d2.getTime();
			     
        	if (d1.getTime() > d2.getTime()) {
            	swal("Warranty start date should be older than warranty end date.");
            	return false;
           	}
        }


		var ddlAcquiredCondition = document.getElementById("selAcquiredConditionIntRecID");
		var selectedAcquiredCondition = ddlAcquiredCondition.options[ddlAcquiredCondition.selectedIndex].value;
		
		if (selectedAcquiredCondition == "addnewconditioncode")
		{
			swal("Invalid Acquired Condition Code selected.");
			return false;
		}
		


		var ddlCurrentCondition = document.getElementById("selCurrentConditionIntRecID");
		var selectedCurrentCondition = ddlCurrentCondition.options[ddlCurrentCondition.selectedIndex].value;
		
		if ((selectedCurrentCondition == "") || (selectedCurrentCondition == "addnewconditioncode"))
		{
			swal("You must select the current condition for a piece of equipment.");
			return false;
		}

        return true;
        
    }
    
// -->
</SCRIPT>   


<style type="text/css">
	
	.ajax-loading {
	    position: relative;
	}
	.ajax-loading::after {
	    background-image: url("/img/loading.gif");
	    background-position: center top;
	    background-repeat: no-repeat;
	    content: "";
	    display: block;
	    height: 100%;
	    min-height: 100px;
	    position: absolute;
	    top: 0;
	    width: 100%;
	}
	.ajaxRowView .visibleRowEdit, .ajaxRowEdit .visibleRowView { display: none; }
	
	
	.styled{
	 	cursor:pointer;
	}

	.plus-button{
		cursor:pointer;
	}
	
	.beatpicker-clear{
		display: none;
	}
	.nav-tabs{
		font-size: 12px;
	}
	
	.the-tabs .nav>li>a{
		padding: 5px 10px;
		font-weight: bold;
	}


	.tab-content{
		margin-top:20px;
		font-size:12px;
	}
	
	   
	.tab-content .split-arrows{
		 text-align:left;
		 margin-top:10px;
		 margin-bottom: 10px;
	 }
	 
	.tab-content .split-arrows a{
		display:inline-block;
		background:#f5f5f5;
		padding:5px;
	}
	
	.tab-content .split-arrows a:hover{
		background:#ccc;
		text-decoration:none;
	}

	.inside {
		position:absolute;
		text-indent:8px;
		margin-top:7px;
		color:green;
		font-size:20px;
	}
	
	.inp {
		text-indent:15px;
	}
	.select-line{
		margin-bottom: 15px;
	}
	
	.row-line{
		margin-bottom: 25px;
	}
	
	.table th, tr, td{
		font-weight: normal;
	}
	
	.table>thead>tr>th{
		border: 0px;
	}
	.table thead>tr>th,.table tbody>tr>th,.table tfoot>tr>th,.table thead>tr>td,.table tbody>tr>td,.table tfoot>tr>td{
	border:0px;
	}

	
	.form-control{
		min-width: 100px;
	}
	
	.textarea-box{
		min-width: 260px;
	}
	
	.container {
	    width: 100%;
	}
	
	.control-label{
		font-size:12px;
		font-weight:normal;
		padding-top:10px;
	}
	.control-label-last{
		padding-top:0px;
	}
	
	.required{
		border-left:3px solid red;
	}
	
	.bottom-tabs-section {
	    border: 1px solid #ccc;
	    padding: 10px;
	    margin-top: 20px;
	    float: left;
	    width: 100%;
	}
		
	.btn-custom{
		width:100%;
		text-align:left;
		color: #333;
	    background-color: #f5f5f5;
		border:1px solid #ddd;
		outline:none;
		border-top-left-radius:5px;
		border-top-right-radius:5px;
		font-size:16px;
		padding:10px;
	}
	
	.btn-custom:hover{
		background:#ccc;
	}
	
	
	.bottom-table table thead th{
		padding:6px;
		font-weight:bold;
		border:1px solid #ddd;
		vertical-align:top;
	}
	
	.bottom-table table>tbody>tr>td{
		padding:6px;
		font-weight:normal;
		border:1px solid #ddd;
		vertical-align: middle;
	}
	
	.narrow-results{
		margin-bottom:15px;
	}
	
	#filter-movement{
		width:40%;
		padding:10px;
		height:34px;
	}
	 
	#filter-service{
		width:40%;
		padding:10px;
		height:34px;
	}
		
	.nav-tabs>li>a {
	color: #fff;
	font-size:16px;
	}
	
	
	.EquipmentTabMovementColor{
		background:#cc4125 !important;
	}
	.EquipmentTabServiceColor{
		background:#ff9900 !important;
	}
		
	.fileicon {
		width:40%;
	}
	
	.nav-tabs > li.active > a,
	.nav-tabs > li.active > a:hover,
	.nav-tabs > li.active > a:focus{
	
		color: #fff;
		font-weight:normal;
		font-size:24px;
	    /*background-color: #111 !important;*/
	    border-color: #2e6da4 !important;
	    margin-bottom:20px;
	    margin-top:0px;
	     
	} 
	
	.user-row {
	    margin-bottom: 14px;
	}
	
	.user-row:last-child {
	    margin-bottom: 0;
	}
	
	.dropdown-user {
	    margin: 13px 0;
	    padding: 5px;
	    height: 100%;
	}
	
	.dropdown-user:hover {
	    cursor: pointer;
	}
	
	.table-user-information > tbody > tr {
	    border-top: 1px solid rgb(221, 221, 221);
	}
	
	.table-user-information > tbody > tr:first-child {
	    border-top: 0;
	}
	
	
	.table-user-information > tbody > tr > td {
	    border-top: 0;
	}
	.toppad {
		margin-top:20px;
	}	
	
	.img-rounded{
		border-radius:15px;
	}
	
</style>


<h1 class="page-header"> Edit <%= GetTerm("Equipment") %> Model - <%= GetBrandNameByModelIntRecID(ModelIntRecID) %>&nbsp;<%= GetModelNameByIntRecID(ModelIntRecID) %></h1>

<div class="container">

<form method="POST" action="<%= BaseURL %>equipment/equipment/editEquipment_submit.asp" name="frmEditEquipment" id="frmEditEquipment">

	 <input type="hidden" id="txtInternalRecordIdentifier" name="txtInternalRecordIdentifier" value="<%= InternalRecordIdentifier %>">
		
	  <!--- FIRST ROW------>
      <div class="row">
      
      <!--- FIRST COLUMN OF FIRST ROW------>
      
      	<!--- BEGIN DEFAULT PICTURE PANEL------>
        <div class="col-xs-2 col-sm-2 col-md-2 col-lg-2 toppad">
          <div class="panel panel-info">
            <div class="panel-heading">
              <h3 class="panel-title">Default Equipment Picture</h3>
            </div>
            <div class="panel-body">
              <div class="row">
                <div class="col-md-12 col-lg-12" align="center">
                	<img alt="Default Equipment Pic" src="http://via.placeholder.com/225x300" class="img-rounded">
                </div>
              </div>
            </div>
          </div>
        </div>
        <!--- END DEFAULT PICTURE PANEL------>
        
        <!--- END FIRST COLUMN OF FIRST ROW------>

		<!--- SECOND COLUMN OF FIRST ROW------>
        <div class="col-xs-3 col-sm-3 col-md-3 col-lg-3 toppad">
        
          <!--- BEGIN STATUS/LOCATION/CONDITION PANEL------>
          <div class="panel panel-info">
            <div class="panel-heading">
              <h3 class="panel-title">Status/Location/Condition</h3>
            </div>
            <div class="panel-body">
              <div class="row">
                <div class=" col-md-12 col-lg-12"> 
                  <table class="table table-user-information">
                    <tbody>
                      <tr>
                        <td><strong>Model</strong></td>
                        <td>						  	
                        	<select class="form-control required" name="selModelIntRecID" id="selModelIntRecID">
					  			<option value="">Select Model of Equipment</option>
						      	<% 'Get all equipment models
						      	  	SQLEquipModels = "SELECT * FROM EQ_Models ORDER BY Model ASC"
		
									Set cnnEquipModels = Server.CreateObject("ADODB.Connection")
									cnnEquipModels.open (Session("ClientCnnString"))
									Set rsEquipModels = Server.CreateObject("ADODB.Recordset")
									rsEquipModels.CursorLocation = 3 
									Set rsEquipModels = cnnEquipModels.Execute(SQLEquipModels)
									If not rsEquipModels.EOF Then
										Do
											If cInt(ModelIntRecID) = cInt(rsEquipModels("InternalRecordIdentifier")) Then 
												Response.Write("<option value='" & rsEquipModels("InternalRecordIdentifier") & "' selected='selected'>" & rsEquipModels("Model") & "</option>")
											Else
												Response.Write("<option value='" & rsEquipModels("InternalRecordIdentifier") & "'>" & rsEquipModels("Model") & "</option>")
											End If
											rsEquipModels.movenext
										Loop until rsEquipModels.eof
									End If
									set rsEquipModels = Nothing
									cnnEquipModels.close
									set cnnEquipModels = Nothing
								%>
							</select>
						</td>
                      </tr>
                      <tr>
                        <td><strong>Location</strong></td>
                        <td>
                        	<% 
                        		CustomerLocation = GetCustomerIDByEquipIntRecID(EquipIntRecID)
                        		
                        		If CustomerLocation <> "NOT PLACED AT AN ACCOUNT" Then %>
                        		
                        		<%= GetCustNameByCustNum(CustomerLocation) %><br>Acct. <%= CustomerLocation %>
                        		
                        	<% Else %>
                        		NOT PLACED AT AN ACCOUNT
                        	<% End If %>	
                        </td>
                      </tr>
                      <tr>
                        <td><strong>Status</strong></td>
                        <td>
                        	<select class="form-control required" name="selStatusCodeIntRecID" id="selStatusCodeIntRecID">
					  			<option value="">Select Status Code of Equipment</option>
					  			<% If userCanEditEqpOnFly(Session("UserNo")) Then %>
					  				<option value="addnewstatuscode">+ ADD NEW STATUS CODE</option>
					  			<% End If %>
					  			
						      	<% 'Get all equipment status codes
						      	  	SQLEquipStatusCodes = "SELECT * FROM EQ_StatusCodes ORDER BY statusDesc ASC"
		
									Set cnnEquipStatusCodes = Server.CreateObject("ADODB.Connection")
									cnnEquipStatusCodes.open (Session("ClientCnnString"))
									Set rsEquipStatusCodes = Server.CreateObject("ADODB.Recordset")
									rsEquipStatusCodes.CursorLocation = 3 
									Set rsEquipStatusCodes = cnnEquipStatusCodes.Execute(SQLEquipStatusCodes)
									If not rsEquipStatusCodes.EOF Then
										Do
											If cInt(StatusCodeIntRecID) = cInt(rsEquipStatusCodes("InternalRecordIdentifier")) Then 
												Response.Write("<option value='" & rsEquipStatusCodes("InternalRecordIdentifier") & "' selected='selected'>" & rsEquipStatusCodes("statusDesc") & " (" & rsEquipStatusCodes("statusBackendSystemCode") & ")</option>")
											Else
												Response.Write("<option value='" & rsEquipStatusCodes("InternalRecordIdentifier") & "'>" & rsEquipStatusCodes("statusDesc") & " (" & rsEquipStatusCodes("statusBackendSystemCode") & ")</option>")
											End If
											rsEquipStatusCodes.movenext
										Loop until rsEquipStatusCodes.eof
									End If
									set rsEquipStatusCodes = Nothing
									cnnEquipStatusCodes.close
									set cnnEquipStatusCodes = Nothing
								%>
							</select>
						</td>
                      </tr>
                      <tr>
                        <td><strong>Availability</strong></td>
                        <td><%
                        		availableForPlacement = GetAvailableForPlacementByEquipIntRecID(EquipIntRecID) 
                        		If availableForPlacement = 1 OR availableForPlacement = true Then Response.write ("AVAILABLE FOR PLACEMENT")
                        		If availableForPlacement = 0 OR availableForPlacement = false Then Response.write ("NOT AVAILABLE FOR PLACEMENT")
                        	%>
                        </td>
                      </tr> 
                      <tr>
                        <td><strong>Current Condition</strong></td>
                        <td>
                        	<select class="form-control required" name="selCurrentConditionIntRecID" id="selCurrentConditionIntRecID">
					  			<option value="" <% If CurrentConditionIntRecID = "" Then Respponse.Write("selected='selected'") %>>Select Current Condition of Equipment</option>
					  			
					  			<% If userCanEditEqpOnFly(Session("UserNo")) Then %>
					  				<option value="addnewconditioncode">+ ADD NEW CONDITION CODE</option>
					  			<% End If %>
					  			
						      	<% 'Get all Condition
						      	  	SQLEquipCondition = "SELECT * FROM EQ_Condition ORDER BY Condition ASC"
		
									Set cnnEquipCondition = Server.CreateObject("ADODB.Connection")
									cnnEquipCondition.open (Session("ClientCnnString"))
									Set rsEquipCondition = Server.CreateObject("ADODB.Recordset")
									rsEquipCondition.CursorLocation = 3 
									Set rsEquipCondition = cnnEquipCondition.Execute(SQLEquipCondition)
									If not rsEquipCondition.EOF Then
										Do
											If CurrentConditionIntRecID <> "" Then
												If cInt(CurrentConditionIntRecID) = cInt(rsEquipCondition("InternalRecordIdentifier")) Then 
													Response.Write("<option value='" & rsEquipCondition("InternalRecordIdentifier") & "' selected='selected'>" & rsEquipCondition("Condition") & "</option>")
												Else
													Response.Write("<option value='" & rsEquipCondition("InternalRecordIdentifier") & "'>" & rsEquipCondition("Condition") & "</option>")
												End If
											Else
												Response.Write("<option value='" & rsEquipCondition("InternalRecordIdentifier") & "'>" & rsEquipCondition("Condition") & "</option>")
											End If
											rsEquipCondition.movenext
										Loop until rsEquipCondition.eof
									End If
									set rsEquipCondition = Nothing
									cnnEquipCondition.close
									set cnnEquipCondition = Nothing
								%>
							</select>
                        </td>
                      </tr>                                            
                    </tbody>
                  </table>
                 </div>
              </div>
            </div>
            <!--
         	<div class="panel-footer">
                <a data-original-title="Broadcast Message" data-toggle="tooltip" type="button" class="btn btn-sm btn-primary"><i class="glyphicon glyphicon-envelope"></i></a>
                <span class="pull-right">
                    <a href="edit.html" data-original-title="Edit this user" data-toggle="tooltip" type="button" class="btn btn-sm btn-warning"><i class="glyphicon glyphicon-edit"></i></a>
                    <a data-original-title="Remove this user" data-toggle="tooltip" type="button" class="btn btn-sm btn-danger"><i class="glyphicon glyphicon-remove"></i></a>
                </span>
            </div>
            -->
            
          </div>
          <!--- END STATUS/LOCATION/CONDITION PANEL------>
          
          
          <!--- BEGIN EQUIPMENT IDENTIFICATION PANEL------>
          <div class="panel panel-info">
            <div class="panel-heading">
              <h3 class="panel-title">Equipment Identification</h3>
            </div>
            <div class="panel-body">
              <div class="row">
                <div class=" col-md-12 col-lg-12"> 
                  <table class="table table-user-information">
                    <tbody>
                      <tr>
                        <td><strong>Insight Asset Tag</strong></td>
                        <td>
                        	<div id="InsightAssetTag"><%= GetInsightAssetTagByEquipIntRecID(EquipIntRecID) %></div>	
                        </td>
                      </tr>
                    
                      <tr>
                        <td><strong>Serial Number</strong></td>
                        <td>		    				
                        	<i class="inside fa fa-barcode" aria-hidden="true"></i>
		    				<input type="text" class="form-control inp" id="txtSerialNumber" name="txtSerialNumber" value="<%= SerialNumber %>">
						</td>
                      </tr>
                      <tr>
                        <td><strong>Asset Tag 1</strong></td>
                        <td>
		    				<i class="inside fa fa-tag"></i>
		    				<input type="text" class="form-control required inp" id="txtAssetTag1" name="txtAssetTag1" value="<%= AssetTag1 %>">
		    				<button type="button" class="btn btn-success btn-xs pull-right" style="margin-top:5px;" id="btnGenerateAssetTag1">Generate Asset Tag</button>
                        </td>
                      </tr>
                      <tr>
                        <td><strong>Asset Tag 2</strong></td>
                        <td>
		    				<i class="inside fa fa-tag"></i>
		    				<input type="text" class="form-control inp" id="txtAssetTag2" name="txtAssetTag2" value="<%= AssetTag2 %>">
		    				<button type="button" class="btn btn-success btn-xs pull-right" style="margin-top:5px;" id="btnGenerateAssetTag2">Generate Asset Tag</button>
                        </td>
                      </tr>
                      <tr>
                        <td><strong>Asset Tag 3</strong></td>
                        <td>
		    				<i class="inside fa fa-tag"></i>
		    				<input type="text" class="form-control inp" id="txtAssetTag3" name="txtAssetTag3" value="<%= AssetTag3 %>">
		    				<button type="button" class="btn btn-success btn-xs pull-right" style="margin-top:5px;" id="btnGenerateAssetTag3">Generate Asset Tag</button>
                        </td>
                      </tr>
                      <tr>
                        <td><strong>Asset Tag 4</strong></td>
                        <td>
		    				<i class="inside fa fa-tag"></i>
		    				<input type="text" class="form-control inp" id="txtAssetTag4" name="txtAssetTag4" value="<%= AssetTag4 %>">
		    				<button type="button" class="btn btn-success btn-xs pull-right" style="margin-top:5px;" id="btnGenerateAssetTag4">Generate Asset Tag</button>
                        </td>
                      </tr>
                     
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          </div>
          <!--- END EQUIPMENT IDENTIFICATION PANEL------>
          
        </div>
        <!--- END SECOND COLUMN OF FIRST ROW------>
        
 

		<!--- THIRD COLUMN OF FIRST ROW------>
        <div class="col-xs-3 col-sm-3 col-md-3 col-lg-3 toppad">
        
        
          <!--- BEGIN FINANCIAL INFORMATION PANEL------>
          <div class="panel panel-success">
            <div class="panel-heading">
              <h3 class="panel-title">Financial Information</h3>
            </div>
            <div class="panel-body">
              <div class="row">
                <div class=" col-md-12 col-lg-12"> 
                  <table class="table table-user-information">
                    <tbody>
                      <tr>
                        <td><strong>Purchase Cost</strong></td>
                        <td>
		    				<i class="inside fa fa-usd"></i>
		    				<input type="text" class="form-control required inp" id="txtPurchaseCost" name="txtPurchaseCost" value="<%= PurchaseCost %>">                        
                        </td>
                      </tr>
                      <tr>
                        <td><strong>Replacement Cost</strong></td>
                        <td>
		    				<i class="inside fa fa-usd"></i>
		    				<input type="text" class="form-control required inp" id="txtReplacementCost" name="txtReplacementCost" value="<%= ReplacementCost %>">
                        </td>
                      </tr>
                      <tr>
                        <td><strong>Lifetime Revenue</strong></td>
                        <td>$$$</td>
                      </tr>
                      <tr>
                        <td colspan="2">$X left until cash positive OR Cash Positive</td>
                      </tr>                     
                    </tbody>
                  </table>
                 </div>
              </div>
            </div>
          </div>
          <!--- END FINANCIAL INFORMATION PANEL------>
          
          
         <!--- BEGIN PURCHASE/ACQUISITION PANEL------> 
         <div class="panel panel-info">
            <div class="panel-heading">
              <h3 class="panel-title">Purchase/Acquisition Info</h3>
            </div>
            <div class="panel-body">
              <div class="row">
                <div class=" col-md-12 col-lg-12"> 
                  <table class="table table-user-information">
                    <tbody>
                      <tr>
                        <td><strong>Acquisition Type</strong></td>
                        <td>
						  	<select class="form-control" name="selAcquisitionCodeIntRecID" id="selAcquisitionCodeIntRecID">
						  			<option value="" selected="selected">Select Acquisition Type for Equipment</option>
						  			
						  			<% If userCanEditEqpOnFly(Session("UserNo")) Then %>
						  				<option value="addnewacquisitioncode">+ ADD NEW ACQUISITION CODE</option>
						  			<% End If %>
						  			
							      	<% 'Get all Acquisition Codes
							      	  	SQLEquipAcquisition = "SELECT * FROM EQ_AcquisitionCodes ORDER BY AcquisitionCode ASC"
			
										Set cnnEquipAcquisition = Server.CreateObject("ADODB.Connection")
										cnnEquipAcquisition.open (Session("ClientCnnString"))
										Set rsEquipAcquisition = Server.CreateObject("ADODB.Recordset")
										rsEquipAcquisition.CursorLocation = 3 
										Set rsEquipAcquisition = cnnEquipAcquisition.Execute(SQLEquipAcquisition)
										If not rsEquipAcquisition.EOF Then
											Do
											
												If AcquisitionCodeIntRecID <> "" Then
													If cInt(AcquisitionCodeIntRecID) = cInt(rsEquipAcquisition("InternalRecordIdentifier")) Then 
														Response.Write("<option value='" & rsEquipAcquisition("InternalRecordIdentifier") & "' selected='selected'>" & rsEquipAcquisition("AcquisitionDesc") & " (" & rsEquipAcquisition("AcquisitionCode") & ")</option>")
													Else
														Response.Write("<option value='" & rsEquipAcquisition("InternalRecordIdentifier") & "'>" & rsEquipAcquisition("AcquisitionDesc") & " (" & rsEquipAcquisition("AcquisitionCode") & ")</option>")
													End If
												Else
													Response.Write("<option value='" & rsEquipAcquisition("InternalRecordIdentifier") & "'>" & rsEquipAcquisition("AcquisitionDesc") & " (" & rsEquipAcquisition("AcquisitionCode") & ")</option>")
												End If
												rsEquipAcquisition.movenext
											Loop until rsEquipAcquisition.eof
										End If
										set rsEquipAcquisition = Nothing
										cnnEquipAcquisition.close
										set cnnEquipAcquisition = Nothing
									%>
							</select>
                        </td>
                      </tr>
                      <tr>
                        <td><strong>Purchase Date</strong></td>
                        <td>                       
			                <div class="input-group date" id="datetimepickerPurchaseDate">
			                    <input type="text" class="form-control" name="txtPurchaseDate" id="txtPurchaseDate">
			                    <span class="input-group-addon">
			                        <span class="glyphicon glyphicon-calendar"></span>
			                    </span>
			                </div>
                        </td>
                      </tr>
                      <tr>
                        <td><strong>Purchased From</strong></td>
                        <td>
						  	<select class="form-control" name="selVendorIntRecID" id="selVendorIntRecID">
						  			<option value="">Select Vendor of Equipment</option>
							      	<% 'Get all vendors
							      	  	SQLEquipVendors = "SELECT * FROM AP_Vendor ORDER BY vendorCompanyName ASC"
			
										Set cnnEquipVendors = Server.CreateObject("ADODB.Connection")
										cnnEquipVendors.open (Session("ClientCnnString"))
										Set rsEquipVendors = Server.CreateObject("ADODB.Recordset")
										rsEquipVendors.CursorLocation = 3 
										Set rsEquipVendors = cnnEquipVendors.Execute(SQLEquipVendors)
										If not rsEquipVendors.EOF Then
											Do
												If PurchasedFromVendorID <> "" Then
													If cInt(PurchasedFromVendorID) = cInt(rsEquipVendors("InternalRecordIdentifier")) Then 
														Response.Write("<option value='" & rsEquipVendors("InternalRecordIdentifier") & "' selected='selected'>" & rsEquipVendors("vendorCompanyName") & "</option>")
													Else
														Response.Write("<option value='" & rsEquipVendors("InternalRecordIdentifier") & "'>" & rsEquipVendors("vendorCompanyName") & "</option>")
													End If
												Else
													Response.Write("<option value='" & rsEquipVendors("InternalRecordIdentifier") & "'>" & rsEquipVendors("vendorCompanyName") & "</option>")
												End If
												rsEquipVendors.movenext
											Loop until rsEquipVendors.eof
										End If
										set rsEquipVendors = Nothing
										cnnEquipVendors.close
										set cnnEquipVendors = Nothing
									%>
							</select>
                        </td>
                      </tr>
                      <tr>
                        <td><strong>PO Number</strong></td>
                        <td><input type="text" class="form-control inp" id="txtPurchasedViaPONumber" name="txtPurchasedViaPONumber" value="<%= PurchasedViaPONumber %>"></td>
                      </tr>
                      <tr>
                        <td><strong>Acquired Condition</strong></td>
                        <td>
						  	<select class="form-control" name="selAcquiredConditionIntRecID" id="selAcquiredConditionIntRecID">
						  			<option value="" <% If AcquiredConditionIntRecID = "" Then Respponse.Write("selected='selected'") %>>Select Acquired Condition of Equipment</option>
						  			
						  			<% If userCanEditEqpOnFly(Session("UserNo")) Then %>
						  				<option value="addnewconditioncode">+ ADD NEW CONDITION CODE</option>
						  			<% End If %>
						  			
							      	<% 'Get all Condition
							      	  	SQLEquipCondition = "SELECT * FROM EQ_Condition ORDER BY Condition ASC"
			
										Set cnnEquipCondition = Server.CreateObject("ADODB.Connection")
										cnnEquipCondition.open (Session("ClientCnnString"))
										Set rsEquipCondition = Server.CreateObject("ADODB.Recordset")
										rsEquipCondition.CursorLocation = 3 
										Set rsEquipCondition = cnnEquipCondition.Execute(SQLEquipCondition)
										If not rsEquipCondition.EOF Then
											Do
												If AcquiredConditionIntRecID <> "" Then
													If cInt(AcquiredConditionIntRecID) = cInt(rsEquipCondition("InternalRecordIdentifier")) Then 
														Response.Write("<option value='" & rsEquipCondition("InternalRecordIdentifier") & "' selected='selected'>" & rsEquipCondition("Condition") & "</option>")
													Else
														Response.Write("<option value='" & rsEquipCondition("InternalRecordIdentifier") & "'>" & rsEquipCondition("Condition") & "</option>")
													End If
												Else
													Response.Write("<option value='" & rsEquipCondition("InternalRecordIdentifier") & "'>" & rsEquipCondition("Condition") & "</option>")
												End If
												rsEquipCondition.movenext
											Loop until rsEquipCondition.eof
										End If
										set rsEquipCondition = Nothing
										cnnEquipCondition.close
										set cnnEquipCondition = Nothing
									%>
							</select>
                        </td>
                      </tr>
                     
                    </tbody>
                  </table>
                 </div>
              </div>
            </div>            
          </div>
          <!--- END PURCHASE/ACQUISITION PANEL------>
          
          
        </div>
        <!--- END THIRD COLUMN OF FIRST ROW------>



		<!--- FOURTH COLUMN OF FIRST ROW------>
        <div class="col-xs-3 col-sm-3 col-md-3 col-lg-3 toppad">
        
          <!--- BEGIN WARRANTY/REPAIR PANEL------>
          <div class="panel panel-danger">
            <div class="panel-heading">
              <h3 class="panel-title">Warranty/Repair</h3>
            </div>
            <div class="panel-body">
              <div class="row">
                <div class=" col-md-12 col-lg-12"> 
                  <table class="table table-user-information">
                    <tbody>
                      <tr>
                        <td><strong>Warranty Start Date</strong></td>
                        <td>
			                <div class="input-group date" id="datetimepickerWarrentyStartDate">
			                    <input type="text" class="form-control" name="txtWarrentyStartDate" id="txtWarrentyStartDate">
			                    <span class="input-group-addon">
			                        <span class="glyphicon glyphicon-calendar"></span>
			                    </span>
			                </div>                        
                        </td>
                      </tr>
                      <tr>
                        <td><strong>Warranty End Date</strong></td>
                        <td>
			                <div class="input-group date" id="datetimepickerWarrentyEndDate">
			                    <input type="text" class="form-control" name="txtWarrentyEndDate" id="txtWarrentyEndDate">
			                    <span class="input-group-addon">
			                        <span class="glyphicon glyphicon-calendar"></span>
			                    </span>
			                </div>
                        </td>
                      </tr>                   
                    </tbody>
                  </table>
                 </div>
              </div>
            </div>
          </div>
          <!--- END WARRANTY/REPAIR PANEL------>
         

          <!--- BEGIN OTHER INFORMATION PANEL------> 
          <div class="panel panel-info">
            <div class="panel-heading">
              <h3 class="panel-title">Other Information</h3>
            </div>
            <div class="panel-body">
              <div class="row">
                <div class=" col-md-12 col-lg-12"> 
                  <table class="table table-user-information">
                    <tbody>
                      <tr>
                        <td><strong>Size</strong></td>
                        <td><input type="text" class="form-control inp" id="txtSize" name="txtSize" value="<%= Size %>"></td>
                      </tr>
                      <tr>
                        <td><strong>Color</strong></td>
                        <td><input type="text" class="form-control inp" id="txtColor" name="txtColor" value="<%= Color %>"></td>
                      </tr>    
                      <tr>
                        <td><strong>Comments</strong></td>
                        <td><textarea class="form-control inp" id="txtComments" name="txtComments" rows="10"><%= Comments %></textarea></td>
                      </tr>                                       
                    </tbody>
                  </table>
                 </div>
              </div>
            </div>
          </div>
          <!--- END OTHER INFORMATION PANEL------>
          
        </div>
        <!--- END FOURTH COLUMN OF FIRST ROW------>
       
      </div>
      <!--- END FIRST ROW------>
      
      
      
      
      
      
	<!--- SECOND ROW------>
	<div class="row pull-right">
		<div class="col-lg-12">
		    <!-- cancel / submit !-->
			<div class="row row-line">
				<div class="col-lg-12 alertbutton">
					<div class="col-lg-12">
						<a href="<%= BaseURL %>equipment/equipment/findEquipment.asp">
		    				<button type="button" class="btn btn-default">&lsaquo; Cancel &amp; Go Back To Equipment Search</button>
						</a>
						<button type="button" class="btn btn-primary" id="btnSubmitEditEquipmentForm"><i class="far fa-save"></i> Save Equipment Changes</button>
					</div>
			    </div>
			</div>
		</div>
	</div>
	<!--- END SECOND ROW------>

</form>  
  
	  
	<!--- BEGIN THIRD ROW------>
	<div class="row">
			 
		<!-- tabs start here !-->
		<div class="bottom-table">
			<div class="row">
				<div class="col-lg-12">
					<div class="bottom-tabs-section">
		
						<!-- tab navigation !-->
						<ul class="nav nav-tabs" role="tablist">
							<li role='presentation' class="active"><a href='#movement' class='EquipmentTabMovementColor' aria-controls='movement' role='tab' data-toggle='tab'>Movement</a></li>
							<li role='presentation'><a href='#service' class='EquipmentTabServiceColor' aria-controls='service' role='tab' data-toggle='tab'>Service Calls</a></li>
						</ul>
						<!-- eof tab navigation -->
					
						<div class="tab-content">
							<!--#include file="editEquipment_movement_tab.asp"-->
							<!--#include file="editEquipment_service_tab.asp"-->
						</div>
							
					</div><!-- eof bottom-tabs-section-->
				</div><!-- eof col-lg-12 -->
			</div><!-- eof row -->
		</div><!-- eof bottom-table -->
	</div><!-- eof row -->
	<!--- END THIRD ROW------>

</div><!-- eof content container -->


<!-- tabs js  !-->
<script type="text/javascript">
		 $('a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
		  e.target // newly activated tab
		  e.relatedTarget // previous active tab
	})
</script>

<script>
		$(document).ready(function(){
		  $("#demo").on("hide.bs.collapse", function(){
		    $(".btn-custom").html('<span class="glyphicon glyphicon-collapse-down"></span> Click to Expand');
		  });
		  $("#demo").on("show.bs.collapse", function(){
		    $(".btn-custom").html('<span class="glyphicon glyphicon-collapse-up"></span> Click to Collapse');
		  });
		});
</script>
 <!-- eof tabs js !-->

 <!-- custom table search !-->

<script>

$(document).ready(function () {

    (function ($) {
        
        $('#filter-movement').keyup(function () {

            var rex = new RegExp($(this).val(), 'i');
            $('.searchable-movement tr').hide();
            $('.searchable-movement tr').filter(function () {
                return rex.test($(this).text());
            }).show();
        })
        
        $('#filter-service').keyup(function () {

            var rex = new RegExp($(this).val(), 'i');
            $('.searchable-service tr').hide();
            $('.searchable-service tr').filter(function () {
                return rex.test($(this).text());
            }).show();
        })

        
    }(jQuery));

});
</script>




<!-- eof custom table search !-->


<!-- checkboxes JS !-->
<script type="text/javascript">
    function changeState(el) {
        if (el.readOnly) el.checked=el.readOnly=false;
        else if (!el.checked) el.readOnly=el.indeterminate=true;
    }
</script>
<!-- eof checkboxes JS !-->

<!-- **************************************************************************************************************************** -->
<!-- MODALS FOR EDIT EQUIPMENT ON THE FLY BEGIN HERE !-->
<!-- **************************************************************************************************************************** -->

<!--#include file="editEquipTablesOnTheFlyModals.asp "-->

<!-- **************************************************************************************************************************** -->
<!-- MODALS FOR EDIT EQUIPMENT ON THE FLY END HERE !-->
<!-- **************************************************************************************************************************** -->


<!--#include file="../../inc/footer-main.asp"-->

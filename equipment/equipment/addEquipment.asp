<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/InSightFuncs_Equipment.asp"-->


<!-- datetime picker !-->
<script src="<%= baseURL %>js/moment.min.js" type="text/javascript"></script>
<script type="text/javascript" src="<%= BaseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.min.js"></script>
<link href="<%= BaseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.css" rel="stylesheet">
<!-- end datetime picker !-->


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








	
		
		$("#btnSubmitAddEquipmentForm").on('click', function(e) {
		
		  	event.preventDefault();
		  	
		  	var InternalRecordIdentifier = $("#txtInternalRecordIdentifier").val();
		  	var AssetTag1 = $("#txtAssetTag1").val();
		  	
		  	if (validateAddEquipmentForm()){
		  	
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
							 	swal("Equipment Successfully Added");
							 	$("#frmAddEquipment").submit();
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
	 
		
	    $('#datetimepickerPurchaseDate').datetimepicker({
	    	useCurrent: false,
	        format: 'MM/DD/YYYY',
	        ignoreReadonly: true,
	        sideBySide: true, 
	        defaultDate: moment(new Date(today))
	    	
		});	


	    $('#datetimepickerWarrentyStartDate').datetimepicker({
	    	useCurrent: false,
	        format: 'MM/DD/YYYY',
	        ignoreReadonly: true,
	        sideBySide: true, 
	        //defaultDate: moment(new Date(today))
	    	
		});	

	
	    $('#datetimepickerWarrentyEndDate').datetimepicker({
	    	useCurrent: false,
	        format: 'MM/DD/YYYY',
	        ignoreReadonly: true,
	        sideBySide: true, 
	        //defaultDate: moment(new Date(today))
	    	
		});	
					
	});	
	

	
    function validateAddEquipmentForm()
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
		
		if (selectedValueStatusCode == "")
		{
			swal("You must choose a status code for a piece of equipment.");
			return false;
		}
		

        if (document.frmAddEquipment.txtAssetTag1.value == "") {
            swal("You must enter at least one asset tag.");
            return false;
        }

        if (document.frmAddEquipment.txtPurchaseCost.value == "") {
            swal("You must enter the purchase cost.");
            return false;
        }
		
        if (document.frmAddEquipment.txtPurchaseCost.value != "") {
        
        	if (isNaN(document.frmAddEquipment.txtPurchaseCost.value)) {
            	swal("Please enter numbers only for the purchase cost.");
            	return false;
           	}
        }
        
        if (document.frmAddEquipment.txtReplacementCost.value == "") {
            swal("You must enter the replacement cost.");
            return false;
        }
        
        if (document.frmAddEquipment.txtReplacementCost.value != "") {
        
        	if (isNaN(document.frmAddEquipment.txtReplacementCost.value)) {
            	swal("Please enter numbers only for the replacement cost.");
            	return false;
           	}
        }


        if ((document.frmAddEquipment.txtWarrentyStartDate.value != "") && (document.frmAddEquipment.txtWarrentyEndDate.value != "")) {
   
			var d1 = new Date(document.frmAddEquipment.txtWarrentyStartDate.value);
			var d2 = new Date(document.frmAddEquipment.txtWarrentyEndDate.value);
			
			//var same = d1.getTime() === d2.getTime();
			//var notSame = d1.getTime() !== d2.getTime();
			     
        	if (d1.getTime() > d2.getTime()) {
            	swal("Warranty start date should be older than warranty end date.");
            	return false;
           	}
        }


		var ddlCurrentCondition = document.getElementById("selCurrentConditionIntRecID");
		var selectedCurrentCondition = ddlCurrentCondition.options[ddlCurrentCondition.selectedIndex].value;
		
		if (selectedCurrentCondition == "")
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
	
</style>


<h1 class="page-header"> Add New Equipment</h1>

<div class="container">

<form method="POST" action="<%= BaseURL %>equipment/equipment/addEquipment_submit.asp" name="frmAddEquipment" id="frmAddEquipment">
		

    <div class="row">
    
        <div class="col-md-3">
			
				<div class="row row-line">
					<div class="form-group col-lg-12">
						<label for="selModelIntRecID" class="col-sm-3 control-label">Model</label>	
		    			<div class="col-sm-8">
						  	<select class="form-control required" name="selModelIntRecID" id="selModelIntRecID">
						  			<option value="" selected="selected">Select Model of Equipment</option>
							      	<% 'Get all equipment modeals
							      	  	SQLEquipModels = "SELECT * FROM EQ_Models ORDER BY Model ASC"
			
										Set cnnEquipModels = Server.CreateObject("ADODB.Connection")
										cnnEquipModels.open (Session("ClientCnnString"))
										Set rsEquipModels = Server.CreateObject("ADODB.Recordset")
										rsEquipModels.CursorLocation = 3 
										Set rsEquipModels = cnnEquipModels.Execute(SQLEquipModels)
										If not rsEquipModels.EOF Then
											Do
												Response.Write("<option value='" & rsEquipModels("InternalRecordIdentifier") & "'>" & rsEquipModels("Model") & "</option>")
												rsEquipModels.movenext
											Loop until rsEquipModels.eof
										End If
										set rsEquipModels = Nothing
										cnnEquipModels.close
										set cnnEquipModels = Nothing
									%>
							</select>
		    			</div>
					</div>
				</div>
				
		
				<div class="row row-line">
					<div class="form-group col-lg-12">
						<label for="selStatusCodeIntRecID" class="col-sm-3 control-label">Status Code</label>	
		    			<div class="col-sm-8">
						  	<select class="form-control required" name="selStatusCodeIntRecID" id="selStatusCodeIntRecID">
						  			<option value="" selected="selected">Select Status Code of Equipment</option>
						  			
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
												Response.Write("<option value='" & rsEquipStatusCodes("InternalRecordIdentifier") & "'>" & rsEquipStatusCodes("statusDesc") & " (" & rsEquipStatusCodes("statusBackendSystemCode") & ")</option>")
												rsEquipStatusCodes.movenext
											Loop until rsEquipStatusCodes.eof
										End If
										set rsEquipStatusCodes = Nothing
										cnnEquipStatusCodes.close
										set cnnEquipStatusCodes = Nothing
									%>
							</select>
		    			</div>
					</div>
				</div>
				
				
				
				<div class="row row-line">
					<div class="form-group col-lg-12">
						<label for="txtSerialNumber" class="col-sm-3 control-label">Serial Number</label>	
		    			<div class="col-sm-8">
		    				<i class="inside fa fa-barcode" aria-hidden="true"></i>
		    				<input type="text" class="form-control inp" id="txtSerialNumber" name="txtSerialNumber">
		    			</div>
					</div>
				</div>		
				
				<div class="row row-line">
					<div class="form-group col-lg-12">
						<label for="txtAssetTag1" class="col-sm-3 control-label">Asset Tag 1</label>	
		    			<div class="col-sm-8">
		    				<i class="inside fa fa-tag"></i>
		    				<input type="text" class="form-control required inp" id="txtAssetTag1" name="txtAssetTag1">
		    				<button type="button" class="btn btn-success btn-xs pull-right" style="margin-top:5px;" id="btnGenerateAssetTag1">Generate Asset Tag</button>
		    			</div>
					</div>
				</div>		
				<div class="row row-line">
					<div class="form-group col-lg-12">
						<label for="txtAssetTag2" class="col-sm-3 control-label">Asset Tag 2</label>	
		    			<div class="col-sm-8">
		    				<i class="inside fa fa-tag"></i>
		    				<input type="text" class="form-control inp" id="txtAssetTag2" name="txtAssetTag2">
		    				<button type="button" class="btn btn-success btn-xs pull-right" style="margin-top:5px;" id="btnGenerateAssetTag2">Generate Asset Tag</button>
		    			</div>
					</div>
				</div>	
				<div class="row row-line">
					<div class="form-group col-lg-12">
						<label for="txtAssetTag3" class="col-sm-3 control-label">Asset Tag 3</label>	
		    			<div class="col-sm-8">
		    				<i class="inside fa fa-tag"></i>
		    				<input type="text" class="form-control inp" id="txtAssetTag3" name="txtAssetTag3">
		    				<button type="button" class="btn btn-success btn-xs pull-right" style="margin-top:5px;" id="btnGenerateAssetTag3">Generate Asset Tag</button>
		    			</div>
					</div>
				</div>									
				<div class="row row-line">
					<div class="form-group col-lg-12">
						<label for="txtAssetTag4" class="col-sm-3 control-label">Asset Tag 4</label>	
		    			<div class="col-sm-8">
		    				<i class="inside fa fa-tag"></i>
		    				<input type="text" class="form-control inp" id="txtAssetTag4" name="txtAssetTag4">
		    				<button type="button" class="btn btn-success btn-xs pull-right" style="margin-top:5px;" id="btnGenerateAssetTag4">Generate Asset Tag</button>
		    			</div>
					</div>
				</div>		
				
		</div>	


        <div class="col-md-3">
					
				<div class="row row-line">
					<div class="form-group col-lg-12">
						<label for="selAcquisitionIntRecID" class="col-sm-3 control-label">Acquisition Type</label>	
		    			<div class="col-sm-8">
						  	<select class="form-control" name="selAcquisitionCodeIntRecID" id="selAcquisitionCodeIntRecID">
						  			<option value="" selected="selected">Select Acquisition Type for Equipment</option>
						  									  			
						  			<% If userCanEditEqpOnFly(Session("UserNo")) Then %>
						  				<option value="addnewacquisitioncode">+ ADD NEW ACQUISITION CODE</option>
						  			<% End If %>

							      	<% 'Get all Acquisition Codes
							      	  	SQLEquipAcquisition = "SELECT * FROM EQ_AcquisitionCodes ORDER BY acquisitionCode ASC"
			
										Set cnnEquipAcquisition = Server.CreateObject("ADODB.Connection")
										cnnEquipAcquisition.open (Session("ClientCnnString"))
										Set rsEquipAcquisition = Server.CreateObject("ADODB.Recordset")
										rsEquipAcquisition.CursorLocation = 3 
										Set rsEquipAcquisition = cnnEquipAcquisition.Execute(SQLEquipAcquisition)
										If not rsEquipAcquisition.EOF Then
											Do
												Response.Write("<option value='" & rsEquipAcquisition("InternalRecordIdentifier") & "'>" & rsEquipAcquisition("AcquisitionDesc") & " (" & rsEquipAcquisition("acquisitionCode") & ")</option>")
												rsEquipAcquisition.movenext
											Loop until rsEquipAcquisition.eof
										End If
										set rsEquipAcquisition = Nothing
										cnnEquipAcquisition.close
										set cnnEquipAcquisition = Nothing
									%>
							</select>
		    			</div>
					</div>
				</div>
	
				

				<div class="row row-line">
					<div class="form-group col-lg-12">
						<label for="selVendorIntRecID" class="col-sm-3 control-label">Purchased From Vendor</label>	
		    			<div class="col-sm-8">
						  	<select class="form-control" name="selVendorIntRecID" id="selVendorIntRecID">
						  			<option value="" selected="selected">Select Vendor of Equipment</option>
							      	<% 'Get all vendors
							      	  	SQLEquipVendors = "SELECT * FROM AP_Vendor ORDER BY vendorCompanyName ASC"
			
										Set cnnEquipVendors = Server.CreateObject("ADODB.Connection")
										cnnEquipVendors.open (Session("ClientCnnString"))
										Set rsEquipVendors = Server.CreateObject("ADODB.Recordset")
										rsEquipVendors.CursorLocation = 3 
										Set rsEquipVendors = cnnEquipVendors.Execute(SQLEquipVendors)
										If not rsEquipVendors.EOF Then
											Do
												Response.Write("<option value='" & rsEquipVendors("InternalRecordIdentifier") & "'>" & rsEquipVendors("vendorCompanyName") & "</option>")
												rsEquipVendors.movenext
											Loop until rsEquipVendors.eof
										End If
										set rsEquipVendors = Nothing
										cnnEquipVendors.close
										set cnnEquipVendors = Nothing
									%>
							</select>
		    			</div>
					</div>
				</div>
				
				<div class="row row-line">
					<div class="form-group col-lg-12">
						<label for="txtPurchasedViaPONumber" class="col-sm-3 control-label">Purchase PO Number</label>	
		    			<div class="col-sm-8">
		    				<input type="text" class="form-control inp" id="txtPurchasedViaPONumber" name="txtPurchasedViaPONumber">
		    			</div>
					</div>
				</div>		



				<div class="row row-line">
					<div class="form-group col-lg-12">
						<label for="txtPurchasedViaPONumber" class="col-sm-3 control-label">Purchase Date</label>
						<div class="col-lg-8">								  	
			                <div class="input-group date" id="datetimepickerPurchaseDate">
			                    <input type="text" class="form-control" name="txtPurchaseDate" id="txtPurchaseDate">
			                    <span class="input-group-addon">
			                        <span class="glyphicon glyphicon-calendar"></span>
			                    </span>
			                </div>
			             </div>
					</div>
				</div>
				
			
		
				<div class="row row-line">
					<div class="form-group col-lg-12">
						<label for="txtPurchaseCost" class="col-sm-3 control-label">Purchase Cost</label>	
		    			<div class="col-sm-8">
		    				<i class="inside fa fa-usd"></i>
		    				<input type="text" class="form-control required inp" id="txtPurchaseCost" name="txtPurchaseCost">		    			
		    			</div>
					</div>
				</div>
				
				<div class="row row-line">
					<div class="form-group col-lg-12">
						<label for="txtReplacementCost" class="col-sm-3 control-label">Replacement Cost</label>	
		    			<div class="col-sm-8">
		    				 <i class="inside fa fa-usd"></i>
		    				<input type="text" class="form-control required inp" id="txtReplacementCost" name="txtReplacementCost">
		    			</div>
					</div>
				</div>
				
    </div><!-- eof col-md-3 -->
    
        
		
	    <div class="col-md-3">	
	
	
				<div class="row row-line">
					<div class="form-group col-lg-12">
						<label for="selAcquiredConditionIntRecID" class="col-sm-3 control-label">Acquired Condition</label>	
		    			<div class="col-sm-8">
						  	<select class="form-control" name="selAcquiredConditionIntRecID" id="selAcquiredConditionIntRecID">
						  			<option value="" selected="selected">Select Acquired Condition of Equipment</option>
						  			
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
												Response.Write("<option value='" & rsEquipCondition("InternalRecordIdentifier") & "'>" & rsEquipCondition("Condition") & "</option>")
												rsEquipCondition.movenext
											Loop until rsEquipCondition.eof
										End If
										set rsEquipCondition = Nothing
										cnnEquipCondition.close
										set cnnEquipCondition = Nothing
									%>
							</select>
		    			</div>
					</div>
				</div>
	
	
				<div class="row row-line">
					<div class="form-group col-lg-12">
						<label for="selCurrentConditionIntRecID" class="col-sm-3 control-label">Current Condition</label>	
		    			<div class="col-sm-8">
						  	<select class="form-control required" name="selCurrentConditionIntRecID" id="selCurrentConditionIntRecID">
						  			<option value="" selected="selected">Select Current Condition of Equipment</option>
						  			
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
												Response.Write("<option value='" & rsEquipCondition("InternalRecordIdentifier") & "'>" & rsEquipCondition("Condition") & "</option>")
												rsEquipCondition.movenext
											Loop until rsEquipCondition.eof
										End If
										set rsEquipCondition = Nothing
										cnnEquipCondition.close
										set cnnEquipCondition = Nothing
									%>
							</select>
		    			</div>
					</div>
				</div>
	
				<div class="row row-line">
					<div class="form-group col-lg-12">
						<label for="txtPurchasedViaPONumber" class="col-sm-3 control-label">Warranty Start Date</label>
						<div class="col-lg-8">									  	
			                <div class="input-group date" id="datetimepickerWarrentyStartDate">
			                    <input type="text" class="form-control" name="txtWarrentyStartDate" id="txtWarrentyStartDate">
			                    <span class="input-group-addon">
			                        <span class="glyphicon glyphicon-calendar"></span>
			                    </span>
			                </div>
			             </div>
					</div>
				</div>
	
	
				<div class="row row-line">
					<div class="form-group col-lg-12">
						<label for="txtPurchasedViaPONumber" class="col-sm-3 control-label">Warranty End Date</label>
						<div class="col-lg-8">										  	
			                <div class="input-group date" id="datetimepickerWarrentyEndDate">
			                    <input type="text" class="form-control" name="txtWarrentyEndDate" id="txtWarrentyEndDate">
			                    <span class="input-group-addon">
			                        <span class="glyphicon glyphicon-calendar"></span>
			                    </span>
			                </div>
			             </div>
					</div>
				</div>
	
	
	    </div>
	    
	    
	    <div class="col-md-3">
	     
				
				
				<div class="row row-line">
					<div class="form-group col-lg-12">
						<label for="txtColor" class="col-sm-3 control-label">Color</label>	
		    			<div class="col-sm-8">
		    				<input type="text" class="form-control inp" id="txtColor" name="txtColor">
		    			</div>
					</div>
				</div>		
				
	
	
				<div class="row row-line">
					<div class="form-group col-lg-12">
						<label for="txtSize" class="col-sm-3 control-label">Size</label>	
		    			<div class="col-sm-8">
		    				<input type="text" class="form-control inp" id="txtSize" name="txtSize">
		    			</div>
					</div>
				</div>		
				
			    <!-- cancel / submit !-->
				<div class="row row-line">
					&nbsp;
				</div>
				
	
				<div class="row row-line">
					<div class="form-group col-lg-12">
						<label for="txtComments" class="col-sm-3 control-label">Comments</label>	
		    			<div class="col-sm-8">
		    				<textarea class="form-control inp" id="txtComments" name="txtComments" rows="10"></textarea>
		    			</div>
					</div>
				</div>		
				
				
	
	     </div>
	     
	     
	 
	</div><!-- eof row -->	


	<div class="row pull-right">
		<div class="col-lg-12">
		    <!-- cancel / submit !-->
			<div class="row row-line">
				<div class="col-lg-12 alertbutton">
					<div class="col-lg-12">
						<a href="<%= BaseURL %>equipment/equipment/findEquipment.asp">
		    				<button type="button" class="btn btn-default">&lsaquo; Cancel &amp; Go Back To Equipment Search</button>
						</a>
						<button type="button" class="btn btn-primary" id="btnSubmitAddEquipmentForm"><i class="far fa-save"></i> Add New Equipment</button>
					</div>
			    </div>
			</div>
		</div>
	</div>


</form>  
  
  

</div><!-- eof content container -->

<!-- **************************************************************************************************************************** -->
<!-- MODALS FOR EDIT EQUIPMENT ON THE FLY BEGIN HERE !-->
<!-- **************************************************************************************************************************** -->

<!--#include file="editEquipTablesOnTheFlyModals.asp "-->

<!-- **************************************************************************************************************************** -->
<!-- MODALS FOR EDIT EQUIPMENT ON THE FLY END HERE !-->
<!-- **************************************************************************************************************************** -->

<!--#include file="../../inc/footer-main.asp"-->

<!--#include file="../inc/header-prospecting.asp"-->
<!--#include file="../inc/InSightFuncs_Prospecting.asp"-->
<% InternalRecordIdentifier = Request.QueryString("i") 
If InternalRecordIdentifier = "" Then Response.Redirect("main.asp")

SQLProspect = "SELECT * FROM PR_Prospects where InternalRecordIdentifier = " & InternalRecordIdentifier 

Set cnnProspect = Server.CreateObject("ADODB.Connection")
cnnProspect.open (Session("ClientCnnString"))
Set rsProspect = Server.CreateObject("ADODB.Recordset")
rsProspect.CursorLocation = 3 
Set rsProspect = cnnProspect.Execute(SQLProspect)

If not rsProspect.EOF Then

	Company = rsProspect("Company")
	Street= rsProspect("Street")
	City= rsProspect("City")
	State= rsProspect("State")
	PostalCode = rsProspect("PostalCode")
	Country= rsProspect("Country")
	Suite= rsProspect("Floor_Suite_Room__c")								
	Website= rsProspect("Website")								
	LeadSourceNumber = rsProspect("LeadSourceNumber")
	LeadSource = GetLeadSourceByNum(LeadSourceNumber)				
	StageNumber = GetProspectCurrentStageByProspectNumber(InternalRecordIdentifier)
	IndustryNumber = rsProspect("IndustryNumber")	
	Industry = GetIndustryByNum(IndustryNumber)											
	OwnerUserNo = rsProspect("OwnerUserNo")				
	CreatedDate= rsProspect("CreatedDate")
	CreatedByUserNo= rsProspect("CreatedByUserNo")				
	TelemarketerUserNo = rsProspect("TelemarketerUserNo")
	Telemarketer = GetUserDisplayNameByUserNo(TelemarketerUserNo)
	ProjectedGPSpend= rsProspect("ProjectedGPSpend")
	NumberOfPantries = rsProspect("NumberOfPantries")
	EmployeeRangeNumber = rsProspect("EmployeeRangeNumber")
	NumEmployees = GetEmployeeRangeByNum(EmployeeRangeNumber)
	CreatedDate = rsProspect("CreatedDate")
	FormerCustNum = rsProspect("FormerCustNum")
	CancelDate = rsProspect("CancelDate")
	LeaseExpirationDate = rsProspect("LeaseExpirationDate")	
	ContractExpirationDate = rsProspect("ContractExpirationDate")
	Comments = rsProspect("Comments")
	CurrentOffering = rsProspect("CurrentOffering")			

End If

SQLContacts1 = "SELECT * FROM PR_ProspectContacts WHERE ProspectIntRecID = " & InternalRecordIdentifier & " AND PrimaryContact = 1"

Set cnnContacts1 = Server.CreateObject("ADODB.Connection")
cnnContacts1.open (Session("ClientCnnString"))
Set rsContacts1 = Server.CreateObject("ADODB.Recordset")
rsContacts1.CursorLocation = 3 
Set rsContacts1 = cnnContacts1.Execute(SQLContacts1)

If not rsContacts1.EOF Then

  	primarySuffix = rsContacts1("Suffix")
  	primaryFirstName = rsContacts1("FirstName")
	primaryLastName = rsContacts1("LastName")	
	primaryTitleNumber = rsContacts1("ContactTitleNumber")
	primaryTitle = GetContactTitleByNum(primaryTitleNumber)
	primaryEmail = rsContacts1("Email") 
	primaryPhone = rsContacts1("Phone")
	primaryPhoneExt = rsContacts1("PhoneExt")
	primaryCell = rsContacts1("Cell")
	primaryFax = rsContacts1("Fax")

					
End If
Set rsContacts1 = Nothing
cnnContacts1.Close


PrimaryCompetitorID = GetPrimaryCompetitorIDByProspectNumber(InternalRecordIdentifier)

If PrimaryCompetitorID <> "" Then
	PrimaryCompetitorName = GetCompetitorByNum(PrimaryCompetitorID)

	SQLCompetitors = "SELECT * FROM PR_ProspectCompetitors WHERE CompetitorRecID = " & PrimaryCompetitorID & " AND ProspectRecID = " &  InternalRecordIdentifier
	
	Set cnnCompetitors = Server.CreateObject("ADODB.Connection")
	cnnCompetitors.open (Session("ClientCnnString"))
	Set rsCompetitors = Server.CreateObject("ADODB.Recordset")
	rsCompetitors.CursorLocation = 3 
	Set rsCompetitors = cnnCompetitors.Execute(SQLCompetitors)
	
	If not rsCompetitors.EOF Then
	
		BottledWater = rsCompetitors("BottledWater")
		FilteredWater = rsCompetitors("FilteredWater")
		OCS = rsCompetitors("OCS")
		OCS_Supply = rsCompetitors("OCS_Supply")
		OfficeSupplies = rsCompetitors("OfficeSupplies")
		Vending = rsCompetitors("Vending")
		Micromarket = rsCompetitors("Micromarket")
		Pantry = rsCompetitors("Pantry")
						
	End If
	Set rsCompetitors = Nothing
	cnnCompetitors.Close
	Set cnnCompetitors = Nothing
	
	
	If BottledWater = vbTrue Then BottledWater = "Bottled Water" Else BottledWater = ""
	If FilteredWater = vbTrue Then FilteredWater = "Filtered Water" Else FilteredWater = ""
	If OCS = vbTrue Then OCS = "OCS" Else OCS = ""
	If OCS_Supply = vbTrue Then OCS_Supply = "OCS Supply" Else OCS_Supply = ""
	If OfficeSupplies = vbTrue Then OfficeSupplies = "Office Supplies " Else OfficeSupplies = ""
	If Vending = vbTrue Then Vending = "Vending" Else Vending = ""
	If Micromarket = vbTrue Then Micromarket = "Micromarkets" Else Micromarket = ""
	If Pantry = vbTrue Then Pantry = "Pantry" Else Pantry = ""
Else

	PrimaryCompetitorName = ""
	BottledWater = ""
	FilteredWater = ""
	OCS = ""
	OCS_Supply = ""
	OfficeSupplies = ""
	Vending = ""
	Micromarket = ""
	Pantry = ""

End If


		
SQLProspectActivities = "SELECT * FROM PR_ProspectActivities where ProspectRecID = " & InternalRecordIdentifier & " AND Status IS NULL"

Set cnnProspectActivities = Server.CreateObject("ADODB.Connection")
cnnProspectActivities.open (Session("ClientCnnString"))
Set rsProspectActivities = Server.CreateObject("ADODB.Recordset")
rsProspectActivities.CursorLocation = 3 
Set rsProspectActivities = cnnProspectActivities.Execute(SQLProspectActivities)

If not rsProspectActivities.EOF Then

  	ActivityRecID = rsProspectActivities("ActivityRecID")
  	nextActivityNum = rsProspectActivities("ActivityRecID")
  	nextActivity = GetActivityByNum(nextActivityNum)
  	nextActivityDueDate = FormatDateTime(rsProspectActivities("ActivityDueDate"),2)
  	nextActivityDueDateTime = FormatDateTime(rsProspectActivities("ActivityDueDate"),2) & " " & FormatDateTime(rsProspectActivities("ActivityDueDate"),3)
	daysOld = DateDiff("d",rsProspectActivities("RecordCreationDateTime"),Now())
	daysOverdue = DateDiff("d",rsProspectActivities("ActivityDueDate"),Now())	
				
End If
Set rsProspectActivities = Nothing
cnnProspectActivities.Close
Set cnnProspectActivities = Nothing

%>


<% 'Read edit prospect tab color settings
SQL = "SELECT * FROM Settings_Global"
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
If not rs.EOF Then
	CRMTabLogColor = rs("CRMTabLogColor")
	CRMTabProductsColor = rs("CRMTabProductsColor")
	CRMTabEquipmentColor = rs("CRMTabEquipmentColor")
	CRMTabDocumentsColor = rs("CRMTabDocumentsColor")
	CRMTabLocationColor = rs("CRMTabLocationColor")
	CRMTabContactsColor = rs("CRMTabContactsColor")
	CRMTabCompetitorsColor = rs("CRMTabCompetitorsColor")
	CRMTabOpportunityColor = rs("CRMTabOpportunityColor")
	CRMTabAuditTrailColor	 = rs("CRMTabAuditTrailColor")
	CRMTileOfferingColor = rs("CRMTileOfferingColor")
	CRMTileCompetitorColor = rs("CRMTileCompetitorColor")
	CRMTileDollarsColor = rs("CRMTileDollarsColor")
	CRMTileActivityColor = rs("CRMTileActivityColor")
	CRMTileStageColor = rs("CRMTileStageColor")
	CRMTileOwnerColor = rs("CRMTileOwnerColor")
	CRMTileCommentsColor = rs("CRMTileCommentsColor")
	CRMHideLocationTab = rs("CRMHideLocationTab")
	CRMHideProductsTab = rs("CRMHideProductsTab")
	CRMHideEquipmentTab = rs("CRMHideEquipmentTab")	
End If

SQL = "SELECT * FROM Settings_Prospecting"
Set rs = cnn8.Execute(SQL)
If not rs.EOF Then
	TabSocialMediaColor = rs("TabSocialMediaColor")
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing

If CRMTabLogColor = "" Then CRMTabLogColor = "#D8F9D1"
If IsNull(CRMTabLogColor) Then CRMTabLogColor = "#D8F9D1"

If CRMTabProductsColor = "" Then CRMTabProductsColor = "#D8F9D1"
If IsNull(CRMTabProductsColor) Then CRMTabProductsColor = "#D8F9D1"

If CRMTabEquipmentColor = "" Then CRMTabEquipmentColor = "#FFA500"
If IsNull(CRMTabEquipmentColor) Then CRMTabEquipmentColor = "#FFA500"

If CRMTabDocumentsColor = "" Then CRMTabDocumentsColor = "#F6F6F6"
If IsNull(CRMTabDocumentsColor) Then CRMTabDocumentsColor = "#F6F6F6"

If CRMTabLocationColor = "" Then CRMTabLocationColor = "#D8F9D1"
If IsNull(CRMTabLocationColor) Then CRMTabLocationColor = "#D8F9D1"

If CRMTabContactsColor = "" Then CRMTabContactsColor = "#FCB3B3"
If IsNull(CRMTabContactsColor) Then CRMTabContactsColor = "#FCB3B3"

If CRMTabCompetitorsColor = "" Then CRMTabCompetitorsColor = "#FCB3B3"
If IsNull(CRMTabCompetitorsColor) Then CRMTabCompetitorsColor = "#FCB3B3"

If CRMTabOpportunityColor = "" Then CRMTabOpportunityColor = "#FFA500"
If IsNull(CRMTabOpportunityColor) Then CRMTabOpportunityColor = "#FFA500"

If CRMTabAuditTrailColor = "" Then CRMTabAuditTrailColor = "#FFA500"
If IsNull(CRMTabAuditTrailColor) Then CRMTabAuditTrailColor = "#FFA500"

If CRMTileOfferingColor = "" Then CRMTileOfferingColor = "#3498db"
If IsNull(CRMTileOfferingColor) Then CRMTileOfferingColor = "#3498db"

If CRMTileCompetitorColor = "" Then CRMTileCompetitorColor = "#9b6bcc"
If IsNull(CRMTileCompetitorColor) Then CRMTileCompetitorColor = "#9b6bcc"

If CRMTileDollarsColor = "" Then CRMTileDollarsColor = "#2ecc71"
If IsNull(CRMTileDollarsColor) Then CRMTileDollarsColor = "#2ecc71"

If CRMTileActivityColor = "" Then CRMTileActivityColor = "#f1c40f"
If IsNull(CRMTileActivityColor) Then CRMTileActivityColor = "#f1c40f"

If CRMTileStageColor = "" Then CRMTileStageColor = "#e67e22"
If IsNull(CRMTileStageColor) Then CRMTileStageColor = "#e67e22"

If CRMTileOwnerColor = "" Then CRMTileOwnerColor = "#95a5a6"
If IsNull(CRMTileOwnerColor) Then CRMTileOwnerColor = "#95a5a6"

If CRMTileCommentsColor = "" Then CRMTileCommentsColor = "#d43f3a"
If IsNull(CRMTileCommentsColor) Then CRMTileCommentsColor = "#d43f3a"

%>



<script src="<%= baseURL %>js/moment.min.js" type="text/javascript"></script>
<script type="text/javascript" src="<%= BaseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.min.js"></script>
<link href="<%= BaseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.css" rel="stylesheet">

<script>

	function highlightBlankFields(){
	  $(".input-group .form-control").each(function() {
	     var val = $(this).val();
	     if(val == "" || val == 0) {
	       $(this).css({ backgroundColor:'#ffff99' });
	     }
	     else {
	     	$(this).css({ backgroundColor:'#fff' });
	     }
	  });
	}

	$(document).ready(function() {
	

		highlightBlankFields();
		
	   var phones = [{ "mask": "(###) ###-####"}];
	    $('#txtPhoneNumber').inputmask({ 
	        mask: phones, 
	        greedy: false, 
	        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
	    $('#txtCellPhoneNumber').inputmask({ 
	        mask: phones, 
	        greedy: false, 
	        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
	    $('#txtFaxNumber').inputmask({ 
	        mask: phones, 
	        greedy: false, 
	        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
		
		
		$(".input-group .form-control").blur(function() {
		  highlightBlankFields()
		});	

		
	  $('input, select, textarea').each(
		    function(index){  
		        var input = $(this);
		        //console.log(input.attr('name'));
		    }
	  ); 
		
	  $("#txtNumEmployees").change(function() {
	  
	  		if ($("#txtProjectedGPSpend").val() == '')
	  		{
				intRecID = $("#txtNumEmployees").val();
				projGPSpend = $("#" + intRecID).val();
				$("#txtProjectedGPSpend").val(projGPSpend);
			}
	  });

 
		$("#txtFirstName").on({
		    mouseover: function() {$("#txtFirstNameLabel").stop().show(500);},
		    mouseout: function() {$("#txtFirstNameLabel").stop().hide(500);}
		})	
		$("#txtFirstNameIcon").on({
		    mouseover: function() {$("#txtFirstNameLabel").stop().show(500);},
		    mouseout: function() {$("#txtFirstNameLabel").stop().hide(500);}
		})	
		
		
		$("#txtLastName").on({
		    mouseover: function() {$("#txtLastNameLabel").stop().show(500);},
		    mouseout: function() {$("#txtLastNameLabel").stop().hide(500);}
		})	
		$("#txtLastNameIcon").on({
		    mouseover: function() {$("#txtLastNameLabel").stop().show(500);},
		    mouseout: function() {$("#txtLastNameLabel").stop().hide(500);}
		})	

		
		$("#txtCompanyName").on({
		    mouseover: function() {$("#txtCompanyNameLabel").stop().show(500);},
		    mouseout: function() {$("#txtCompanyNameLabel").stop().hide(500);}
		})
		$("#txtCompanyNameIcon").on({
		    mouseover: function() {$("#txtCompanyNameLabel").stop().show(500);},
		    mouseout: function() {$("#txtCompanyNameLabel").stop().hide(500);}
		})	

			
		$("#txtTitle").on({
		    mouseover: function() {$("#txtTitleLabel").stop().show(500);},
		    mouseout: function() {$("#txtTitleLabel").stop().hide(500);}
		})	
		$("#txtTitleIcon").on({
		    mouseover: function() {$("#txtTitleLabel").stop().show(500);},
		    mouseout: function() {$("#txtTitleLabel").stop().hide(500);}
		})	

		
		$("#txtAddressLine1").on({
		    mouseover: function() {$("#txtAddressLine1Label").stop().show(500);},
		    mouseout: function() {$("#txtAddressLine1Label").stop().hide(500);}
		})	
		$("#txtAddressLine1Icon").on({
		    mouseover: function() {$("#txtAddressLine1Label").stop().show(500);},
		    mouseout: function() {$("#txtAddressLine1Label").stop().hide(500);}
		})	

		
		$("#txtAddressLine2").on({
		    mouseover: function() {$("#txtAddressLine2Label").stop().show(500);},
		    mouseout: function() {$("#txtAddressLine2Label").stop().hide(500);}
		})	
		$("#txtAddressLine2Icon").on({
		    mouseover: function() {$("#txtAddressLine2Label").stop().show(500);},
		    mouseout: function() {$("#txtAddressLine2Label").stop().hide(500);}
		})	

		
		$("#txtCity").on({
		    mouseover: function() {$("#txtCityLabel").stop().show(500);},
		    mouseout: function() {$("#txtCityLabel").stop().hide(500);}
		})	
		$("#txtCityIcon").on({
		    mouseover: function() {$("#txtCityLabel").stop().show(500);},
		    mouseout: function() {$("#txtCityLabel").stop().hide(500);}
		})	

		
		$("#txtState").on({
		    mouseover: function() {$("#txtStateLabel").stop().show(500);},
		    mouseout: function() {$("#txtStateLabel").stop().hide(500);}
		})	
		$("#txtStateIcon").on({
		    mouseover: function() {$("#txtStateLabel").stop().show(500);},
		    mouseout: function() {$("#txtStateLabel").stop().hide(500);}
		})	

		
		$("#txtZipCode").on({
		    mouseover: function() {$("#txtZipCodeLabel").stop().show(500);},
		    mouseout: function() {$("#txtZipCodeLabel").stop().hide(500);}
		})	
		$("#txtZipCodeIcon").on({
		    mouseover: function() {$("#txtZipCodeLabel").stop().show(500);},
		    mouseout: function() {$("#txtZipCodeLabel").stop().hide(500);}
		})	

		
		$("#txtCountry").on({
		    mouseover: function() {$("#txtCountryLabel").stop().show(500);},
		    mouseout: function() {$("#txtCountryLabel").stop().hide(500);}
		})	
		$("#txtCountryIcon").on({
		    mouseover: function() {$("#txtCountryLabel").stop().show(500);},
		    mouseout: function() {$("#txtCountryLabel").stop().hide(500);}
		})	

		
		$("#txtEmailAddress").on({
		    mouseover: function() {$("#txtEmailAddressLabel").stop().show(500);},
		    mouseout: function() {$("#txtEmailAddressLabel").stop().hide(500);}
		})	
		$("#txtEmailAddressIcon").on({
		    mouseover: function() {$("#txtEmailAddressLabel").stop().show(500);},
		    mouseout: function() {$("#txtEmailAddressLabel").stop().hide(500);}
		})	

		
		$("#txtWebsiteURL").on({
		    mouseover: function() {$("#txtWebsiteURLLabel").stop().show(500);},
		    mouseout: function() {$("#txtWebsiteURLLabel").stop().hide(500);}
		})	
		$("#txtWebsiteURLIcon").on({
		    mouseover: function() {$("#txtWebsiteURLLabel").stop().show(500);},
		    mouseout: function() {$("#txtWebsiteURLLabel").stop().hide(500);}
		})	

		
		$("#txtPhoneNumber").on({
		    mouseover: function() {$("#txtPhoneNumberLabel").stop().show(500);},
		    mouseout: function() {$("#txtPhoneNumberLabel").stop().hide(500);}
		})	
		$("#txtPhoneNumberIcon").on({
		    mouseover: function() {$("#txtPhoneNumberLabel").stop().show(500);},
		    mouseout: function() {$("#txtPhoneNumberLabel").stop().hide(500);}
		})	
		
		
		$("#txtPhoneNumberExt").on({
		    mouseover: function() {$("#txtPhoneNumberExtLabel").stop().show(500);},
		    mouseout: function() {$("#txtPhoneNumberExtLabel").stop().hide(500);}
		})	
		$("#txtPhoneNumberExtIcon").on({
		    mouseover: function() {$("#txtPhoneNumberExtLabel").stop().show(500);},
		    mouseout: function() {$("#txtPhoneNumberExtLabel").stop().hide(500);}
		})	
		
		
		$("#txtCellPhoneNumber").on({
		    mouseover: function() {$("#txtCellPhoneNumberLabel").stop().show(500);},
		    mouseout: function() {$("#txtCellPhoneNumberLabel").stop().hide(500);}
		})	
		$("#txtCellPhoneNumberIcon").on({
		    mouseover: function() {$("#txtCellPhoneNumberLabel").stop().show(500);},
		    mouseout: function() {$("#txtCellPhoneNumberLabel").stop().hide(500);}
		})	

		
		$("#txtFaxNumber").on({
		    mouseover: function() {$("#txtFaxNumberLabel").stop().show(500);},
		    mouseout: function() {$("#txtFaxNumberLabel").stop().hide(500);}
		})	
		$("#txtFaxNumberIcon").on({
		    mouseover: function() {$("#txtFaxNumberLabel").stop().show(500);},
		    mouseout: function() {$("#txtFaxNumberLabel").stop().hide(500);}
		})	

		
		$("#txtIndustry").on({
		    mouseover: function() {$("#txtIndustryLabel").stop().show(500);},
		    mouseout: function() {$("#txtIndustryLabel").stop().hide(500);}
		})	
		$("#txtIndustryIcon").on({
		    mouseover: function() {$("#txtIndustryLabel").stop().show(500);},
		    mouseout: function() {$("#txtIndustryLabel").stop().hide(500);}
		})	

		
		$("#txtProjectedGPSpend").on({
		    mouseover: function() {$("#txtProjectedGPSpendLabel").stop().show(500);},
		    mouseout: function() {$("#txtProjectedGPSpendLabel").stop().hide(500);}
		})	
		$("#txtProjectedGPSpendIcon").on({
		    mouseover: function() {$("#txtProjectedGPSpendLabel").stop().show(500);},
		    mouseout: function() {$("#txtProjectedGPSpendLabel").stop().hide(500);}
		})	

		
		$("#txtNumEmployees").on({
		    mouseover: function() {$("#txtNumEmployeesLabel").stop().show(500);},
		    mouseout: function() {$("#txtNumEmployeesLabel").stop().hide(500);}
		})	
		$("#txtNumEmployeesIcon").on({
		    mouseover: function() {$("#txtNumEmployeesLabel").stop().show(500);},
		    mouseout: function() {$("#txtNumEmployeesLabel").stop().hide(500);}
		})	

		
		$("#txtNumPantries").on({
		    mouseover: function() {$("#txtNumPantriesLabel").stop().show(500);},
		    mouseout: function() {$("#txtNumPantriesLabel").stop().hide(500);}
		})	
		$("#txtNumPantriesIcon").on({
		    mouseover: function() {$("#txtNumPantriesLabel").stop().show(500);},
		    mouseout: function() {$("#txtNumPantriesLabel").stop().hide(500);}
		})	

		
		$("#txtTelemarketerUserNo").on({
		    mouseover: function() {$("#txtTelemarketerUserNoLabel").stop().show(500);},
		    mouseout: function() {$("#txtTelemarketerUserNoLabel").stop().hide(500);}
		})	
		$("#txtTelemarketerUserNoIcon").on({
		    mouseover: function() {$("#txtTelemarketerUserNoLabel").stop().show(500);},
		    mouseout: function() {$("#txtTelemarketerUserNoLabel").stop().hide(500);}
		})	

		
		$("#txtPrimaryCompetitor").on({
		    mouseover: function() {$("#txtPrimaryCompetitorLabel").stop().show(500);},
		    mouseout: function() {$("#txtPrimaryCompetitorLabel").stop().hide(500);}
		})	
		$("#txtPrimaryCompetitorIcon").on({
		    mouseover: function() {$("#txtPrimaryCompetitorLabel").stop().show(500);},
		    mouseout: function() {$("#txtPrimaryCompetitorLabel").stop().hide(500);}
		})	

		
		$("#txtLeadSource").on({
		    mouseover: function() {$("#txtLeadSourceLabel").stop().show(500);},
		    mouseout: function() {$("#txtLeadSourceLabel").stop().hide(500);}
		})	
		$("#txtLeadSourceIcon").on({
		    mouseover: function() {$("#txtLeadSourceLabel").stop().show(500);},
		    mouseout: function() {$("#txtLeadSourceLabel").stop().hide(500);}
		})	

		
		$("#txtFormerCustomerNumber").on({
		    mouseover: function() {$("#txtFormerCustomerNumberLabel").stop().show(500);},
		    mouseout: function() {$("#txtFormerCustomerNumberLabel").stop().hide(500);}
		})	
		$("#txtFormerCustomerNumberIcon").on({
		    mouseover: function() {$("#txtFormerCustomerNumberLabel").stop().show(500);},
		    mouseout: function() {$("#txtFormerCustomerNumberLabel").stop().hide(500);}
		})	
		

		
	});
</script>


<script language="JavaScript">
<!--
	function isValidPhone(p) {
	  //var phoneRe = /^[2-9]\d{2}[2-9]\d{2}\d{4}$/;
	  //var phoneRe = /^[\+]?[(]?[0-9]{3}[)]?[-\s\.]?[0-9]{3}[-\s\.]?[0-9]{4,6}$/;
	  var phoneRe = /^(1\s|1|)?((\(\d{3}\))|\d{3})(\-|\s)?(\d{3})(\-|\s)?(\d{4})$/;
	  var digits = p.replace(/\D/g, "");
	  return phoneRe.test(digits);
	}
	
	function isValidEmail(email) 
	{
	    var re = /\S+@\S+\.\S+/;
	    return re.test(email);
	}	

   function validateAddProspectForm()
    {
    
       if (document.frmEditProspect.txtCompanyName.value == "") {
            swal("Company name cannot be blank.");
            return false;
       }
       if ((document.frmEditProspect.txtEmailAddress.value !== "") && (isValidEmail(document.frmEditProspect.txtEmailAddress.value) == false)) {
           swal("The email address is invalid. Please enter any format like the following: anystring@anystring.anystring.");
           return false;
       }
       if ((document.frmEditProspect.txtPhoneNumber.value !== "") && (isValidPhone(document.frmEditProspect.txtPhoneNumber.value) == false)) {
           swal("The phone number is invalid. Please enter the number in the following format: (555) 555-5555.");
           return false;
       }
       if ((document.frmEditProspect.txtPhoneNumberExt.value !== "") && (document.frmEditProspect.txtPhoneNumber.value == "")) {
           swal("A phone extension was added with no phone number. Please enter a phone number or clear the extension.");
           return false;
       }       
       if ((document.frmEditProspect.txtCellPhoneNumber.value !== "") && (isValidPhone(document.frmEditProspect.txtCellPhoneNumber.value) == false)) {
           swal("The cell phone number is invalid. Please enter the number in the following format: (555) 555-5555.");
           return false;
       }
       if ((document.frmEditProspect.txtFaxNumber.value !== "") && (isValid(document.frmEditProspect.txtFaxNumber.value) == false)) {
           swal("The fax number is invalid. Please enter the number in the following format: (555) 555-5555.");
           return false;
       }
	   
	   var radio = document.getElementsByName('radStage'); // get all radio buttons
	   var isChecked = 0; // default is 0 
	   for(var i=0; i<radio.length;i++) { // go over all the radio buttons with name 'radStage'
			if(radio[i].checked) isChecked = 1; // if one of them is checked - tell me
		}

 		if(isChecked == 0) { // if the default value stayed the same, check the first radio button
   			swal("Please select a stage.");
   			return false;
   		}
          
       if (document.frmEditProspect.txtNextActivity.value == "") {
            swal("Next activity cannot be blank.");
            return false;
       }
       if (document.frmEditProspect.txtNextActivityDueDate.value == "") {
            swal("Next activity due date cannot be blank.");
            return false;
       }
       if (document.frmEditProspect.txtPrimaryCompetitor.value == "") {
            swal("Primary competitor cannot be blank.");
            return false;
       }
       
 		var chkd = document.frmEditProspect.chkBottledWater.checked || +
 			document.frmEditProspect.chkFilteredWater.checked|| +
 			document.frmEditProspect.chkOCS.checked|| +
 			document.frmEditProspect.chkOCS_Supply.checked|| +
 			document.frmEditProspect.chkOfficeSupplies.checked|| +
 			document.frmEditProspect.chkVending.checked|| +
 			document.frmEditProspect.chkMicroMarket.checked|| +
 			document.frmEditProspect.chkPantry.checked;
			
			if (chkd == true)
			{
			}
			else
			{
	            swal("You must select at least one offering for the primary competitor.");
	            return false;
			}     
						 
       
       return true;

    }
// -->
</script>   
 

<style type="text/css">

/*Colored Content Boxes
------------------------------------*/
	.showme{ 
	display: none;
	}
	.showhim:hover .showme{
	display : block;
	}
	.container{
		width: 100%;
	}
	
	.quick-info-block {
	  padding: 3px 20px;
	  text-align: center;
	  margin-bottom: 20px;
	  border-radius: 7px;
	}
	
	.quick-info-block p{
	  color: #fff;
	  font-size:16px;
	}
	.quick-info-block h2 {
	  color: #fff;
	  font-size:20px;
	  margin-bottom:25px;
	}

	.quick-info-block h2.black {
	  color: #000;
	  font-size:20px;
	  margin-bottom:25px;
	}
	
	.quick-info-block h2 a:hover{
	  text-decoration: none;
	}
	
	.quick-info-block-light,
	.quick-info-block-default {
	  background: #fafafa;
	  border: solid 1px #eee; 
	}
	
	.quick-info-block-default:hover {
	  box-shadow: 0 0 8px #eee;
	}
	
	.quick-info-block-light p,
	.quick-info-block-light h2,
	.quick-info-block-default p,
	.quick-info-block-default h2 {
	  color: #555;
	}

	.quick-info-block-u {
	  background: #72c02c;
	}
	.quick-info-block-blue {
	  background: #3498db;
	}
	.quick-info-block-red {
	  background: #e74c3c;
	}
	.quick-info-block-sea {
	  background: #1abc9c;
	}
	.quick-info-block-grey {
	  background: #f8f8f8;
	}
	.quick-info-block-yellow {
	  background: #f1c40f;
	}
	.quick-info-block-orange {
	  background: #e67e22;
	}
	.quick-info-block-green {
	  background: #2ecc71;
	}
	.quick-info-block-purple {
	  background: #9b6bcc;
	}
	.quick-info-block-aqua {
	  background: #27d7e7;
	}
	.quick-info-block-brown {
	  background: #9c8061;
	}
	.quick-info-block-dark-blue {
	  background: #4765a0;
	}
	.quick-info-block-light-green {
	  background: #79d5b3;
	}
	.quick-info-block-dark {
	  background: #555;
	}
	.quick-info-block-light {
	  background: #ecf0f1;
	}
	
	textarea.form-control {
    	height: 100px; !important;
    	width:385px !important;
    	border-radius:3px !important;
	}
		
	hr.tile {
	    border: 0;
	    height: 3px;
	    background-image: linear-gradient(to right, rgba(0, 0, 0, 0), rgba(255, 255, 255, 0.95), rgba(0, 0, 0, 0));
	}
	
	.CRMTabLogColor{
		<% Response.Write("background:" & CRMTabLogColor & " !important;") %>
	}
	.CRMTabProductsColor{
		<% Response.Write("background:" & CRMTabProductsColor & " !important;") %>
	}
	.CRMTabEquipmentColor{
		<% Response.Write("background:" & CRMTabEquipmentColor & " !important;") %>
	}
	.CRMTabDocumentsColor{
		<% Response.Write("background:" & CRMTabDocumentsColor & " !important;") %>
	}
	.CRMTabLocationColor{
		<% Response.Write("background:" & CRMTabLocationColor & " !important;") %>
	}
	.CRMTabContactsColor{
		<% Response.Write("background:" & CRMTabContactsColor & " !important;") %>
	}
	.CRMTabCompetitorsColor{
		<% Response.Write("background:" & CRMTabCompetitorsColor & " !important;") %>
	}
	.CRMTabOpportunityColor{
		<% Response.Write("background:" & CRMTabOpportunityColor & " !important;") %>
	}
	.CRMTabAuditTrailColor{
		<% Response.Write("background:" & CRMTabAuditTrailColor & " !important;") %>
	}
	
	
	.CRMTileOfferingColor {
		<% Response.Write("background:" & CRMTileOfferingColor & " !important;") %>
	}
	.CRMTileCompetitorColor {
		<% Response.Write("background:" & CRMTileCompetitorColor & " !important;") %>
	}
	.CRMTileDollarsColor {
		<% Response.Write("background:" & CRMTileDollarsColor & " !important;") %>
	}
	.CRMTileActivityColor {
		<% Response.Write("background:" & CRMTileActivityColor & " !important;") %>
	}
	.CRMTileStageColor {
		<% Response.Write("background:" & CRMTileStageColor & " !important;") %>
	}
	
	.CRMTileOwnerColor {
		<% Response.Write("background:" & CRMTileOwnerColor & " !important;") %>
	}
	.CRMTileCommentsColor {
		<% Response.Write("background:" & CRMTileCommentsColor & " !important;") %>
	}


	#txtFirstNameLabel,
	#txtLastNameLabel,
	#txtCompanyNameLabel,
	#txtTitleLabel,
	#txtAddressLine1Label,
	#txtAddressLine2Label,
	#txtCityLabel,
	#txtStateLabel,
	#txtZipCodeLabel,
	#txtCountryLabel,
	#txtEmailAddressLabel,
	#txtWebsiteURLLabel,
	#txtPhoneNumberLabel,
	#txtPhoneNumberExtLabel,
	#txtCellPhoneNumberLabel,
	#txtFaxNumberLabel,
	#txtIndustryLabel{
	    display: none;
	    text-align:left;
	    color:#000;
	    font-size:16px;
	    margin-bottom:2px;
	}

	#txtProjectedGPSpendLabel,
	#txtNumEmployeesLabel,
	#txtNumPantriesLabel,
	#txtTelemarketerUserNoLabel,
	#txtPrimaryCompetitorLabel,
	#txtLeadSourceLabel,
	#txtFormerCustomerNumberLabel{
	    display: none;
	    text-align:left;
	   	color:#fff;
	    font-size:16px;
	    margin-bottom:2px;
	}
	
	.red-line{
		border-left:3px solid red;
	}   
	
</style>
<!-- eof css !-->
<h1 class="page-header"><i class="fa fa-fw fa-asterisk"></i> Edit Prospect <%= Company %>
	<!-- customize !-->
	<div class="col pull-right">
	</div>
	<!-- eof customize !-->
</h1>

		
<form autocomplete="off" action="<%= BaseURL %>prospecting/addProspect_submit.asp" method="POST" name="frmEditProspect" id="frmEditProspect" onsubmit="return validateEditProspectForm();" class="form-horizontal track-event-form bv-form">
<input autocomplete="false" name="hidden" type="text" style="display:none;">
<div class="container pull-left">
<div class="row">
      <div class="col-md-3">

		<div class="quick-info-block quick-info-block-grey">
		<h2 class="heading-md black"><i class="icon-2x color-light fa fa-user-circle"></i>&nbsp;Business Card - <%= Company %></h2>

              <div class="form-group">
 
	                <div class="col-sm-6">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-user"></i></div>
                    		<select data-placeholder="Choose Suffix, Mr., Mrs., etc." class="C_Country_Modal form-control" id="txtSuffix" name="txtSuffix">  
                    			<option value="">Salutation, Mr., Mrs., etc.</option>  
                    			<option value="Mr." <% If primarySuffix = "Mr." Then Response.write("selected") %>>Mr.</option>
								<option value="Mrs." <% If primarySuffix = "Mrs." Then Response.write("selected") %>>Mrs.</option>
								<option value="Miss" <% If primarySuffix = "Miss" Then Response.write("selected") %>>Miss</option>
								<option value="Dr." <% If primarySuffix = "Dr." Then Response.write("selected") %>>Dr.</option>
								<option value="Ms." <% If primarySuffix = "Ms." Then Response.write("selected") %>>Ms.</option>                     
							</select>
	                    	
	                   </div>
	                </div> 
              
               </div>

              <div class="form-group">
	                            
	                <div class="col-sm-6">
	                  <p id="txtFirstNameLabel">First Name</p>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtFirstNameIcon"><i class="fa fa-user"></i></div>
	                    	<input type="text" class="form-control" id="txtFirstName" name="txtFirstName" value="<%= primaryFirstName %>">
	                   </div>
	                </div> 
	                
	                <div class="col-sm-6">
	                  <p id="txtLastNameLabel">Last Name</p>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtLastNameIcon"><i class="fa fa-user"></i></div>
	                    	<input type="text" class="form-control" id="txtLastName" name="txtLastName" value="<%= primaryLastName %>">
	                   </div>
	                </div> 
	 
               </div>
               

              <div class="form-group">
	                            	                
	                <div class="col-sm-12">
	                  <p id="txtCompanyNameLabel">Company Name</p>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtCompanyNameIcon"><i class="fa fa-briefcase"></i></div>
	                    	<input type="text" class="form-control red-line" id="txtCompanyName" name="txtCompanyName" value="<%= Company %>">
	                   </div>
	                </div> 
	 
               </div>
               

              <div class="form-group">
	                            
	                <div class="col-sm-12">
	                  <p id="txtTitleLabel">Job Title</p>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtTitleIcon"><i class="fa fa-id-card-o"></i></div>
                    		<select data-placeholder="Choose Job Title" class="C_Country_Modal form-control" id="txtTitle" name="txtTitle">  
                    			<option value="">Select Job Title</option>                          
								<%
								SQLContactTitles = "SELECT *, InternalRecordIdentifier as id FROM PR_ContactTitles ORDER BY ContactTitle"
								Set cnnContactTitles = Server.CreateObject("ADODB.Connection")
								cnnContactTitles.open (Session("ClientCnnString"))
								Set rsContactTitles = Server.CreateObject("ADODB.Recordset")
								rsContactTitles.CursorLocation = 3 
								Set rsContactTitles = cnnContactTitles.Execute(SQLContactTitles)
								If not rsContactTitles.EOF Then
									Do While Not rsContactTitles.EOF
											%><option value="<%= rsContactTitles("id") %>" <% if rsContactTitles("id") = primaryTitleNumber Then Response.write("selected") %>><%= rsContactTitles("ContactTitle") %></option><%
										rsContactTitles.MoveNext						
									Loop
								End If
								Set rsContactTitles = Nothing
								cnnContactTitles.Close
								Set cnnContactTitles = Nothing
								
								%> 
							</select>
   	                   </div>
	                </div> 
	 
               </div>

              <div class="form-group">
	                            
	                <div class="col-sm-12">
	                  <p id="txtAddressLine1Label">Street Address</p>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtAddressLine1Icon"><i class="fa fa-address-card"></i></div>
	                    	<input type="text" class="form-control" id="txtAddressLine1" name="txtAddressLine1" value="<%= Street %>">
	                   </div>
	                </div> 
	           </div>     
	                
	                
	          <div class="form-group">
	                
	                <div class="col-sm-12">
	                  <p id="txtAddressLine2Label">Suite, Floor #, etc.</p>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtAddressLine2Icon"><i class="fa fa-address-card-o"></i></div>
	                    	<input type="text" class="form-control" id="txtAddressLine2" name="txtAddressLine2" value="<%= Suite %>">
	                   </div>
	                </div> 
	 
               </div>



              <div class="form-group">
	                            
	                <div class="col-sm-9">
	                  <p id="txtCityLabel">City</p>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtCityIcon"><i class="fa fa-address-book"></i></div>
	                    	<input type="text" class="form-control" id="txtCity" name="txtCity" value="<%= City %>">
	                   </div>
	                </div> 
	                
	          </div>     
	          <div class="form-group">

	                <div class="col-sm-7">
	                  <p id="txtStateLabel">State</p>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtStateIcon"><i class="fa fa-address-book"></i></div>
                    		<select data-placeholder="Choose State" class="C_Country_Modal form-control" id="txtState" name="txtState"> 
                    			<option value="">State</option>
								<!--#include file="statelist.asp"-->
							</select>				
		
	                   </div>
	                </div> 
	                <div class="col-sm-5">
	                  <p id="txtZipCodeLabel">Zip Code</p>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtZipCodeIcon"><i class="fa fa-address-book"></i></div>
	                    	<input type="text" class="form-control" id="txtZipCode" name="txtZipCode" value="<%= PostalCode %>">
	                   </div>
	                </div> 
	 
               </div>
               
                
              <div class="form-group">
	                            	                
	                <div class="col-sm-6">
	                  <p id="txtCountryLabel">Choose Country</p>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtCountryIcon"><i class="fa fa-globe"></i></div>
                    		<select data-placeholder="Choose Country" class="C_Country_Modal form-control" id="txtCountry" name="txtCountry"> 
								<!--#include file="countrylist.asp"-->
							</select>

	                   </div>
	                </div> 
	                	 
               </div>
              
              <div class="form-group">
	                            
	                <div class="col-sm-12">
	                  <p id="txtEmailAddressLabel">Email Address</p>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtEmailAddressIcon"><i class="fa fa-envelope"></i></div>
	                    	<input type="text" class="form-control" id="txtEmailAddress" name="txtEmailAddress" value="<%= primaryEmail %>">
	                   </div>
	                </div> 
	          </div>    
	          <div class="form-group">

	                <div class="col-sm-12">
	                  <p id="txtWebsiteURLLabel">Company Website URL</p>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtWebsiteURLIcon"><i class="fa fa-globe"></i></div>
	                    	<input type="text" class="form-control" id="txtWebsiteURL" name="txtWebsiteURL" value="<%= Website %>">
	                   </div>
	                </div> 
	                
               </div>
               

              <div class="form-group">

	                <div class="col-sm-6">
	                  <p id="txtPhoneNumberLabel">Phone Number</p>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtPhoneNumberIcon"><i class="fa fa-phone"></i></div>
	                    	<input type="text" class="form-control" id="txtPhoneNumber" name="txtPhoneNumber" value="<%= primaryPhone %>">
	                   </div>
	                </div> 
	          </div>
	          
	          <div class="form-group">
	                <div class="col-sm-6">
	                  <p id="txtPhoneNumberExtLabel">Extension</p>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtPhoneNumberExtIcon"><i class="fa fa-phone"></i></div>
	                    	<input type="text" class="form-control" id="txtPhoneNumberExtLabel" name="txtPhoneNumberExtLabel" value="<%= primaryPhoneExt %>">
	                   </div>
	                </div> 
	 
               </div>

                      
               <div class="form-group">
               
	                <div class="col-sm-6">
	                  <p id="txtCellPhoneNumberLabel">Cell Phone Number</p>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtCellPhoneNumberIcon"><i class="fa fa-mobile"></i></div>
	                    	<input type="text" class="form-control" id="txtCellPhoneNumber" name="txtCellPhoneNumber" value="<%= primaryCell %>">
	                   </div>
	                </div> 
	                	 
	                <div class="col-sm-6">
	                  <p id="txtFaxNumberLabel">Fax Number</p>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtFaxNumberIcon"><i class="fa fa-fax"></i></div>
	                    	<input type="text" class="form-control" id="txtFaxNumber" name="txtFaxNumber" value="<%= primaryFax %>">
	                   </div>
	                </div> 
	                	 
               </div>
               <div class="form-group">
	                            
	                <div class="col-sm-12">
	                  <p id="txtIndustryLabel">Select Industry</p>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtIndustryIcon"><i class="fa fa-building"></i></div>
                    		<select data-placeholder="Choose Industry" class="C_Country_Modal form-control" id="txtIndustry" name="txtIndustry"> 
                    		<option value="">Select Industry</option>
			  	  			<%
			  	  			'Get all industries
					      	  	SQL9 = "SELECT * FROM PR_Industries ORDER BY Industry "
			
								Set cnn9 = Server.CreateObject("ADODB.Connection")
								cnn9.open (Session("ClientCnnString"))
								Set rs9 = Server.CreateObject("ADODB.Recordset")
								rs9.CursorLocation = 3 
								Set rs9 = cnn9.Execute(SQL9)
									
								If not rs9.EOF Then
									Do
										IndustryNumber = rs9("InternalRecordIdentifier")

										If IndustryNumber = 0 Then
											%><option value="<%= rs9("InternalRecordIdentifier") %>" <% if IndustryNumber = IndustryNumber Then Response.write("selected") %>>-- Not Specified --</option><%
										Else
											%><option value="<%= rs9("InternalRecordIdentifier") %>" <% if IndustryNumber = IndustryNumber Then Response.write("selected") %>><%= rs9("Industry") %></option><%
										End If
										rs9.movenext
									Loop until rs9.eof
								End If
								set rs9 = Nothing
								cnn9.close
								set cnn9 = Nothing
							%>
						</select>


	                   </div>
	                </div> 
	                	 
               </div>
              
                    
			</div>
        <!-- END QUICK INFO BOX -->
        
        
      </div><!-- end col-md-6 -->



      <div class="col-md-3">

		<div class="quick-info-block CRMTileCommentsColor">
		<h2 class="heading-md"><i class="icon-2x color-light fa fa-comment"></i>&nbsp;<%= GetTerm("Comments") %></h2>
	          <div class="form-group">        	                
	                <div class="col-sm-12">
	                  <div class="input-group">
							<textarea class="form-control" id="txtComments" name="txtComments"><%= Comments %></textarea>
	                   </div>
	                </div> 	 
	           </div>
		</div>
	  

		<div class="quick-info-block CRMTileDollarsColor">
		<h2 class="heading-md"><i class="icon-2x color-light fa fa-usd"></i>&nbsp;<%= GetTerm("Opportunity") %></h2>

              <div class="form-group">
	                            	                
	                <div class="col-sm-12">
	                	<p id="txtProjectedGPSpendLabel">Projected GP Spend (numbers only)</p>
	                  	<div class="input-group">
	                    	<div class="input-group-addon" id="txtProjectedGPSpendIcon"><i class="fa fa-credit-card-alt"></i></div>
	                    	<input type="text" class="form-control showhim" id="txtProjectedGPSpend" name="txtProjectedGPSpend" value="<%= FormatCurrency(ProjectedGPSpend,2) %>">
	                   </div>
	                </div> 
	                	 
               </div>

              <div class="form-group">
	                            	                
	                <div class="col-sm-12">
	                  <p id="txtNumEmployeesLabel">Select # Employees</p>	                
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtNumEmployeesIcon"><i class="fa fa-users"></i></div>
                    		<select data-placeholder="Select # Employees" class="C_Country_Modal form-control" id="txtNumEmployees" name="txtNumEmployees"> 
                    			<option value="">Select # Employees</option>
				  	  			<%
				  	  			'Get employee ranges
									SQL9 = "SELECT *, Cast(LEFT(Range,CHARINDEX('-',Range)-1) as int) as Expr1 FROM PR_EmployeeRangeTable "
									SQL9 = SQL9 & "order by Expr1"

									Set cnn9 = Server.CreateObject("ADODB.Connection")
									cnn9.open (Session("ClientCnnString"))
									Set rs9 = Server.CreateObject("ADODB.Recordset")
									rs9.CursorLocation = 3 
									Set rs9 = cnn9.Execute(SQL9)
										
									If not rs9.EOF Then
										Do
											%><option value="<%= rs9("InternalRecordIdentifier") %>" <% If rs9("InternalRecordIdentifier") = EmployeeRangeNumber Then Response.write("selected") %>><%= rs9("Range") %></option><%
											rs9.movenext
										Loop until rs9.eof
									End If
									set rs9 = Nothing
									cnn9.close
									set cnn9 = Nothing
								%>
							</select>
				  	  			<%
				  	  			'Get GP Spend From Employee Range
									SQL9 = "SELECT *, Cast(LEFT(Range,CHARINDEX('-',Range)-1) as int) as Expr1 FROM PR_EmployeeRangeTable "
									SQL9 = SQL9 & "order by Expr1"

									Set cnn9 = Server.CreateObject("ADODB.Connection")
									cnn9.open (Session("ClientCnnString"))
									Set rs9 = Server.CreateObject("ADODB.Recordset")
									rs9.CursorLocation = 3 
									Set rs9 = cnn9.Execute(SQL9)
										
									If not rs9.EOF Then
										Do
											%><input type="hidden" value="<%= rs9("ProjectedGPSpend") %>" id="<%= rs9("InternalRecordIdentifier") %>"><%
											rs9.movenext
										Loop until rs9.eof
									End If
									set rs9 = Nothing
									cnn9.close
									set cnn9 = Nothing
								%>
							

	                   </div>
	                </div> 
	                	 
               </div>

              <div class="form-group">
	                            	                
	                <div class="col-sm-12">
	                  <p id="txtNumPantriesLabel">Select # Pantries</p>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtNumPantriesIcon"><i class="fa fa-apple"></i></div>
                    		<select data-placeholder="Select # Pantries" class="C_Country_Modal form-control" id="txtNumPantries" name="txtNumPantries" value="<%= NumberOfPantries %>"> 
                    			<option value="">Select # Pantries</option>
								<% For i = 0 To 50 %>
								  <option value="<%= i %>"><%= i %></option>
								<% Next %>							
							</select>

	                   </div>
	                </div> 
	                	 
               </div>
               
               <div class="form-group">
               		<div class="col-sm-6"><p>Bldg. Lease Expires Date</p></div>
					<div class="col-sm-6">
				        <div class="input-group date" id="datetimepickerLeaseExpiresDate">
				            <input type="text" class="form-control" id="txtLeaseExpirationDate" name="txtLeaseExpirationDate" value="<%= LeaseExpirationDate %>" readonly="readonly" />
				            <span class="input-group-addon">
				                <span class="glyphicon glyphicon-calendar"></span>
				            </span>
				        </div>
				    </div>
				</div>   
				<script type="text/javascript">
		            $(function () {
		            	
		                $('#datetimepickerLeaseExpiresDate').datetimepicker({
		                   //minDate: moment(),
		                   format: 'MM/DD/YYYY',
		                   ignoreReadonly: true
		                
		                });  
		            });
		        </script>    
                
                
               <div class="form-group">
               		<div class="col-sm-6"><p>Contract Expiration Date</p></div>
					<div class="col-sm-6">
				        <div class="input-group date" id="datetimepickerContractExpireDate">
				            <input type="text" class="form-control" id="txtContractExpirationDate" name="txtContractExpirationDate" value="<%= ContractExpirationDate %>" readonly="readonly" />
				            <span class="input-group-addon">
				                <span class="glyphicon glyphicon-calendar"></span>
				            </span>
				        </div>
				    </div>
				</div>   
				<script type="text/javascript">
		            $(function () {
		            	
		                $('#datetimepickerContractExpireDate').datetimepicker({
		                   //minDate: moment(),
		                   format: 'MM/DD/YYYY',
		                   ignoreReadonly: true
		                
		                });  
		            });
		        </script>                     
			</div>
        <!-- END QUICK INFO BOX -->
        
 		<div class="quick-info-block CRMTileOwnerColor">
		<h2 class="heading-md"><i class="icon-2x color-light fa fa-user"></i>&nbsp;<%= GetTerm("Owner") %></h2>

              <div class="form-group">
	                            	                
	                <div class="col-sm-12">
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-user"></i></div>
                    		<select data-placeholder="Choose Owner" class="C_Country_Modal form-control" id="txtOwner" name="txtOwner"> 
							<option value="<%=Session("UserNo")%>"><%=GetUserFirstAndLastNameByUserNo(Session("UserNo"))%></option>
					      	<%'Owner dropdown
					      	 
				      	 
				      	  	SQL = "SELECT UserNo, userFirstName, userLastName, userType FROM " & MUV_Read("SQL_Owner") & ".tblUsers "
				      	  	SQL = SQL & "WHERE userArchived <> 1 and UserNo <> " & Session("UserNo") & " AND userEnabled = 1 AND "
							SQL = SQL & "(userType='Outside Sales' OR userType='Outside Sales Manager' OR userType='Admin' "
							SQL = SQL & "OR userType='Inside Sales' OR userType='Inside Sales Manager' OR userType='CSR' "
							SQL = SQL & "OR userType='CSR Manager') "
				      	  	SQL = SQL & "ORDER BY userFirstName, userLastName"
			
							Set cnn8 = Server.CreateObject("ADODB.Connection")
							cnn8.open (Session("ClientCnnString"))
							Set rs = Server.CreateObject("ADODB.Recordset")
							rs.CursorLocation = 3 
							Set rs = cnn8.Execute(SQL)
						
							If not rs.EOF Then
								Do
									FullName = rs("userFirstName") & " " & rs("userLastName")
									If rs("UserNo") = OwnerUserNo Then
										Response.Write("<option value='" & rs("UserNo") & "' selected>" & FullName & "</option>")
									Else
										Response.Write("<option value='" & rs("UserNo") & "'>" & FullName & "</option>")
									End If
									rs.movenext
								Loop until rs.eof
							End If
							set rs = Nothing
							cnn8.close
							set cnn8 = Nothing
		      				%>
						</select>

	                   </div>
	                </div> 
	                	 
               </div>

   			</div>

   			
        <!-- END QUICK INFO BOX -->
      </div><!-- end col-md-6 -->
	
	<div class="col-md-3">
        
   			
    			
		<div class="quick-info-block CRMTileOfferingColor">
		<h2 class="heading-md"><i class="icon-2x color-light fa fa-clock-o"></i>&nbsp;<%= GetTerm("Current Supplier Info") %></h2>
              <div class="form-group">        	                
	                <div class="col-sm-12">
	                  <div class="input-group">
							<textarea class="form-control" id="txtCurrentOffering" name="txtCurrentOffering"><%= CurrentOffering %></textarea>
	                   </div>
	                </div> 	 
               </div>
   		</div>
   			
  
       

            <div class="quick-info-block CRMTileCompetitorColor">
            <h2 class="heading-md"><i class="icon-2x color-light fa fa-user-circle-o"></i>&nbsp;<%= GetTerm("Primary Competitor") %></h2>
            
              <div class="form-group">
	                            	                
	                <div class="col-sm-12">
	                  <p id="txtTelemarketerUserNoLabel">Choose Telemarketer</p>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtTelemarketerUserNoIcon"><i class="fa fa-user"></i></div>
                    		<select data-placeholder="Choose Telemarketer" class="C_Country_Modal form-control" id="txtTelemarketerUserNo" name="txtTelemarketerUserNo"> 
							<option value="">Select a Telemarketer</option>
					      	<%'Telemarketer dropdown

				      	  	SQL = "SELECT UserNo, userFirstName, userLastName, userType FROM " & MUV_Read("SQL_Owner") & ".tblUsers "
				      	  	SQL = SQL & "WHERE userArchived <> 1 and UserNo <> " & Session("UserNo") & " AND userEnabled = 1"
  					      	SQL = SQL & " AND userType = 'Telemarketing' "
				      	  	SQL = SQL & "ORDER BY userFirstName, userLastName"
			
							Set cnn8 = Server.CreateObject("ADODB.Connection")
							cnn8.open (Session("ClientCnnString"))
							Set rs = Server.CreateObject("ADODB.Recordset")
							rs.CursorLocation = 3 
							Set rs = cnn8.Execute(SQL)
						
							If not rs.EOF Then
								Do
									FullName = rs("userFirstName") & " " & rs("userLastName")
									If rs("UserNo") = TelemarketerUserNo Then
										Response.Write("<option value='" & rs("UserNo") & "' selected>" & FullName & "</option>")
									Else
										Response.Write("<option value='" & rs("UserNo") & "'>" & FullName & "</option>")
									End If
									rs.movenext
								Loop until rs.eof
							End If
							set rs = Nothing
							cnn8.close
							set cnn8 = Nothing
		      				%>
						</select>

	                   </div>
	                </div> 
	                	 
               </div>
            

              <div class="form-group">
	                            	                
	                <div class="col-sm-12">
	                  <p id="txtLeadSourceLabel">Select Lead Source</p>
	                  <div class="input-group">
                    	<div class="input-group-addon" id="txtLeadSourceIcon"><i class="fa fa-external-link"></i></div>
                		<select data-placeholder="Choose Lead Source" class="C_Country_Modal form-control" id="txtLeadSource" name="txtLeadSource"> 
                    		<option value="">Select Lead Source</option>
				  	  			<%
									SQL9 = "SELECT * FROM PR_LeadSources ORDER BY LeadSource"

									Set cnn9 = Server.CreateObject("ADODB.Connection")
									cnn9.open (Session("ClientCnnString"))
									Set rs9 = Server.CreateObject("ADODB.Recordset")
									rs9.CursorLocation = 3 
									Set rs9 = cnn9.Execute(SQL9)
										
									If not rs9.EOF Then
										Do
											%><option value="<%= rs9("InternalRecordIdentifier") %>" <% If rs9("InternalRecordIdentifier") = LeadSourceNumber Then Response.Write("selected") %>><%= rs9("LeadSource") %></option><%
											rs9.movenext
										Loop until rs9.eof
									End If
									set rs9 = Nothing
									cnn9.close
									set cnn9 = Nothing
								%>
						</select>

	                   </div>
	                </div> 
	                	 
               </div>


              <div class="form-group">
	                            	                
	                <div class="col-sm-12">
	                  <p id="txtPrimaryCompetitorLabel">Select Primary Competitor</p>
	                  <div class="input-group">
                    	<div class="input-group-addon" id="txtPrimaryCompetitorIcon"><i class="fa fa-coffee"></i></div>
                		<select data-placeholder="Choose Primary Competitor" class="C_Country_Modal form-control red-line" id="txtPrimaryCompetitor" name="txtPrimaryCompetitor"> 
                    		<option value="">Select Primary Competitor</option>
							<%
							SQLCompetitorNames = "SELECT * FROM PR_Competitors ORDER BY CompetitorName"
							Set cnnCompetitorNames = Server.CreateObject("ADODB.Connection")
							cnnCompetitorNames.open (Session("ClientCnnString"))
							Set rsCompetitorNames = Server.CreateObject("ADODB.Recordset")
							rsCompetitorNames.CursorLocation = 3 
							Set rsCompetitorNames = cnnCompetitorNames.Execute(SQLCompetitorNames)
							
							If not rsCompetitorNames.EOF Then
								sep = ""
								Do While Not rsCompetitorNames.EOF
										%><option value="<%= rsCompetitorNames("InternalRecordIdentifier") %>" <% If rsCompetitorNames("InternalRecordIdentifier") = PrimaryCompetitorID Then Response.Write("selected") %>><%= rsCompetitorNames("CompetitorName") %></option><%
									rsCompetitorNames.MoveNext						
								Loop
							End If
							Set rsCompetitorNames = Nothing
							cnnCompetitorNames.Close
							Set cnnCompetitorNames = Nothing
							%> 
						</select>

	                   </div>
	                </div> 
	                	                
	                	 
               </div>

			   <style type="text/css">
					
					fieldset.group  { 
					  margin: 0; 
					  padding: 0; 
					  margin-bottom: 1.25em; 
					  padding: .125em; 
					} 
					
					fieldset.group legend { 
					  margin: 0; 
					  padding: 0; 
					  font-weight: bold; 
					  margin-left: 20px; 
					  font-size: 100%; 
					  color: black; 
					} 
					
					
					ul.checkbox  { 
					  margin: 0; 
					  padding: 0; 
					  margin-left: 20px !important; 
					  list-style: none; 
					  text-align: left !important;
					} 
					
					ul.checkbox li input { 
					  margin-right: .25em; 
					} 
					
					ul.checkbox li { 
					  border: 1px transparent solid; 
					  display:inline-block;
					  width:12em;
					} 
					
					ul.checkbox li label { 
					  margin-left: ; 
					} 
					
					.checkbox label{
						/*color:#fff !important;*/
					}
					
					.checkbox label, .radio label {
					    min-height: 20px;
					    padding-left: 20px;
					    margin-bottom: 0;
					    font-weight: 400;
					    cursor: pointer;
					    color: #fff;
					}					
			  </style>


              <div class="form-group">
              
              		<h2 class="heading-md" style="margin-bottom:5px;"><i class="icon-2x color-light fa fa-industry"></i>&nbsp;Offerings</h2>
					
					<div class="col-sm-12">
					
						<fieldset class="group"> 
							<ul class="checkbox"> 
							  <li><input type="checkbox" id="chkBottledWater" name="chkBottledWater" <% If BottledWater = "Bottled Water" Then Response.Write("checked") %>><label for="chkBottledWater">Bottled Water</label></li> 
							  <li><input type="checkbox" id="chkFilteredWater" name="chkFilteredWater" <% If FilteredWater = "Filtered Water" Then Response.Write("checked") %>><label for="chkFilteredWater">Filtered Water</label></li> 
							  <li><input type="checkbox" id="chkOCS" name="chkOCS" <% If OCS = "OCS" Then Response.Write("checked") %>><label for="chkOCS">OCS</label></li> 
							  <li><input type="checkbox" id="chkOCS_Supply" name="chkOCS_Supply" <% If OCS_Supply = "OCS Supply" Then Response.Write("checked") %>><label for="chkOCS_Supply">OCS Supply</label></li> 
							  <li><input type="checkbox" id="chkOfficeSupplies" name="chkOfficeSupplies" <% If OfficeSupplies = "Office Supplies" Then Response.Write("checked") %>><label for="chkOfficeSupplies">Office Supplies</label></li> 
							  <li><input type="checkbox" id="chkVending" name="chkVending" <% If Vending = "Vending" Then Response.Write("checked") %>><label for="chkVending">Vending</label></li> 
							  <li><input type="checkbox" id="chkMicroMarket" name="chkMicroMarket" <% If Micromarkets = "Micromarkets" Then Response.Write("checked") %>><label for="chkMicroMarket">Micromarket</label></li>
							  <li><input type="checkbox" id="chkPantry" name="chkPantry" <% If Pantry = "Pantry"  Then Response.Write("checked") %>><label for="chkPantry">Pantry</label></li>
							</ul> 
						</fieldset> 
					</div>	 
					
               </div>

               
 				<div class="form-group">
	                            	                
	                <div class="col-sm-12">
	                  <p id="txtFormerCustomerNumberLabel">Former Customer #</p>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtFormerCustomerNumberIcon"><i class="fa fa-id-card-o"></i></div>
	                    	<input type="text" class="form-control" id="txtFormerCustomerNumber" name="txtFormerCustomerNumber" value="<%= FormerCustNum %>">
	                   </div>
	                </div> 
   	 
               </div>              
               
               <div class="form-group">
               		<div class="col-sm-6"><p>Former Customer Cancel Date</p></div>
					<div class="col-sm-6">
				        <div class="input-group date" id="datetimepickerCancelDate">
				            <input type="text" class="form-control" id="txtFormerCustomerCancelDate" name="txtFormerCustomerCancelDate" value="<%= CancelDate %>" readonly="readonly" />
				            <span class="input-group-addon">
				                <span class="glyphicon glyphicon-calendar"></span>
				            </span>
				        </div>
				    </div>
				</div>   
 
               

   			</div>
   			
   			
   		</div>
   		
   	<div class="col-md-3">
   						
		<div class="quick-info-block CRMTileActivityColor">
			<h2 class="heading-md"><i class="icon-2x color-light fa fa-arrow-right"></i>&nbsp;<%= GetTerm("Next Activity") %></h2>
	
		    <div class="form-group">          	                
	            <div class="col-sm-12">
	            <% If nextActivity <> "" Then %>
		              <p><strong>Next Activity</strong>: <%= nextActivity %></p>
		              <p><strong>Due Date</strong>: <%= nextActivityDueDateTime %></p>
	             <% Else %>
		              <p><strong>Next Activity</strong>: No Next Activity</p>
		              <p><strong>Due Date</strong>: NA</p>	             
	             <% End If %>
	        	</div>  
		     </div>
		</div>
   			
   			
   	
		
        <div class="quick-info-block CRMTileStageColor">
	        <h2 class="heading-md"><i class="icon-2x color-light fa fa-tasks"></i>&nbsp;<%= GetTerm("Stage") %></h2>
	
			<div class="form-group">
				<div class="col-sm-12" style="width:400px; margin-left:0px; margin-right:0px; text-align: center;">
					<p><strong>Current Stage</strong>: <%= GetStageByNum(GetProspectCurrentStageByProspectNumber(InternalRecordIdentifier)) %></p>
					<%
					 	SQLCurrentStageInfo = "SELECT TOP 1 * FROM PR_ProspectStages Where ProspectRecID = " & InternalRecordIdentifier & " AND " & " StageRecID = " & GetProspectCurrentStageByProspectNumber(InternalRecordIdentifier) & " ORDER BY RecordCreationDateTime DESC"
					
						Set cnnCurrentStageInfo = Server.CreateObject("ADODB.Connection")
						cnnCurrentStageInfo.open (Session("ClientCnnString"))
						Set rsCurrentStageInfo = Server.CreateObject("ADODB.Recordset")
						rsCurrentStageInfo.CursorLocation = 3 
						Set rsCurrentStageInfo = cnnCurrentStageInfo.Execute(SQLCurrentStageInfo)
						If not rsCurrentStageInfo.EOF Then
							ProspectCurrentStageNotes = rsCurrentStageInfo("Notes")
						End If
						set rsCurrentStageInfo = Nothing
						cnnCurrentStageInfo.close
						set cnnCurrentStageInfo = Nothing
	

					%>
					<p><strong>Current Stage Notes</strong>: <%= ProspectCurrentStageNotes %></p>								
				</div>
	        </div> 
	        	 
       </div>
       
       
				
               
      </div><!-- end col-md-3 -->
      

 </div> <!-- end row -->
        
        
<div class="form-group pull-right">
	<div class="col-sm-12">
		<button class="btn btn-primary btn-lg btn-block" href="main.asp" role="button" type="submit">SAVE THIS PROSPECT <i class="fa fa-floppy-o" aria-hidden="true"></i></button>
	</div>
</div>
        
</div> <!-- end container -->

</form>
 
 <!-- tabs js !-->
 <script type="text/javascript">
 $('a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
  e.target // newly activated tab
  e.relatedTarget // previous active tab
})
 </script>
 <!-- eof tabs js !-->

<!--#include file="../inc/footer-main.asp"-->
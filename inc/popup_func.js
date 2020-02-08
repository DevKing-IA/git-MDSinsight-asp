function initAddPriceGroupPopUp()
{
	$('#addPriceGroupWindow').dialog({ bgiframe: true,modal: true, autoOpen: false, title: 'Add Product Price Group', width:560, 
	buttons: {
		Save: function() {
			if (checkprodGroupForm()){
				var prodGroupDescription = document.getElementById("prodGroupDescription").value;
				$.ajax({
					type:"POST",
					url: "../inc/AjaxFuncs.asp",
					dataType: "application/x-www-form-urlencoded",
					data: "action=addPriceGroup&prodGroupDescription="+encodeURIComponent(prodGroupDescription),
					async: false,
					success: function(msg){document.getElementById("selPriceGroup").innerHTML = msg;}
				})
				$(this).dialog('close');
			}
		},
		Cancel: function() {
			$(this).dialog('close');
		}
	}
	});
}

function showAddPriceGroupPopUp()
{
	document.getElementById("prodGroupDescription").value = "";
	$('#addPriceGroupWindow').dialog('open');
}

function initAddManufacturerPopUp()
{
	$('#addManufacturerWindow').dialog({ bgiframe: true,modal: true, autoOpen: false, title: 'Add Manufacturer', width:540, 
	buttons: {
		Save: function() {
			if (checkManufacturerForm()){
				var manufacturerName = document.getElementById("manufacturerName").value;
				$.ajax({
					type:"POST",
					url: "AjaxFuncs.asp",
					dataType: "application/x-www-form-urlencoded",
					data: "action=addManufacturer&manufacturerName="+encodeURIComponent(manufacturerName),
					async: false,
					success: function(msg){document.getElementById("selManufacturer").innerHTML = msg;}
				})
				$(this).dialog('close');
			}
		},
		Cancel: function() {
			$(this).dialog('close');
		}
	},
	close: function(){
		$(this).dialog("destroy"); 
	}
	});
}

function showAddManufacturerPopUp()
{
	initAddManufacturerPopUp();
	document.getElementById("manufacturerName").value = "";
	$('#addManufacturerWindow').dialog('open');
}

function initAddTermPopUp()
{
	$('#addTermWindow').dialog({ bgiframe: true,modal: true, autoOpen: false, title: 'Add Term', width:540, top:90, 
	buttons: {
		Save: function() {
			if (checkTermForm()){
				var termsDescription = document.getElementById("termsDescription").value;
				var termsDays = document.getElementById("termsDays").value;
				
				$.ajax({
					type:"POST",
					url: "../inc/AjaxFuncs.asp",
					dataType: "application/x-www-form-urlencoded",
					data: "action=addTerm&termsDescription="+encodeURIComponent(termsDescription)+"&termsDays="+termsDays,
					async: false,
					success: function(msg){document.getElementById("selTerms").innerHTML = msg;}
				})
				$(this).dialog('close');
			}
		},
		Cancel: function() {
			$(this).dialog('close');
		}
	}
	});
}

function showAddTermPopUp()
{
	document.getElementById("termsDescription").value = "";
	document.getElementById("termsDays").value = "";
	$('#addTermWindow').dialog('open');
}

function initAddTaxCodePopUp()
{
	$('#addTaxCodeWindow').dialog({ bgiframe: true,modal: true, autoOpen: false, title: 'Add Tax Code', width:560, 
	buttons: {
		Save: function() {
			if (checkTaxCodeForm()){
				var taxDescription = document.getElementById("taxDescription").value;
				var taxPercent = document.getElementById("taxPercent").value;
				var taxFreight;
				if (document.getElementById("taxFreight").checked == true) 
				{
					taxFreight = 1;
				}
				else
				{
					taxFreight = 0;
				}
				$.ajax({
					type:"POST",
					url: "../inc/AjaxFuncs.asp",
					dataType: "application/x-www-form-urlencoded",
					data: "action=addTaxCode&taxDescription="+encodeURIComponent(taxDescription)+"&taxPercent="+taxPercent+"&taxFreight="+taxFreight,
					async: false,
					success: function(msg){
						document.getElementById("selTaxCodes").innerHTML = msg;
						document.getElementById("lstTaxCode2").add(new Option(taxDescription,taxDescription), document.getElementById("lstTaxCode1").selectedIndex);
						document.getElementById("lstTaxCode3").add(new Option(taxDescription,taxDescription), document.getElementById("lstTaxCode1").selectedIndex);
					}
				})
				$(this).dialog('close');
			}
		},
		Cancel: function() {
			$(this).dialog('close');
		}
	}
	});
}

function showAddTaxCodePopUp()
{
	document.getElementById("taxDescription").value = "";
	document.getElementById("taxPercent").value = "";
	document.getElementById("taxFreight").checked = true;
	$('#addTaxCodeWindow').dialog('open');
}

function checkAddUser() {
	var email = document.getElementById("txtEmail").value;
	var password = document.getElementById("txtPassword1").value;
	var custID = document.getElementById("CustID").value;
	var userNo = document.getElementById("txtUserNo").value;
	initCheckAddUserPopUp();	
	$.ajax({
		type:"POST",
		url:"../inc/AjaxFuncs.asp",
		dataType:"json",
		data:"action=checkAddUser&email="+email + "&password="+password+"&custID="+custID+"&userNo="+userNo,
		async:false,
		success:checkAddUserHandler})
}

function checkAddUserHandler(response) {
	switch(response.status) {
		case "ALLOW":
			if (checkform(document.frmAddUser)) {
			btnSaveUserAdd();
			document.frmAddUser.submit();
			}
			break;
		case "CONFIRM":
			if (checkform(document.frmAddUser)) {
				btnSaveUserAdd();
				document.getElementById("dupCustNames").innerHTML = response.custNames;
				$('#checkAddUserPopUp').dialog('open');
			}
			break;
		case "DENY":
			alert("There is already a user with that email address and password!");
			break;
	}
}


function initCheckAddUserPopUp()
{
	$('#checkAddUserPopUp').dialog({ bgiframe: true,modal: true, autoOpen: false, title: 'Confirmation page', width:560, top:50,
	buttons: {
		Yes: function() {
			document.frmAddUser.submit();
			$(this).dialog('close');
		},
		No: function() {
			$(this).dialog('close');
		}
	}
	});
}





function checkAddRouteselUser() {
	var email = document.getElementById("txtEmail").value;
	var password = document.getElementById("txtPassword1").value;
	initCheckAddRouteselUserPopUp();	
	$.ajax({
		type:"POST",
		url:"../inc/AjaxFuncs.asp",
		dataType:"json",
		data:"action=checkAddRouteselUser&email="+email + "&password="+password,
		async:false,
		success:checkAddRouteselUserHandler})
}

function checkAddRouteselUserHandler(response) {
	switch(response.status) {
		case "ALLOW":
			btnSaveUserAdd();
			document.frmAddRouteselUser.submit();
			break;
		case "DENY":
			alert("There is already a routesel user with that email address and password!");
			break;
	}
}


function initCheckAddRouteselUserPopUp()
{
	$('#checkAddRouteselUserPopUp').dialog({ bgiframe: true,modal: true, autoOpen: false, title: 'Confirmation page', width:560, top:50,
	buttons: {
		Yes: function() {
			document.frmAddRouteselUser.submit();
			$(this).dialog('close');
		},
		No: function() {
			$(this).dialog('close');
		}
	}
	});
}


function initFTMCCEmailsPopUp(userNo)
{
	$.ajax({
        type:"POST",
        url: "../inc/AjaxFuncs.asp",
        dataType: "application/x-www-form-urlencoded",
        data: "action=getFTMCCEmailsToModify&userNo="+encodeURIComponent(userNo),
        async: false,
        success: function(msg){document.getElementById("ftmCCEmailsWindow").innerHTML = msg;}
    })
	$('#ftmCCEmailsWindow').dialog({ bgiframe: true,modal: true, autoOpen: false, title: 'Modify CC Emails', width:560, 
	buttons: {
		Save: function() {
				var ftmCCEmail1 = document.getElementById("txtFTMCCEmail1").value;
				var ftmCCEmail2 = document.getElementById("txtFTMCCEmail2").value;
				var ftmCCEmail3 = document.getElementById("txtFTMCCEmail3").value;
				var ftmCCEmail4 = document.getElementById("txtFTMCCEmail4").value;
				var ftmCCEmail5 = document.getElementById("txtFTMCCEmail5").value;
				
				$.ajax({
					type:"POST",
					url: "../inc/AjaxFuncs.asp",
					dataType: "application/x-www-form-urlencoded",
					data: "action=updateFTMCCEmails&userNo="+encodeURIComponent(userNo)+"&ftmCCEmail1="+encodeURIComponent(ftmCCEmail1)+"&ftmCCEmail2="+encodeURIComponent(ftmCCEmail2)+"&ftmCCEmail3="+encodeURIComponent(ftmCCEmail3)+"&ftmCCEmail4="+encodeURIComponent(ftmCCEmail4)+"&ftmCCEmail5="+encodeURIComponent(ftmCCEmail5),
					async: false,
					success: function(msg){
						document.getElementById("reload").value = ","; //updated values will be displayed on mouse over
					}
				})
				
				$(this).dialog('close');
		},
		Cancel: function() {
			$(this).dialog('close');
		}		
	},
	close: function(){
		$(this).dialog("destroy"); 
	}
	});
}

function showFTMCCEmailsPopUp()
{
	$('#ftmCCEmailsWindow').dialog('open');
}

function initEditQuotedPricingPopUp(prodSKU,custID)
{
	$.ajax({
        type:"POST",
        url: "../inc/AjaxFuncs.asp",
        dataType: "application/x-www-form-urlencoded",
        data: "action=showEditQuotedPricingPopup&prodSKU="+encodeURIComponent(prodSKU)+"&custID="+encodeURIComponent(custID),
        async: false,
        success: function(msg){document.getElementById("editQuotedPricingWindow").innerHTML = msg;}
    })
	$('#editQuotedPricingWindow').dialog({ bgiframe: true,modal: true, autoOpen: false, title: 'Edit Quoted Pricing', width:560, 
	buttons: {
		Save: function() {
				var umCase = "";
				var umUnit = "";
				if (document.getElementById("umCase")){
					umCase=document.getElementById("umCase").value;
				}
				if (document.getElementById("umUnit")){
					umUnit=document.getElementById("umUnit").value;
				}
				if ((umCase=="")&&(umUnit=="")){
					var result = confirm('Leaving field(s) empty will result in deletion of quote pricing for this product. Do you want to continue?');
					if (result == true) {
						$.ajax({
							type:"POST",
							url: "../inc/AjaxFuncs.asp",
							dataType: "application/x-www-form-urlencoded",
							data: "action=updateQuotedPricing&prodSKU="+encodeURIComponent(prodSKU)+"&custID="+encodeURIComponent(custID)+"&umCase="+encodeURIComponent(umCase)+"&umUnit="+encodeURIComponent(umUnit),
							async: false,
							success: function(msg){
								document.frmQuotedPricing.submit();
							}
						})
						$(this).dialog('close');
						
					}
				}else{
					var validInput = true;
					if ((umUnit!="")&&(!isNumber(umUnit))){
						alert("Please enter valid quoted pricing value");
						document.getElementById("umUnit").focus();
						validInput = false;
					}
					if (validInput){
						if ((umCase!="")&&(!isNumber(umCase))){
							alert("Please enter valid quoted pricing value");
							document.getElementById("umCase").focus();
							validInput = false;
						}
					}
					if (validInput){				
						$.ajax({
							type:"POST",
							url: "../inc/AjaxFuncs.asp",
							dataType: "application/x-www-form-urlencoded",
							data: "action=updateQuotedPricing&prodSKU="+encodeURIComponent(prodSKU)+"&custID="+encodeURIComponent(custID)+"&umCase="+encodeURIComponent(umCase)+"&umUnit="+encodeURIComponent(umUnit),
							async: false,
							success: function(msg){
								document.frmQuotedPricing.submit();
							}
						})
						$(this).dialog('close');	
					}			
				}		
				
		},
		Cancel: function() {
			$(this).dialog('close');
		}		
	},
	close: function(){
		$(this).dialog("destroy"); 
	}
	});
}

function showEditQuotedPricingPopUp()
{
	$('#editQuotedPricingWindow').dialog('open');
}

function initAddRoutePopUp()
{
	$('#addRouteWindow').dialog({ bgiframe: true,modal: true, autoOpen: false, title: 'Add Route', width:560, 
	buttons: {
		Save: function() {
			if (checkRouteForm()){
				var routeDescription = document.getElementById("routeDescription").value;
				var routeDriver = document.getElementById("routeDriver").value;
				var routeVehicle = document.getElementById("routeVehicle").value;
				var routeemailaddress = document.getElementById("routeemailaddress").value;
				
				$.ajax({
					type:"POST",
					url: "../inc/AjaxFuncs.asp",
					dataType: "application/x-www-form-urlencoded",
					data: "action=addRoute&routeDescription="+encodeURIComponent(routeDescription)+"&routeDriver="+routeDriver+"&routeVehicle="+routeVehicle+"&routeemailaddress="+routeemailaddress,
					async: false,
					success: function(msg){document.getElementById("selRoutes").innerHTML = msg;}
				})
				$(this).dialog('close');
			}
		},
		Cancel: function() {
			$(this).dialog('close');
		}
	}
	});
}

function showAddRoutePopUp()
{
	document.getElementById("routeDescription").value = "";
	document.getElementById("routeDriver").value = "";
	document.getElementById("routeVehicle").value = "";
	document.getElementById("routeemailaddress").value = "";
	$('#addRouteWindow').dialog('open');
}

/*manage attributes*/

function initAddAttrPopUp(parentID)
{
	$.ajax({
        type:"POST",
        url: "../inc/AjaxFuncs.asp",
        dataType: "application/x-www-form-urlencoded",
        data: "action=showAddAttrPopup&parentID="+encodeURIComponent(parentID),
        async: false,
        success: function(msg){document.getElementById("addAttrWindow").innerHTML = msg;}
    })
	$('#addAttrWindow').dialog({ bgiframe: true,modal: true, autoOpen: false, title: 'Add Attribute', width:560, 
	buttons: {
		Save: function() {
				var attrName = document.getElementById("newtxtAttrName").value;
				if (attrName==''){
					alert("Please fill attribute name");
					document.getElementById("newtxtAttrName").focus();
				}
				else if ((document.getElementById("txtAttrThumbnail").value!="")&&(!CheckExtension(document.getElementById("txtAttrThumbnail"))))
				{} 
				else
				{
					$(this).dialog('close');
					document.getElementById("addAttrForm").submit();				
				}
		},
		Cancel: function() {
			$(this).dialog('close');
		}		
	},
	close: function(){
		$(this).dialog("destroy"); 
	}
	});
}

function showAddAttrPopUp()
{
	$('#addAttrWindow').dialog('open');
}

function initDeleteAttrPopUp(attributeID, parentID)
{
	$.ajax({
        type:"POST",
        url: "../inc/AjaxFuncs.asp",
        dataType: "application/x-www-form-urlencoded",
        data: "action=showDeleteAttrPopup&attributeID="+encodeURIComponent(attributeID)+"&parentID="+encodeURIComponent(parentID),
        async: false,
        success: function(msg){document.getElementById("deleteAttrWindow").innerHTML = msg;}
    })
	$('#deleteAttrWindow').dialog({ bgiframe: true,modal: true, autoOpen: false, title: 'Delete Attribute', width:560, 
	buttons: {
		Delete: function() {		
			$.ajax({
				type:"POST",
				url: "../inc/AjaxFuncs.asp",
				dataType: "application/x-www-form-urlencoded",
				data: "action=deleteAttribute&attributeID="+encodeURIComponent(attributeID)+"&parentID="+encodeURIComponent(parentID),
				async: false,
				success: function(msg){
					$(this).dialog('close');
					document.getElementById("deleteAttrForm").submit();
				}
			})
				
		},
		Cancel: function() {
			$(this).dialog('close');
		}		
	},
	close: function(){
		$(this).dialog("destroy"); 
	}
	});
}

function showDeleteAttrPopUp()
{
	$('#deleteAttrWindow').dialog('open');
}

function initEditAttrPopUp(attributeID)
{
	$.ajax({
        type:"POST",
        url: "../inc/AjaxFuncs.asp",
        dataType: "application/x-www-form-urlencoded",
        data: "action=showEditAttrPopup&attributeID="+encodeURIComponent(attributeID),
        async: false,
        success: function(msg){document.getElementById("editAttrWindow").innerHTML = msg;}
    })
	$('#editAttrWindow').dialog({ bgiframe: true,modal: true, autoOpen: false, title: 'Edit Attribute', width:560, 
	buttons: {
		Save: function() {
				var attrName = document.getElementById("edittxtAttrName").value;
				if (attrName==''){
					alert("Please fill attribute name");
					document.getElementById("edittxtAttrName").focus();
				}
				else if ((document.getElementById("txtAttrThumbnail").value!="")&&(!CheckExtension(document.getElementById("txtAttrThumbnail"))))
				{} 
				else{
					$(this).dialog('close');
					document.getElementById("editAttrForm").submit();				
				}
		},
		Cancel: function() {
			$(this).dialog('close');
		}		
	},
	close: function(){
		$(this).dialog("destroy"); 
	}
	});
}

function showEditAttrPopUp()
{
	$('#editAttrWindow').dialog('open');
}

/*Edit prodSKU*/
function initEditProdSKUPopUp(prodSKU)
{	
	document.getElementById("txtNewProdSKU").value = prodSKU;
	document.getElementById("editProdSKUError").innerHTML = "";
	
	$('#editProdSKUWindow').dialog({ bgiframe: true,modal: true, autoOpen: false, title: 'Edit ProdSKU', width:560, 
	buttons: {
		Save: function() {
				var newProdSKU = document.getElementById("txtNewProdSKU").value;
				if (newProdSKU==''){
					alert("Please fill prodSKU field");
					document.getElementById("txtNewProdSKU").focus();
				}
				else{
					if (newProdSKU!=prodSKU){				
						$.ajax({
							type:"POST",
							url: "../inc/AjaxFuncs.asp",
							dataType: "application/x-www-form-urlencoded",
							data: "action=checkProdSKUExistence&prodSKU="+encodeURIComponent(newProdSKU),
							async: false,
							success: function(msg){
								if (msg=='exists'){
									document.getElementById("editProdSKUError").innerHTML = "There is already a product with such Item# (SKU).";
								}else{
									$.ajax({
										type:"POST",
										url: "../inc/AjaxFuncs.asp",
										dataType: "application/x-www-form-urlencoded",
										data: "action=updateProdSKU&prodSKU="+encodeURIComponent(prodSKU)+"&newProdSKU="+encodeURIComponent(newProdSKU),
										async: false,
										success: function(msg){
											if (msg=='success'){
												$('#editProdSKUWindow').dialog('close');
												document.getElementById("prodSKUDiv").innerHTML = newProdSKU+"&nbsp;<a href='javascript:void(0)' style='font-size:13px' onclick=editProdSKU('"+newProdSKU+"')>Edit</a>";												
											}
										}
									})
								}
							}
						})
					}else{
						$(this).dialog('close');
					}
				}
		},
		Cancel: function() {
			$(this).dialog('close');
		}		
	},
	close: function(){
		$(this).dialog("destroy"); 
	}
	});
}

function showEditProdSKUPopUp()
{
	$('#editProdSKUWindow').dialog('open');
}

function showPreview(){
var data = CKEDITOR.instances.txtProdDesc.getData();    
var childWin = window.open("preview.html", "_blank");
   if (childWin.addEventListener)  // W3C DOM
        childWin.addEventListener("load",function() {childWin.document.getElementById("previewContent").innerHTML = data;},false);
   else if (childWin.attachEvent) { // IE DOM
         childWin.attachEvent("onload", function() {childWin.document.getElementById("previewContent").innerHTML = data;});
   } 
  
}

function saveEquivalentSKU(index) {

	var partnerid = document.getElementById("txtPartnerID"+index).value;
	var corpesssku = document.getElementById("txtCorpessSKU"+index).value;
	var partnersku = document.getElementById("txtPartnerSKU"+index).value;
	
	$.ajax({
		type:"POST",
		url:"../inc/AjaxFuncs.asp",
		dataType: "application/x-www-form-urlencoded",
		data:"action=saveEquivalentSKU&partnerid="+encodeURIComponent(partnerid)+"&corpesssku="+encodeURIComponent(corpesssku)+"&partnersku="+encodeURIComponent(partnersku),
		async:false,
		success: function(msg){
		if (msg=='success'){
			document.getElementById("txtPartnerSKU"+index).className = "";
			document.getElementById("txtPartnerSKU"+index).className = "partnerInputSaved";											
		}
		else if (msg=='skipped') {
			document.getElementById("txtPartnerSKU"+index).className = "";
			document.getElementById("txtPartnerSKU"+index).className = "partnerInputNotSaved";
		}
		else {
			document.getElementById("txtPartnerSKU"+index).className = "";
			document.getElementById("txtPartnerSKU"+index).className = "partnerInputSavedError";		
		}
		}
})

}

function clearEquivalentSKUCSS(index) {

	var partnersku = document.getElementById("txtPartnerSKU"+index).value;
	document.getElementById("txtPartnerSKU"+index).className = "";
	document.getElementById("txtPartnerSKU"+index).className = "partnerInputNotSaved";

}

function checkSKU(index) {

	var partnersku = document.getElementById("txtPartnerSKU"+index).value;
	var reg = new RegExp('^[-_a-zA-Z0-9]+$');

	if (!reg.test(partnersku) && (partnersku && partnersku.length != 0)) {
		alert('Please provide a valid SKU. Numbers, letters, underscores and hyphens are allowed. No special characters.');
		sku_input_field = document.getElementById("txtPartnerSKU"+index);
		//sku_input_field.value = "ERROR";
		sku_input_field.focus();
		return false;
	}
	else {
		saveEquivalentSKU(index);
	}
}





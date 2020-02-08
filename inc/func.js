//***************************************************
//***************************************************

var ShipName = "";
var ShipCompany = "";
var ShipAddress1 = "";
var ShipAddress2 = "";
var ShipCity = "";
var ShipState = "";
var ShipStateIndex = 0;
var ShipZip = "";
var ShipPhone = "";


//***************************************************
//***************************************************

function InitSaveVariables(form) 
{

ShipName = form.txtName.value;
ShipCompany = form.txtCompany.value;
ShipAddress1 = form.txtAddress1.value;
ShipAddress2 = form.txtAddress2.value;
ShipCity = form.txtCity.value;
ShipZip = form.txtZip.value;
ShipStateIndex = form.txtState.selectedIndex;
ShipState = form.txtState[ShipStateIndex].value;
ShipPhone = form.txtPhone.value;

}



//***************************************************
//***************************************************

function ShipToBillPerson(form) 
{
	if (form.chkSameShipping.checked) 
	{
		InitSaveVariables(form);
		form.txtShipToName.value = form.txtName.value;
		form.txtShipToCompany.value = form.txtCompany.value;
		form.txtShipToAddress1.value = form.txtAddress1.value;
		form.txtShipToAddress2.value = form.txtAddress2.value;
		form.txtShipToCity.value = form.txtCity.value;
		form.txtShipToZip.value = form.txtZip.value;
		form.txtShipToState.selectedIndex = form.txtState.selectedIndex;
		form.txtShipToPhone.value = form.txtPhone.value;
	}
	else 
	{
		form.txtShipToName.value = ShipName;
		form.txtShipToCompany.value = ShipCompany;
		form.txtShipToAddress1.value = ShipAddress1;
		form.txtShipToAddress2.value = ShipAddress2;
		form.txtShipToCity.value = ShipCity;
		form.txtShipToZip.value = ShipZip;       
		form.txtShipToState.selectedIndex = ShipStateIndex;
		form.txtShipToPhone.value = ShipPhone;
   	}
}


//***************************************************
//***************************************************
function isInteger(theField)
{
  if (theField.value == "")
  {
		return(true);
  }
    var i;
    var sString = theField.value;
    for (i = 0; i < sString.length; i++)
    {
        var c = sString.charAt(i);
        if (!((c <= "9") || (c == "0")))
        {
			alert('Please enter an integer for the quantity.');
			theField.focus();
			theField.select();
            return (false);
        }
    }
  if ((sString > 0) || (sString == 0))
  {
  return (true);
  }
  {
	alert('Please enter an integer for the quantity');
	theField.focus();
	theField.select();
	return (false);
  }
}

function isNumber(str){ 
	if(str.length==0) 
	{
		return false;
	} 
	numdecs = 0; 
	for (i2 = 0; i2 < str.length; i2++) 
	{
		mychar = str.charAt(i2); 
		if ((mychar >= "0" && mychar <= "9") || mychar == "." ){ 
		if (mychar == ".") 
			numdecs++; 
		} 
		else
		{ 
			return false; 
		}
	} 
	if (numdecs > 1){return false;} 
	return true; 
}

function handleParent(parentId)
{
	
	myParent = document.getElementById("collapse" + parentId)

	if (myParent.style.display=="none") {
		myParent.style.display="block";
		document.all("txtExpCol" + parentId).innerText = "-";
	} else {
		myParent.style.display="none";
		document.all("txtExpCol" + parentId).innerText ="+";
	}
}



function checkLogin (form) {
if (form.loginEmail.value == "") { alert("Please enter your Email Address."); form.loginEmail.focus(); return false;}
if (form.loginPassword.value == "") { alert("Please enter your Password."); form.loginPassword.focus(); return false;} return true;}

function checkEmail (form) { if (form.txtEmail.value == "") { alert("Please enter your Email Address."); form.txtEmail.focus(); return false;} return true;}

function Checkout ( form ){
if (form.txtFirstName.value == ""){alert("First Name must be filled in.");  form.txtFirstName.focus();  return false;}
if (form.txtLastName.value == ""){alert("Last Name must be filled in.");  form.txtLastName.focus();  return false;}
if (form.txtEmail.value == ""){alert("Email Address must be filled in.");  form.txtEmail.focus();  return false;}
if (form.txtuserPhone.value == ""){alert("Phone Number must be filled in.");  form.txtuserPhone.focus();  return false;}
if (form.txtPassword.value == ""){alert("Password must be filled in.");  form.txtPassword.focus();  return false;}
if (form.txtPassword2.value == ""){alert("Password Confirmation must be filled in.");  form.txtPassword2.focus();  return false;}
if (form.txtPassword.value !== form.txtPassword2.value){alert("Passwords Don\'t Match.");  form.txtPassword2.focus();  return false;}
if (form.txtAddress1.value == ""){alert("Address must be filled in.");  form.txtAddress1.focus();  return false;}
if (form.txtCity.value == ""){alert("City must be filled in.");  form.txtCity.focus();  return false;}
if (form.txtState.value == ""){alert("State must be filled in.");  form.txtState.focus();  return false;}
if (form.txtZip.value == ""){alert("Zip must be filled in.");  form.txtZip.focus();  return false;}
if (form.txtContact1.value == ""){alert("Contact must be filled in.");  form.txtContact1.focus();  return false;}
if (form.txtName.value == ""){alert("Company Name must be filled in.");  form.txtName.focus();  return false;} return true ;}

//***************************************************
//***************************************************

function verifyDollarAmount(objStartName, index, objEndName){

	var dollarValue = /^\b\d+[.]\d{2}\b$/;
	var message;
	var i;
	
	if ( objEndName == "Charge" )    {message = "shipping charge!";}
	if ( objEndName == "Start" )     {message = "order start amount!";}
	if ( objEndName == "End" )       {message = "order end amount!";}
	if ( objEndName == "OverCharge" ){message = "shipping charge!";}
	if ( objEndName == "Over" )      {message = "order amount!";}
	
	if ( objEndName == "OverCharge" || objEndName == "Over"){
		if ( !dollarValue.test(document.all(objStartName + objEndName).value) ){
			alert("Line 6 contains an invalid " + message);
			document.all(objStartName + objEndName).select();
			return false;
		}
	}
	else{
		for ( i=0; i<index; i++ ){
			if ( !dollarValue.test(document.all(objStartName + (i+1) + objEndName).value) ){
				alert("Line " + (i+1) + " contains an invalid " + message);
				document.all(objStartName + (i+1) + objEndName).select();
				return false;
			}
		}
	}
	return true;
}

//***************************************************
//***************************************************

function compareStartEnd(txtRatesStart, txtRatesEnd, objStartName, objEndName){

	var i;

	for ( i=0; i<txtRatesStart.length; i++ ){
		if ( txtRatesStart[i] != 0.00 && txtRatesEnd[i] != 0.00 ){
				if ( txtRatesStart[i] < txtRatesEnd[i] )
					continue;
				else{
					alert("Order end amount cannot be lower or equal to start amount!");
					document.all(objStartName + (i+1) + objEndName).select();
					return false
					break;
				}
			}
	}
}

//***************************************************
//***************************************************

function checkEndStartGap(txtRatesStart, txtRatesEnd, objStartName, objEndName){

	var i;
	var prevEndValue;
	var nextStartValue;

	for ( i=0; i<((txtRatesStart.length)-1); i++ ){
		if ( txtRatesEnd[i] != 0.00 && txtRatesStart[i+1] != 0.00){
			prevEndValue = txtRatesEnd[i];
			nextStartValue = ((parseFloat(prevEndValue)) + 0.01).toFixed(2);
			if( txtRatesStart[i+1] != nextStartValue ){
				alert("Order start amount must be $0.01 greater than last order end amount!");
				document.all(objStartName + (i+2) + objEndName).select();
				return false;
			}
		}
	}
}

//***************************************************
//***************************************************

function partialAmountFill(txtRatesStart, txtRatesEnd, objStartName, objEndName){

	var i;

	for ( i=1; i<txtRatesStart.length; i++ ){
		if ( objEndName == "Start" ){
			if ( txtRatesStart[i] == 0.00 && txtRatesEnd[i] != 0.00 ){
				alert("Line " + (i+1) + " order start amount cannot be $0.00!");
				document.all(objStartName + (i+1) + objEndName).select();
				return false;
			}
		}
		else if ( objEndName == "End" ){
			if ( txtRatesStart[i] != 0.00 && txtRatesEnd[i] == 0.00 ){
				alert("Line " + (i+1) + " order end amount cannot be $0.00!");
				document.all(objStartName + (i+1) + objEndName).select();
				return false;
			}
		}
		else if ( objEndName == "Charge" && objStartName != "txtRateOver" ){	
			if ( (txtRatesStart[i] == 0.00) && (txtRatesEnd[i] == 0.00) && ( (document.all(objStartName + (i+1) + objEndName).value) != 0.00 ) ){
				alert("Line " + (i+1) + " should have order amounts filled!");
				document.all(objStartName + (i+1) + objEndName).select();
				return false;
			}
		}
		if ( objStartName == "txtRateOver" ){
			if ( (document.all(objStartName + objEndName).value != 0.00) && (document.all(objStartName).value == 0.00) ){
				alert("Line 6 should have order amount filled!");
				document.all(objStartName + objEndName).select();
				return false;
			}
		}
	}
	if ( objEndName == "Charge" && objStartName != "txtRateOver" ){
		if ( (txtRatesEnd[0] == 0.00) && (document.all(objStartName + "1" + objEndName).value != 0.00 ) ){
			alert("Line 1 should have order amount filled!");
			document.all(objStartName + "1" + objEndName).select();
			return false;
		}
	}
}

//***************************************************
//***************************************************

function verifyOrderOver(txtRatesEnd, objStartName, objEndName){
	
	var i;
	var maxValue = 0.00;

	for ( i=0; i<txtRatesEnd.length; i++ ){
		if ( txtRatesEnd[i] >= maxValue ){
			maxValue = txtRatesEnd[i];
		}
	}

	if ( (document.all(objStartName + objEndName).value != 0.00) && (document.all(objStartName + objEndName).value <= maxValue) ){
		alert("Line 6 order amount cannot be of a lesser or equal amount then previous end amount!");
		document.all(objStartName + objEndName).select();
		return false;
	}
}

//***************************************************
//***************************************************

function checkShipRates(form){

	var txtRatesStart = new Array();
	var txtRatesEnd = new Array();
	var i;

	//Verify that the dollar amounts filled in are valid numerical values

	if ( verifyDollarAmount("txtRate", 5, "Charge") == false ) {return false;}
	if ( verifyDollarAmount("txtRate", 0, "OverCharge") == false ) {return false;}
	if ( verifyDollarAmount("txtRate", 0, "Over") == false ) {return false;}	

	if ( verifyDollarAmount("txtRate", 5, "Start") == false ) {return false;}
	else if ( verifyDollarAmount("txtRate", 5, "End") == false ) {return false;}
	else{
		for ( i=0; i<5; i++ ){
			txtRatesStart[i] = parseFloat(document.all("txtRate" + (i+1) + "Start").value)
			txtRatesEnd[i]   = parseFloat(document.all("txtRate" + (i+1) + "End").value)
		}
	}
	
	//Verify that end value is greater then start value
	
	if ( compareStartEnd(txtRatesStart, txtRatesEnd, "txtRate", "End") == false ) {return false;}

			
	//Verify that gap from end to start is of $0.01

	if ( checkEndStartGap(txtRatesStart, txtRatesEnd, "txtRate", "Start") == false ) {return false;}

	//Verify that any one line does not have partial entries
	
	if ( partialAmountFill(txtRatesStart, txtRatesEnd, "txtRate", "Start") == false ) {return false;}
	if ( partialAmountFill(txtRatesStart, txtRatesEnd, "txtRate", "End" ) == false) {return false;}
	if ( partialAmountFill(txtRatesStart, txtRatesEnd, "txtRate", "Charge") == false ) {return false;}
	if ( partialAmountFill(txtRatesStart, txtRatesEnd, "txtRateOver", "Charge") == false ) {return false;}


	//Verify that over order amount cannot be smaller then the last order end amount

	if ( verifyOrderOver(txtRatesEnd, "txtRate", "Over") == false ) {return false;}
}


//******* Get the correct shipTo info ********//
//******* based on selected location  ********//

function loadEditShipTo(userNo){
	btnShip();
	document.frmAddUser.action = "edit_user.asp?id=" + userNo;
	document.frmAddUser.submit();
}

function btnSave(userNo){
	document.getElementById("btnPush").value = "save";
	return true;		
}

function btnSaveUserAdd(){
	document.getElementById("btnPush").value = "save";
}

function btnSavePartnerAdd(){
	document.getElementById("btnPush").value = "save";
}

function btnShip(){
	document.getElementById("btnPush").value = "ship";
}

function loadAddShipTo(){
	btnShip();
	document.frmAddUser.action = "add_user.asp";
	document.frmAddUser.submit();
}

function allDigits(theField)
{
	return inValidCharSet(theField,"0123456789");
}

function inValidCharSet(theField,charset)
{
	var result = true;
	var str = theField.value

	for (var i=0;i<str.length;i++)
		if (charset.indexOf(str.substr(i,1))<0)
		{
			alert('Please enter an integer value.');
			theField.focus();
			theField.select();
			result = false;
			break;
		}
	
	return result;
}

/************************/
/* Add and Edit customer pages */
/************************/
function setShippingInfo()
{
	if (document.getElementById("chkShippingSameAsBill").checked == true){
		document.getElementById("txtName").value = document.getElementById("txtBillToName").value;
		document.getElementById("txtAddress1").value = document.getElementById("txtBillToAddress1").value;
		document.getElementById("txtAddress2").value = document.getElementById("txtBillToAddress2").value;
		document.getElementById("txtCity").value = document.getElementById("txtBillToCity").value;
		document.getElementById("txtState").selectedIndex = document.getElementById("txtBillToState").selectedIndex;
		document.getElementById("txtZip").value = document.getElementById("txtBillToZip").value;
	}
	else
	{
		document.getElementById("txtName").value = "";
		document.getElementById("txtAddress1").value = "";
		document.getElementById("txtAddress2").value = "";
		document.getElementById("txtCity").value = "";
		document.getElementById("txtState").selectedIndex = defaultStateIndex;
		document.getElementById("txtZip").value = "";
	}
}

function setShippingField(fieldName){
	if (document.getElementById("chkShippingSameAsBill").checked == true){
		if (fieldName == "State")
		{
			document.getElementById("txtState").selectedIndex = document.getElementById("txtBillToState").selectedIndex;
		}
		else
		{
			document.getElementById("txt" + fieldName).value = document.getElementById("txtBillTo"+fieldName).value;
		}
	}
}

function resetShippingAsBillSetting()
{
	document.getElementById("chkShippingSameAsBill").checked = false;
}

function checkPassword( form ){
	if (form.txtAddInfoPassword.value == ""){alert("Password can't be empty.");  form.txtAddInfoPassword.focus();  return false;}
return true ;}

function checkUserLimit(){
	var returnStatus;
	$.ajax({
        type:"POST",
        url: "../inc/AjaxFuncs.asp",
        dataType: "application/x-www-form-urlencoded",
        data: "action=checkUserLimitAjax",
        async: false,
        success: function(msg){
			if (msg=="ALLOW"){
				returnStatus = true;
			}else if (msg=="DENY"){	
				returnStatus = false;
				$.ajax({
					type:"POST",
					url: "../inc/AjaxFuncs.asp",
					dataType: "application/x-www-form-urlencoded",
					data: "action=showUserSubscriptionPopup",
					success: function(msg){
						document.getElementById("userSubscriptionPopup").innerHTML = msg;
						openUserSubscriptionPopup();
					}
				});				
			}else{
				returnStatus = false; // some error occured
			}
        }
    })
    return returnStatus;
}

function openUserSubscriptionPopup(){
	$('#userSubscriptionPopup').dialog({ bgiframe: true,modal: true, autoOpen: false, title: 'Subscription upgrade required', width:560, 
	buttons: {
		Save: function() {
			var userName = document.getElementById("txtUserName").value;
			var confirmation = document.getElementById("txtConfirmation").value;
			if (userName==''){
				alert("Please fill your name");
				document.getElementById("txtUserName").focus();
			}else if (confirmation.toLowerCase()!="agree"){
				alert("Please type AGREE");
				document.getElementById("txtConfirmation").focus();
			}else{
				$.ajax({
					type:"POST",
					url: "../inc/AjaxFuncs.asp",
					dataType: "application/x-www-form-urlencoded",
					data: "action=updateSubscriptionPlan&userName="+encodeURIComponent(userName),
					success: function(msg){
						document.getElementById("userSubscriptionPopup").innerHTML = "Your subscription plan has been updated successfully. You are being redirected to Add User page, please wait until page refreshes.";
						$(":button:contains('Save')").hide();
						$(":button:contains('Cancel')").hide();
						window.location = "../users/add_user.asp";	
					}
				});	
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
	$('#userSubscriptionPopup').dialog('open');
}

function selectCategory(level, id,page,pageBtn,numperpage, sortField, sortOrder, searchSKU, searchName){
	document.getElementById("currentCategoryParams").value = "id="+encodeURIComponent(id)+"&PageNum="+encodeURIComponent(page)+"&pageBtn="+encodeURIComponent(pageBtn)+"&numRecPerPage="+encodeURIComponent(numperpage)+"&O="+encodeURIComponent(sortOrder)+"&F="+encodeURIComponent(sortField)+"&searchSKU="+encodeURIComponent(searchSKU)+"&searchName="+encodeURIComponent(searchName);
	document.getElementById("currentCategory").value = id;
	$.ajax({		
		type:"POST",
		url: "../inc/AjaxFuncs.asp",
		dataType: "application/x-www-form-urlencoded",
		data: "action=displayCategorySettings&id="+encodeURIComponent(id),
		success: function(msg){document.getElementById("categorySettings").innerHTML=msg;}
	})
	if ((id!='newtop')&&(id!='newsub')){
		$.ajax({		
			type:"POST",
			url: "../inc/AjaxFuncs.asp",
			dataType: "application/x-www-form-urlencoded",
			data: "action=updateCategoryHeader&id="+encodeURIComponent(id),
			success: function(msg){
				document.getElementById("categoryHeader").innerHTML=msg;			
			}
		})
	}
	selectProducts(level, id,page,pageBtn,numperpage, sortField, sortOrder, searchSKU, searchName);
}

function addCategory(id){	
	$('#tabs ul li').removeClass('active');
	$('#tabs .divtabs').hide();
	$('#tabs .divtabs:first').show();
	$('#tabs ul li:first').addClass('active');
	
	document.getElementById("currentCategory").value = id;
	var parentCatID = $('div.tree-node-selected .catID').val();
	var parentCatClass = $('div.tree-node-selected .catID').attr('class');
	if (id=='newtop'){
		document.getElementById("categoryHeader").innerHTML = "Add new root category";		
		selectCategory(1,id,1,'','', '', '', '', '');
	}else{
		if (parentCatID){
			if (parentCatClass.indexOf("lev3")>0){
				alert("Only 3 levels of sub categories are allowed");
			}else{
				var level;
				if (parentCatClass.indexOf("lev1")>0){
					level = 2;
				}
				if (parentCatClass.indexOf("lev2")>0){
					level = 3;
				}
				$.ajax({		
					type:"POST",
					url: "../inc/AjaxFuncs.asp",
					dataType: "application/x-www-form-urlencoded",
					data: "action=updateCategoryHeader&id="+encodeURIComponent(parentCatID),
					success: function(msg){
						document.getElementById("categoryHeader").innerHTML="Add new sub category for &quot;" + msg + "&quot; category";			
					}
				})
				selectCategory(level,id,1,'','', '', '', '', '');
			}
		}else{
			alert("Please select category to add sub category");
			document.getElementById("currentCategory").value = '';
		}
	}
}

function selectProducts(level, id,page,pageBtn,numperpage, sortField, sortOrder, searchSKU, searchName) {
	document.getElementById("productsList").innerHTML="<table><tr><td width='550px' align='center'><img src='../images/loading.gif' /></td></tr></table>";
	$.ajax({		
		type:"POST",
		url: "../inc/AjaxFuncs.asp",
		dataType: "application/x-www-form-urlencoded",
		data: "action=selectCategoryProducts&id="+encodeURIComponent(id)+"&level="+encodeURIComponent(level)+"&PageNum="+encodeURIComponent(page)+"&pageBtn="+encodeURIComponent(pageBtn)+"&numRecPerPage="+encodeURIComponent(numperpage)+"&O="+encodeURIComponent(sortOrder)+"&F="+encodeURIComponent(sortField)+"&searchSKU="+encodeURIComponent(searchSKU)+"&searchName="+encodeURIComponent(searchName),
		success: function(msg){document.getElementById("productsList").innerHTML=msg;}
	})
}

function selectAllProducts(){
	if (document.getElementById("chbAll").checked == true){
		$(".prodChb").attr('checked','checked');
	}else{
		$('.prodChb').removeAttr('checked');
	}
}

function saveCategory(){ 
	var id = document.getElementById("currentCategory").value;
	
	if (id==''){
		alert("Please press 'Add root category' button to start adding new category");
	}else{
		var params = document.getElementById("currentCategoryParams").value;
		var productsNum = document.getElementById("productsNum").value;
		var productsToRemove = "";
		var productsToAdd = "";
		var rankToAdd = "";
		var productsToUpdate = "";
		var rankToUpdate = "";
	
		var prodChb, prodChbCopy, prodRank, prodSKU;
		for (i=0;i<=productsNum;i++){
			if (document.getElementById("prodChb"+i).checked == true){
				prodChb = 1;
			}else{
				prodChb = 0;
			}
			prodChbCopy = document.getElementById("prodChbCopy"+i).value;
			if (prodChb==prodChbCopy){
				prodRank = document.getElementById("prodRank"+i).value;
				if (prodRank != document.getElementById("prodRankCopy"+i).value){
					prodSKU = document.getElementById("prodSKU"+i).value;
					productsToUpdate = productsToUpdate + prodSKU + ",";
					rankToUpdate = rankToUpdate + prodRank + ",";
				} 
			}else{
				prodSKU = document.getElementById("prodSKU"+i).value;
				if (prodChb==0){
					productsToRemove = productsToRemove + prodSKU + ",";
				}else{
					prodRank = document.getElementById("prodRank"+i).value;
					productsToAdd = productsToAdd + prodSKU + ",";
					rankToAdd = rankToAdd + prodRank + ",";
				}
			}
		}
		
		//alert("productsToRemove:"+productsToRemove+" productsToAdd:" + productsToAdd + " rankToAdd:" + rankToAdd+" productsToUpdate:" + productsToUpdate + " rankToUpdate:" + rankToUpdate);
		
		var categoryDescription = document.getElementById("categoryDescription").value;
		var categoryRank = document.getElementById("categoryRank").value;
		var categoryDisplayOnWeb;
		if (document.getElementById("categoryDisplayOnWeb").checked == true) 
		{
			categoryDisplayOnWeb = 1;
		}
		else
		{
			categoryDisplayOnWeb = 0;
		}
		if (categoryDescription == "")
		{
			alert("Description must be filled in.");
			document.getElementById('categoryDescription').focus();
		}  
		else
		{
		
			if ((id=='newtop')||(id=='newsub')){
				var parentCatID = $('div.tree-node-selected .catID').val();
				document.getElementById("manageCategories").innerHTML="Category addition is in progress.. Please wait until page refreshes..";
				$.ajax({		
					type:"POST",
					url: "../inc/AjaxFuncs.asp",
					dataType: "application/x-www-form-urlencoded",
					data: "action=addCategory&id="+encodeURIComponent(id)+"&parentid="+encodeURIComponent(parentCatID)+"&categoryDescription="+encodeURIComponent(categoryDescription)+"&categoryRank="+categoryRank+"&categoryDisplayOnWeb="+categoryDisplayOnWeb+"&productsToAdd="+encodeURIComponent(productsToAdd)+"&rankToAdd="+encodeURIComponent(rankToAdd),
					success: function(msg){
						if (msg!=''){
							document.getElementById("productsInOtherCategoriesPopup").innerHTML = msg;
							showProductsInOtherCategoriesPopup();
						}else{
							window.location = "manage_categories.asp";
						}
					}
				})
			}else{
				document.getElementById("manageCategories").innerHTML="Category update is in progress.. Please wait until page refreshes..";
				$.ajax({		
					type:"POST",
					url: "../inc/AjaxFuncs.asp",
					dataType: "application/x-www-form-urlencoded",
					data: "action=updateCategory&id="+encodeURIComponent(id)+"&categoryDescription="+encodeURIComponent(categoryDescription)+"&categoryRank="+categoryRank+"&categoryDisplayOnWeb="+categoryDisplayOnWeb+"&productsToAdd="+encodeURIComponent(productsToAdd)+"&rankToAdd="+encodeURIComponent(rankToAdd)+"&productsToUpdate="+encodeURIComponent(productsToUpdate)+"&rankToUpdate="+encodeURIComponent(rankToUpdate)+"&productsToRemove="+encodeURIComponent(productsToRemove),
					success: function(msg){
						if (msg!=''){
							document.getElementById("productsInOtherCategoriesPopup").innerHTML = msg;
							showProductsInOtherCategoriesPopup();
						}else{
							window.location = "manage_categories.asp";
						}
					}
				})
			}
		}
	}
}

function showProductsInOtherCategoriesPopup(){
	$('#productsInOtherCategoriesPopup').dialog({ bgiframe: true,modal: true, autoOpen: false, title: 'Products exist in other categories', width:560, 
	buttons: {
		Continue: function() {
		
			var productsNum = document.getElementById("prodsNum").value;
			var productsToRemove = "";
			var cat1Str = "";
			var cat2Str = "";
			var cat3Str = "";
		
			for (i=0;i<=productsNum;i++){
				if (document.getElementById("chbRemove"+i).checked == true){
					productsToRemove = productsToRemove + document.getElementById("prodSKU"+i).value+",";
					cat1Str = cat1Str + document.getElementById("Category1_"+i).value+",";
					cat2Str = cat2Str + document.getElementById("Category2_"+i).value+",";
					cat3Str = cat3Str + document.getElementById("Category3_"+i).value+",";
				}
			}	
			if (productsToRemove!=""){		
				$.ajax({		
					type:"POST",
					url: "../inc/AjaxFuncs.asp",
					dataType: "application/x-www-form-urlencoded",
					data: "action=deleteProductsFromOtherCategories&productsToRemove="+encodeURIComponent(productsToRemove)+"&cat1Str="+encodeURIComponent(cat1Str)+"&cat2Str="+encodeURIComponent(cat2Str)+"&cat3Str="+encodeURIComponent(cat3Str),
					success: function(msg){}
				})
			}	
			$(this).dialog('close');	
		},
		Cancel: function() {
			$(this).dialog('close');
		}		
	},
	close: function(){
		$(this).dialog("destroy");
		window.location = "manage_categories.asp"; 
	}
	});


	$('#productsInOtherCategoriesPopup').dialog('open');
}

function checkSearchFields(){
	if ((document.getElementById('searchSKU').value=='')&&(document.getElementById('searchName').value=='')){
		alert("Please specify criteria for search");
		return false;
	}
	return true;
}

function deleteCategory(){
	var id = document.getElementById("currentCategory").value;
	if ((id!='')&&(id!='newtop')&&(id!='newsub')){
		var level = $('div.tree-node-selected .catID').attr('class');
		$.ajax({		
			type:"POST",
			url: "../inc/AjaxFuncs.asp",
			dataType: "application/x-www-form-urlencoded",
			data: "action=getDeleteCategoryConfirmation&id="+encodeURIComponent(id)+"&level="+encodeURIComponent(level),
			success: function(msg){
				document.getElementById("deleteCategoryPopup").innerHTML = msg;
				showDeleteCategoryPopup(id);
			}
		})
	}
}

function showDeleteCategoryPopup(id){
	$('#deleteCategoryPopup').dialog({ bgiframe: true,modal: true, autoOpen: false, title: 'Delete Category', width:560, 
		buttons: {
			Delete: function() {			
				$.ajax({		
					type:"POST",
					url: "../inc/AjaxFuncs.asp",
					dataType: "application/x-www-form-urlencoded",
					data: "action=deleteCategory&id="+encodeURIComponent(id),
					success: function(msg){
						document.getElementById("manageCategories").innerHTML="Category deletion is in progress.. Please wait until page refreshes..";
						window.location = "manage_categories.asp";
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
    
	$('#deleteCategoryPopup').dialog('open');
}

function resetCategory(){
	var id = document.getElementById("currentCategory").value;
	if (id!=''){
		var levelStr = $('div.tree-node-selected .catID').attr('class');
		var level;
		if (levelStr.indexOf("lev1")>0){level = 1;}
		if (levelStr.indexOf("lev2")>0){level = 2;}
		if (levelStr.indexOf("lev3")>0){level = 3;}
		selectCategory(level, id,1,'','', '', '', '', '');
	}
}

/* RouteSel account functions */

function showHideAccountSettings() {
	if (document.getElementById("txtUserTypeR").checked == true) {
		document.getElementById("divAccountSettings").style.visibility = "visible";
		document.getElementById("divAccountSettings").style.display = "block";
	} else {
		document.getElementById("divAccountSettings").style.visibility = "hidden";
		document.getElementById("divAccountSettings").style.display = "none";
	}
}

function compareNames(a, b) {
	var nameA = a.Val.toLowerCase();
	var nameB = b.Val.toLowerCase();
	if (nameA < nameB) { return -1 }
	if (nameA > nameB) { return 1 }
	return 0;
}

function addAccount() {
	var selCustID = document.frmAddUser.lstAllCustIDs.options[document.frmAddUser.lstAllCustIDs.selectedIndex];
	var count = arrSelectedCustIDs.length;
	arrSelectedCustIDs[count] = { Val: selCustID.value, Text: selCustID.text };
	arrSelectedCustIDs.sort(compareNames);

	//populate listbox of selected ids from array
	document.frmAddUser.lstSelectedCustIds.length = arrSelectedCustIDs.length;
	for (i = 0; i < arrSelectedCustIDs.length; i++) {
		document.frmAddUser.lstSelectedCustIds.options[i].text = arrSelectedCustIDs[i].Text;
		document.frmAddUser.lstSelectedCustIds.options[i].value = arrSelectedCustIDs[i].Val;
	}
	//remove selected id from the list of available ids
	document.frmAddUser.lstAllCustIDs.options[document.frmAddUser.lstAllCustIDs.selectedIndex] = null;
	document.frmAddUser.lstAllCustIDs.selectedIndex = 0;
	//update array with available ids
	arrAvailableCustIDs.length = document.frmAddUser.lstAllCustIDs.length;
	for (i = 0; i < document.frmAddUser.lstAllCustIDs.length; i++) {
		arrAvailableCustIDs[i] = { Val: document.frmAddUser.lstAllCustIDs.options[i].value, Text: document.frmAddUser.lstAllCustIDs.options[i].text };
	}
}


function deleteAccount() {
	var delAccount = document.frmAddUser.lstSelectedCustIds.options[document.frmAddUser.lstSelectedCustIds.selectedIndex];
	var count = arrAvailableCustIDs.length;
	//return deleted id to array of available ids
	arrAvailableCustIDs[count] = { Val: delAccount.value, Text: delAccount.text };
	arrAvailableCustIDs.sort(compareNames);
	//populate the list of available ids from array
	document.frmAddUser.lstAllCustIDs.length = arrAvailableCustIDs.length;
	for (i = 0; i < arrAvailableCustIDs.length; i++) {
		document.frmAddUser.lstAllCustIDs.options[i].text = arrAvailableCustIDs[i].Text;
		document.frmAddUser.lstAllCustIDs.options[i].value = arrAvailableCustIDs[i].Val;
	}
	document.frmAddUser.lstAllCustIDs.selectedIndex = 0;
	//remove deleted id from the list of selected ids
	document.frmAddUser.lstSelectedCustIds.options[document.frmAddUser.lstSelectedCustIds.selectedIndex] = null;
	//update array with selected ids
	arrSelectedCustIDs.length = document.frmAddUser.lstSelectedCustIds.length;
	for (i = 0; i < document.frmAddUser.lstSelectedCustIds.length; i++) {
		arrSelectedCustIDs[i].Val = document.frmAddUser.lstSelectedCustIds.options[i].value;
		arrSelectedCustIDs[i].Text = document.frmAddUser.lstSelectedCustIds.options[i].text;
	}
}

function setAccounts() {
	var strCustomerIDs = "";
	for (i = 0; i < document.frmAddUser.lstSelectedCustIds.length; i++) {
		strCustomerIDs = strCustomerIDs + document.frmAddUser.lstSelectedCustIds.options[i].value + ",";
	}
	if (strCustomerIDs.length >= 1) {
		strCustomerIDs = strCustomerIDs.substring(0, strCustomerIDs.length - 1);
	}
	document.frmAddUser.accountsList.value = strCustomerIDs;
}

function addRouteselAccount() {
	var selCustID = document.frmAddRouteselUser.lstAllCustIDs.options[document.frmAddRouteselUser.lstAllCustIDs.selectedIndex];
	var count = arrSelectedCustIDs.length;
	arrSelectedCustIDs[count] = { Val: selCustID.value, Text: selCustID.text };
	arrSelectedCustIDs.sort(compareNames);

	//populate listbox of selected ids from array
	document.frmAddRouteselUser.lstSelectedCustIds.length = arrSelectedCustIDs.length;
	for (i = 0; i < arrSelectedCustIDs.length; i++) {
		document.frmAddRouteselUser.lstSelectedCustIds.options[i].text = arrSelectedCustIDs[i].Text;
		document.frmAddRouteselUser.lstSelectedCustIds.options[i].value = arrSelectedCustIDs[i].Val;
	}
	//remove selected id from the list of available ids
	document.frmAddRouteselUser.lstAllCustIDs.options[document.frmAddRouteselUser.lstAllCustIDs.selectedIndex] = null;
	document.frmAddRouteselUser.lstAllCustIDs.selectedIndex = 0;
	//update array with available ids
	arrAvailableCustIDs.length = document.frmAddRouteselUser.lstAllCustIDs.length;
	for (i = 0; i < document.frmAddRouteselUser.lstAllCustIDs.length; i++) {
		arrAvailableCustIDs[i] = { Val: document.frmAddRouteselUser.lstAllCustIDs.options[i].value, Text: document.frmAddRouteselUser.lstAllCustIDs.options[i].text };
	}
}

function deleteRouteselAccount() {
	var delAccount = document.frmAddRouteselUser.lstSelectedCustIds.options[document.frmAddRouteselUser.lstSelectedCustIds.selectedIndex];
	var count = arrAvailableCustIDs.length;
	//return deleted id to array of available ids
	arrAvailableCustIDs[count] = { Val: delAccount.value, Text: delAccount.text };
	arrAvailableCustIDs.sort(compareNames);
	//populate the list of available ids from array
	document.frmAddRouteselUser.lstAllCustIDs.length = arrAvailableCustIDs.length;
	for (i = 0; i < arrAvailableCustIDs.length; i++) {
		document.frmAddRouteselUser.lstAllCustIDs.options[i].text = arrAvailableCustIDs[i].Text;
		document.frmAddRouteselUser.lstAllCustIDs.options[i].value = arrAvailableCustIDs[i].Val;
	}
	document.frmAddRouteselUser.lstAllCustIDs.selectedIndex = 0;
	//remove deleted id from the list of selected ids
	document.frmAddRouteselUser.lstSelectedCustIds.options[document.frmAddRouteselUser.lstSelectedCustIds.selectedIndex] = null;
	//update array with selected ids
	arrSelectedCustIDs.length = document.frmAddRouteselUser.lstSelectedCustIds.length;
	for (i = 0; i < document.frmAddRouteselUser.lstSelectedCustIds.length; i++) {
		arrSelectedCustIDs[i].Val = document.frmAddRouteselUser.lstSelectedCustIds.options[i].value;
		arrSelectedCustIDs[i].Text = document.frmAddRouteselUser.lstSelectedCustIds.options[i].text;
	}
}

function setRouteselAccounts() {
	var strCustomerIDs = "";
	for (i = 0; i < document.frmAddRouteselUser.lstSelectedCustIds.length; i++) {
		strCustomerIDs = strCustomerIDs + document.frmAddRouteselUser.lstSelectedCustIds.options[i].value + ",";
	}
	if (strCustomerIDs.length >= 1) {
		strCustomerIDs = strCustomerIDs.substring(0, strCustomerIDs.length - 1);
	}
	document.frmAddRouteselUser.accountsList.value = strCustomerIDs;
}



<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/InsightFuncs_Equipment.asp"-->
<!--#include file="../../inc/InsightFuncs_BizIntel.asp"-->

<%
Server.ScriptTimeout = 900000 'Default value

ClientKeyForFileNames = MUV_READ("ClientKeyForFileNames")

ClearSearchForm = Request.QueryString("s")

If ClearSearchForm = "clear" Then

		frmEquipmentAssetFullSubmitted = ""
		frmEquipmentAssetPartialSubmitted = ""
		frmCustomerIDSubmitted = ""
		frmClassManfBrandModelSubmitted = ""
		frmGroupSubmitted = ""
		
		EquipFullAssetSerial = ""
		EquipPartialAssetSerial = ""
		EquipCustomerID = ""
		EquipClassIntRecID = ""
		EquipManfIntRecID = ""
		EquipBrandIntRecID = ""
		EquipModelIntRecID = ""
		EquipShowOnlyAvailForPlacement1 = 0
		EquipGroupIntRecID = ""
		EquipShowOnlyAvailForPlacement2 = 0		

Else

	'******************************************************************************
	'********************Check If Form 1 Submittited*******************************
	'******************************************************************************
	
	frmEquipmentAssetFullSubmitted = Request.Form("frmEquipmentAssetFullSubmitted")
	
	If frmEquipmentAssetFullSubmitted <> "" Then
		EquipFullAssetSerial = Request.Form("txtEquipIDToPass")
					
		Response.Write("<div id=""PleaseWaitPanel"">")
		Response.Write("<br><br><strong>Analyzing Matching Assets, please wait...</strong><br><br>")
		Response.Write("<img src='" & baseURL & "/img/loading.gif'/>")
		Response.Write("</div>")
		Response.Flush()
		
	Else
		EquipFullAssetSerial = ""
	End If
	
	
	'******************************************************************************
	'********************Check If Form 2 Submittited*******************************
	'******************************************************************************
	
	frmEquipmentAssetPartialSubmitted = Request.Form("frmEquipmentAssetPartialSubmitted")
	
	If frmEquipmentAssetPartialSubmitted <> "" Then
		EquipPartialAssetSerial = Request.Form("txtPartialEquip")
		
		Response.Write("<div id=""PleaseWaitPanel"">")
		Response.Write("<br><br><strong>Analyzing Matching Assets, please wait...</strong><br><br>")
		Response.Write("<img src='" & baseURL & "/img/loading.gif'/>")
		Response.Write("</div>")
		Response.Flush()
		
	Else
		EquipPartialAssetSerial = ""
	End If
	
	
	'******************************************************************************
	'********************Check If Form 3 Submittited*******************************
	'******************************************************************************
	
	frmCustomerIDSubmitted = Request.Form("frmCustomerIDSubmitted")
	
	If frmCustomerIDSubmitted <> "" Then
	
		EquipCustomerID = Request.Form("txtCustomerIDToPass")
		
		CustName = GetCustNameByCustNum(EquipCustomerID)			
		Response.Write("<div id=""PleaseWaitPanel"">")
		Response.Write("<br><br><strong>Analyzing " & CustName & ", please wait...</strong><br><br>")
		Response.Write("<img src='" & baseURL & "/img/loading.gif'/>")
		Response.Write("</div>")
		Response.Flush()
	Else
		EquipCustomerID = ""
	End If
	

	
	'******************************************************************************
	'********************Check If Form 4 Submittited*******************************
	'******************************************************************************
	
	frmClassManfBrandModelSubmitted = Request.Form("frmClassManfBrandModelSubmitted")
	
	If frmClassManfBrandModelSubmitted <> "" Then
		EquipClassIntRecID = Request.Form("selClassIntRecID")
		EquipManfIntRecID = Request.Form("selManfIntRecID")
		EquipBrandIntRecID = Request.Form("selBrandIntRecID")
		EquipModelIntRecID = Request.Form("selModelIntRecID")
		EquipShowOnlyAvailForPlacement1 = Request.Form("chkShowOnlyAvailForPlacement1")
		If EquipShowOnlyAvailForPlacement1 = "on" then EquipShowOnlyAvailForPlacement1 = 1 Else EquipShowOnlyAvailForPlacement1 = 0
		
		Response.Write("<div id=""PleaseWaitPanel"">")
		Response.Write("<br><br><strong>Analyzing Matching Classes, Manuf. Brands and Models, please wait...</strong><br><br>")
		Response.Write("<img src='" & baseURL & "/img/loading.gif'/>")
		Response.Write("</div>")
		Response.Flush()
		
	Else
		EquipClassIntRecID = ""
		EquipManfIntRecID = ""
		EquipBrandIntRecID = ""
		EquipModelIntRecID = ""
		EquipShowOnlyAvailForPlacement1 = 0
	End If
	
	
	
	'******************************************************************************
	'********************Check If Form 5 Submittited*******************************
	'******************************************************************************
	
	frmGroupSubmitted = Request.Form("frmGroupSubmitted")
	
	If frmGroupSubmitted <> "" Then
		EquipGroupIntRecID = Request.Form("selGroupIntRecID")
		EquipShowOnlyAvailForPlacement2 = Request.Form("chkShowOnlyAvailForPlacement2")
		If EquipShowOnlyAvailForPlacement2 = "on" then EquipShowOnlyAvailForPlacement2 = 1 Else EquipShowOnlyAvailForPlacement2 = 0
		
		Response.Write("<div id=""PleaseWaitPanel"">")
		Response.Write("<br><br><strong>Analyzing Matching Groups, please wait...</strong><br><br>")
		Response.Write("<img src='" & baseURL & "/img/loading.gif'/>")
		Response.Write("</div>")
		Response.Flush()
		
	Else
		EquipGroupIntRecID = ""
		EquipShowOnlyAvailForPlacement2 = 0
	End If
	
End If
	
'******************************************************************************
'************************Form Testing Script***********************************
'******************************************************************************

'Response.Write "<br>"
'Response.Write "<br>"
'Response.Write "<br>"
'Response.Write "<br>"

Dim Item, fieldName, fieldValue
Dim a, b, c, d

Set d = Server.CreateObject("Scripting.Dictionary")

For Each Item In Request.Form
    fieldName = Item
    fieldValue = Request.Form(Item)

    d.Add fieldName, fieldValue
Next

' Rest of the code is for going through the Dictionary
a = d.Keys  ' Field names  '
b = d.Items ' Field values '

For c = 0 To d.Count - 1
    'Response.Write a(c) & " = " & b(c)
    'Response.Write "<br>"
Next

'******************************************************************************
'************************END Form Testing Script*******************************
'******************************************************************************

%>
<!---------------------------------------------------------------------------------------------------------->
<!----------THIS IS A CUSTOM STYLESHEET ADDED FOR THE AUTOCOMPLETE SEARCH FOR CATEGORY ANALYSIS ONLY-------->
<!-----------IT OVERRIDES THE STYLES THAT ARE STILL LOADED IN HEADER.ASP------------------------------------>

<!---------------------------------------------------------------------------------------------------------->
<link rel="stylesheet" href="<%= BaseURL %>js/easyautocomplete/EasyAutocomplete-1.4.0/easy-autocomplete-cat-analysis.css"> 

<!---------------------------------------------------------------------------------------------------------->
<!---------------------------------------------------------------------------------------------------------->
<!--
	js/easyautocomplete/EasyAutocomplete-1.4.0/easy-autocomplete.themes.css ALSO CONTAINS A CUSTOM STYLE
	SET CALLED "easy-autocomplete.eac-cat-analysis"
-->
<!---------------------------------------------------------------------------------------------------------->
<!---------------------------------------------------------------------------------------------------------->

<script type="text/javascript">


$(document).ready(function(){	

	$("#PleaseWaitPanel").hide();
		
	$('section h4').click(function(event) {
	  event.preventDefault();
	  $(this).addClass('active');
	  $(this).siblings().removeClass('active');
	
	  var ph = $(this).parent().height();
	  var ch = $(this).next().height();
	
	  if (ch > ph) {
	    $(this).parent().css({
	      'min-height': ch + 'px'
	    });
	  } else {
	    $(this).parent().css({
	      'height': 'auto'
	    });
	  }
	});
	
	function tabParentHeight() {
	  var ph = $('section').height();
	  var ch = $('section ul').height();
	  if (ch > ph) {
	    $('section').css({
	      'height': ch + 'px'
	    });
	  } else {
	    $(this).parent().css({
	      'height': 'auto'
	    });
	  }
	}
	
	$(window).resize(function() {
	  tabParentHeight();
	});
	
	$(document).resize(function() {
	  tabParentHeight();
	});
	
	tabParentHeight();	
	
	var randomNumberBetween0and100 = Math.floor(Math.random() * 100);
	
	var autocompleteJSONFileURLEquipment = "../../clientfiles/<%= ClientKeyForFileNames %>/autocomplete/equipment_list_<%= ClientKeyForFileNames %>.json?v=" + randomNumberBetween0and100;
	var autocompleteJSONFileURLAccount = "../../clientfiles/<%= ClientKeyForFileNames %>/autocomplete/customer_account_list_CSZ_Equipment.json?v=" + randomNumberBetween0and100;

	var optionsEquipment = {
	  url: autocompleteJSONFileURLEquipment,
	  placeholder: "Search for equipment by FULL serial number or asset tag",
	  getValue: "name",
	  list: {	
        onChooseEvent: function() {
        
            var EquipIntRecID = $("#txtEquipID").getSelectedItemData().code;
            $("#txtEquipIDToPass").val(EquipIntRecID);
            //window.location.href = "editEquipment.asp?i=" + EquipIntRecID;
            
    	},		  
	    match: {
	      enabled: true
		},
		maxNumberOfElements: 30		
	  },
	  theme: "cat-analysis"
	};
	
	$("#txtEquipID").easyAutocomplete(optionsEquipment);
	
	
	
  	
  	var optionsCustomer = {
	  url: autocompleteJSONFileURLAccount,
	  placeholder: "Search for a customer by name, account, city, state, zip",
	  getValue: "name",
	  list: {	
        onChooseEvent: function() {
        
            var custID = $("#txtCustomerID").getSelectedItemData().code;
            $("#txtCustomerIDToPass").val(custID);
            //window.location.href = "editEquipment.asp?i=" + custID;
            
    	},		  
	    match: {
	      enabled: true
		},
		maxNumberOfElements: 30		
	  },
	  theme: "cat-analysis"
	};
	
	$("#txtCustomerID").easyAutocomplete(optionsCustomer);


        
    $('#filter-equipment').keyup(function () {

        var rex = new RegExp($(this).val(), 'i');
        $('.searchable-equipment tr').hide();
        $('.searchable-equipment tr').filter(function () {
            return rex.test($(this).text());
        }).show();
    })
	

});



</script>


 
<style type="text/css">
	
	.table-responsive{
		width:1775px;
	}

	
	table.sortable th:not(.sorttable_sorted):not(.sorttable_sorted_reverse):not(.sorttable_nosort):after { 
    	content: " \25B4\25BE" 
	}
	
	.nav-tabs>li>a{
		background: #f5f5f5;
		border: 1px solid #ccc;
		color: #000;
	}
	
	.nav-tabs>li>a:hover{
		border: 1px solid #ccc;
	}
	
	.nav-tabs>li.active>a, .nav-tabs>li.active>a:focus, .nav-tabs>li.active>a:hover{
		color: #000;
		border: 1px solid #ccc;
	}
	

	.narrow-results{
		margin:0px 0px 20px 0px;
	}
	
	#filter{
		width:60%;
	}

	.modal-link{
		cursor:pointer;
	}
	
	.modal-content{
		max-height:360px;
		overflow-y:auto;
	}

	 .modal-content .row{
		 padding-bottom:20px;
	 }
	
	 .modal-content p{
		 margin-bottom:20px;
		 white-space:normal;
	 }
	
	
	#searchIcon {
	    left: auto !important;
	    float: right;
	    margin-right: 685px;
	}	


	/********************************/
	/* BEGIN CODE FOR VERTICAL TABS */
	/********************************/
	html {
	  box-sizing: border-box;
	}
	*,
	*:before,
	*:after {
	  box-sizing: border-box;
	}

	.container {
	  width: 100%;
	  line-height: 1.8em;
	  /*letter-spacing: 0.1em;*/
	  color: #000;
	  max-width:1800px;
	  margin-left:0px;
	  font-family:Arial, Helvetica, sans-serif;
	  /*background: #E9E9E9;*/
	  /* font-family: 'Helvetica Neue', serif;*/
	}

	h1.search-header {
	  font-size: 2em;
	  line-height: 1.4em;
	  text-align: center;
	  padding: 0.5em;
	  margin-top:0px;
	  margin-bottom:0px;
	}
	
	h2.search-subheader {
	  font-size: 1.2em;
	  line-height: 1.4em;
	  text-align: center;
	  padding: 0.5em;
	  margin-top:-20px;
	  margin-bottom:0px;
	  color:#585858;
	}

	section.search-criteria {
	  zoom: 1;
	  position: relative;
	  height: auto;
	  background: #E9E9E9;
	}
	
	section.search-criteria:after,
	section.search-criteria:before {
	  content: "";
	  display: table;
	}
	
	section.search-criteria:after {
	  clear: both;
	}
	
	section.search-criteria h4 {
	  background: rgba(0,0,0,0.1);
	  cursor: pointer;
	  border: 1px solid rgba(0,0,0,0.2);
	  border-top: none;
	  padding: 15px 20px;
	  margin-top: 0px !important;
	  margin-bottom: 0px !important;
	}
	
	section.search-criteria h4:first-child {
	  border-top: 1px solid rgba(0,0,0,0.2);
	}
	
	@media screen and (min-width: 600px) {
	  section.search-criteria h4 {
	    position: relative;
	    width: 33.333333333333336%;
	    height: 20%;
	    display: block;
	  }
	}
	
	section.search-criteria ul {
	  zoom: 1;
	  position: relative;
	  height: auto;
	  min-height: 100%;
	  border: 1px solid rgba(0,0,0,0.2);
	  border-left: none;
	  display: none;
	}
	
	section.search-criteria ul:after,
	section.search-criteria ul:before {
	  content: "";
	  display: table;
	}
	
	section.search-criteria ul:after {
	  clear: both;
	}
	
	section.search-criteria ul li {
	  list-style: none;
	}
	
	@media screen and (min-width: 600px) {
	  section.search-criteria ul {
	    position: absolute;
	    width: 66.66666666666667%;
	    right: 0;
	    top: 0;
	    padding: 15px 30px;
	  }
	}
	
	section.search-criteria .active {
	  cursor: default;
	  border-bottom: 1px solid rgba(0,0,0,0.2);
	  border-right: none;
	}
	
	@media screen and (min-width: 600px) {
	  section.search-criteria .active {
	    background: rgba(0,0,0,0);
	    border-right: 1px solid rgba(0,0,0,0);
	  }
	}
	
	section.search-criteria .active + ul {
	  display: block;
	}
	/********************************/
	/* END CODE FOR VERTICAL TABS */
	/********************************/
	
	
	.form-control-partial-search{
	    border-color: #ccc;
	    border-radius: 7px;
	    border-style: solid;
	    border-width: 1px;
	    box-shadow: 0 1px 2px rgba(0, 0, 0, 0.1) inset;
	    color: #2e6da4;
	    float: none;
	    padding: 6px 12px;
	    width: 450px;
	    height: 45px;
	}
	
	.select50 {
		width: 450px;
	    height: 45px;
	    border-radius: 6px;
	    display: inline;
	}
	
	.btn{
		line-height: 1.8em !important;
	}
		
	.easy-autocomplete.eac-cat-analysis ul {
	    background: none repeat scroll 0 0 #ffffff;
	    border-top: 1px dotted #ccc;
	    display: none;
	    margin-top: 0;
	    padding-bottom: 0;
	    padding-left: 0;
	    position: relative;
	    top: -1px;
	    max-height: 400px;
	    overflow-y: scroll;
	    width:440px !important;
	    
	}
	
	.easy-autocomplete.eac-cat-analysis ul li{
		border-right:0px !important;
	}
	
	.searchrow {
	    /* margin-right: -15px; */
	    margin-left:0px;
	    margin-top:20px;
	}	

</style>

<!--- eof on/off scripts !-->

<h1 class="page-header">Find / Edit <%= GetTerm("Equipment") %></h1>

 	<div class="row">
	 	<div class="col-lg-12">
		 	<p><a href="findEquipment.asp?s=c"><button type="button" class="btn btn-success">Clear All Search Results</button></a></p>
	 	</div>
	</div>
   
    <section class="container">
    
		<h1 class="search-header">Select Equipment Search Criteria</h1>
		<section class="search-criteria">
		
		  <h4 <% If frmEquipmentAssetPartialSubmitted = "" AND frmClassManfBrandModelSubmitted = "" AND frmGroupSubmitted = "" AND frmCustomerIDSubmitted = "" Then Response.Write("class='active'") %>>Search By Full Asset Tag, Serial #</h4>
		  
		  <ul>
		  
		  	<li style="margin-bottom:15px;">
		  		<strong>Note</strong>: This search box will provide matching search terms. Please select/click the one you want.
		  	</li>
		  
		    <li>	
		    	<form id="frmEquipmentAssetFull" name="frmEquipmentAssetFull" action="<%= BaseURL %>equipment/equipment/findEquipment.asp" method="post">	  
		    	
		    		<input type="hidden" name="frmEquipmentAssetFullSubmitted" id="frmEquipmentAssetFullSubmitted" value="frmEquipmentAssetFullSubmitted">
		    		
		    		<!-- select equipment record !-->
					<input id="txtEquipID" name="txtEquipID">
					<input type="hidden" id="txtEquipIDToPass" name="txtEquipIDToPass" value="<%= EquipFullAssetSerial %>">
					<i id="searchIcon" class="fa fa-search fa-2x"></i>
					
					<div class="form-group" style="margin-top:15px;">
						<input type="submit" id="btnSubmit" name="btnSubmit" class="btn btn-primary" value="Submit">
					</div>
					<!-- eof select equipment record !-->
				</form>
		  	</li>
		  	
		  </ul>

		  <h4 <% If frmEquipmentAssetPartialSubmitted <> "" Then Response.Write("class='active'") %>>Search By Partial Asset Tag, Serial #</h4>
		  
		  <ul>
		    <li>
	    		<form id="frmEquipmentAssetPartial" name="frmEquipmentAssetPartial" action="<%= BaseURL %>equipment/equipment/findEquipment.asp" method="post">
	    		
	    			<input type="hidden" name="frmEquipmentAssetPartialSubmitted" id="frmEquipmentAssetPartialSubmitted" value="frmEquipmentAssetPartialSubmitted">
                	
                    <div class="form-group">
                        <input type="text" name="txtPartialEquip" id="txtPartialEquip" value="<%= EquipPartialAssetSerial %>" class="form-control-partial-search" placeholder="Search for equipment by PARTIAL serial number or asset tag">
                        <i id="searchIcon" class="fa fa-search fa-2x"></i>
                    </div>
                                            
                    <div class="form-group">
                        <input type="submit" id="btnSubmit" name="btnSubmit" class="btn btn-primary" value="Search By Partial Asset Tag, Serial #">
                    </div>
                
                </form>
		    </li>
		  </ul>
		  
		  <h4 <% If frmCustomerIDSubmitted <> "" Then Response.Write("class='active'") %>>Search By Customer</h4>
		  
		  <ul>
		  
		  	<li style="margin-bottom:15px;">
		  		<strong>Note</strong>: This search box will provide matching search terms. Please select/click the one you want.
		  	</li>
		  
		    <li>
		    	<form id="frmCustomerID" name="frmCustomerID" action="<%= BaseURL %>equipment/equipment/findEquipment.asp" method="post">
		    	
		    		<input type="hidden" name="frmCustomerIDSubmitted" id="frmCustomerIDSubmitted" value="frmCustomerIDSubmitted">
		    		
		    		<!-- select customer record !-->
					<input id="txtCustomerID" name="txtCustomerID">
					<input type="hidden" id="txtCustomerIDToPass" name="txtCustomerIDToPass" value="<%= EquipCustomerID %>">
					<i id="searchIcon" class="fa fa-search fa-2x"></i>
					
					<div class="form-group" style="margin-top:15px;">
						<input type="submit" id="btnSubmit" name="btnSubmit" class="btn btn-primary" value="Submit">
					</div>
					<!-- eof select customer record !-->
				</form>
		    </li>
		  </ul>
		  
		  <h4 <% If frmClassManfBrandModelSubmitted <> "" Then Response.Write("class='active'") %>>Search by Class, Manufacturer, Brand, or Model</h4>
		  
		  <ul>
		    <li>
	                <form id="frmClassManfBrandModel" name="frmClassManfBrandModel" action="<%= BaseURL %>equipment/equipment/findEquipment.asp" method="post">
	                
	                <input type="hidden" name="frmClassManfBrandModelSubmitted" id="frmClassManfBrandModelSubmitted" value="frmClassManfBrandModelSubmitted">
 
                    <div class="form-group">
					  	<select class="form-control select50" name="selClassIntRecID" id="selClassIntRecID">
					  			<option value="">Select Class of Equipment</option>
						      	<% 'Get all equipment classes
						      	  	SQLEquipClasses = "SELECT * FROM EQ_Classes ORDER BY Class ASC"
		
									Set cnnEquipClasses = Server.CreateObject("ADODB.Connection")
									cnnEquipClasses.open (Session("ClientCnnString"))
									Set rsEquipClasses = Server.CreateObject("ADODB.Recordset")
									rsEquipClasses.CursorLocation = 3 
									Set rsEquipClasses = cnnEquipClasses.Execute(SQLEquipClasses)
									If not rsEquipClasses.EOF Then
										Do
											If EquipClassIntRecID <> "" Then
												If cInt(EquipClassIntRecID) = cInt(rsEquipClasses("InternalRecordIdentifier")) Then
													Response.Write("<option value='" & rsEquipClasses("InternalRecordIdentifier") & "' selected='selected'>" & rsEquipClasses("Class") & "</option>")
												Else
													Response.Write("<option value='" & rsEquipClasses("InternalRecordIdentifier") & "'>" & rsEquipClasses("Class") & "</option>")
												End If
											Else
												Response.Write("<option value='" & rsEquipClasses("InternalRecordIdentifier") & "'>" & rsEquipClasses("Class") & "</option>")
											End If
												
											rsEquipClasses.movenext
										Loop until rsEquipClasses.eof
									End If
									set rsEquipClasses = Nothing
									cnnEquipClasses.close
									set cnnEquipClasses = Nothing
								%>
						</select>

					  	<select class="form-control select50" name="selManfIntRecID" id="selManfIntRecID">
					  			<option value="">Select Manufacturer of Equipment</option>
						      	<% 'Get all equipment classes
						      	  	SQLEquipManufacturers = "SELECT * FROM EQ_Manufacturers ORDER BY ManufacturerName ASC"
		
									Set cnnEquipManufacturers = Server.CreateObject("ADODB.Connection")
									cnnEquipManufacturers.open (Session("ClientCnnString"))
									Set rsEquipManufacturers = Server.CreateObject("ADODB.Recordset")
									rsEquipManufacturers.CursorLocation = 3 
									Set rsEquipManufacturers = cnnEquipManufacturers.Execute(SQLEquipManufacturers)
									If not rsEquipManufacturers.EOF Then
										Do
											If EquipManfIntRecID <> "" Then
												If cInt(EquipManfIntRecID) = cInt(rsEquipManufacturers("InternalRecordIdentifier")) Then
													Response.Write("<option value='" & rsEquipManufacturers("InternalRecordIdentifier") & "' selected='selected'>" & rsEquipManufacturers("ManufacturerName") & "</option>")
												Else
													Response.Write("<option value='" & rsEquipManufacturers("InternalRecordIdentifier") & "'>" & rsEquipManufacturers("ManufacturerName") & "</option>")
												End If
											Else
												Response.Write("<option value='" & rsEquipManufacturers("InternalRecordIdentifier") & "'>" & rsEquipManufacturers("ManufacturerName") & "</option>")
											End If
											
											rsEquipManufacturers.movenext
										Loop until rsEquipManufacturers.eof
									End If
									set rsEquipManufacturers = Nothing
									cnnEquipManufacturers.close
									set cnnEquipManufacturers = Nothing
								%>
						</select>
                    </div>
                    
  
                    <div class="form-group">
                    
					  	<select class="form-control select50" name="selBrandIntRecID" id="selBrandIntRecID">
					  			<option value="">Select Brand of Equipment</option>
						      	<% 'Get all equipment brands
						      	  	SQLEquipBrands = "SELECT * FROM EQ_Brands ORDER BY Brand ASC"
		
									Set cnnEquipBrands = Server.CreateObject("ADODB.Connection")
									cnnEquipBrands.open (Session("ClientCnnString"))
									Set rsEquipBrands = Server.CreateObject("ADODB.Recordset")
									rsEquipBrands.CursorLocation = 3 
									Set rsEquipBrands = cnnEquipBrands.Execute(SQLEquipBrands)
									If not rsEquipBrands.EOF Then
										Do
											If EquipBrandIntRecID <> "" Then
												If cInt(EquipBrandIntRecID) = cInt(rsEquipBrands("InternalRecordIdentifier")) Then
													Response.Write("<option value='" & rsEquipBrands("InternalRecordIdentifier") & "' selected='selected'>" & rsEquipBrands("Brand") & "</option>")
												Else
													Response.Write("<option value='" & rsEquipBrands("InternalRecordIdentifier") & "'>" & rsEquipBrands("Brand") & "</option>")
												End If
											Else
												Response.Write("<option value='" & rsEquipBrands("InternalRecordIdentifier") & "'>" & rsEquipBrands("Brand") & "</option>")
											End If
												
											rsEquipBrands.movenext
										Loop until rsEquipBrands.eof
									End If
									set rsEquipBrands = Nothing
									cnnEquipBrands.close
									set cnnEquipBrands = Nothing
								%>
						</select>

					  	<select class="form-control select50" name="selModelIntRecID" id="selModelIntRecID">
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
											If EquipModelIntRecID <> "" Then
												If cInt(EquipModelIntRecID) = cInt(rsEquipModels("InternalRecordIdentifier")) Then
													Response.Write("<option value='" & rsEquipModels("InternalRecordIdentifier") & "' selected='selected'>" & rsEquipModels("Model") & "</option>")
												Else
													Response.Write("<option value='" & rsEquipModels("InternalRecordIdentifier") & "'>" & rsEquipModels("Model") & "</option>")
												End If
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
                    </div>
                    
	                <div class="form-group">
	                
						<% If EquipShowOnlyAvailForPlacement1 = 1 Then %>
							<input type="checkbox" checked="checked" id="chkShowOnlyAvailForPlacement1" name="chkShowOnlyAvailForPlacement1">
						<% Else %>
							<input type="checkbox" id="chkShowOnlyAvailForPlacement1" name="chkShowOnlyAvailForPlacement1">		    
						<% End If %>

	                    &nbsp;&nbsp;Only Show Equipment Available for Placement
	                </div>                    
                    
                    <div class="form-group">
                        <input type="submit" id="btnSubmit" name="btnSubmit" class="btn btn-primary" value="Search by Class, Manufacturer, Brand, or Model">
                    </div>
                
                </form>
	    
		    </li>
		  </ul>
		  
		  <h4 <% If frmGroupSubmitted <> "" Then Response.Write("class='active'") %>>Search By Equipment Group</h4>
		  
		  <ul>
		    <li>
                <form id="frmGroup" name="frmGroup" action="<%= BaseURL %>equipment/equipment/findEquipment.asp" method="post">
                
                	<input type="hidden" name="frmGroupSubmitted" id="frmGroupSubmitted" value="frmGroupSubmitted">
            	
	                <div class="form-group">
					  	<select class="form-control select50" name="selGroupIntRecID" id="selGroupIntRecID">
					  			<option value="">Select Group of Equipment</option>
						      	<% 'Get all equipment groups
						      	  	SQLEquipGroups = "SELECT * FROM EQ_Groups ORDER BY GroupName ASC"
		
									Set cnnEquipGroups = Server.CreateObject("ADODB.Connection")
									cnnEquipGroups.open (Session("ClientCnnString"))
									Set rsEquipGroups = Server.CreateObject("ADODB.Recordset")
									rsEquipGroups.CursorLocation = 3 
									Set rsEquipGroups = cnnEquipGroups.Execute(SQLEquipGroups)
									If not rsEquipGroups.EOF Then
										Do
											If EquipGroupIntRecID <> "" Then
												If cInt(EquipGroupIntRecID) = cInt(rsEquipGroups("InternalRecordIdentifier")) Then
													Response.Write("<option value='" & rsEquipGroups("InternalRecordIdentifier") & "' selected='selected'>" & rsEquipGroups("GroupName") & "</option>")
												Else
													Response.Write("<option value='" & rsEquipGroups("InternalRecordIdentifier") & "'>" & rsEquipGroups("GroupName") & "</option>")
												End If
											Else
												Response.Write("<option value='" & rsEquipGroups("InternalRecordIdentifier") & "'>" & rsEquipGroups("GroupName") & "</option>")
											End If
												
											rsEquipGroups.movenext
										Loop until rsEquipGroups.eof
									End If
									set rsEquipGroups = Nothing
									cnnEquipGroups.close
									set cnnEquipGroups = Nothing
								%>
						</select>
	                </div>

	                <div class="form-group">
	                
						<% If EquipShowOnlyAvailForPlacement2 = 1 Then %>
							<input type="checkbox" checked="checked" id="chkShowOnlyAvailForPlacement2" name="chkShowOnlyAvailForPlacement2">
						<% Else %>
							<input type="checkbox" id="chkShowOnlyAvailForPlacement2" name="chkShowOnlyAvailForPlacement2">		    
						<% End If %>

	                    &nbsp;&nbsp;Only Show Equipment Available for Placement
	                </div>    
	                	                                    
	                <div class="form-group">
	                    <input type="submit" id="btnSubmit" name="btnSubmit" class="btn btn-primary" value="Search By Equipment Group">
	                </div>
	            
	            </form>
		    </li>
		  </ul>
		  
		</section>   
		

		<div id="equipmentSearchResults">
		
		<% If EquipCustomerID = "" Then %>
			
			<% If frmEquipmentAssetFullSubmitted <> "" OR frmEquipmentAssetPartialSubmitted <> "" OR frmClassManfBrandModelSubmitted <> "" OR frmGroupSubmitted <> "" Then %>
			
				<div class="row searchrow">
					<h1 class="search-header">Equipment Search Results</h1>

					<% If frmEquipmentAssetFullSubmitted <> "" Then %>
						<h2 class="search-subheader">Equipment With Serial # or Asset Tag "<strong><%= EquipFullAssetSerial %></strong>"</h2>
					<% End If %>
					
					<% If frmEquipmentAssetPartialSubmitted <> "" Then %>
						<h2 class="search-subheader">Equipment With Serial # or Asset Tag Matching "<strong><%= EquipPartialAssetSerial %></strong>"</h2>
					<% End If %>
					
					<% If frmClassManfBrandModelSubmitted <> "" Then %>
					
						<% If EquipShowOnlyAvailForPlacement1 = 1 AND EquipClassIntRecID <> "" Then %>
							<h2 class="search-subheader">Available Equipment In The Class <strong><%= GetClassNameByIntRecID(EquipClassIntRecID) %></strong></h2>
						<% Else %>
							<% If EquipClassIntRecID <> "" Then %>
								<h2 class="search-subheader">Equipment In The Class <strong><%= GetClassNameByIntRecID(EquipClassIntRecID) %></strong></h2>
							<% End If %>
						<% End If %>
					

						<% If EquipShowOnlyAvailForPlacement1 = 1 AND EquipManfIntRecID <> "" Then %>
							<h2 class="search-subheader">Available Equipment with Manufacturer <strong><%= GetManufacturerNameByIntRecID(EquipManfIntRecID) %></strong></h2>
						<% Else %>
							<% If EquipManfIntRecID <> "" Then %>
								<h2 class="search-subheader">Equipment with Manufacturer <strong><%= GetManufacturerNameByIntRecID(EquipManfIntRecID) %></strong></h2>
							<% End If %>
						<% End If %>


						<% If EquipShowOnlyAvailForPlacement1 = 1 AND EquipBrandIntRecID <> "" Then %>
							<h2 class="search-subheader">Available Equipment with Brand <strong><%= GetBrandNameByIntRecID(EquipBrandIntRecID) %></strong></h2>
						<% Else %>
							<% If EquipBrandIntRecID <> "" Then %>
								<h2 class="search-subheader">Equipment with Brand <strong><%= GetBrandNameByIntRecID(EquipBrandIntRecID) %></strong></h2>
							<% End If %>
						<% End If %>


						<% If EquipShowOnlyAvailForPlacement1 = 1 AND EquipModelIntRecID <> "" Then %>
							<h2 class="search-subheader">Available Equipment with Model <strong><%= GetModelNameByIntRecID(EquipModelIntRecID) %></strong></h2>
						<% Else %>
							<% If EquipModelIntRecID <> "" Then %>
								<h2 class="search-subheader">Equipment with Model <strong><%= GetModelNameByIntRecID(EquipModelIntRecID) %></strong></h2>
							<% End If %>
						<% End If %>
					
					<% End If %>
					
					<% If frmGroupSubmitted <> "" Then %>
					
						<% If EquipShowOnlyAvailForPlacement2 = 1 Then %>
							<h2 class="search-subheader">Available Equipment In The Group <strong><%= GetGroupNameByIntRecID(EquipGroupIntRecID) %></strong></h2>
						<% Else %>
							<h2 class="search-subheader">Equipment In The Group <strong><%= GetGroupNameByIntRecID(EquipGroupIntRecID) %></strong></h2>
						<% End If %>
					
					<% End If %>
					
					
					<div class="input-group narrow-results"> 
						<span class="input-group-addon">Narrow Results</span>
						<input id="filter-equipment" type="text" class="form-control filter-search-width" placeholder="Type here...">
					</div>
				</div>
			
				<div class="table-responsive">
			            <table class="table table-striped table-condensed table-hover table-bordered sortable searchable-equipment">
			              <thead>
			                <tr>
								<th>Customer</th>
								<th>Acct. #</th>
							  	<th>Description/Type</th>
							  	<th>Status</th>
							  	<th>Movement Code</th>
							  	<th>Frequency</th>
							  	<th>Install Date</th>
							  	<th>Equip. Value</th>
							  	<th>Serial #</th>
							  	<th>Asset #</th>
			                </tr>
			              </thead>
			              <tbody class='searchable'>
			              
							<%
							
							If frmEquipmentAssetFullSubmitted <> "" Then
							
								SQLCustomerEquipment = "SELECT * FROM EQ_CustomerEquipment INNER JOIN "
								SQLCustomerEquipment = SQLCustomerEquipment & " EQ_Equipment ON EQ_Equipment.InternalRecordIdentifier = EQ_CustomerEquipment.EquipIntRecID "
								SQLCustomerEquipment = SQLCustomerEquipment & " WHERE EQ_Equipment.InternalRecordIdentifier = '" & EquipFullAssetSerial & "' "

							ElseIf frmEquipmentAssetPartialSubmitted <> "" Then
							
								SQLCustomerEquipment = "SELECT * FROM EQ_CustomerEquipment INNER JOIN "
								SQLCustomerEquipment = SQLCustomerEquipment & " EQ_Equipment ON EQ_Equipment.InternalRecordIdentifier = EQ_CustomerEquipment.EquipIntRecID "
								SQLCustomerEquipment = SQLCustomerEquipment & " WHERE EQ_Equipment.AssetTag1 LIKE '%" & EquipPartialAssetSerial & "%' OR "
								SQLCustomerEquipment = SQLCustomerEquipment & " EQ_Equipment.AssetTag2 LIKE '%" & EquipPartialAssetSerial & "%' OR "
								SQLCustomerEquipment = SQLCustomerEquipment & " EQ_Equipment.AssetTag3 LIKE '%" & EquipPartialAssetSerial & "%' OR "
								SQLCustomerEquipment = SQLCustomerEquipment & " EQ_Equipment.AssetTag4 LIKE '%" & EquipPartialAssetSerial & "%' OR "
								SQLCustomerEquipment = SQLCustomerEquipment & " EQ_Equipment.SerialNumber LIKE '%" & EquipPartialAssetSerial & "%' "
								SQLCustomerEquipment = SQLCustomerEquipment & " ORDER BY EQ_Equipment.InternalRecordIdentifier "
							
							ElseIf frmClassManfBrandModelSubmitted <> "" Then
							
								SQLCustomerEquipment = "SELECT * FROM EQ_CustomerEquipment INNER JOIN "
								SQLCustomerEquipment = SQLCustomerEquipment & " EQ_Equipment ON EQ_Equipment.InternalRecordIdentifier = EQ_CustomerEquipment.EquipIntRecID "
								
								If EquipClassIntRecID <> "" OR EquipManfIntRecID <> "" OR EquipBrandIntRecID <> "" OR EquipModelIntRecID <> "" Then
									SQLCustomerEquipment = SQLCustomerEquipment & " INNER JOIN "
									SQLCustomerEquipment = SQLCustomerEquipment & " EQ_Models ON EQ_Equipment.ModelIntRecID = EQ_Models.InternalRecordIdentifier "
								End If
								
								If EquipClassIntRecID <> "" Then
									SQLCustomerEquipment = SQLCustomerEquipment & " INNER JOIN "
									SQLCustomerEquipment = SQLCustomerEquipment & " EQ_Classes ON EQ_Models.ClassIntRecID = EQ_Classes.InternalRecordIdentifier "
								End If

								If EquipManfIntRecID <> "" Then
									SQLCustomerEquipment = SQLCustomerEquipment & " INNER JOIN "
									SQLCustomerEquipment = SQLCustomerEquipment & " EQ_Manufacturers ON EQ_Models.ManufacIntRecID = EQ_Manufacturers.InternalRecordIdentifier "
								End If
	
								If EquipBrandIntRecID <> "" Then
									SQLCustomerEquipment = SQLCustomerEquipment & " INNER JOIN "
									SQLCustomerEquipment = SQLCustomerEquipment & " EQ_Brands ON EQ_Models.BrandIntRecID = EQ_Brands.InternalRecordIdentifier "
								End If
							
								If EquipShowOnlyAvailForPlacement1 = 1 Then
									SQLCustomerEquipment = SQLCustomerEquipment & " INNER JOIN "
									SQLCustomerEquipment = SQLCustomerEquipment & " EQ_StatusCodes ON EQ_Equipment.StatusCodeIntRecID = EQ_StatusCodes.InternalRecordIdentifier "
								End If
								
								SQLCustomerEquipment = SQLCustomerEquipment & " WHERE "
								
								If EquipClassIntRecID <> "" Then
									SQLCustomerEquipment = SQLCustomerEquipment & " EQ_Classes.InternalRecordIdentifier = " & EquipClassIntRecID & " "
								End If

								If EquipManfIntRecID <> "" Then
									If EquipClassIntRecID <> "" Then
										SQLCustomerEquipment = SQLCustomerEquipment & " AND EQ_Manufacturers .InternalRecordIdentifier = " & EquipManfIntRecID & " "
									Else
										SQLCustomerEquipment = SQLCustomerEquipment & " EQ_Manufacturers .InternalRecordIdentifier = " & EquipManfIntRecID & " "
									End If
								End If

								If EquipBrandIntRecID <> "" Then
									If EquipClassIntRecID <> "" OR EquipManfIntRecID <> "" Then
										SQLCustomerEquipment = SQLCustomerEquipment & " AND EQ_Brands.InternalRecordIdentifier = " & EquipBrandIntRecID & " "
									Else
										SQLCustomerEquipment = SQLCustomerEquipment & " EQ_Brands.InternalRecordIdentifier = " & EquipBrandIntRecID & " "
									End If
								End If

								If EquipModelIntRecID <> "" Then
									If EquipClassIntRecID <> "" OR EquipManfIntRecID <> "" OR EquipBrandIntRecID <> "" Then
										SQLCustomerEquipment = SQLCustomerEquipment & " AND EQ_Models.InternalRecordIdentifier = " & EquipModelIntRecID & " "
									Else
										SQLCustomerEquipment = SQLCustomerEquipment & " EQ_Models.InternalRecordIdentifier = " & EquipModelIntRecID & " "
									End If
								End If
								
								
								If EquipShowOnlyAvailForPlacement1 = 1 Then
									If EquipClassIntRecID <> "" OR EquipManfIntRecID <> "" OR EquipBrandIntRecID <> "" OR EquipModelIntRecID <> "" Then
										SQLCustomerEquipment = SQLCustomerEquipment & " AND EQ_StatusCodes.statusAvailableForPlacement = 1 "
									Else
										SQLCustomerEquipment = SQLCustomerEquipment & " EQ_StatusCodes.statusAvailableForPlacement = 1 "
									End If
								End If
								
								SQLCustomerEquipment = SQLCustomerEquipment & " ORDER BY EQ_Equipment.InternalRecordIdentifier "
							
							
							ElseIf frmGroupSubmitted <> "" Then
							
								SQLCustomerEquipment = "SELECT * FROM EQ_CustomerEquipment  "
								SQLCustomerEquipment = SQLCustomerEquipment & " INNER JOIN EQ_Equipment ON EQ_Equipment.InternalRecordIdentifier = EQ_CustomerEquipment.EquipIntRecID "
								SQLCustomerEquipment = SQLCustomerEquipment & " INNER JOIN EQ_Models ON EQ_Equipment.ModelIntRecID = EQ_Models.InternalRecordIdentifier "
								SQLCustomerEquipment = SQLCustomerEquipment & " INNER JOIN EQ_Groups ON EQ_Models.GroupIntRecID = EQ_Groups.InternalRecordIdentifier "
								
								
								If EquipShowOnlyAvailForPlacement2 = 1 Then
									SQLCustomerEquipment = SQLCustomerEquipment & " INNER JOIN "
									SQLCustomerEquipment = SQLCustomerEquipment & " EQ_StatusCodes ON EQ_Equipment.StatusCodeIntRecID = EQ_StatusCodes.InternalRecordIdentifier "
								End If
								
								SQLCustomerEquipment = SQLCustomerEquipment & " WHERE EQ_Groups.InternalRecordIdentifier = " & EquipGroupIntRecID
								
								If EquipShowOnlyAvailForPlacement2 = 1 Then
									SQLCustomerEquipment = SQLCustomerEquipment & " AND EQ_StatusCodes.statusAvailableForPlacement = 1 "
								End If
								
								SQLCustomerEquipment = SQLCustomerEquipment & " ORDER BY EQ_Equipment.InternalRecordIdentifier "
							
							
							End If			
							
							Response.write("SQLCustomerEquipment : " & SQLCustomerEquipment & "<br>")				
						
							Set cnnCustomerEquipment = Server.CreateObject("ADODB.Connection")
							cnnCustomerEquipment.open (Session("ClientCnnString"))
							Set rsCustomerEquipment = cnnCustomerEquipment.Execute(SQLCustomerEquipment)
					
							If NOT rsCustomerEquipment.EOF Then
			
								Do While Not rsCustomerEquipment.EOF
			      
									EquipIntRecID = rsCustomerEquipment("EquipIntRecID") 
							      	CustID = rsCustomerEquipment("CustID")
							      	RecordSource = rsCustomerEquipment("RecordSource")
									InstallDate = rsCustomerEquipment("InstallDate")
									StatusCodeIntRecID = rsCustomerEquipment("StatusCodeIntRecID")
									
									SQLEquipStatusCode = "SELECT * FROM EQ_StatusCodes WHERE InternalRecordIdentifier = " & StatusCodeIntRecID
										
									Set cnnEquipStatusCode = Server.CreateObject("ADODB.Connection")
									cnnEquipStatusCode.open (Session("ClientCnnString"))
									Set rsEquipStatusCode = cnnEquipStatusCode.Execute(SQLEquipStatusCode)
									
									If NOT rsEquipStatusCode.EOF Then
										InstallType = rsEquipStatusCode("statusBackendSystemCode")
										InstallTypeFullName = rsEquipStatusCode("statusDesc")
									Else
										InstallType = ""
										InstallTypeFullName = ""
									End If
															
									
									If InstallType = "R" then
									
										RentalFrequencyType = rsCustomerEquipment("RentalFrequencyType")
										
										Select Case RentalFrequencyType
										Case "D"
											RentalFrequencyFullName = "DAYS"
										Case "M"
											RentalFrequencyFullName = "MONTH(S)"
										Case "Y"
											RentalFrequencyFullName = "YEAR(S)"
										End Select
										
										RentalFrequencyNumber = rsCustomerEquipment("RentalFrequencyNumber")
										RentAmt = rsCustomerEquipment("RentAmt")
										
										If RentAmt <> "" Then
											RentAmt = FormatCurrency(RentAmt,2)
										Else
											RentAmt = ""
										End If
										
									Else
										RentalFrequencyFullName = ""
										RentalFrequencyType = ""
										RentalFrequencyNumber = ""
										RentAmt = ""
									End If
									
									
									MovementCodeIntRecID = rsCustomerEquipment("MovementCodeIntRecID")
									
									If MovementCodeIntRecID <> "" Then
									
										SQLEquipMovementCode = "SELECT * FROM EQ_MovementCodes WHERE InternalRecordIdentifier = " & MovementCodeIntRecID
											
										Set cnnEquipMovementCode = Server.CreateObject("ADODB.Connection")
										cnnEquipMovementCode.open (Session("ClientCnnString"))
										Set rsEquipMovementCode = cnnEquipMovementCode.Execute(SQLEquipMovementCode)
										
										If NOT rsEquipMovementCode.EOF Then
											MovementCode = rsEquipMovementCode("movementCode")
											MovementCodeDesc = rsEquipMovementCode("movementCodeDesc")
										Else
											MovementCode = ""
											MovementCodeDesc = ""
										End If
										
									Else
										MovementCode = ""
										MovementCodeDesc = ""
									End If
			
									
									SerialNumber = rsCustomerEquipment("SerialNumber")
									PurchaseCost = rsCustomerEquipment("PurchaseCost")
									
									If PurchaseCost <> "" then
										TotalPurchaseCost = TotalPurchaseCost + PurchaseCost
										PurchaseCost = FormatCurrency(PurchaseCost,2)
									End If
									
									ModelIntRecID = rsCustomerEquipment("ModelIntRecID")
									
									If ModelIntRecID <> 0 Then
										BrandName = GetBrandNameByModelIntRecID(ModelIntRecID)
									Else
										BrandName = ""
									End If
									
									AssetTag1 = rsCustomerEquipment("AssetTag1")
									Description = "DESC NEEDED"
									Description  = GetModelNameByIntRecID(rsCustomerEquipment("ModelIntRecID"))
									
									ModelCount = GetTotalNumberOfModelsForCustomer(CustID,ModelIntRecID)
			
						        %>
									<!-- table line !-->
									<tr>
										
										<td><a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= GetCustNameByCustNum(CustID) %></td>
										
										<% If InStr(CustID,"<") OR InStr(CustID,">") Then %>
											<% 
												CustID = Replace(CustID, "<", "&#60;")
												CustID = Replace(CustID, ">", "&#62;")
											%>
										<% End If %>
										
										<td><a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= CustID %></a></td>
										<% If BrandName <> "" Then %>
											<td><a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= UCASE(BrandName) %>&nbsp;<%= Description %></a></td>
										<% Else %>
											<td><a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= Description %></a></td>
										<% End If %>
										<td><a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= InstallTypeFullName %></a></td>
										<td><a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= MovementCode %> - <%= MovementCodeDesc %></a></td>
										<% If InstallType = "R" Then %>
											<td><a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= RentAmt %>&nbsp;/&nbsp;<%= RentalFrequencyNumber %>&nbsp;<%= RentalFrequencyFullName %></a></td>
										<% Else %>
											<td>&nbsp;</td>
										<% End If %>
										<td align="right"><a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= InstallDate %></a></td>
										<% If PurchaseCost <> "" Then %>
											<td align="right"><a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= FormatCurrency(PurchaseCost,0) %></a></td>
										<% End If %>
										<td align="center"><a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= SerialNumber %></a></td>
										<td align="center"><a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= AssetTag1 %></a></td>
								   	</tr>
								<%
									rsCustomerEquipment.movenext
								loop
							End If
							set rsCustomerEquipment= Nothing
							cnnCustomerEquipment.close
							set cnnCustomerEquipment = Nothing
				            %>
						</tbody>
					</table>
				</div>
			<% End If %>
		<% End If %>
		</div>
	
	
	
	
	
		
		
		<div id="customerSearchResults">
		
			<%
			If EquipCustomerID <> "" Then
			
				CustIDPassed = EquipCustomerID
				CustName = GetCustNameByCustNum(CustIDPassed)
				TotalEquipmentValue = GetTotalValueOfEquipmentForCustomer(CustIDPassed)

				%>

				<h1 class="search-header">Equipment Search Results for <%= CustName %>, Acct. <%= CustIDPassed %></h1>
				<div align="right">
					<h3 style="margin-top:20px;margin-bottom:-20px;">Total Value <span style="color:green; font-weight:bold"><%= FormatCurrency(TotalEquipmentValue,2) %></span></h3>
				</div>
				<%
				
				Set rsCustomerEquipmentByClass = Server.CreateObject("ADODB.Recordset")
				rsCustomerEquipmentByClass.CursorLocation = 3 
			
				Set rsCustomerEquipment = Server.CreateObject("ADODB.Recordset")
				rsCustomerEquipment.CursorLocation = 3 
				
				
				Set rsEquipStatusCode = Server.CreateObject("ADODB.Recordset")
				rsEquipStatusCode.CursorLocation = 3 
				
					
				SQLCustomerEquipmentByClass = "SELECT EQ_Classes.Class, EQ_Classes.InternalRecordIdentifier, SUM(EQ_Equipment.PurchaseCost) AS Expr1 "
				SQLCustomerEquipmentByClass  = SQLCustomerEquipmentByClass & " FROM EQ_CustomerEquipment INNER JOIN "
				SQLCustomerEquipmentByClass  = SQLCustomerEquipmentByClass & " EQ_Equipment ON EQ_CustomerEquipment.EquipIntRecID = EQ_Equipment.InternalRecordIdentifier INNER JOIN "
				SQLCustomerEquipmentByClass  = SQLCustomerEquipmentByClass & " EQ_Models ON EQ_Equipment.ModelIntRecID = EQ_Models.InternalRecordIdentifier INNER JOIN "
				SQLCustomerEquipmentByClass  = SQLCustomerEquipmentByClass & " EQ_Classes ON EQ_Models.ClassIntRecID = EQ_Classes.InternalRecordIdentifier "
				SQLCustomerEquipmentByClass  = SQLCustomerEquipmentByClass & " WHERE        (EQ_CustomerEquipment.CustID = " & CustIDPassed & ") "
				SQLCustomerEquipmentByClass  = SQLCustomerEquipmentByClass & " GROUP BY EQ_Classes.Class, EQ_Classes.InternalRecordIdentifier "
				SQLCustomerEquipmentByClass  = SQLCustomerEquipmentByClass & " ORDER BY Expr1 DESC"
					
				Set cnnCustomerEquipmentByClass = Server.CreateObject("ADODB.Connection")
				cnnCustomerEquipmentByClass.open (Session("ClientCnnString"))
				Set rsCustomerEquipmentByClass = cnnCustomerEquipmentByClass.Execute(SQLCustomerEquipmentByClass)
				
				If NOT rsCustomerEquipmentByClass.EOF Then
				
					Do While NOT rsCustomerEquipmentByClass.EOF
					
						ClassName = rsCustomerEquipmentByClass("Class")
						ClassIntRecID = rsCustomerEquipmentByClass("InternalRecordIdentifier")
						ClassTotalEquipValue = rsCustomerEquipmentByClass("Expr1")
				
				
						%>	
						<h3><%= ClassName %>&nbsp;<span style="color:green;"><%= FormatCurrency(ClassTotalEquipValue,2) %></span></h3>
						<table class="table table-condensed table-hover large-table">			
							<thead>
							  <tr style="background-color: #EEE;">
							  	<th style="width: 3%;">+</th>
							  	<th style="width: 25%;">Description/Type</th>
							  	<th>Status</th>
							  	<th>Frequency</th>
							  	<th style="text-align: center;">Install Date</th>
							  	<th style="text-align: center;">Equip. Value</th>
							  	<th style="text-align: center;">Serial #</th>
							  	<th style="text-align: center;">Asset #</th>
							  </tr>
							</thead>
							<tbody>
							
							<%	
							TotalPurchaseCost = 0 
							
							SQLCustomerEquipment = " SELECT        EQ_Equipment.ModelIntRecID, MAX(EQ_Equipment.PurchaseCost) AS purchsum "
							SQLCustomerEquipment = SQLCustomerEquipment & " FROM            EQ_CustomerEquipment INNER JOIN "
							SQLCustomerEquipment = SQLCustomerEquipment & " EQ_Equipment ON EQ_Equipment.InternalRecordIdentifier = EQ_CustomerEquipment.EquipIntRecID INNER JOIN "
							SQLCustomerEquipment = SQLCustomerEquipment & " EQ_Models ON EQ_Equipment.ModelIntRecID = EQ_Models.InternalRecordIdentifier "
							SQLCustomerEquipment = SQLCustomerEquipment & " WHERE        (EQ_CustomerEquipment.CustID = " & CustIDPassed & ") AND (EQ_Models.ClassIntRecID = " & ClassIntRecID & ") "
							SQLCustomerEquipment = SQLCustomerEquipment & " GROUP BY EQ_Equipment.ModelIntRecID "
							SQLCustomerEquipment = SQLCustomerEquipment & " ORDER BY purchsum DESC "		
													
							Set cnnCustomerEquipment = Server.CreateObject("ADODB.Connection")
							cnnCustomerEquipment.open (Session("ClientCnnString"))
							Set rsCustomerEquipment = cnnCustomerEquipment.Execute(SQLCustomerEquipment)
							
							'***************************************************************************************
							'BUILD THE MASTER ORDER BY CLAUSE HERE
							'****************************************************************************************
							If Not rsCustomerEquipment.EOF Then
							
								EqpOrderByClauseCustom = " ORDER BY CASE ModelIntRecID "
								SortCount = 0
							
								Do While NOT rsCustomerEquipment.EOF
							
									EqpOrderByClauseCustom = EqpOrderByClauseCustom & " WHEN " & rsCustomerEquipment("ModelIntRecID") & " THEN " & Trim(SortCount) & " "
									SortCount = SortCount + 1
							
									rsCustomerEquipment.MoveNext
								Loop
								
								EqpOrderByClauseCustom = EqpOrderByClauseCustom & " END "
							
							End If
							
							'Response.write(EqpOrderByClauseCustom & "<br>")
			
							SQLCustomerEquipment = "SELECT * FROM EQ_CustomerEquipment INNER JOIN "
							SQLCustomerEquipment = SQLCustomerEquipment & " EQ_Equipment ON EQ_Equipment.InternalRecordIdentifier = EQ_CustomerEquipment.EquipIntRecID "
							SQLCustomerEquipment = SQLCustomerEquipment & " INNER JOIN "
							SQLCustomerEquipment = SQLCustomerEquipment & " EQ_Models ON EQ_Equipment.ModelIntRecID = EQ_Models.InternalRecordIdentifier "
							SQLCustomerEquipment = SQLCustomerEquipment & " WHERE "
							SQLCustomerEquipment = SQLCustomerEquipment & " (EQ_CustomerEquipment.CustID = " & CustIDPassed & ") AND (EQ_Models.ClassIntRecID = " & ClassIntRecID & ") "
							SQLCustomerEquipment = SQLCustomerEquipment & EqpOrderByClauseCustom 
							
							'Response.write(SQLCustomerEquipment & "<br>")
			
							Set rsCustomerEquipment = cnnCustomerEquipment.Execute(SQLCustomerEquipment)
			
							If NOT rsCustomerEquipment.EOF Then
							
								FirstPassOnModel = True
								ModelLoopCounter = 1
							
								Do While NOT rsCustomerEquipment.EOF
								
									EquipIntRecID = rsCustomerEquipment("EquipIntRecID")
									InstallDate = rsCustomerEquipment("InstallDate")
									StatusCodeIntRecID = rsCustomerEquipment("StatusCodeIntRecID")
									
									SQLEquipStatusCode = "SELECT * FROM EQ_StatusCodes WHERE InternalRecordIdentifier = " & StatusCodeIntRecID
										
									Set cnnEquipStatusCode = Server.CreateObject("ADODB.Connection")
									cnnEquipStatusCode.open (Session("ClientCnnString"))
									Set rsEquipStatusCode = cnnEquipStatusCode.Execute(SQLEquipStatusCode)
									
									If NOT rsEquipStatusCode.EOF Then
										InstallType = rsEquipStatusCode("statusBackendSystemCode")
										InstallTypeFullName = rsEquipStatusCode("statusDesc")
									Else
										InstallType = ""
										InstallTypeFullName = ""
									End If
															
									
									If InstallType = "R" then
									
										RentalFrequencyType = rsCustomerEquipment("RentalFrequencyType")
										
										Select Case RentalFrequencyType
										Case "D"
											RentalFrequencyFullName = "DAYS"
										Case "M"
											RentalFrequencyFullName = "MONTH(S)"
										Case "Y"
											RentalFrequencyFullName = "YEAR(S)"
										End Select
										
										RentalFrequencyNumber = rsCustomerEquipment("RentalFrequencyNumber")
										RentAmt = rsCustomerEquipment("RentAmt")
										
										If RentAmt <> "" Then
											RentAmt = FormatCurrency(RentAmt,2)
										Else
											RentAmt = ""
										End If
										
									Else
										RentalFrequencyFullName = ""
										RentalFrequencyType = ""
										RentalFrequencyNumber = ""
										RentAmt = ""
									End If
															
									SerialNumber = rsCustomerEquipment("SerialNumber")
									PurchaseCost = rsCustomerEquipment("PurchaseCost")
									
									If PurchaseCost <> "" then
										TotalPurchaseCost = TotalPurchaseCost + PurchaseCost
										PurchaseCost = FormatCurrency(PurchaseCost,2)
									End If
									
									ModelIntRecID = rsCustomerEquipment("ModelIntRecID")
									
									If ModelIntRecID <> 0 Then
										BrandName = GetBrandNameByModelIntRecID(ModelIntRecID)
									Else
										BrandName = ""
									End If
									
									AssetTag1 = rsCustomerEquipment("AssetTag1")
									Description = "DESC NEEDED"
									Description  = GetModelNameByIntRecID(rsCustomerEquipment("ModelIntRecID"))
									
									ModelCount = GetTotalNumberOfModelsForCustomer(CustIDPassed,ModelIntRecID)
									
									%>
								
									<% If cInt(ModelCount) = 1 Then %>
									
										<tr>
											<td>&nbsp;</td>
											<% If BrandName <> "" Then %>
												<td><a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= UCASE(BrandName) %>&nbsp;<%= Description %></a></td>
											<% Else %>
												<td><a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= Description %></a></td>
											<% End If %>
											<td><a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= InstallTypeFullName %></a></td>
											<% If InstallType = "R" Then %>
												<td><a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= RentAmt %>&nbsp;/&nbsp;<%= RentalFrequencyNumber %>&nbsp;<%= RentalFrequencyFullName %></a></td>
											<% Else %>
												<td>&nbsp;</td>
											<% End If %>
											<td align="right"><a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= InstallDate %></a></td>
											<td align="right"><a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= FormatCurrency(PurchaseCost,0) %></a></td>
											<td align="center"><a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= SerialNumber %></a></td>
											<td align="center"><a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= AssetTag1 %></a></td>
										</tr>
										
									<% ElseIf (cInt(ModelCount) > 1) AND (cInt(ModelLoopCounter) <= cInt(ModelCount)) Then %>
									
										<% If FirstPassOnModel = True Then %>
										
											<% ModelLoopCounter = 1 %>
																		
											<tr class="accordion-toggle">
												<% If BrandName <> "" Then %>
													<td data-toggle="collapse" data-target=".equip<%= ModelIntRecID %>"><i class="fa fa-plus-circle fa-lg" aria-hidden="true" style="color:#009800"></i></td>
													<td colspan="4"><%= UCASE(BrandName) %>&nbsp;<%= Description %>&nbsp;<span class="equip_qty">(<%= ModelCount %>)</span></td>
													<td align="right"><%= FormatCurrency(GetTotalValueOfModelsForCustomer(CustIDPassed,ModelIntRecID),0) %></td>
													<td>&nbsp;</td>
													<td>&nbsp;</td>
												<% Else %>
													<td data-toggle="collapse" data-target=".equip<%= ModelIntRecID %>" colspan="4"><%= Description %>&nbsp;(<%= ModelCount %>)&nbsp;<i class="fa fa-plus-circle" aria-hidden="true" style="color:#009800"></i></td>
													<td align="right"><%= FormatCurrency(GetTotalValueOfModelsForCustomer(CustIDPassed,ModelIntRecID),0) %></td>
													<td>&nbsp;</td>
													<td>&nbsp;</td>										
												<% End If %>
											</tr>		  
											<tr class="collapse equip<%= ModelIntRecID %>" style="background-color:#e5ffe5">
												<td>&nbsp;</td>
												<% If BrandName <> "" Then %>
													<td style="padding-left:20px;"><i class="fa fa-folder-open-o" aria-hidden="true"></i>&nbsp;<a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= UCASE(BrandName) %>&nbsp;<%= Description %></a></td>
												<% Else %>
													<td style="padding-left:20px;"><i class="fa fa-folder-open-o" aria-hidden="true"></i>&nbsp;<a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= Description %></a></td>
												<% End If %>
												<td><a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= InstallTypeFullName %></a></td>
												<% If InstallType = "R" Then %>
													<td><a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= RentAmt %>&nbsp;/&nbsp;<%= RentalFrequencyNumber %>&nbsp;<%= RentalFrequencyFullName %></a></td>
												<% Else %>
													<td>&nbsp;</td>
												<% End If %>
												<td align="right"><a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= InstallDate %></a></td>
												<td align="right"><a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= FormatCurrency(PurchaseCost,0) %></a></td>
												<td align="center"><a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= SerialNumber %></a></td>
												<td align="center"><a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= AssetTag1 %></a></td>		  	
											</tr>
											
											<% FirstPassOnModel = False %>
											<% ModelLoopCounter = ModelLoopCounter + 1 %>
											
										<% Else %>
										
											<tr class="collapse equip<%= ModelIntRecID %>" style="background-color:#e5ffe5">
												<td>&nbsp;</td>
												<% If BrandName <> "" Then %>
													<td style="padding-left:20px;"><i class="fa fa-folder-open-o" aria-hidden="true"></i>&nbsp;<a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= UCASE(BrandName) %>&nbsp;<%= Description %></a></td>
												<% Else %>
													<td style="padding-left:20px;"><i class="fa fa-folder-open-o" aria-hidden="true"></i>&nbsp;<a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= Description %></a></td>
												<% End If %>
												<td><a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= InstallTypeFullName %></a></td>
												<% If InstallType = "R" Then %>
													<td><a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= RentAmt %>&nbsp;/&nbsp;<%= RentalFrequencyNumber %>&nbsp;<%= RentalFrequencyFullName %></a></td>
												<% Else %>
													<td>&nbsp;</td>
												<% End If %>
												<td align="right"><a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= InstallDate %></a></td>
												<td align="right"><a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= FormatCurrency(PurchaseCost,0) %></a></td>
												<td align="center"><a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= SerialNumber %></a></td>
												<td align="center"><a href="<%= BaseURL %>equipment/equipment/editEquipment.asp?i=<%= EquipIntRecID %>"><%= AssetTag1 %></a></td>
											</tr>
											
											<% 
												ModelLoopCounter = ModelLoopCounter + 1
											
												If cInt(ModelLoopCounter) > cInt(ModelCount) Then
													FirstPassOnModel = True
													ModelLoopCounter = 1
												End If
											 %>
											
										<% End If %>
										
									<% End If %>
									<%
								
									rsCustomerEquipment.MoveNext
								
								Loop			
							End If
							
							%>
									  	
						</tbody>
						
						<tfoot>
						  <tr>
						  	<td colspan="2">TOTAL</td>
						  	<td>---</td>
						  	<td>---</td>
						  	<td>---</td>
						  	<td align="right"><%= FormatCurrency(TotalPurchaseCost,2) %></td>
						  	<td>---</td>
						  	<td>---</td>			  	
						  </tr>
						</tfoot>
					</table>
				<%
				
					rsCustomerEquipmentByClass.MoveNext
					Loop
					
					Set rsCustomerEquipment = Nothing
					cnnCustomerEquipment.Close
					Set cnnCustomerEquipment = Nothing
					
				End If
			
				Set rsEquipStatusCode = Nothing
				cnnEquipStatusCode.Close
				Set cnnEquipStatusCode = Nothing
			
				Set rsCustomerEquipmentByClass = Nothing
				cnnCustomerEquipmentByClass.Close
				Set cnnCustomerEquipmentByClass = Nothing
			
			End If
			
			%>	
		</div>
	                                        
    </section>
								

<!--#include file="../../inc/footer-main.asp"-->
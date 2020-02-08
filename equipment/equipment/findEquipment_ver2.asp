<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/InsightFuncs_Equipment.asp"-->
<%
Server.ScriptTimeout = 900000 'Default value

ClientKeyForFileNames = MUV_READ("ClientKeyForFileNames")

EquipIDForDetail = Request.QueryString("EID")

If EquipIDForDetail = "" Then 
	EquipIDForDetail = Request.Form("txtEquipIDToPass")
End If

CustomerIDForDetail = Request.QueryString("CID")

If CustomerIDForDetail = "" Then 
	CustomerIDForDetail = Request.Form("txtCustomerIDToPass")
End If


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
	
	var randomNumberBetween0and100 = Math.floor(Math.random() * 100);
	
	var autocompleteJSONFileURLEquipment = "../../clientfiles/<%= ClientKeyForFileNames %>/autocomplete/equipment_list_<%= ClientKeyForFileNames %>.json?v=" + randomNumberBetween0and100;
	var autocompleteJSONFileURLAccount = "../../clientfiles/<%= ClientKeyForFileNames %>/autocomplete/equipment_list_<%= ClientKeyForFileNames %>.json?v=" + randomNumberBetween0and100;

	var optionsEquipment = {
	  url: autocompleteJSONFileURLEquipment,
	  placeholder: "Search for equipment by FULL serial number or asset tag",
	  getValue: "name",
	  list: {	
        onChooseEvent: function() {
        
            var EquipIntRecID = $("#txtEquipID").getSelectedItemData().code;
            $("#txtEquipIDToPass").val(EquipIntRecID);
            window.location.href = "editEquipment.asp?i=" + EquipIntRecID;
            
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
        
            var custID = $("#txtCustID").getSelectedItemData().code;
            $("#txtCustIDToPass").val(custID);
            window.location.href = "editEquipment.asp?i=" + custID;
            
    	},		  
	    match: {
	      enabled: true
		},
		maxNumberOfElements: 30		
	  },
	  theme: "cat-analysis"
	};
	
	$("#txtCustomerID").easyAutocomplete(optionsCustomer);
	

});



</script>

<%

		'*********************************************************
		' Begin Auto Complete Equipment List
		'*********************************************************

		'SQL = "SELECT InternalRecordIdentifier, ModelIntRecID, SerialNumber, AssetTag1 FROM EQ_Equipment WHERE ModelIntRecID <> '' ORDER BY SerialNumber"
		'Set cnnAutoComplete = Server.CreateObject("ADODB.Connection")
		'cnnAutoComplete.open (Session("ClientCnnString"))
		'Set rsAutoComplete = Server.CreateObject("ADODB.Recordset")
		'rsAutoComplete.CursorLocation = 3 
		'Set rsAutoComplete = cnnAutoComplete.Execute(SQL)
		
		'If not rsAutoComplete.EOF Then
		'strAuto = "["
		'Do While Not rsAutoComplete.EOF
		   ' strAuto = strAuto & "{""name"":""" & GetModelNameByIntRecID(rsAutoComplete("ModelIntRecID")) & " --- " & rsAutoComplete("SerialNumber") & " --- " & rsAutoComplete("AssetTag1") & """, ""code"":""" & rsAutoComplete("InternalRecordIdentifier") & """},"
		    'rsAutoComplete.MoveNext
		'Loop
		'End If
		
		'If right(strAuto,1)= "," Then strAuto = left(strAuto,len(strAuto)-1) 
		
		'strAuto = trim(strAuto) & "]"
		
		'set fs=Server.CreateObject("Scripting.FileSystemObject")
		'set fs2=Server.CreateObject("Scripting.FileSystemObject")
		
		'set tfile=fs.CreateTextFile(Server.MapPath("..\..\..\") & "\clientfiles\"  & ClientKeyForFileNames & "\autocomplete\equipment_list_" & ClientKeyForFileNames & ".json")
		'tfile.WriteLine(strAuto)
		'tfile.close
		'set tfile=nothing
		'set fs=nothing
		
		'Set rsAutoComplete = Nothing
		'cnnAutoComplete.Close
		'Set AutoComplete = nothing

		
		'*********************************************************
		' END Auto Complete Equipment List
		'*********************************************************


%>

 
<style type="text/css">
 	.email-table{
		width:46%;
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
	
	.container{
		max-width:1800px;
		margin-left:0px;
	}

	.narrow-results{
		margin:0px 0px 20px 0px;
	}
	
	#filter{
		width:40%;
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
	
	.searchrow {
		width:1600px !important;
	}

	hr.formdivider { 
	  border: 0; 
	  height: 1px; 
	  background-image: -webkit-linear-gradient(left, #f0f0f0, #8c8b8b, #f0f0f0);
	  background-image: -moz-linear-gradient(left, #f0f0f0, #8c8b8b, #f0f0f0);
	  background-image: -ms-linear-gradient(left, #f0f0f0, #8c8b8b, #f0f0f0);
	  background-image: -o-linear-gradient(left, #f0f0f0, #8c8b8b, #f0f0f0); 
	}
	
	h2.search {
	   font-size:19px;
	}
	
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
	
	#searchIcon {
	    left: auto !important;
	    float: right;
	    margin-right: 45px;
	}	
	
	.vertical-divider {
	  position: absolute;
	  z-index: 10;
	  top: 50%;
	  left: 33.33%;
	  margin: 0;
	  padding: 0;
	  width: auto;
	  height: 50%;
	  line-height: 0;
	  font-size:28px;
	  font-weight:bold;	  
	  text-align:center;
	  text-transform: uppercase;
	  transform: translateX(-50%);
	}
	
	.vertical-divider:before, 
	.vertical-divider:after {
	  position: absolute;
	  left: 50%;
	  content: '';
	  z-index: 9;
	  border-left: 1px solid rgba(34,36,38,.15);
	  border-right: 1px solid rgba(255,255,255,.1);
	  width: 0;
	  height: calc(100% - 1rem);
	}
	
	.row-divided > .vertical-divider {
	  height: calc(50% - 1rem);    
	}
	
	.vertical-divider:before {
	  top: -100%;
	}
	
	.vertical-divider:after {
	  top: auto;
	  bottom: 0;
	}
	
	
	.vertical-divider2 {
	  position: absolute;
	  z-index: 10;
	  top: 50%;
	  left: 66.67%;
	  margin: 0;
	  padding: 0;
	  width: auto;
	  height: 50%;
	  line-height: 0;
	  font-size:28px;
	  font-weight:bold;
	  text-align:center;
	  text-transform: uppercase;
	  transform: translateX(-50%);
	}
	
	.vertical-divider2:before, 
	.vertical-divider2:after {
	  position: absolute;
	  left: 50%;
	  content: '';
	  z-index: 9;
	  border-left: 1px solid rgba(34,36,38,.15);
	  border-right: 1px solid rgba(255,255,255,.1);
	  width: 0;
	  height: calc(100% - 1rem);
	}
	
	.row-divided > .vertical-divider2 {
	  height: calc(50% - 1rem);    
	}
	
	.vertical-divider2:before {
	  top: -100%;
	}
	
	.vertical-divider2:after {
	  top: auto;
	  bottom: 0;
	}

	
	.row-divided {
	  position:relative;
	}
	
	.row-divided > [class^="col-"],
	.row-divided > [class*=" col-"] {
	  padding-left: 30px;  /* gutter width (give a little extra room) 2x default */
	  padding-right: 30px; /* gutter width (give a little extra room) 2x default */
	}
	
	
	
	
	/* just to set different column heights - not needed to function */          
	.column-one {
	  height: 500px; 
	  background-color: #EBFFF9;
	}
	.column-two {
	  height: 400px;
	  background-color: #F7F3FF;
	}	
	.column-three {
	  height: 400px;
	  background-color: #CAEDFE;
	}	
		 
</style>

<!--- eof on/off scripts !-->

<h1 class="page-header">Find / Edit <%= GetTerm("Equipment") %></h1>

    
    <section class="container">
    
	    <!-- row requires "row-divided" class -->
	    <div class="row row-divided searchrow">
	        <div class="col-xs-4 column-one">
	            <h2>Search By Asset Tag, Serial #, or Customer</h2>
	            <p>6 column wide (col-xs-6)</p>
	            
                  <h2 class="search">Search By Complete Asset Tag or Serial Number</h2>
                    
		    		<!-- select equipment record !-->
						<input id="txtEquipID" name="txtEquipID">
						<input type="hidden" id="txtEquipIDToPass" name="txtEquipIDToPass" value="<%= EquipIDForDetail %>" >
						<i id="searchIcon" class="fa fa-search fa-2x"></i>
					<!-- eof select equipment record !-->
					
	                   <div class="form-group">
                        <input type="submit" id="btnSubmit" name="btnSubmit" class="btn btn-primary" value="Search">
                    </div>
                
				
                    
                    <hr class="formdivider">

					
					<h2 class="search">Search By Customer</h2>
                    
 
		    		<!-- select equipment record !-->
						<input id="txtCustomerID" name="txtCustomerID">
						<input type="hidden" id="txtCustomerIDToPass" name="txtCustomerIDToPass" value="<%= CustomerIDForDetail %>" >
						<i id="searchIcon" class="fa fa-search fa-2x"></i>
					<!-- eof select equipment record !-->
					
                    <div class="form-group">
                        <input type="submit" id="btnSubmit" name="btnSubmit" class="btn btn-primary" value="Search">
                    </div>
                

					<hr class="formdivider">
					
					<h2 class="search">Search By Partial Asset Tag or Serial Number</h2>
					
					 <form id="frmContact" name="frmContact" action="#" method="post">
                    	
                        <div class="form-group">
                            <input type="text" name="txtPartialEquip" id="txtPartialEquip" class="form-control-partial-search" placeholder="Search for equipment by PARTIAL serial number or asset tag">
                            <i id="searchIcon" class="fa fa-search fa-2x"></i>
                        </div>
                                                
                        <div class="form-group">
                            <input type="submit" id="btnSubmit" name="btnSubmit" class="btn btn-primary" value="Partial Search">
                        </div>
                    
                    </form>
            
	        </div>
	        
	      	<div class="vertical-divider">or</div>
	      	
	    	<div class="col-xs-4 column-two">
	            <h2>Search by Class, Manufacturer, Brand, or Model</h2>
	            <p>6 column wide (col-xs-6)</p>   
                
                <form id="frmContact2" name="frmContact2" action="#" method="post">
 
                    <div class="form-group">
					  	<select class="form-control required" name="selClassIntRecID" id="selClassIntRecID">
					  			<option value="" selected="selected">Select Class of Equipment</option>
						      	<% 'Get all equipment classes
						      	  	SQLEquipClasses = "SELECT * FROM EQ_Classes ORDER BY Class ASC"
		
									Set cnnEquipClasses = Server.CreateObject("ADODB.Connection")
									cnnEquipClasses.open (Session("ClientCnnString"))
									Set rsEquipClasses = Server.CreateObject("ADODB.Recordset")
									rsEquipClasses.CursorLocation = 3 
									Set rsEquipClasses = cnnEquipClasses.Execute(SQLEquipClasses)
									If not rsEquipClasses.EOF Then
										Do
											Response.Write("<option value='" & rsEquipClasses("InternalRecordIdentifier") & "'>" & rsEquipClasses("Class") & "</option>")
											rsEquipClasses.movenext
										Loop until rsEquipClasses.eof
									End If
									set rsEquipClasses = Nothing
									cnnEquipClasses.close
									set cnnEquipClasses = Nothing
								%>
						</select>
                    </div>



                    <div class="form-group">
					  	<select class="form-control required" name="selManfIntRecID" id="selManfIntRecID">
					  			<option value="" selected="selected">Select Manufacturer of Equipment</option>
						      	<% 'Get all equipment classes
						      	  	SQLEquipManufacturers = "SELECT * FROM EQ_Manufacturers ORDER BY ManufacturerName ASC"
		
									Set cnnEquipManufacturers = Server.CreateObject("ADODB.Connection")
									cnnEquipManufacturers.open (Session("ClientCnnString"))
									Set rsEquipManufacturers = Server.CreateObject("ADODB.Recordset")
									rsEquipManufacturers.CursorLocation = 3 
									Set rsEquipManufacturers = cnnEquipManufacturers.Execute(SQLEquipManufacturers)
									If not rsEquipManufacturers.EOF Then
										Do
											Response.Write("<option value='" & rsEquipManufacturers("InternalRecordIdentifier") & "'>" & rsEquipManufacturers("ManufacturerName") & "</option>")
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
					  	<select class="form-control required" name="selBrandIntRecID" id="selBrandIntRecID">
					  			<option value="" selected="selected">Select Brand of Equipment</option>
						      	<% 'Get all equipment brands
						      	  	SQLEquipBrands = "SELECT * FROM EQ_Brands ORDER BY Brand ASC"
		
									Set cnnEquipBrands = Server.CreateObject("ADODB.Connection")
									cnnEquipBrands.open (Session("ClientCnnString"))
									Set rsEquipBrands = Server.CreateObject("ADODB.Recordset")
									rsEquipBrands.CursorLocation = 3 
									Set rsEquipBrands = cnnEquipBrands.Execute(SQLEquipBrands)
									If not rsEquipBrands.EOF Then
										Do
											Response.Write("<option value='" & rsEquipBrands("InternalRecordIdentifier") & "'>" & rsEquipBrands("Brand") & "</option>")
											rsEquipBrands.movenext
										Loop until rsEquipBrands.eof
									End If
									set rsEquipBrands = Nothing
									cnnEquipBrands.close
									set cnnEquipBrands = Nothing
								%>
						</select>
                    </div>

             	
                    <div class="form-group">
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
                    
                    <div class="form-group">
                        <input type="submit" id="btnSubmit" name="btnSubmit" class="btn btn-primary" value="Search">
                    </div>
                
                </form>
	    	</div>
	    	
	      	<div class="vertical-divider2">or</div>
	      	
	    	<div class="col-xs-4 column-three">
	            <h2>Search By Equipment Group</h2>
	            <p>6 column wide (col-xs-6)</p>   
                    
                <form id="frmContact22" name="frmContact22" action="#" method="post">
                	
                    <div class="form-group">
					  	<select class="form-control required" name="selGroupIntRecID" id="selGroupIntRecID">
					  			<option value="" selected="selected">Select Group of Equipment</option>
						      	<% 'Get all equipment groups
						      	  	SQLEquipGroups = "SELECT * FROM EQ_Groups ORDER BY GroupName ASC"
		
									Set cnnEquipGroups = Server.CreateObject("ADODB.Connection")
									cnnEquipGroups.open (Session("ClientCnnString"))
									Set rsEquipGroups = Server.CreateObject("ADODB.Recordset")
									rsEquipGroups.CursorLocation = 3 
									Set rsEquipGroups = cnnEquipGroups.Execute(SQLEquipGroups)
									If not rsEquipGroups.EOF Then
										Do
											Response.Write("<option value='" & rsEquipGroups("InternalRecordIdentifier") & "'>" & rsEquipGroups("GroupName") & "</option>")
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
                        <input type="submit" id="btnSubmit" name="btnSubmit" class="btn btn-primary" value="Search">
                    </div>
                
                </form>
	            
	    	</div>
	    	
	    </div>
                                           
    </section>
								

<!--#include file="../../inc/footer-main.asp"-->
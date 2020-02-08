<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/InSightFuncs_Equipment.asp"-->

<link rel="stylesheet" type="text/css" href="<%= BaseURL %>css/tabs.css" />

<SCRIPT LANGUAGE="JavaScript">
<!--
    function validateModelForm()
    {

        if (document.frmAddModel.txtModel.value == "") {
            swal("Model can not be blank.");
            return false;
        }

		var ddlBrand = document.getElementById("selBrandIntRecID");
		var selectedValueBrand = ddlBrand.options[ddlBrand.selectedIndex].value;
		
		if (selectedValueBrand == "")
		{
			swal("Brand must be selected for this model.");
			return false;
		}

		var ddlGroup = document.getElementById("selGroupIntRecID");
		var selectedValueGroup = ddlGroup.options[ddlGroup.selectedIndex].value;
		
		if (selectedValueGroup == "")
		{
			swal("Group must be selected for this model.");
			return false;
		}

		var ddlClass = document.getElementById("selClassIntRecID");
		var selectedValueClass= ddlClass.options[ddlClass.selectedIndex].value;
		
		if (selectedValueClass == "")
		{
			swal("Class must be selected for this model.");
			return false;
		}
        
        if (document.frmAddModel.txtDefaultRentalPrice.value != "") {
        
        	if (isNaN(document.frmAddModel.txtDefaultRentalPrice.value)) {
            	swal("Please enter numbers only for the default rental price.");
            	return false;
           	}
        }
        if (document.frmAddModel.txtDefaultCost.value != "") {
        
        	if (isNaN(document.frmAddModel.txtDefaultCost.value)) {
            	swal("Please enter numbers only for the default cost.");
            	return false;
           	}
        }
        if (document.frmAddModel.txtReplacementCost.value != "") {
        
        	if (isNaN(document.frmAddModel.txtReplacementCost.value)) {
            	swal("Please enter numbers only for the replacement cost.");
            	return false;
           	}
        }
        if (document.frmAddModel.txtInsightAssetTagPrefix.value == "") {
            swal("The Insight asset tag prefix cannot be blank.");
            return false;
        }
        

        return true;
        

    }
// -->
</SCRIPT>   


<!-- password strength meter !-->

<style type="text/css">

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

.pass-strength h5{
	margin-top: 0px;
	color: #000;
}
.popover.primary {
    border-color:#337ab7;
}
.popover.primary>.arrow {
    border-top-color:#337ab7;
}
.popover.primary>.popover-title {
    color:#fff;
    background-color:#337ab7;
    border-color:#337ab7;
}
.popover.success {
    border-color:#d6e9c6;
}
.popover.success>.arrow {
    border-top-color:#d6e9c6;
}
.popover.success>.popover-title {
    color:#3c763d;
    background-color:#dff0d8;
    border-color:#d6e9c6;
}
.popover.info {
    border-color:#bce8f1;
}
.popover.info>.arrow {
    border-top-color:#bce8f1;
}
.popover.info>.popover-title {
    color:#31708f;
    background-color:#d9edf7;
    border-color:#bce8f1;
}
.popover.warning {
    border-color:#faebcc;
}
.popover.warning>.arrow {
    border-top-color:#faebcc;
}
.popover.warning>.popover-title {
    color:#8a6d3b;
    background-color:#fcf8e3;
    border-color:#faebcc;
}
.popover.danger {
    border-color:#ebccd1;
}
.popover.danger>.arrow {
    border-top-color:#ebccd1;
}
.popover.danger>.popover-title {
    color:#a94442;
    background-color:#f2dede;
    border-color:#ebccd1;
}

.select-line{
	margin-bottom: 15px;
}

.enable-disable{
	margin-top:20px;
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

.when-col{
	width: 10%;
}

.reference-col{
	width: 45%;
}

.has-more-col{
	width: 12%;
}

.form-control{
	min-width: 100px;
}

.textarea-box{
	min-width: 260px;
}

.custom-container{
	max-width: 1200px;
	margin:0 auto;
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
	</style>
<!-- eof password strength meter !-->

<h1 class="page-header"> Add New <%= GetTerm("Equipment") %> Model</h1>

<div class="custom-container">

	<form method="POST" action="addModel_submit.asp" name="frmAddModel" id="frmAddModel" onsubmit="return validateModelForm();">

	<div class="col-lg-6">
	
		<div class="row row-line">
			<div class="form-group col-lg-12">
				<label for="txtModel" class="col-sm-3 control-label">Model</label>	
    			<div class="col-sm-6">
    				<input type="text" class="form-control required" id="txtModel" name="txtModel">
    			</div>
			</div>
		</div>
		

		<div class="row row-line">
			<div class="form-group col-lg-12">
				<label for="txtBrand" class="col-sm-3 control-label">Brand</label>	
    			<div class="col-sm-6">
    			
    			<% totalBrands = GetTotalNumberOfBrands() %>
    			
    			<% If totalBrands > 0 Then %>
				  	<select class="form-control" name="selBrandIntRecID" id="selBrandIntRecID">
				  			<option value="">Select Brand For Model</option>
					      	<% 'Get all Manufacturers 
					      	  	SQL9 = "SELECT * FROM EQ_Brands ORDER BY Brand ASC"
	
								Set cnn9 = Server.CreateObject("ADODB.Connection")
								cnn9.open (Session("ClientCnnString"))
								Set rs9 = Server.CreateObject("ADODB.Recordset")
								rs9.CursorLocation = 3 
								Set rs9 = cnn9.Execute(SQL9)
								If not rs9.EOF Then
									Do
										Response.Write("<option value='" & rs9("InternalRecordIdentifier") & "'>" & rs9("Brand") & "</option>")
										rs9.movenext
									Loop until rs9.eof
								End If
								set rs9 = Nothing
								cnn9.close
								set cnn9 = Nothing
							%>
					</select>
				<% Else %>
					<p>You do not have any brands defined. A model must have a brand. 
					<strong>Please add a brand to start adding models</strong>.
					<a href="<%= BaseURL %>equipment/brands/addBrand.asp">
	    				<button type="button" class="btn btn-primary"><i class="fa fa-plus"></i> Add a Brand</button>
					</a>
					</p>
				<% End If %>

    			</div>
			</div>
		</div>
		
		
		
		<div class="row row-line">
			<div class="form-group col-lg-12">
				<label for="txtBrand" class="col-sm-3 control-label">Groups</label>	
    			<div class="col-sm-6">
    			
    			<% totalGroups = GetTotalNumberOfGroups() %>
    			
    			<% If totalGroups > 0 Then %>
				  	<select class="form-control" name="selGroupIntRecID" id="selGroupIntRecID">
				  			<option value="">Select Group For Model</option>
					      	<% 'Get all Types 
					      	  	SQL9 = "SELECT * FROM EQ_Groups ORDER BY GroupName ASC"
	
								Set cnn9 = Server.CreateObject("ADODB.Connection")
								cnn9.open (Session("ClientCnnString"))
								Set rs9 = Server.CreateObject("ADODB.Recordset")
								rs9.CursorLocation = 3 
								Set rs9 = cnn9.Execute(SQL9)
								If not rs9.EOF Then
									Do
										Response.Write("<option value='" & rs9("InternalRecordIdentifier") & "'>" & rs9("GroupName") & "</option>")
										rs9.movenext
									Loop until rs9.eof
								End If
								set rs9 = Nothing
								cnn9.close
								set cnn9 = Nothing
							%>
					</select>
				<% Else %>
					<p>You do not have any Groups defined. A model must have a Group (i.e. Single Cup Brewer, Glass Pot Brewer, etc.). 
					<strong>Please add a Group to start adding models</strong>.
					<a href="<%= BaseURL %>equipment/groups/addGroup.asp">
	    				<button type="button" class="btn btn-primary"><i class="fa fa-plus"></i> Add a Group</button>
					</a>
					</p>
				<% End If %>

    			</div>
			</div>
		</div>
		
		
		
		<div class="row row-line">
			<div class="form-group col-lg-12">
				<label for="txtBrand" class="col-sm-3 control-label">Class</label>	
    			<div class="col-sm-6">
    			
    			<% totalClasses = GetTotalNumberOfClasses() %>
    			
    			<% If totalClasses > 0 Then %>
				  	<select class="form-control" name="selClassIntRecID" id="selClassIntRecID">
				  			<option value="">Select Class For Model</option>
					      	<% 'Get all Classes 
					      	  	SQL9 = "SELECT * FROM EQ_Classes ORDER BY Class ASC"
	
								Set cnn9 = Server.CreateObject("ADODB.Connection")
								cnn9.open (Session("ClientCnnString"))
								Set rs9 = Server.CreateObject("ADODB.Recordset")
								rs9.CursorLocation = 3 
								Set rs9 = cnn9.Execute(SQL9)
								If not rs9.EOF Then
									Do
										Response.Write("<option value='" & rs9("InternalRecordIdentifier") & "'>" & rs9("Class") & "</option>")
										rs9.movenext
									Loop until rs9.eof
								End If
								set rs9 = Nothing
								cnn9.close
								set cnn9 = Nothing
							%>
					</select>
				<% Else %>
					<p>You do not have any classes defined. A model must have a class (i.e. OCS, Vending, Water, etc.). 
					<strong>Please add a class to start adding models</strong>.
					<a href="<%= BaseURL %>equipment/classes/addClass.asp">
	    				<button type="button" class="btn btn-primary"><i class="fa fa-plus"></i> Add a Class</button>
					</a>
					</p>
				<% End If %>

    			</div>
			</div>
		</div>
		
	</div>	
	<div class="col-lg-6">	

		<div class="row row-line">
			<div class="form-group col-lg-12">
				<label for="txtModel" class="col-sm-3 control-label">Default Rental Price</label>	
    			<div class="col-sm-6">
    				<i class="inside fa fa-usd"></i>
    				<input type="text" class="form-control required inp" id="txtDefaultRentalPrice" name="txtDefaultRentalPrice">
    			</div>
			</div>
		</div>
		
		<div class="row row-line">
			<div class="form-group col-lg-12">
				<label for="txtModel" class="col-sm-3 control-label">Default Cost</label>	
    			<div class="col-sm-6">
    				 <i class="inside fa fa-usd"></i>
    				<input type="text" class="form-control required inp" id="txtDefaultCost" name="txtDefaultCost">
    			</div>
			</div>
		</div>
		
		
		<div class="row row-line">
			<div class="form-group col-lg-12">
				<label for="txtModel" class="col-sm-3 control-label">Replacement Cost</label>	
    			<div class="col-sm-6">
    				 <i class="inside fa fa-usd"></i>
    				<input type="text" class="form-control required inp" id="txtReplacementCost" name="txtReplacementCost">
    			</div>
			</div>
		</div>
		
		<div class="row row-line">
			<div class="form-group col-lg-12">
				<label for="txtModel" class="col-sm-3 control-label">Backend System Code</label>	
    			<div class="col-sm-6">
    				<input type="text" class="form-control inp" id="txtBackendSystemCode" name="txtBackendSystemCode">
    			</div>
			</div>
		</div>		

		<div class="row row-line">
			<div class="form-group col-lg-12">
				<label for="txtModel" class="col-sm-3 control-label">Insight Asset Tag Prefix</label>	
    			<div class="col-sm-6">
    				<input type="text" class="form-control required inp" id="txtInsightAssetTagPrefix" name="txtInsightAssetTagPrefix">
    			</div>
			</div>
		</div>			
		
	    <!-- cancel / submit !-->
		<div class="row row-line">
			&nbsp;
		</div>
		
		
	    <!-- cancel / submit !-->
		<div class="row row-line">
			<div class="col-lg-12 alertbutton">
				<div class="col-lg-12">
					<a href="<%= BaseURL %>equipment/models/main.asp">
	    				<button type="button" class="btn btn-default">&lsaquo; Cancel &amp; Go Back To Models List</button>
					</a>
					<% If totalBrands > 0 AND totalGroups > 0 AND totalClasses > 0 Then %>
						<button type="submit" class="btn btn-primary"><i class="far fa-save"></i> Save</button>
					<% End If %>
				</div>
		    </div>
		</div>
		
	</div>
		
	</form>
</div>

<!--#include file="../../inc/footer-main.asp"-->

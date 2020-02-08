<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/InSightFuncs_Equipment.asp"-->



<link rel="stylesheet" type="text/css" href="<%= BaseURL %>css/tabs.css" />

<SCRIPT LANGUAGE="JavaScript">
<!--
    function validateBrandForm()
    {

        if (document.frmAddBrand.txtBrand.value == "") {
            swal("Brand name cannot be blank.");
            return false;
        }
        
		var ddlManufacturer = document.getElementById("selManufacturerIntRecID");
		var selectedValue = ddlManufacturer.options[ddlManufacturer.selectedIndex].value;
		
		if (selectedValue == "")
		{
			swal("Brand manufacturer must be selected.");
			return false;
		}

        if (document.frmAddBrand.txtInsightAssetTagPrefix.value == "") {
            swal("The Insight asset tag prefix cannot be blank.");
            return false;
        }

        return true;

    }
// -->
</SCRIPT>   


<!-- password strength meter !-->

<style type="text/css">

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
	max-width:600px;
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

<h1 class="page-header"> Add New <%= GetTerm("Equipment") %> Brand</h1>

<div class="custom-container">

	<form method="POST" action="addBrand_submit.asp" name="frmAddBrand" id="frmAddBrand" onsubmit="return validateBrandForm();">

		<div class="row row-line">

			<div class="form-group col-lg-12">
				<label for="txtBrand" class="col-sm-3 control-label">Brand</label>	
    			<div class="col-sm-6">
    				<input type="text" class="form-control required" id="txtBrand" name="txtBrand" >
    			</div>
			</div>
			
		</div>

		<div class="row row-line">

			<div class="form-group col-lg-12">
				<label for="txtBrand" class="col-sm-3 control-label">Manufacturer</label>	
    			<div class="col-sm-6">
    			
    			<% totalManufacturers = GetTotalNumberOfManufacturers() %>
    			
    			<% If totalManufacturers > 0 Then %>
				  	<select class="form-control" name="selManufacturerIntRecID" id="selManufacturerIntRecID">
				  			<option value="">Select Manufacturer For Brand</option>
					      	<% 'Get all Manufacturers 
					      	  	SQL9 = "SELECT * FROM EQ_Manufacturers ORDER BY ManufacturerName ASC"
	
								Set cnn9 = Server.CreateObject("ADODB.Connection")
								cnn9.open (Session("ClientCnnString"))
								Set rs9 = Server.CreateObject("ADODB.Recordset")
								rs9.CursorLocation = 3 
								Set rs9 = cnn9.Execute(SQL9)
								If not rs9.EOF Then
									Do
										Response.Write("<option value='" & rs9("InternalRecordIdentifier") & "'>" & rs9("ManufacturerName") & "</option>")
										rs9.movenext
									Loop until rs9.eof
								End If
								set rs9 = Nothing
								cnn9.close
								set cnn9 = Nothing
							%>
					</select>
				<% Else %>
					<p>You do not have any manufacturers defined. A brand must have a manufacturer. <strong>Please add a manufacturer to start adding brands</strong>.
					
					<a href="<%= BaseURL %>equipment/manufacturers/addManufacturer.asp">
	    				<button type="button" class="btn btn-primary"><i class="fa fa-plus"></i> Add a Manufacturer</button>
					</a>
					</p>
				<% End If %>

    			</div>
			</div>
			
		</div>
		
		<div class="row row-line">

			<div class="form-group col-lg-12">
				<label for="txtBrand" class="col-sm-3 control-label">Insight Asset Tag Prefix</label>	
    			<div class="col-sm-6">
    				<input type="text" class="form-control required" id="txtInsightAssetTagPrefix" name="txtInsightAssetTagPrefix">
    			</div>
			</div>
			
		</div>
		
		
	    <!-- cancel / submit !-->
		<div class="row row-line">
			<div class="col-lg-12 alertbutton">
				<div class="col-lg-12">
					<a href="<%= BaseURL %>equipment/brands/main.asp">
	    				<button type="button" class="btn btn-default">&lsaquo; Cancel &amp; Go Back To Brands List</button>
					</a>
					<% If totalManufacturers > 0 Then %>
						<button type="submit" class="btn btn-primary"><i class="far fa-save"></i> Save</button>
					<% End If %>
				</div>
		    </div>
		</div>
		
	</form>
</div>

<!--#include file="../../inc/footer-main.asp"-->

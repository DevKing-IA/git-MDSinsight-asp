<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/InSightFuncs_Equipment.asp"-->

<% InternalRecordIdentifier = Request.QueryString("i") 
If InternalRecordIdentifier = "" Then Response.Redirect("main.asp")
%>

<link rel="stylesheet" type="text/css" href="<%= BaseURL %>css/tabs.css" />


<SCRIPT LANGUAGE="JavaScript">
<!--
    function validateEditBrandForm()
    {

        if (document.frmEditBrand.txtBrand.value == "") {
            swal("Lead source can not be blank.");
            return false;
        }

		var ddlManufacturer = document.getElementById("selManufacturerIntRecID");
		var selectedValue = ddlManufacturer.options[ddlManufacturer.selectedIndex].value;
		
		if (selectedValue == "")
		{
			swal("Brand manufacturer must be selected.");
			return false;
		}

        if (document.frmEditBrand.txtInsightAssetTagPrefix.value == "") {
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



<%
SQL = "SELECT * FROM EQ_Brands where InternalRecordIdentifier = " & InternalRecordIdentifier 

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	Brand = rs("Brand")
	ManufacturerIntRecID = rs("ManufacIntRecID")
	InsightAssetTagPrefix = rs("InsightAssetTagPrefix")
End If
set rs = Nothing
cnn8.close
set cnn8 = Nothing

%>


<h1 class="page-header"> Edit <%= GetTerm("Equipment") %> Brand</h1>

<div class="custom-container">

	<form method="POST" action="editBrand_submit.asp" name="frmEditBrand" id="frmEditBrand" onsubmit="return validateEditBrandForm();">

		<div class="row row-line">

			<input type="hidden" id="txtInternalRecordIdentifier" name="txtInternalRecordIdentifier" value="<%=InternalRecordIdentifier%>">
    				
			<div class="form-group col-lg-12">
				<label for="txtBrand" class="col-sm-3 control-label">Brand</label>	
    			<div class="col-sm-6">
    				<input type="text" class="form-control required" id="txtBrand" name="txtBrand" value="<%= Brand %>">
    			</div>
			</div>

		</div>


		<div class="row row-line">

			<div class="form-group col-lg-12">
				<label for="txtBrand" class="col-sm-3 control-label">Manufacturer</label>	
    			<div class="col-sm-6">
				  	<select class="form-control" name="selManufacturerIntRecID" id="selManufacturerIntRecID">
				      	<% 'Get all Manufacturers 
				      	  	SQL9 = "SELECT * FROM EQ_Manufacturers ORDER BY ManufacturerName ASC"

							Set cnn9 = Server.CreateObject("ADODB.Connection")
							cnn9.open (Session("ClientCnnString"))
							Set rs9 = Server.CreateObject("ADODB.Recordset")
							rs9.CursorLocation = 3 
							Set rs9 = cnn9.Execute(SQL9)
							If not rs9.EOF Then
								Do
									If cInt(ManufacturerIntRecID) = cInt(rs9("InternalRecordIdentifier")) Then
										Response.Write("<option value='" & rs9("InternalRecordIdentifier") & "' selected='selected'>" & rs9("ManufacturerName") & "</option>")
									Else
										Response.Write("<option value='" & rs9("InternalRecordIdentifier") & "'>" & rs9("ManufacturerName") & "</option>")
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
		
		
		<div class="row row-line">

			<div class="form-group col-lg-12">
				<label for="txtBrand" class="col-sm-3 control-label">Insight Asset Tag Prefix</label>	
    			<div class="col-sm-6">
    				<input type="text" class="form-control required" id="txtInsightAssetTagPrefix" name="txtInsightAssetTagPrefix" value="<%= InsightAssetTagPrefix %>">
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
					<button type="submit" class="btn btn-primary"><i class="far fa-save"></i> Save</button>
				</div>
		    </div>
		</div>
		
	</form>
</div>

<!--#include file="../../inc/footer-main.asp"-->

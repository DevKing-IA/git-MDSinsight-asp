<!--#include file="../../inc/header.asp"-->
<!-- function that gets the value of the tab when it is clicked and then
updates the value of a hidden form field so when the page posts, it returns
back to the tab that was previously opened -->

<script type="text/javascript">
	$(function () {
		$('a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
		var target = $(e.target).attr("href");
		$('input[name="txtTab"]').val(target);
		//alert(target);
		});
	})
	
	$(window).load(function()
	{
	   var phones = [{ "mask": "(###) ###-####"}];
	    $('#txtPhone1').inputmask({ 
	        mask: phones, 
	        greedy: false, 
	        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
	    $('#txtPhone2').inputmask({ 
	        mask: phones, 
	        greedy: false, 
	        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
	    $('#txtPhone3').inputmask({ 
	        mask: phones, 
	        greedy: false, 
	        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
		$('#txtFax').inputmask({ 
	        mask: phones, 
	        greedy: false, 
	        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
        
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

   function validateCompanyIdentityForm()
    {
    
       if (document.frmCompanyID.txtCompanyName.value == "") {
            swal("Company name cannot be blank.");
            return false;
       }
       if (document.frmCompanyID.txtAddress1.value == "") {
            swal("Address 1 cannot be blank.");
            return false;
       }
       if (document.frmCompanyID.txtCity.value == "") {
            swal("City cannot be blank.");
            return false;
       }
       if (document.frmCompanyID.selState.value == "") {
            swal("State cannot be blank.");
            return false;
       }
       if (document.frmCompanyID.txtZip.value == "") {
            swal("Zip Code cannot be blank.");
            return false;
       }
       if (document.frmCompanyID.selCountry.value == "") {
            swal("Country cannot be blank.");
            return false;
       }
       if (document.frmCompanyID.txtEmail.value == "") {
            swal("Email address cannot be blank.");
            return false;
       }
       if ((document.frmCompanyID.txtEmail.value !== "") && (isValidEmail(document.frmCompanyID.txtEmail.value) == false)) {
           swal("The email address is invalid. Please enter any format like the following: anystring@anystring.anystring.");
           return false;
       }
       if ((document.frmCompanyID.txtPhone1.value !== "") && (isValidPhone(document.frmCompanyID.txtPhone1.value) == false)) {
           swal("Phone number 1 is invalid. Please enter the number in the following format: (555) 555-5555.");
           return false;
       } 
       if ((document.frmCompanyID.txtPhone2.value !== "") && (isValidPhone(document.frmCompanyID.txtPhone2.value) == false)) {
           swal("Phone number 2 is invalid. Please enter the number in the following format: (555) 555-5555.");
           return false;
       }
       if ((document.frmCompanyID.txtPhone3.value !== "") && (isValidPhone(document.frmCompanyID.txtPhone3.value) == false)) {
           swal("Phone number 3 is invalid. Please enter the number in the following format: (555) 555-5555.");
           return false;
       }
       
       if ((document.frmCompanyID.txtFax.value !== "") && (isValid(document.frmCompanyID.txtFax.value) == false)) {
           swal("The fax number is invalid. Please enter the number in the following format: (555) 555-5555.");
           return false;
       }      
       return true;

    }
// -->
</script>   
<script src="<%= BaseURL %>js/bootstrap-yearly-calendar/bootstrap-year-calendar.js"></script>
<link rel="stylesheet" type="text/css" href="<%= BaseURL %>js/bootstrap-yearly-calendar/bootstrap-year-calendar.css">

<script type="text/javascript" src="<%= BaseURL %>js/bootstrap-clockpicker/clockpicker.js"></script>
<link rel="stylesheet" type="text/css" href="<%= BaseURL %>js/bootstrap-clockpicker/clockpicker.css" />

<script src="http://cdnjs.cloudflare.com/ajax/libs/moment.js/2.10.3/moment.js"></script>

<script type="text/javascript" src="<%= BaseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.js"></script>
<link rel="stylesheet" type="text/css" href="<%= BaseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.css" />


<!-- spectrum color picker !-->
<script src="<%= BaseURL %>/js/spectrum-color-picker/spectrum.js"></script>
<link rel="stylesheet" type="text/css" href="<%= BaseURL %>/js/spectrum-color-picker/spectrum.css">
<style  type="text/css">
.full-spectrum .sp-palette {
max-width: 200px;
}

 
</style>
<!-- eof spectrum color picker !-->
	

<script language="javascript">

var valid_ext = /(.png)$/i;

function CheckExtension(path_field)
{
	if (valid_ext.test(path_field.value)) return true;
    swal({
        title: "Company Logo Upload Error",
        text: "Only .png files are allowed for the company logo - The background must be transparent.",
        imageUrl: "../../img/alert-icons/transparent-logo-error-image.png",
        confirmButtonColor: "#337ab7",
        confirmButtonText: 'OK'
	});
}

</script>

<!-- time picker !-->
<link rel="stylesheet" href="<%= BaseURL %>js/timepicker/timepicker/ui-1.10.0/ui-lightness/jquery-ui-1.10.0.custom.min.css" type="text/css" />
<link rel="stylesheet" href="<%= BaseURL %>js/timepicker/timepicker/jquery.ui.timepicker.css?v=0.3.3" type="text/css" />
<script type="text/javascript" src="<%= BaseURL %>js/timepicker/timepicker/ui-1.10.0/jquery.ui.core.min.js"></script>
<script type="text/javascript" src="<%= BaseURL %>js/timepicker/timepicker/ui-1.10.0/jquery.ui.widget.min.js"></script>
<script type="text/javascript" src="<%= BaseURL %>js/timepicker/timepicker/ui-1.10.0/jquery.ui.tabs.min.js"></script>
<script type="text/javascript" src="<%= BaseURL %>js/timepicker/timepicker/ui-1.10.0/jquery.ui.position.min.js"></script>

<script type="text/javascript" src="<%= BaseURL %>js/timepicker/timepicker/jquery.ui.timepicker.js?v=0.3.3"></script>
<!-- eof time picker !-->

<%

If (Request.ServerVariables("REQUEST_METHOD") = "POST") Then

	
	'*******************************************
	' Prepare ASPUpload to receive company logo
	'*******************************************

	Set Upload = Server.CreateObject("Persits.Upload.1")
	Upload.Save
	
	CompanyName = Upload.Form("txtCompanyName")
	Address1 = Upload.Form("txtAddress1")
	Address2 = Upload.Form("txtAddress2")
	City = Upload.Form("txtCity")
	State = Upload.Form("selState")
	Zip = Upload.Form("txtZip")
	Country = Upload.Form("selCountry")
	Phone1 = Upload.Form("txtPhone1")
	Phone2 = Upload.Form("txtPhone2")
	Phone3 = Upload.Form("txtPhone3")
	Fax = Upload.Form("txtFax")
	Email = Upload.Form("txtEmail")
	Attention = Upload.Form("txtAttention")
	MessageToPrint = Upload.Form("txtMessageToPrint")
	CompanyLogo = Upload.Form("txtCompanyLogo")
	CompanyIdentityColor1 = Upload.Form("txtCompanyIdentityColor1")
	CompanyIdentityColor2 = Upload.Form("txtCompanyIdentityColor2")
	Timezone = Upload.Form("selTimezone")
	BusinessDayStartz = Upload.Form("txtBusinessDayStart")
	BusinessDayEndz = Upload.Form("txtBusinessDayEnd")
	PeriodsOrMonths = Upload.Form("optCompanyPeriodsOrMonths")
	PointOfServiceLogic= Upload.Form("chkPointOfServiceLogicOnOff")
	
	If PointOfServiceLogic="on" Then
		PointOfServiceLogicValue = 1
	Else
		PointOfServiceLogicValue = 0
	End if	

	'******************************************************************
	' Overwrite the existing logo or upload a new one as logo.png
	'******************************************************************
	
	'Rename the files
	' Construct the save path
	Pth ="../../clientfiles/" & trim(MUV_READ("ClientID")) & "/logos/"

	For Each File in Upload.Files
	   File.SaveAsVirtual  Pth & "logo" & File.Ext
	Next

	'*******************************************
	' This code to write to the audit trail file
	'*******************************************
	SQL = "SELECT * FROM Settings_CompanyID"
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)

	If not rs.EOF Then
		If CompanyName <> rs("Stmt_CompanyName") Then
			CreateAuditLogEntry "Company ID Setting Change", "Company ID Setting Change", "Minor", 1, MUV_Read("DisplayName") & " changed company name from " & rs("Stmt_CompanyName") & " to " & CompanyName 
		End If
		If Address1 <> rs("Stmt_Address1") Then
			CreateAuditLogEntry "Company ID Setting Change", "Company ID Setting Change", "Minor", 1, MUV_Read("DisplayName") & " changed address 1 from " & rs("Stmt_Address1") & " to " & Address1 
		End If
		If Address2 <> rs("Stmt_Address2") Then
			CreateAuditLogEntry "Company ID Setting Change", "Company ID Setting Change", "Minor", 1, MUV_Read("DisplayName") & " changed address 2 from " & rs("Stmt_Address2") & " to " & Address2 
		End If
		If City <> rs("Stmt_City") Then
			CreateAuditLogEntry "Company ID Setting Change", "Company ID Setting Change", "Minor", 1, MUV_Read("DisplayName") & " changed city from " & rs("Stmt_City") & " to " & City 
		End If
		If State <> rs("Stmt_State") Then
			CreateAuditLogEntry "Company ID Setting Change", "Company ID Setting Change", "Minor", 1, MUV_Read("DisplayName") & " changed state from " & rs("Stmt_State") & " to " & State 
		End If
		If Zip <> rs("Stmt_Zip") Then
			CreateAuditLogEntry "Company ID Setting Change", "Company ID Setting Change", "Minor", 1, MUV_Read("DisplayName") & " changed zip code from " & rs("Stmt_Zip") & " to " & Zip 
		End If
		If Country <> rs("Stmt_Country") Then
			CreateAuditLogEntry "Company ID Setting Change", "Company ID Setting Change", "Minor", 1, MUV_Read("DisplayName") & " changed country from " & rs("Stmt_Country") & " to " & Country
		End If
		If Phone1 <> rs("Stmt_Phone1") Then
			CreateAuditLogEntry "Company ID Setting Change", "Company ID Setting Change", "Minor", 1, MUV_Read("DisplayName") & " changed phone 1 from " & rs("Stmt_Phone1") & " to " & Phone1 
		End If
		If Phone2 <> rs("Stmt_Phone2") Then
			CreateAuditLogEntry "Company ID Setting Change", "Company ID Setting Change", "Minor", 1, MUV_Read("DisplayName") & " changed phone 2 from " & rs("Stmt_Phone2") & " to " & Phone2 
		End If
		If Phone3 <> rs("Stmt_Phone3") Then
			CreateAuditLogEntry "Company ID Setting Change", "Company ID Setting Change", "Minor", 1, MUV_Read("DisplayName") & " changed phone 3 from " & rs("Stmt_Phone3") & " to " & Phone3 
		End If
		If Fax <> rs("Stmt_Fax") Then
			CreateAuditLogEntry "Company ID Setting Change", "Company ID Setting Change", "Minor", 1, MUV_Read("DisplayName") & " changed fax from " & rs("Stmt_Fax") & " to " & Fax 
		End If
		If Email <> rs("Stmt_Email") Then
			CreateAuditLogEntry "Company ID Setting Change", "Company ID Setting Change", "Minor", 1, MUV_Read("DisplayName") & " changed email from " & rs("Stmt_Email") & " to " & Email 
		End If
		If Attention <> rs("Stmt_Attn") Then
			CreateAuditLogEntry "Company ID Setting Change", "Company ID Setting Change", "Minor", 1, MUV_Read("DisplayName") & " changed attention from " & rs("Stmt_Attn") & " to " & Attention 
		End If
		If SpecialMessage <> rs("Stmt_Message") Then
			CreateAuditLogEntry "Company ID Setting Change", "Company ID Setting Change", "Minor", 1, MUV_Read("DisplayName") & " changed message to print from " & rs("Stmt_Message") & " to " & MessageToPrint 
		End If
		
		'*******************************************************
		'Special Code to see if new logo was uploaded
		'*******************************************************
		serverName = Request.ServerVariables("SERVER_NAME")
		If serverName = "www.mdsinsight.com" Then serverName = "mdsinsight.com"
		pathForLogoFile = "C:\home\" & serverName & "\wwwroot\clientfilesV\" & MUV_Read("ClientID") & "\logos\logo.png"
		
		Set objFSO = CreateObject ("Scripting.FileSystemObject")
		
		If objFSO.FileExists(pathForLogoFile) Then
		
			Set objFile = objFSO.GetFile(pathForLogoFile)
			lastModifiedFileDate = FormatDateTime(objFile.DateLastModified)
			todaysDate = FormatDateTime(Now())
			
			yearDiff = DateDiff("y",lastModifiedFileDate,todaysDate)
			monthDiff = DateDiff("m",lastModifiedFileDate,todaysDate)
			dayDiff = DateDiff("d",lastModifiedFileDate,todaysDate)
			hourDiff = DateDiff("h",lastModifiedFileDate,todaysDate)
			minuteDiff = DateDiff("n",lastModifiedFileDate,todaysDate)
			
			If yearDiff = 0 AND monthDiff = 0 AND dayDiff = 0 AND hourDiff = 0 AND minuteDiff = 0 Then
				
				CreateAuditLogEntry "Company ID Setting Change", "Company ID Setting Change", "Minor", 1, MUV_Read("DisplayName") & " uploaded a new company logo last modified at " & lastModifiedFileDate  & " on " & todaysDate
				
				'Response.Write("<br>LastModifiedDate: " & lastModifiedFileDate & "<br>")
				'Response.Write("<br>todaysDate: " & todaysDate & "<br>")
				'Response.Write("<br>DateDiff Years: " & DateDiff("y",lastModifiedFileDate,todaysDate) & "<br>")
				'Response.Write("<br>DateDiff Months: " & DateDiff("m",lastModifiedFileDate,todaysDate) & "<br>")
				'Response.Write("<br>DateDiff Days: " & DateDiff("d",lastModifiedFileDate,todaysDate) & "<br>")
				'Response.Write("<br>DateDiff Hours: " & DateDiff("h",lastModifiedFileDate,todaysDate) & "<br>")
				'Response.Write("<br>DateDiff Minutes: " & DateDiff("n",lastModifiedFileDate,todaysDate) & "<br>")
				
			End If
			
			Set objFile = Nothing

		End If
		'*******************************************************
	
		If CompanyIdentityColor1 <> rs("CompanyIdentityColor1") Then
			CreateAuditLogEntry "Company ID Setting Change", "Company ID Setting Change", "Minor", 1, MUV_Read("DisplayName") & " changed Company Identity Color 1 from " & rs("CompanyIdentityColor1") & " to " & CompanyIdentityColor1
		End If
		If CompanyIdentityColor2 <> rs("CompanyIdentityColor2") Then
			CreateAuditLogEntry "Company ID Setting Change", "Company ID Setting Change", "Minor", 1, MUV_Read("DisplayName") & " changed Company Identity Color 2 from " & rs("CompanyIdentityColor2") & " to " & CompanyIdentityColor2
		End If
		If Timezone <> rs("Timezone") Then
			CreateAuditLogEntry "Company ID Setting Change", "Company ID Setting Change", "Minor", 1, MUV_Read("DisplayName") & " changed Timezone from " & rs("Timezone") & " to " & Timezone
		End If
		If BusinessDayStartz  <> rs("BusinessDayStart") Then
			CreateAuditLogEntry "Company ID Setting Change", "Company ID Setting Change", "Minor", 1, MUV_Read("DisplayName") & " changed Business Day Start from " & rs("BusinessDayStart") & " to " & BusinessDayStartz 
		End If
		If BusinessDayEndz <> rs("BusinessDayEnd") Then
			CreateAuditLogEntry "Company ID Setting Change", "Company ID Setting Change", "Minor", 1, MUV_Read("DisplayName") & " changed Business Day End from " & rs("BusinessDayEnd") & " to " & BusinessDayEndz
		End If
		
		If PeriodsOrMonths <> rs("PeriodsOrMonths") Then
			CreateAuditLogEntry "Company ID Setting Change", "Company ID Setting Change", "Minor", 1, MUV_Read("DisplayName") & " changed Periods or Months from " & rs("PeriodsOrMonths") & " to " & PeriodsOrMonths
		End If

		If PointOfServiceLogicValue <> rs("PointOfServiceLogicOnOff") Then
			CreateAuditLogEntry "Company ID Setting Change", "Company ID Setting Change", "Minor", 1, MUV_Read("DisplayName") & " changed Point Of Service Logic from " & rs("PointOfServiceLogicOnOff") & " to " & PointOfServiceLogicValue
		End If
		
	End If
	'******************************************
	' End code to write to the audit trail file
	'******************************************
	
	SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_CompanyID SET "
	SQL = SQL & "Stmt_CompanyName = '" & CompanyName & "', "
	SQL = SQL & "Stmt_Address1 = '" & Address1 & "', "
	SQL = SQL & "Stmt_Address2 = '" & Address2 & "', "
	SQL = SQL & "Stmt_City = '"& City & "', "
	SQL = SQL & "Stmt_State = '"& State & "', "
	SQL = SQL & "Stmt_Zip = '" & Zip & "', "
	SQL = SQL & "Stmt_Country = '" & Country & "', "
	SQL = SQL & "Stmt_Phone1 = '" & Phone1 & "',"
	SQL = SQL & "Stmt_Phone2 = '" & Phone2 & "',"
	SQL = SQL & "Stmt_Phone3 = '" & Phone3 & "',"
	SQL = SQL & "Stmt_Fax = '" & Fax & "',"
	SQL = SQL & "Stmt_Email = '" & Email & "',"
	SQL = SQL & "Stmt_Attn = '" & Attention & "',"
	SQL = SQL & "CompanyIdentityColor1 = '" & CompanyIdentityColor1 & "',"
	SQL = SQL & "CompanyIdentityColor2 = '" & CompanyIdentityColor2 & "',"
	SQL = SQL & "Stmt_Message = '" & MessageToPrint & "',"
	SQL = SQL & "Timezone = '" & Timezone & "',"
	SQL = SQL & "BusinessDayStart = '" & BusinessDayStartz  & "',"
	SQL = SQL & "BusinessDayEnd  = '" & BusinessDayEndz & "', "
	SQL = SQL & "PeriodsOrMonths = '" & PeriodsOrMonths & "',"
	SQL = SQL & "PointOfServiceLogicOnOff = " & PointOfServiceLogicValue & ""

	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	
End If

SQL = "SELECT * FROM Settings_CompanyID"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)

If not rs.EOF Then
	CompanyName = rs("Stmt_CompanyName")
	Address1 = rs("Stmt_Address1")
	Address2 = rs("Stmt_Address2")
	City = rs("Stmt_City")
	State = rs("Stmt_State")
	Zip = rs("Stmt_Zip")
	Country = rs("Stmt_Country")
	Phone1 = rs("Stmt_Phone1")
	Phone2 = rs("Stmt_Phone2")
	Phone3 = rs("Stmt_Phone3")
	Fax = rs("Stmt_Fax")
	Email = rs("Stmt_Email")
	Attention = rs("Stmt_Attn")
	MessageToPrint = rs("Stmt_Message")
	CompanyIdentityColor1 = rs("CompanyIdentityColor1")
	CompanyIdentityColor2 = rs("CompanyIdentityColor2")
	Timezone = rs("Timezone")
	BusinessDayStartz  = rs("BusinessDayStart")
	BusinessDayEndz = rs("BusinessDayEnd")
	PeriodsOrMonths = rs("PeriodsOrMonths")
	PointOfServiceLogicOnOff = rs("PointOfServiceLogicOnOff")
End If

If PointOfServiceLogicOnOff = 0 Then
	PointOfServiceLogicStatus = ""
Else
	PointOfServiceLogicStatus = "checked"
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing
%>


<!-- local custom css !-->
<style type="text/css">
	.form-control{
		overflow-x: hidden;
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

	.post-labels{
 		padding-top: 5px;
 	}
 	
 	.row-margin{
	 	margin-bottom: 20px;
	 	margin-top: 20px;
 	}
 	
 	h3{
	 	margin-top: 0px;
 	}
 	
 	.table-size .category{
	 	width: 35%;
	 	font-weight: normal;
 	}
 	
 	.table-size .group-name{
	 	width: 40%
 	}
 	
 	.table-size .sort-order{
	 	width: 10%;
 	}
 	
 	.table-size .display{
	 	width: 15%;
 	}
 
	 .col-line{
		 margin-bottom: 20px;
	  }

	 #ui-timepicker-div{
	font-family: Arial;
}
 
	
.ui-timepicker-table td a{
	padding: 3px;
	width:auto;
	text-align: left;
	font-size: 11px;
}	

.ui-timepicker-table .ui-timepicker-title{
	font-size: 13px;
}

.ui-timepicker-table th.periods{
	font-size: 13px;
}

.ui-widget-header{
	background: #193048;
	border: 1px solid #193048;
}
.m-l-23{
	margin-left:23px;
}
.m-l-14{
	margin-left:-14px;
}
</style>
<!-- eof local custom css !-->

<h1 class="page-header"><i class="fa fa-list-alt"></i> Company Settings</h1>

<input type="hidden" name="txtTab" id="txtTab" value="">

<!-- tabs start here !-->
<div class="row ">
	<div class="col-lg-12">
	<%
	'ActiveTab="identity1"
	%>
	<!-- tabs navigation !-->
	<ul class="nav nav-tabs" role="tablist">
		    <li role="presentation" <% If ActiveTab = "" OR ActiveTab="identity1" Then Response.write("class='active'") %>><a href="#identity1" aria-controls="manage" role="tab" data-toggle="tab">Company Identity</a></li>
		    <li role="presentation" <% If ActiveTab = "calendar1" Then Response.write("class='active'") %>><a href="#calendar1" aria-controls="manage" role="tab" data-toggle="tab">Company Calendar</a></li>
		    <li role="presentation" <% If ActiveTab = "reportperiod" Then Response.write("class='active'") %>><a href="#reportperiod" aria-controls="manage" role="tab" data-toggle="tab">Company Report Period</a></li>
		    <li role="presentation" <% If ActiveTab = "accountingperiod" Then Response.write("class='active'") %>><a href="#accountingperiod" aria-controls="manage" role="tab" data-toggle="tab">Company Accounting Period</a></li>
	</ul>
	<!-- eof tabs navigation !-->
			
	<!-- tabs content !-->
	<div class="tab-content row-margin">

				<!-- Leakage tab !-->
     			 <div class="tab-content">
      				<div role="tabpanel" class="tab-pane fade in active" id="identity1"> 
						<form method="post" action="main.asp" name="frmCompanyID" id="frmCompanyID" ENCTYPE="multipart/form-data" onsubmit="return validateCompanyIdentityForm();">
	        			<div  class="row">
							<div class="col-lg-5">
	  		          			<div class="col-lg-12 col-line">
									<div class="form-group">
										<label   class="col-lg-4 control-label">Company Name</label>
										<div class="col-lg-8">
											<input type="text" id="txtCompanyName" name="txtCompanyName" class="form-control" value="<%=CompanyName%>">
									    </div>
									 </div>
								</div>
	          
								<div class="col-lg-12 col-line">
									<div class="form-group">
										<label   class="col-lg-4 control-label">Address 1</label>
										<div class="col-lg-8">
											<input type="text"  id="txtAddress1" name="txtAddress1" class="form-control" value="<%=Address1%>">
										</div>
									</div>
								</div>

								<div class="col-lg-12 col-line">
									<div class="form-group">
										<label   class="col-lg-4 control-label">Address 2</label>
										<div class="col-lg-8">
											<input type="text"  id="txtAddress2" name="txtAddress2" class="form-control" value="<%=Address2%>">
										</div>
									</div>
								</div>

								<div class="col-lg-7 col-line"> 
									<div class="form-group">
										<label   class="col-lg-4 control-label">City</label>
										<div class="col-lg-8">
											<input type="text"   id="txtCity" name="txtCity" class="form-control" value="<%=City%>">
										</div>
									</div>
								</div>

								<div class="col-lg-5 col-line"> 
									<div class="form-group">
										<label   class="col-lg-3 control-label">Zip</label>
										<div class="col-lg-9">
											<input type="text" id="txtZip" name="txtZip" class="form-control" value="<%=Zip%>">
										</div>
									</div>
								</div>
								
								<div class="col-lg-7 col-line"> 
									<div class="form-group">
										<label  class="col-lg-4 control-label">Country</label>
										<div class="col-lg-8">
											<select class="form-control" id="selCountry" name="selCountry">
												<option value="United States" <% If Country="United States" Then Response.Write(" selected ")%>>United States</option>
												<option value="Canada"<% If Country="Canada" Then Response.Write(" selected ")%>>Canada</option>
 											</select>
										</div>
									</div>
								</div>	

								<div class="col-lg-5 col-line"> 
									<div class="form-group">
										<label   class="col-lg-3 control-label">State</label>
										<div class="col-lg-9">
											<!--<input type="text" id="txtState" name="txtState" class="form-control" value="<%=State%>">-->
											<select class="form-control" id="selState" name="selState">
												<option value="AA"<% If State="AA" Then Response.Write(" selected ")%>>AA</option>
												<option value="AE"<% If State="AE" Then Response.Write(" selected ")%>>AE</option>
												<option value="AK"<% If State="AK" Then Response.Write(" selected ")%>>AK</option>
												<option value="AL"<% If State="AL" Then Response.Write(" selected ")%>>AL</option>
												<option value="AP"<% If State="AP" Then Response.Write(" selected ")%>>AP</option>
												<option value="AR"<% If State="AR" Then Response.Write(" selected ")%>>AR</option>
												<option value="AS"<% If State="AS" Then Response.Write(" selected ")%>>AS</option>
												<option value="AZ"<% If State="AZ" Then Response.Write(" selected ")%>>AZ</option>
												<option value="CA"<% If State="CA" Then Response.Write(" selected ")%>>CA</option>
												<option value="CO"<% If State="CO" Then Response.Write(" selected ")%>>CO</option>
												<option value="CT"<% If State="CT" Then Response.Write(" selected ")%>>CT</option>
												<option value="DC"<% If State="DC" Then Response.Write(" selected ")%>>DC</option>
												<option value="DE"<% If State="DE" Then Response.Write(" selected ")%>>DE</option>
												<option value="FL"<% If State="FL" Then Response.Write(" selected ")%>>FL</option>
												<option value="FM"<% If State="FM" Then Response.Write(" selected ")%>>FM</option>
												<option value="GA"<% If State="GA" Then Response.Write(" selected ")%>>GA</option>
												<option value="GU"<% If State="GU" Then Response.Write(" selected ")%>>GU</option>
												<option value="HI"<% If State="HI" Then Response.Write(" selected ")%>>HI</option>
												<option value="IA"<% If State="IA" Then Response.Write(" selected ")%>>IA</option>
												<option value="ID"<% If State="ID" Then Response.Write(" selected ")%>>ID</option>
												<option value="IL"<% If State="IL" Then Response.Write(" selected ")%>>IL</option>
												<option value="IN"<% If State="IN" Then Response.Write(" selected ")%>>IN</option>
												<option value="KS"<% If State="KS" Then Response.Write(" selected ")%>>KS</option>
												<option value="KY"<% If State="KY" Then Response.Write(" selected ")%>>KY</option>
												<option value="LA"<% If State="LA" Then Response.Write(" selected ")%>>LA</option>
												<option value="MA"<% If State="MA" Then Response.Write(" selected ")%>>MA</option>
												<option value="MD"<% If State="MD" Then Response.Write(" selected ")%>>MD</option>
												<option value="ME"<% If State="ME" Then Response.Write(" selected ")%>>ME</option>
												<option value="MI"<% If State="MI" Then Response.Write(" selected ")%>>MI</option>
												<option value="MN"<% If State="MN" Then Response.Write(" selected ")%>>MN</option>
												<option value="MO"<% If State="MO" Then Response.Write(" selected ")%>>MO</option>
												<option value="MP"<% If State="MP" Then Response.Write(" selected ")%>>MP</option>
												<option value="MS"<% If State="MS" Then Response.Write(" selected ")%>>MS</option>
												<option value="MT"<% If State="MT" Then Response.Write(" selected ")%>>MT</option>
												<option value="NC"<% If State="NC" Then Response.Write(" selected ")%>>NC</option>
												<option value="ND"<% If State="ND" Then Response.Write(" selected ")%>>ND</option>
												<option value="NE"<% If State="NE" Then Response.Write(" selected ")%>>NE</option>
												<option value="NH"<% If State="NH" Then Response.Write(" selected ")%>>NH</option>
												<option value="NJ"<% If State="NJ" Then Response.Write(" selected ")%>>NJ</option>
												<option value="NY"<% If State="NY" Then Response.Write(" selected ")%>>NY</option>
												<option value="NM"<% If State="NM" Then Response.Write(" selected ")%>>NM</option>
												<option value="NV"<% If State="NV" Then Response.Write(" selected ")%>>NV</option>
												<option value="OH"<% If State="OH" Then Response.Write(" selected ")%>>OH</option>
												<option value="OK"<% If State="OK" Then Response.Write(" selected ")%>>OK</option>
												<option value="OR"<% If State="OR" Then Response.Write(" selected ")%>>OR</option>
												<option value="PA"<% If State="PA" Then Response.Write(" selected ")%>>PA</option>
												<option value="PR"<% If State="PR" Then Response.Write(" selected ")%>>PR</option>
												<option value="PW"<% If State="PW" Then Response.Write(" selected ")%>>PW</option>
												<option value="RI"<% If State="RI" Then Response.Write(" selected ")%>>RI</option>
												<option value="SC"<% If State="SC" Then Response.Write(" selected ")%>>SC</option>
												<option value="SD"<% If State="SD" Then Response.Write(" selected ")%>>SD</option>
												<option value="TN"<% If State="TN" Then Response.Write(" selected ")%>>TN</option>
												<option value="TX"<% If State="TX" Then Response.Write(" selected ")%>>TX</option>
												<option value="UT"<% If State="UT" Then Response.Write(" selected ")%>>UT</option>
												<option value="VA"<% If State="VA" Then Response.Write(" selected ")%>>VA</option>
												<option value="VI"<% If State="VI" Then Response.Write(" selected ")%>>VI</option>
												<option value="VT"<% If State="VT" Then Response.Write(" selected ")%>>VT</option>
												<option value="WA"<% If State="WA" Then Response.Write(" selected ")%>>WA</option>
												<option value="WV"<% If State="WV" Then Response.Write(" selected ")%>>WV</option>
												<option value="WI"<% If State="WI" Then Response.Write(" selected ")%>>WI</option>
												<option value="WY"<% If State="WY" Then Response.Write(" selected ")%>>WY</option>
												<option value="AB"<% If State="AB" Then Response.Write(" selected ")%>>AB</option>
												<option value="BC"<% If State="BC" Then Response.Write(" selected ")%>>BC</option>
												<option value="MB"<% If State="MB" Then Response.Write(" selected ")%>>MB</option>
												<option value="NB"<% If State="NB" Then Response.Write(" selected ")%>>NB</option>
												<option value="NL"<% If State="NL" Then Response.Write(" selected ")%>>NL</option>
												<option value="NS"<% If State="NS" Then Response.Write(" selected ")%>>NS</option>
												<option value="NU"<% If State="NU" Then Response.Write(" selected ")%>>NU</option>
												<option value="ON"<% If State="ON" Then Response.Write(" selected ")%>>ON</option>
												<option value="PE"<% If State="PE" Then Response.Write(" selected ")%>>PE</option>
												<option value="QC"<% If State="QC" Then Response.Write(" selected ")%>>QC</option>
												<option value="SK"<% If State="SK" Then Response.Write(" selected ")%>>SK</option>
												<option value="NT"<% If State="NT" Then Response.Write(" selected ")%>>NT</option>
												<option value="YT"<% If State="YT" Then Response.Write(" selected ")%>>YT</option>												
 											</select>
										</div>
									</div>
								</div>
								
		  		          		<div class="col-lg-12 col-line">
									<div class="form-group">
										<label   class="col-lg-4 control-label">Company Logo</label>
										<div class="col-lg-8">
											<img src="<%= BaseURL %>clientfilesV/<%= MUV_Read("ClientID") %>/logos/logo.png">
											<input type="file" name="txtCompanyLogo" id="txtCompanyLogo" class="form-control" maxlength="50" accept="image/*" onchange="CheckExtension(this)" />
											<strong>Image must be in .png format with a transparent background</strong>
									    </div>
									 </div>
								</div>		
								
								<div class="col-lg-12 col-line">
									<div class="form-group">
										<label   class="col-lg-4 control-label">Email</label>
										<div class="col-lg-8">
											<input type="text"  id="txtEmail" name="txtEmail" class="form-control" value="<%=Email%>">
										</div>
									</div>
								</div>

								<div class="col-lg-12 col-line">
									<div class="form-group">
										<label   class="col-lg-4 control-label">A/R Attention</label>
										<div class="col-lg-8">
											<input type="text"  id="txtAttention" name="txtAttention" class="form-control" value="<%=Attention%>">
										</div>
									</div>
								</div>

								<div class="col-lg-12 col-line">
									<div class="form-group">
										<label   class="col-lg-4 control-label">Message to print</label>
										<div class="col-lg-8">
											<textarea class="form-control"  id="txtMessageToPrint" name="txtMessageToPrint" rows="3"><%=MessageToPrint%></textarea>
										</div>
									</div>
								</div>

												
								
							</div>
							<!---first div end --->
							
							
							

							<div class="col-lg-7">
							<div class="col-lg-6 col-line">
								<div class="form-group">
									<label   class="col-lg-5 control-label">Phone 1</label>
									<div class="col-lg-7">
										<input type="text"  id="txtPhone1" name="txtPhone1" class="form-control" value="<%=Phone1%>">
									</div>
								</div>
							</div>
			          
							<div class="col-lg-6 col-line">
								<div class="form-group">
									<label   class="col-lg-5 control-label">Phone2</label>
									<div class="col-lg-7">
										<input type="text"  id="txtPhone2" name="txtPhone2" class="form-control" value="<%=Phone2%>">
									</div>
								</div>
							</div>

							<div class="col-lg-6 col-line">
								<div class="form-group">
									<label   class="col-lg-5 control-label">Phone3</label>
									<div class="col-lg-7">
										<input type="text"  id="txtPhone3" name="txtPhone3" class="form-control" value="<%=Phone3%>">
									</div>
								</div>
							</div>

							<div class="col-lg-6 col-line">
								<div class="form-group">
									<label   class="col-lg-5 control-label">Fax</label>
									<div class="col-lg-7">
										<input type="text"  id="txtFax" name="txtFax" class="form-control" value="<%=Fax%>">
									</div>
								</div>
							</div>
							
							<div class="col-lg-12 col-line"> 
								<div class="form-group">
									<label  class="col-lg-2 control-label">Time Zone</label>
									<div class="col-lg-3 m-l-23">
										<select class="form-control" id="selTimezone" name="selTimezone">
											<option value="Eastern" <% If Timezone="Eastern" Then Response.Write(" selected ")%>>Eastern</option>
											<option value="Central"<% If Timezone="Central" Then Response.Write(" selected ")%>>Central</option>
											<option value="Mountain"<% If Timezone="Mountain" Then Response.Write(" selected ")%>>Mountain</option>
											<option value="Pacific"<% If Timezone="Pacific" Then Response.Write(" selected ")%>>Pacific</option>
										</select>
									</div>
									
									<label  class="col-lg-6 control-label"></label>
								</div>
							</div>

							<div class="col-lg-12 col-line">
								<div class="form-group">
									<label   class="col-lg-4 control-label">Business Day Start Time</label>
									<div class="col-lg-2 m-l-14">
										<input type="text" id="txtBusinessDayStart"  name="txtBusinessDayStart" value="<%= BusinessDayStartz  %>" class="form-control"  />
									</div>
										<label  class="col-lg-6 control-label"></label>
								 </div>
							</div>
							<div class="col-lg-12 col-line">
								<div class="form-group">
									<label   class="col-lg-4 control-label">Business Day End Time</label>
									<div class="col-lg-2 m-l-14">
										<input type="text" id="txtBusinessDayEnd"  name="txtBusinessDayEnd" value="<%= BusinessDayEndz %>" class="form-control"  />
									</div>
										<label  class="col-lg-6 control-label"></label>
								 </div>
							</div>
							

							<div class="col-lg-12 col-line">
								<div class="form-group">
									<label class="col-lg-3 control-label">Periods or Months</label>
									<div class="col-lg-4 m-l-23">
										<input type="radio" name="optCompanyPeriodsOrMonths" id="optCompanyPeriods" value="P" <% If PeriodsOrMonths ="P" Then Response.Write("checked")%>>&nbsp;&nbsp;Periods &nbsp;&nbsp;&nbsp;&nbsp;
										<input type="radio" name="optCompanyPeriodsOrMonths" id="optCompanyMonths" value="M" <% If PeriodsOrMonths ="M" OR PeriodsOrMonths = "" Then Response.Write("checked")%>>&nbsp;&nbsp;Months
									</div>
									<div class="col-lg-4">
										
									</div>
									
								 </div>
							</div>	
							
							
							<div class="col-lg-12 col-line">
								<div class="form-group">
									<label   class="col-lg-7 control-label">Company Identity Color 1 (darker color - used for text)</label>
									<div class="col-lg-5">
										<input type='text' id="txtCompanyIdentityColor1" name="txtCompanyIdentityColor1" value="<%= CompanyIdentityColor1 %>">
									</div>
								</div>
							</div>
							
							
							<div class="col-lg-12 col-line">
								<div class="form-group">
									<label   class="col-lg-7 control-label">Company Identity Color 2 (dark or light color)</label>
									<div class="col-lg-5">
										<input type='text' id="txtCompanyIdentityColor2" name="txtCompanyIdentityColor2" value="<%= CompanyIdentityColor2 %>">
									</div>
								</div>
							</div>

							
							<div class="col-lg-12 col-line">
								<div class="form-group">
									<label class="col-lg-3 control-label">Point Of Service Logic</label>
									<div class="col-lg-4 m-l-23">
										<input type="checkbox" id="chkPointOfServiceLogicOnOff" name="chkPointOfServiceLogicOnOff" <%=PointOfServiceLogicStatus%>>
									</div>
									<div class="col-lg-4">
										
									</div>
									
								 </div>
							</div>
							
						</div>
					</div>

					<Br>  	 
					<a href="#" onClick="window.location.reload();"><button type="button" class="btn btn-default">Cancel</button></a> 
            
					<button type="submit" class="btn btn-primary"><i class="far fa-save"></i> Save</button>       
 
                   
					<!-- splitter !-->
					<div class="row">
						<div class="col-lg-12">
						<hr />
						</div>
					</div>
					<!-- eof splitter !-->
					</form>
					<!-- spectrum color picker js !-->
					<script>
					$("#txtCompanyIdentityColor1").spectrum({
						color: '<%= CompanyIdentityColor1 %>',
						showInput: true,
						className: "full-spectrum",
						showInitial: true,
						showPalette: true,
						showSelectionPalette: true,
						maxSelectionSize: 10,
						preferredFormat: "hex",
						localStorageKey: "spectrum.demo",
						move: function (color) {
							
						},
						show: function () {
						
						},
						beforeShow: function () {
						
						},
						hide: function () {
						
						},
						change: function() {
							
						},
						palette: [
							["rgb(0, 0, 0)", "rgb(67, 67, 67)", "rgb(102, 102, 102)",
							"rgb(204, 204, 204)", "rgb(217, 217, 217)","rgb(255, 255, 255)"],
							["rgb(152, 0, 0)", "rgb(255, 0, 0)", "rgb(255, 153, 0)", "rgb(255, 255, 0)", "rgb(0, 255, 0)",
							"rgb(0, 255, 255)", "rgb(74, 134, 232)", "rgb(0, 0, 255)", "rgb(153, 0, 255)", "rgb(255, 0, 255)"], 
							["rgb(230, 184, 175)", "rgb(244, 204, 204)", "rgb(252, 229, 205)", "rgb(255, 242, 204)", "rgb(217, 234, 211)", 
							"rgb(208, 224, 227)", "rgb(201, 218, 248)", "rgb(207, 226, 243)", "rgb(217, 210, 233)", "rgb(234, 209, 220)", 
							"rgb(221, 126, 107)", "rgb(234, 153, 153)", "rgb(249, 203, 156)", "rgb(255, 229, 153)", "rgb(182, 215, 168)", 
							"rgb(162, 196, 201)", "rgb(164, 194, 244)", "rgb(159, 197, 232)", "rgb(180, 167, 214)", "rgb(213, 166, 189)", 
							"rgb(204, 65, 37)", "rgb(224, 102, 102)", "rgb(246, 178, 107)", "rgb(255, 217, 102)", "rgb(147, 196, 125)", 
							"rgb(118, 165, 175)", "rgb(109, 158, 235)", "rgb(111, 168, 220)", "rgb(142, 124, 195)", "rgb(194, 123, 160)",
							"rgb(166, 28, 0)", "rgb(204, 0, 0)", "rgb(230, 145, 56)", "rgb(241, 194, 50)", "rgb(106, 168, 79)",
							"rgb(69, 129, 142)", "rgb(60, 120, 216)", "rgb(61, 133, 198)", "rgb(103, 78, 167)", "rgb(166, 77, 121)",
							"rgb(91, 15, 0)", "rgb(102, 0, 0)", "rgb(120, 63, 4)", "rgb(127, 96, 0)", "rgb(39, 78, 19)", 
							"rgb(12, 52, 61)", "rgb(28, 69, 135)", "rgb(7, 55, 99)", "rgb(32, 18, 77)", "rgb(76, 17, 48)"]
						]
					});

					$("#txtCompanyIdentityColor2").spectrum({
						color: '<%= CompanyIdentityColor2 %>',
						showInput: true,
						className: "full-spectrum",
						showInitial: true,
						showPalette: true,
						showSelectionPalette: true,
						maxSelectionSize: 10,
						preferredFormat: "hex",
						localStorageKey: "spectrum.demo",
						move: function (color) {
							
						},
						show: function () {
						
						},
						beforeShow: function () {
						
						},
						hide: function () {
						
						},
						change: function() {
							
						},
						palette: [
							["rgb(0, 0, 0)", "rgb(67, 67, 67)", "rgb(102, 102, 102)",
							"rgb(204, 204, 204)", "rgb(217, 217, 217)","rgb(255, 255, 255)"],
							["rgb(152, 0, 0)", "rgb(255, 0, 0)", "rgb(255, 153, 0)", "rgb(255, 255, 0)", "rgb(0, 255, 0)",
							"rgb(0, 255, 255)", "rgb(74, 134, 232)", "rgb(0, 0, 255)", "rgb(153, 0, 255)", "rgb(255, 0, 255)"], 
							["rgb(230, 184, 175)", "rgb(244, 204, 204)", "rgb(252, 229, 205)", "rgb(255, 242, 204)", "rgb(217, 234, 211)", 
							"rgb(208, 224, 227)", "rgb(201, 218, 248)", "rgb(207, 226, 243)", "rgb(217, 210, 233)", "rgb(234, 209, 220)", 
							"rgb(221, 126, 107)", "rgb(234, 153, 153)", "rgb(249, 203, 156)", "rgb(255, 229, 153)", "rgb(182, 215, 168)", 
							"rgb(162, 196, 201)", "rgb(164, 194, 244)", "rgb(159, 197, 232)", "rgb(180, 167, 214)", "rgb(213, 166, 189)", 
							"rgb(204, 65, 37)", "rgb(224, 102, 102)", "rgb(246, 178, 107)", "rgb(255, 217, 102)", "rgb(147, 196, 125)", 
							"rgb(118, 165, 175)", "rgb(109, 158, 235)", "rgb(111, 168, 220)", "rgb(142, 124, 195)", "rgb(194, 123, 160)",
							"rgb(166, 28, 0)", "rgb(204, 0, 0)", "rgb(230, 145, 56)", "rgb(241, 194, 50)", "rgb(106, 168, 79)",
							"rgb(69, 129, 142)", "rgb(60, 120, 216)", "rgb(61, 133, 198)", "rgb(103, 78, 167)", "rgb(166, 77, 121)",
							"rgb(91, 15, 0)", "rgb(102, 0, 0)", "rgb(120, 63, 4)", "rgb(127, 96, 0)", "rgb(39, 78, 19)", 
							"rgb(12, 52, 61)", "rgb(28, 69, 135)", "rgb(7, 55, 99)", "rgb(32, 18, 77)", "rgb(76, 17, 48)"]
						]
					});


					</script>

					<!-- time picker js !-->
					<script type="text/javascript">
						$('#txtBusinessDayStart').timepicker();
						$('#txtBusinessDayEnd').timepicker();
					</script>
					<!-- eof time picker js !-->  					
				</div>
				
				
				<div role="tabpanel" class="tab-pane fade" id="calendar1"> 
				
<%


Server.ScriptTimeout = 500

SQLBuildCalendarDataSource = "SELECT * FROM Settings_CompanyCalendar ORDER BY YearNum"

Set cnnBuildCalendarDataSource = Server.CreateObject("ADODB.Connection")
cnnBuildCalendarDataSource.open (Session("ClientCnnString"))
Set rsBuildCalendarDataSource = Server.CreateObject("ADODB.Recordset")
rsBuildCalendarDataSource.CursorLocation = 3 

Set rsBuildCalendarDataSource = cnnBuildCalendarDataSource.Execute(SQLBuildCalendarDataSource)

If not rsBuildCalendarDataSource.EOF Then

	DayCount = 0
	jsonDataCalendar = ""
	
	Do While Not rsBuildCalendarDataSource.EOF
	
		MonthNam = rsBuildCalendarDataSource("MonthNam")
		MonthNum = rsBuildCalendarDataSource("MonthNum")
		MonthNum = cInt(MonthNum) - 1
		DayNum = rsBuildCalendarDataSource("DayNum")
		YearNum = rsBuildCalendarDataSource("YearNum")
		OpenClosedCloseEarly = rsBuildCalendarDataSource("OpenClosedCloseEarly")
		AlterDate=rsBuildCalendarDataSource("AlternateDeliveryDate")
		If OpenClosedCloseEarly = "Closed" Then
			ClosingTime = ""
			closingEarlyTime12Hour = ""
		Else
		
			ClosingTime = rsBuildCalendarDataSource("ClosingTime")
			ClosingTime = FormatDateTime(ClosingTime, 4)
			
			closingEarlyHour = cInt(hour(ClosingTime))
			closingEarlyMinute = cInt(minute(ClosingTime))
			
			
			If (closingEarlyHour = 0) Then
			     closingEarlyHour12Hour = 12
			     closingEarlyAMPM = "PM"
			ElseIf (closingEarlyHour > 12) Then
			     closingEarlyHour12Hour = closingEarlyHour - 12
			     closingEarlyAMPM = "PM"
			ElseIf (closingEarlyHour < 12) Then
				closingEarlyHour12Hour = closingEarlyHour
				closingEarlyAMPM = "AM"
			ElseIf (closingEarlyHour = 12) Then
				closingEarlyHour12Hour = closingEarlyHour
				closingEarlyAMPM = "PM"
			End If
			
			If closingEarlyMinute = 0 Then
				closingEarlyMinute12Hour = "00"
			Else
				closingEarlyMinute12Hour = closingEarlyMinute
			End If
			
			closingEarlyTime12Hour = closingEarlyHour12Hour & ":" & closingEarlyMinute12Hour & " " & closingEarlyAMPM
			'closingEarlyTime12Hour = FormatDateTime(closingEarlyTime12Hour, 4) 
			
		End If
		
		Description = rsBuildCalendarDataSource("Description")
		Description = Replace(Description,"'","\'")
		
		'FullDate = MonthNum & "/" & DayNum & "/" & YearNum
		'FullDate = FormatDateTime(FullDate,2)

		If DayCount = 0 Then
			jsonDataCalendar = "["
		End If

		jsonDataCalendar = jsonDataCalendar & "{businessDayID:" & DayCount & ","
		jsonDataCalendar = jsonDataCalendar & "businessDayDescription:'" & Description & "',"
		jsonDataCalendar = jsonDataCalendar & "businessDayStatus:'" & OpenClosedCloseEarly & "',"
		jsonDataCalendar = jsonDataCalendar & "closeEarlyTime:'" & closingEarlyTime12Hour & "',"
        IF NOT IsNull(AlterDate) THEN
            IF LEN(CSTR(AlterDate))>0 THEN
		        jsonDataCalendar = jsonDataCalendar & "alterDate:'" &FormatDateTime(AlterDate,2) & "',"
                ELSE
                    jsonDataCalendar = jsonDataCalendar & "alterDate:'',"
            END IF
            ELSE
            jsonDataCalendar = jsonDataCalendar & "alterDate:'',"
        END IF
		jsonDataCalendar = jsonDataCalendar & "startDate:new Date(" & YearNum & "," & MonthNum & "," & DayNum & "),"
		jsonDataCalendar = jsonDataCalendar & "endDate:new Date(" & YearNum & "," & MonthNum & "," & DayNum & ")},"

		DayCount = DayCount + 1
		rsBuildCalendarDataSource.MoveNext
		
	Loop
	
	If Len(jsonDataCalendar)>0 Then jsonDataCalendar = Left(jsonDataCalendar,Len(jsonDataCalendar)-1)
	jsonDataCalendar = jsonDataCalendar & "]"
	
	
End If

'************************************************************************************************
'Get the minimum date to show based on  the oldest year is in Settings_CompanyCalendar
'************************************************************************************************
SQLBuildCalendarDataSource = "SELECT * FROM Settings_CompanyCalendar ORDER BY YearNum ASC"

Set cnnBuildCalendarDataSource = Server.CreateObject("ADODB.Connection")
cnnBuildCalendarDataSource.open (Session("ClientCnnString"))
Set rsBuildCalendarDataSource = Server.CreateObject("ADODB.Recordset")
rsBuildCalendarDataSource.CursorLocation = 3 

Set rsBuildCalendarDataSource = cnnBuildCalendarDataSource.Execute(SQLBuildCalendarDataSource)

If not rsBuildCalendarDataSource.EOF Then
	MinYearNum = rsBuildCalendarDataSource("YearNum")
Else
	MinYearNum = Year(Date())
End If

MinDateToShow = "1/1/" & MinYearNum 
'************************************************************************************************
					
Set rsBuildCalendarDataSource = Nothing
cnnBuildCalendarDataSource.Close
Set BuildCalendarDataSource = nothing

'Response.write("<br><br><br>MinDateToShow : " & MinDateToShow)
%>


<!-- local custom css !-->
<style type="text/css">
	.form-control{
		overflow-x: hidden;
		}
		
	.post-labels{
 		padding-top: 5px;
 	}
 	.row-margin{
	 	margin-bottom: 20px;
	 	margin-top: 20px;
 	}
 	
 	h3{
	 	margin-top: 0px;
 	}
 	
 	.table-size .category{
	 	width: 35%;
	 	font-weight: normal;
 	}
 	
 	.table-size .group-name{
	 	width: 40%
 	}
 	
 	.table-size .sort-order{
	 	width: 10%;
 	}
 	
 	.table-size .display{
	 	width: 15%;
 	}
 
	 .col-line{
		 margin-bottom: 20px;
	  }
     #alterdate {    padding: 6px 12px;}
     .ui-datepicker {
   background: #ffffff;
   border: 1px solid #555;
   color: #000000;
 }

</style>
<!-- eof local custom css !-->

<!--<h1 class="page-header"><i class="fa fa-calendar" aria-hidden="true"></i> Company Calendar</h1>-->


<!-- tabs start here !-->
<div class="row ">
	<div class="col-lg-12">
			<div class="row">
			<div class="col-lg-12 col-line">
				<div class="panel panel-default" style="margin:10px;">
					<div class="panel-heading">Choose dates that your company is closed or will be Close Early. By default, weekends are disabled and all weekdays are considerd open, until you click the date to change its status.</div>
					<div class="panel-body">
						<div id="calendar"></div>
					</div>
				</div>
			</div>
		</div>
		
	</div>
</div>


<script type="text/javascript">


$(function() {
    
	$('#updateCompanyCalendarModal').on('shown.bs.modal', function (e) {
	
	 	var businessDayStatus = $('#businessDayStatusHidden').val();

	 	if (businessDayStatus == 'Open') {
	 		 $("#radOpen").prop("checked",true);
	 		 $("#txtBusinessDayDescription").val('');
	 		 $("#closeEarlyTimepicker").val('');
	 		 $("#closingEarlyTimeDiv").hide();
	 	}
	 	else if (businessDayStatus == 'Closed') {
	 		$("#radClosed").prop("checked",true);
	 		$("#closingEarlyTimeDiv").hide();
	 	}
	 	else if (businessDayStatus == 'Close Early') {
	 		$("#radClosingEarly").prop("checked",true);
	 		$("#closingEarlyTimeDiv").show();
	 	}
	
	});		
	
	
	//var firstOfThisYear = new Date(new Date().getFullYear(), 0, 1);
	var firstOfThisYear = '<%= MinDateToShow %>';
	var currentYear = new Date().getFullYear();
	
	    

    $('#calendar').calendar({ 
    	/* Set options here */ 
    	
    	//This disables Saturday and Sunday as business days
    	disabledWeekDays: [0,6],
    	minDate:firstOfThisYear,
    	startYear:currentYear,
    	style: 'custom',
    	
        //This is used to color weekends gray, since they are disabled
        //It colors all other days, green, which means the company is open
        //The customDataSourceRenderer that follows will color closed and Close Early days
        customDayRenderer: function(element, date) {
        
        	var day = date.getDay();
			var isWeekend = (day == 6) || (day == 0);    // 6 = Saturday, 0 = Sunday
			
			if (isWeekend) {
                $(element).css('font-weight', 'normal');
                $(element).css('font-size', '14px');
                $(element).css('color', '#DCDCDC');
            }
            else {
            	$(element).css('font-weight', 'normal');
                $(element).css('font-size', '14px');
                $(element).css('color', '#5cb85c');

            }
        },	 
    	
    	//This will style the calendar days as red for closed and yellow for Close Early
		customDataSourceRenderer: function(element, date, events) {
		
                for(var i in events) {	                                
				    if (events[i].businessDayStatus == 'Closed') {
						$(element).css('background-color', 'red');
						$(element).css('color', 'white');
						$(element).css('border-radius', '5px');
					}   
				    if (events[i].businessDayStatus == 'Close Early') {
		                $(element).css('background-color', '#F5BB00');
		                $(element).css('color', 'white');
		                $(element).css('border-radius', '5px');
					}    
				    else if ((events[i].businessDayStatus !== 'Close Early') && (events[i].businessDayStatus !== 'Closed')){
		                $(element).css('font-weight', 'normal');
		                $(element).css('font-size', '14px');
		                $(element).css('color', '#5cb85c');
					}    												        
                }
		},
		enableContextMenu: false,
				 
		//When the user clicks on a calendar date, this function passes the needed information to the modal window.
		//A day has an "event" when it is either closed or Close Early.       
        clickDay: function(e) {
            if(e.events.length > 0) {
                for(var i in e.events) {
					
					clickedDateFormatted = moment(e.events[i].startDate).format('MM/DD/YYYY');
						                                
				    $('#updateCompanyCalendarModal input[name="businessDayID"]').val(e.events[i].businessDayID);
				    $('#updateCompanyCalendarModal #txtBusinessDayDescription').val(e.events[i].businessDayDescription);
				    $('#updateCompanyCalendarModal #selectedDate').html(clickedDateFormatted);
				    $('#updateCompanyCalendarModal input[name="dateToEdit"]').val(clickedDateFormatted);
				    $('#updateCompanyCalendarModal #closingEarlyTime').val(e.events[i].closeEarlyTime);
				    $('#updateCompanyCalendarModal input[id="closeEarlyTimepicker"]').val(e.events[i].closeEarlyTime);
				    $('#updateCompanyCalendarModal #businessDayStatus').html(e.events[i].businessDayStatus);   
                    $('#updateCompanyCalendarModal input[name="businessDayStatusHidden"]').val(e.events[i].businessDayStatus);
                    if (e.events[i].alterDate.length > 0) $("#alterdate").val(moment(e.events[i].alterDate).format('MM/DD/YYYY'));
                    else $("#alterdate").val("");
                    if (e.events[i].businessDayStatus == 'Close Early') $(".date-alter>label").html("Orders received after the cutoff time specified above should have their delivery date set to");
                    else $(".date-alter>label").html("Reschedule this day's deliveries for");
                    $(".date-alter").removeClass("hidden");
                }
            }
            else {
				    clickedDateFormatted = moment(e.date).format('MM/DD/YYYY');
				    $(".date-alter").addClass("hidden");
				    $('#updateCompanyCalendarModal #selectedDate').html(clickedDateFormatted);
				    $('#updateCompanyCalendarModal input[name="dateToEdit"]').val(clickedDateFormatted);
				    $('#updateCompanyCalendarModal input[name="businessDayStatusHidden"]').val('Open');
				    $('#updateCompanyCalendarModal #businessDayStatus').html('Open');
    

            }
            $('#updateCompanyCalendarModal').modal();
			    

        },
        
        //This function is used to show a popover div on a closed or Close Early business day
        //It will show the description, status and Close Early time (if applicable) when the user mouses over the calendar
        
        mouseOnDay: function(e) {
            if(e.events.length > 0) {
                var content = '';
                
                for(var i in e.events) {
                
                	if (e.events[i].businessDayStatus == 'Close Early') {
                	
                        content += '<div class="event-tooltip-content">'
                            + '<div class="event-description" style="color:' + e.events[i].color + '">' + e.events[i].businessDayDescription + '</div>'
                            + '<div class="event-status">' + e.events[i].businessDayStatus + ' ' + e.events[i].closeEarlyTime + '</div>';
                        if (e.events[i].alterDate.length > 0) content += '<div class="event-status">Alternate delivery date:' + e.events[i].alterDate + '</div>';

                                content += '</div>';
                	}
                	else {
                        content += '<div class="event-tooltip-content">'
                            + '<div class="event-description" style="color:' + e.events[i].color + '">' + e.events[i].businessDayDescription + '</div>'
                            + '<div class="event-status">' + e.events[i].businessDayStatus + '</div>';
                            if (e.events[i].alterDate.length > 0) content += '<div class="event-status">Alternate delivery date:' + e.events[i].alterDate + '</div>';

                                content += '</div>';
                               
                	}
                }
            
                $(e.element).popover({ 
                    trigger: 'manual',
                    container: 'body',
                    html:true,
                    content: content
                });
                
                $(e.element).popover('show');
            }
        },
        mouseOutDay: function(e) {
            if(e.events.length > 0) {
                $(e.element).popover('hide');
            }
        },
        dayContextMenu: function(e) {
            $(e.element).popover('hide');
        },
    dataSource:<%= jsonDataCalendar %>
});


});
</script>


<!-- splitter !-->
<div class="row">
	<div class="col-lg-12">
	<hr />
	</div>
</div>
<!-- eof splitter !-->


<div class="modal modal-fade" id="updateCompanyCalendarModal" style="display: none;">
	<div class="modal-dialog">
		<div class="modal-content">
			<script language="JavaScript">
			<!--

			   function validateCalendarChange()
			    {
				   var selectedBusinessDayStatus = $("input[name=radUpdatedDateStatus]:checked").val()
				   var enteredBusinessDayDesc = $("#txtBusinessDayDescription").val();
				   var enteredCloseEarlyTime = $("#closeEarlyTimepicker").val();
				   		    
			       if ((selectedBusinessDayStatus == "Closed" || selectedBusinessDayStatus == "Close Early") && enteredBusinessDayDesc == "") {
			            swal("Please enter a description for the calendar date.");
			            return false;
			       }

			       if (selectedBusinessDayStatus == "Close Early" && enteredCloseEarlyTime == "") {
			            swal("Please enter the time you will be closing early.");
			            return false;
			       }
			
			       return true;
			    }
			// -->
			</script>  
		
			<script>
			
				$(document).ready(function() {
					
					$("#closingEarlyTimeDiv").hide();
                    $("#alterdate").datepicker().next("button").button({
                    icons: {
                        primary: "glyphicon glyphicon-calendar"
                    }});

                    $("input[name='radUpdatedDateStatus']").on("click",function(){
                
                        if ($(this).val()=="Open") {
                            $(".date-alter").addClass("hidden");
                            $("#alterdate").val("");

                        }
                        else {
                            if ($(this).val()=="Closed") $(".date-alter>label").html("Reschedule this day's deliveries for");
                            else $(".date-alter>label").html("Orders received after the cutoff time specified above should have their delivery date set to");
                            $(".date-alter").removeClass("hidden");

                        }
                    });

                    $('.clockpicker').clockpicker({
					    placement: 'top',
					    align: 'left',
					    donetext: 'Done',
					    twelvehour:true
					});
						 
					$('input[type=radio][name=radUpdatedDateStatus]').on('change', function() {
					     if ($(this).val() == 'Close Early') {
					     	$("#closingEarlyTimeDiv").show();				     	
					     }
					     else if ($(this).val() == 'Open') {
					     	$("#closingEarlyTimeDiv").hide();
					     	$("#txtBusinessDayDescription").val('');
					     }
					     else if ($(this).val() == 'Closed') {
					     	$("#closingEarlyTimeDiv").hide();
						 } 
					});	  
					
				}); //end document.ready() function
			
			</script>
		
			<div class="modal-header">
				<button type="button" class="close" data-dismiss="modal"><span aria-hidden="true">×</span><span class="sr-only">Close</span></button>
				<h4 class="modal-title">
					Update Company Calendar for <span id="selectedDate"></span>
				</h4>
			</div>
			<form class="form-horizontal" method="post" name="frmUpdateCompanyCalendar" id="updateCompanyCalendar" action="calendarSaveFromModal.asp" onsubmit="return validateCalendarChange()">
			<div class="modal-body clearfix">
					<input type="hidden" name="businessDayID" value="">
					<input type="hidden" name="dateToEdit" id="dateToEdit" value="">
					<input type="hidden" name="closingEarlyTime" id="closingEarlyTime" value="">
					<input type="hidden" name="businessDayStatusHidden" id="businessDayStatusHidden" value="">
					
					<div class="form-group clearfix">
						<label for="min-date" class="col-sm-4 control-label">Description</label>
						<div class="col-sm-8">
							<input name="txtBusinessDayDescription" id="txtBusinessDayDescription" type="text" class="form-control">
						</div>
					</div>
					<div class="form-group clearfix">
						<label for="min-date" class="col-sm-4 control-label">Current Status</label>
						<div class="col-sm-8">
							<span id="businessDayStatus"></span>
						</div>
					</div>
					<div class="form-group clearfix">
						<label for="min-date" class="col-sm-4 control-label">Change Status To (Open, Closed, Close Early)</label>
						<div class="col-sm-8">
				
							<div class="radio">
							  <label><input type="radio" name="radUpdatedDateStatus" value="Open" id="radOpen">Open</label>
							</div>
							<div class="radio">
							  <label><input type="radio" name="radUpdatedDateStatus" value="Closed" id="radClosed">Closed</label>
							</div>
							<div class="radio">
							  <label><input type="radio" name="radUpdatedDateStatus" value="Close Early" id="radClosingEarly">Close Early</label>
							</div>
						</div>
					</div>
					<div class="form-group clearfix" id="closingEarlyTimeDiv">
						<label for="min-date" class="col-sm-4 control-label">Early Close Time</label>
						<div class="col-sm-8">
						
							<div class="input-group clockpicker" style="width:150px">
							    <input type="text" class="form-control" name="closeEarlyTimepicker" id="closeEarlyTimepicker" value="">
							    <span class="input-group-addon">
							        <span class="glyphicon glyphicon-time"></span>
							    </span>
							</div>
						</div>
					</div>
                    <div class="form-group clearfix date-alter hidden">
                        <label for="alter-date" class="col-sm-4 control-label">Reschedule this day's deliveries for</label>

                        <div class="col-sm-8">
                            <label for="alter-date" class="col-sm-12 control-label">Alternate delivery date</label>
                            <div class=" input-group">
                                
                                <input type="text" id="alterdate" name="alterdate" class="col-md-12">
                                <span class="input-group-addon">
                                    <span class="glyphicon glyphicon-calendar"></span>
                                </span>
                            </div>
                        </div>

                       
                    </div>
			</div>
			<div class="modal-footer">
				<button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
				<button type="submit" class="btn btn-primary" id="save-event">
					Save
				</button>
			</div>
			</form>
		</div>
	</div>
</div>
<div id="context-menu">
</div>
<style>
.event-tooltip-content:not(:last-child) {
	border-bottom:1px solid #ddd;
	padding-bottom:5px;
	margin-bottom:5px;
}

.event-tooltip-content .event-title {
	font-size:18px;
}

.event-tooltip-content .event-status {
	font-size:12px;
}
</style>

				
				</div>



<div role="tabpanel" class="tab-pane fade" id="reportperiod">
<!-- function that gets the value of the tab when it is clicked and then
updates the value of a hidden form field so when the page posts, it returns
back to the tab that was previously opened -->

<script type="text/javascript">
	
	$(document).ready(function(){
	
		$("#mytable #checkall").click(function () {
	        if ($("#mytable #checkall").is(':checked')) {
	            $("#mytable input[type=checkbox]").each(function () {
	                $(this).prop("checked", true);
	            });
	
	        } else {
	            $("#mytable input[type=checkbox]").each(function () {
	                $(this).prop("checked", false);
	            });
	        }
	    });
		    
		 $("[data-toggle=tooltip]").tooltip();

		$('#periodStartDateEdit').datetimepicker({
		    format: 'MM/DD/YYYY',
		    widgetPositioning: {
		        vertical: 'bottom',
		        horizontal: 'auto'
		    }
		});
		
		$('#periodEndDateEdit').datetimepicker({
		    format: 'MM/DD/YYYY',
		    widgetPositioning: {
		        vertical: 'bottom',
		        horizontal: 'auto'
		    }
		});
		

		$('#periodStartDateAdd').datetimepicker({
		    format: 'MM/DD/YYYY',
		    widgetPositioning: {
		        vertical: 'bottom',
		        horizontal: 'auto'
		    }
		});
		
		$('#periodEndDateAdd').datetimepicker({
		    format: 'MM/DD/YYYY',
		    widgetPositioning: {
		        vertical: 'bottom',
		        horizontal: 'auto'
		    }
		});
		
        $('#periodYearAdd').datetimepicker({
            viewMode: 'years',
            format: 'YYYY'
        });		
	
		
		$('#addCompanyReportPeriod').on('shown.bs.modal', function (e) {
			document.location.hash = 'reportperiod';
		 	
		});		


		
		$('#editCompanyReportPeriod').on('shown.bs.modal', function (e) {
			
			document.location.hash = 'reportperiod';
		
		 	var periodYear = $(e.relatedTarget).attr('data-period-year');
		 	var periodNum = $(e.relatedTarget).attr('data-period-num');
		 	var periodStartDate = $(e.relatedTarget).attr('data-period-start');
		 	var periodEndDate = $(e.relatedTarget).attr('data-period-end');
		 	var periodIntRecID = $(e.relatedTarget).attr('data-record-id');
		 	
	    	var $modal = $(this);

			 $modal.find('#periodYearEdit').empty().append("<strong>Period Year</strong>: " + periodYear);
			 $modal.find('#periodNumEdit').empty().append("<strong>Period</strong>: " + periodNum);
			 
			 $("#txtPeriodNumEdit").val(periodNum);
			 $("#txtPeriodYearEdit").val(periodYear);
			 $("#txtIntRecID").val(periodIntRecID);
		 	 $("#periodStartDateEdit").datetimepicker("defaultDate", periodStartDate);
		 	 $("#txtPeriodStartDateEdit").val(periodStartDate);
		 	 $("#periodEndDateEdit").datetimepicker("defaultDate", periodEndDate);
		 	 $("#txtPeriodEndDateEdit").val(periodEndDate);
			 $("#txtIntRecID").val(periodIntRecID);
		});		


		$('#deleteCompanyReportPeriod').on('show.bs.modal', function(e) {

	    	var $modal = $(this);
			var chkBoxArray = [];
			$(".checkthis:checked").each(function() {
			    chkBoxArray.push(this.id);
			});			
	    	
	    	if (chkBoxArray.length > 0) {
		    	$.ajax({
					type:"POST",
					url: "../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
					cache: false,
					data: "action=GetReportPeriodDeleteInformationForModal&reportPeriodsArray="+encodeURIComponent(chkBoxArray),
					success: function(response)
					 {
		               	 $modal.find('#deleteReportPeriodsInfo').html(response);	               	 
		             },
		             failure: function(response)
					 {
					   $modal.find('#deleteReportPeriodsInfo').html("Failed");
		             }
				});	
			}
			else {
				swal("Please select at least one reporting period to delete.");
			}    
		});
	
	 
	});	
	
	
</script>


<!-- local custom css !-->
<style type="text/css">
	.form-control{
		overflow-x: hidden;
		}
		
	.post-labels{
 		padding-top: 5px;
 	}
 	.row-margin{
	 	margin-bottom: 20px;
	 	margin-top: 20px;
 	}
 	
 	h3{
	 	margin-top: 0px;
 	}
 	
 	.table-size .category{
	 	width: 35%;
	 	font-weight: normal;
 	}
 	
 	.table-size .group-name{
	 	width: 40%
 	}
 	
 	.table-size .sort-order{
	 	width: 10%;
 	}
 	
 	.table-size .display{
	 	width: 15%;
 	}
 
	 .col-line{
		 margin-bottom: 20px;
	  }

</style>
<!-- eof local custom css !-->

<!--<h1 class="page-header"><i class="fa fa-calendar-o" aria-hidden="true"></i> Company Report Periods</h1>-->


		<div class="row">
			<div class="col-lg-12 col-line">
				<div class="panel panel-default" style="margin:10px;">
					<div class="panel-heading">Build your custom period date ranges for each year's reports.</div>
					<div class="panel-body">
						<div class="container">
						<div class="row">
					        <div class="col-md-12">
					        <h4>
					        <p data-placement="top" data-toggle="tooltip" title="Add Report Period"><a class="btn btn-success btn-large" data-title="Add Report Period" data-toggle="modal" data-target="#addCompanyReportPeriod"><i class="fa fa-plus-circle" aria-hidden="true"></i> Add New Report Period</a></p>					        
				            </h4>  
				            <div class="table-responsive">
				            <table id="mytable" class="table table-bordred table-striped">
				                   <thead>
				                   <th><input type="checkbox" id="checkall" /></th>
										<th>Year</th>
										<th>Period</th>
										<th>Begin Date</th>
										<th>End Date</th>
										<th>Edit</th>
										<th>Delete</th>
				                   </thead>
					    <tbody>
						<%
						
						Server.ScriptTimeout = 500
						
						SQLBuildPeriodsDataSource = "SELECT * FROM Settings_CompanyPeriods ORDER BY Year DESC, Period DESC"
						
						Set cnnBuildPeriodsDataSource = Server.CreateObject("ADODB.Connection")
						cnnBuildPeriodsDataSource.open (Session("ClientCnnString"))
						Set rsBuildPeriodsDataSource = Server.CreateObject("ADODB.Recordset")
						rsBuildPeriodsDataSource.CursorLocation = 3 
						
						Set rsBuildPeriodsDataSource = cnnBuildPeriodsDataSource.Execute(SQLBuildPeriodsDataSource)
						
						If not rsBuildPeriodsDataSource.EOF Then
						
							
							Do While Not rsBuildPeriodsDataSource.EOF
							
								IntRecID = rsBuildPeriodsDataSource("InternalRecordIdentifier")
								PeriodYear = rsBuildPeriodsDataSource("Year")
								Period = rsBuildPeriodsDataSource("Period")
								PeriodBeginDate = formatDateTime(rsBuildPeriodsDataSource("BeginDate"),2)
								PeriodEndDate = formatDateTime(rsBuildPeriodsDataSource("EndDate"),2)
								
								%>
							    <tr>
							    <td><input type="checkbox" class="checkthis" id="<%= IntRecID %>"></td>
							    <td><%= PeriodYear %></td>
							    <td><%= Period %></td>
							    <td><%= PeriodBeginDate %></td>
							    <td><%= PeriodEndDate %></td>
							    <td><p data-placement="top" data-toggle="tooltip" title="Edit Report Period"><a class="btn btn-primary btn-xs" data-period-year="<%= PeriodYear %>" data-period-num="<%= Period %>" data-period-start="<%= PeriodBeginDate %>" data-period-end="<%= PeriodEndDate %>" data-record-id="<%= IntRecID %>" data-title="Edit Report Period" data-toggle="modal" data-target="#editCompanyReportPeriod"><span class="glyphicon glyphicon-pencil"></span></a></p></td>
							    <td><p data-placement="top" data-toggle="tooltip" title="Delete Report Period"><a class="btn btn-danger btn-xs" data-record-id="<%= IntRecID %>" data-title="Delete Report Period" data-toggle="modal" data-target="#deleteCompanyReportPeriod" ><span class="glyphicon glyphicon-trash"></span></a></p></td>
							    </tr>
							    
								<%
								rsBuildPeriodsDataSource.MoveNext
							Loop
						Else
							%><tr><td colspan="7">No Reporting Periods Have Been Added. Please Click The Green Button Above To Start Building Your Periods.</td></tr><%							
							
						End If
											
						Set rsBuildPeriodsDataSource = Nothing
						cnnBuildPeriodsDataSource.Close
						Set cnnBuildPeriodsDataSource = nothing
						
						%>					    
					    </tbody>
					        
					</table>
					                
					            </div>
					            
					        </div>
						</div>
					</div>

					</div>
				</div>
			</div>
		</div>
		
<!---------------------------------------------------------------------------------------------------------------------------------->
<!---------------------------------------------------------------------------------------------------------------------------------->
<!-- ADD, EDIT AND DELETE MODALS FOR COMPANY REPORT PERIODS                                                                       -->
<!---------------------------------------------------------------------------------------------------------------------------------->
<!---------------------------------------------------------------------------------------------------------------------------------->


<div class="modal fade" id="addCompanyReportPeriod" tabindex="-1" role="dialog" aria-labelledby="edit" aria-hidden="true">
	<div class="modal-dialog">
		<div class="modal-content">

		    <script>
		    
				function validateAddNewPeriodFields()
			    {
							    				       
				   var selectedPeriodAdd = $("#selPeriodNumAdd option:selected").val();
				   var selectedYearAdd = $("#txtPeriodYearAdd").val();
				   var selectedStartDateAdd = $("#txtPeriodStartDateAdd").val();
				   var selectedEndDateAdd = $("#txtPeriodEndDateAdd").val();
				   		    
			       if (selectedPeriodAdd == "") {
			            swal("Please select a reporting period number.");
			            return false;
			       }	
			       if (selectedYearAdd == "") {
			            swal("Please select a reporting year.");
			            return false;
			       }						       			       
				   if (selectedStartDateAdd == "") {
			            swal("Please select a start date for this reporting period.");
			            return false;
			       }	
				   if (selectedEndDateAdd == "") {
			            swal("Please select an end date for this reporting period.");
			            return false;
			       }	
			       
					var d1 = Date.parse(selectedStartDateAdd);
					var d2 = Date.parse(selectedEndDateAdd);
					
					if (d1 > d2) {
			            swal("The end date must occur AFTER the start date.");
			            return false;
			       }	
		       			       	
			       return true;
			    }

		    
				$(document).ready(function(){
		
					$('#periodYearAdd').on("dp.change", function (e){
					
					    var selectedPeriodAdd = $("#selPeriodNumAdd").val();
					  	var selectedYearAdd = $("#txtPeriodYearAdd").val();
					  	
				  		if (selectedPeriodAdd == "" || selectedPeriodAdd == null) {
				  			selectedPeriodAdd = 1;
				  		}
				  		
				    	$.ajax({
							type:"POST",
							url: "../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
							cache: false,
							data: "action=WritePeriodsInUseDropdownForReportYearAdd&periodYear=" + encodeURIComponent(selectedYearAdd) + "&periodNum=" + encodeURIComponent(selectedPeriodAdd),
							success: function(response)
							 {
				               	 $('#selPeriodNumContainerDivAdd').empty().append(response);	               	 
				             },
				             failure: function(response)
							 {
							    $('#selPeriodNumContainerDivAdd').empty().append("Failed");
				             }
						});
					});	
				    
				
					$('#btnAddNewPeriod').on("click", function (e){
					
					    var selectedPeriodAdd = $("#selPeriodNumAdd option:selected").val();
					  	var selectedYearAdd = $("#txtPeriodYearAdd").val();
					  	var selectedStartDateAdd = $("#txtPeriodStartDateAdd").val();
					  	var selectedEndDateAdd = $("#txtPeriodEndDateAdd").val();
					    
					    
					    if (validateAddNewPeriodFields()) {
						    
					    	$.ajax({
								type:"POST",
								url: "../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
								cache: false,
								data: "action=ValidateAndAddReportPeriod&periodYear=" + encodeURIComponent(selectedYearAdd) + "&periodNum=" + encodeURIComponent(selectedPeriodAdd) + "&periodStartDate=" + encodeURIComponent(selectedStartDateAdd) + "&periodEndDate=" + encodeURIComponent(selectedEndDateAdd),
								success: function(response)
								 {
					               	 if (response == 'Success') {
					               	 	location.reload();
									 }	 
									 else {
									 	swal(response);
									 }              	 
					             },
					             failure: function(response)
								 {
								    swal("Failed");
					             }
							});
						}
					});
					

				});	//end document.ready() function
				
		    </script>
	    <div class="modal-header">
	        <button type="button" class="close" data-dismiss="modal" aria-hidden="true"><span class="glyphicon glyphicon-remove" aria-hidden="true"></span></button>
	        <h4 class="modal-title custom_align" id="Heading">Add Report Period</h4>
	    </div>
	    <div class="modal-body">
    
	        <div class="form-group">
	        	Year
	            <div class='input-group date' id='periodYearAdd'>
	                <input type='text' class="form-control" id="txtPeriodYearAdd">
	                <span class="input-group-addon">
	                    <span class="glyphicon glyphicon-calendar">
	                    </span>
	                </span>
	            </div>
	        </div>
	
			<div class="form-group" id="selPeriodNumContainerDivAdd">
				<label for="selPeriodNum">Period</label>
				<select class="form-control" id="selPeriodNum" name="selPeriodNumAdd">				
					<%
					For i = 1 To 100
					  	%><option value="<%= i %>"><%= i %></option><%
					Next
					%>				
				</select>
			</div>
				
	        <div class="form-group">
	        	Period Start Date
	            <div class='input-group date' id='periodStartDateAdd'>
	                <input type='text' class="form-control" id="txtPeriodStartDateAdd">
	                <span class="input-group-addon">
	                    <span class="glyphicon glyphicon-calendar">
	                    </span>
	                </span>
	            </div>
	        </div>
	
	        <div class="form-group">
	        	Period End Date
	            <div class='input-group date' id='periodEndDateAdd'>
	                <input type='text' class="form-control" id="txtPeriodEndDateAdd">
	                <span class="input-group-addon">
	                    <span class="glyphicon glyphicon-calendar">
	                    </span>
	                </span>
	            </div>
	        </div>
				
	      </div>
      
		<div class="modal-footer ">
			<button type="button" class="btn btn-success btn-lg" style="width: 100%;" id="btnAddNewPeriod"><i class="fa fa-plus" aria-hidden="true"></i> Add New Period</button>
		</div>

       </div>
	<!-- /.modal-content --> 
	</div>
<!-- /.modal-dialog --> 
</div>
    



<div class="modal fade" id="editCompanyReportPeriod" tabindex="-1" role="dialog" aria-labelledby="edit" aria-hidden="true">
	<div class="modal-dialog">
		<div class="modal-content">
			<form name="frmEditReportPeriods" id="frmEditReportPeriods" method="post" action="reportPeriodsUpdateFromModal.asp">
			    <script>
			    
					function validateEditPeriodFields()
				    {
								    				       
					   var selectedStartDateEdit = $("#txtPeriodStartDateEdit").val();
					   var selectedEndDateEdit = $("#txtPeriodEndDateEdit").val();
					   		    						       			       
					   if (selectedStartDateEdit == "") {
				            swal("Please select a start date for this reporting period.");
				            return false;
				       }	
					   if (selectedEndDateEdit == "") {
				            swal("Please select an end date for this reporting period.");
				            return false;
				       }	
				       
						var d1 = Date.parse(selectedStartDateEdit);
						var d2 = Date.parse(selectedEndDateEdit);
						
						if (d1 > d2) {
				            swal("The end date must occur AFTER the start date.");
				            return false;
				       }	
			       			       	
				       return true;
				    }
			    
					$(document).ready(function(){						
						
						$('#btnEditPeriod').on("click", function (e){
						
						    var selectedPeriodEdit = $("#txtPeriodNumEdit").val();
						  	var selectedYearEdit = $("#txtPeriodYearEdit").val();
						  	var selectedStartDateEdit = $("#txtPeriodStartDateEdit").val();
						  	var selectedEndDateEdit = $("#txtPeriodEndDateEdit").val();
						    var selectedIntRecIDEdit = $("#txtIntRecID").val();
						    
						    if (validateEditPeriodFields()) {
							    
						    	$.ajax({
									type:"POST",
									url: "../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
									cache: false,
									data: "action=UpdateReportPeriod&periodIntRecID=" + encodeURIComponent(selectedIntRecIDEdit) + "&periodYear=" + encodeURIComponent(selectedYearEdit) + "&periodNum=" + encodeURIComponent(selectedPeriodEdit) + "&periodStartDate=" + encodeURIComponent(selectedStartDateEdit) + "&periodEndDate=" + encodeURIComponent(selectedEndDateEdit),
									success: function(response)
									 {
						               	 if (response == 'Success') {
						               	 	location.reload();
										 }	 
										 else {
										 	swal(response);
										 }              	 
						             },
						             failure: function(response)
									 {
									    swal("Failed");
						             }
								});
							}
						});
				
	
					});	
			    </script>
		    <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-hidden="true"><span class="glyphicon glyphicon-remove" aria-hidden="true"></span></button>
		        <h4 class="modal-title custom_align" id="Heading">Edit Report Period</h4>
		    </div>
		    <div class="modal-body">
	    		<input type="hidden" name="txtIntRecID" id="txtIntRecID">
	    		<input type="hidden" name="txtPeriodYearEdit" id="txtPeriodYearEdit">
	    		<input type="hidden" name="txtPeriodNumEdit" id="txtPeriodNumEdit">
	    		
	    		<div class="form-group" id="periodYearEdit"></div>
	    				
				<div class="form-group" id="periodNumEdit"></div>
					
		        <div class="form-group">
		        	Period Start Date
		            <div class='input-group date' id='periodStartDateEdit'>
		                <input type='text' class="form-control" id="txtPeriodStartDateEdit">
		                <span class="input-group-addon">
		                    <span class="glyphicon glyphicon-calendar">
		                    </span>
		                </span>
		            </div>
		        </div>
		
		        <div class="form-group">
		        	Period End Date
		            <div class='input-group date' id='periodEndDateEdit'>
		                <input type='text' class="form-control" id="txtPeriodEndDateEdit">
		                <span class="input-group-addon">
		                    <span class="glyphicon glyphicon-calendar">
		                    </span>
		                </span>
		            </div>
		        </div>
					
		      </div>
	      
				<div class="modal-footer ">
					<button type="button" class="btn btn-primary btn-lg" style="width: 100%;" id="btnEditPeriod"><span class="fa fa-pencil"></span> Update Period</button>
				</div>
			</form>
       </div>
	<!-- /.modal-content --> 
	</div>
<!-- /.modal-dialog --> 
</div>

    
<div class="modal fade" id="deleteCompanyReportPeriod" tabindex="-1" role="dialog" aria-labelledby="edit" aria-hidden="true">
	<div class="modal-dialog">
		<div class="modal-content">
			<form name="frmDeleteReportPeriods" id="frmDeleteReportPeriods" method="post" action="reportPeriodsDeleteFromModal.asp">
			
			  	<div class="modal-header">
					<button type="button" class="close" data-dismiss="modal" aria-hidden="true"><span class="fas fa-trash-alt" aria-hidden="true"></span></button>
					<h4 class="modal-title custom_align" id="Heading">Delete Period</h4>
				</div>
				
				<div class="modal-body">
					<div class="alert alert-danger"><span class="glyphicon glyphicon-warning-sign"></span> Are you sure you want to delete the following period(s)?</div>
					<div id="deleteReportPeriodsInfo"></div>
				</div>
				
				<div class="modal-footer ">
					<button type="button" class="btn btn-success" onclick="frmDeleteReportPeriods.submit()"><i class="fas fa-trash-alt" aria-hidden="true"></i> Yes, Delete</button>
					<button type="button" class="btn btn-default" data-dismiss="modal"><i class="fa fa-ban" aria-hidden="true"></i> Nevermind, Do Not Delete</button>
				</div>
			
			</form>
		</div>
	<!-- /.modal-content --> 
</div>
<!-- /.modal-dialog --> 

<!---------------------------------------------------------------------------------------------------------------------------------->
<!---------------------------------------------------------------------------------------------------------------------------------->
<!-- END ADD, EDIT AND DELETE MODALS FOR COMPANY REPORT PERIODS                                                                       -->
<!---------------------------------------------------------------------------------------------------------------------------------->
<!---------------------------------------------------------------------------------------------------------------------------------->



</div>				
				</div>



<div role="tabpanel" class="tab-pane fade" id="accountingperiod">
<script type="text/javascript">
	
	$(document).ready(function(){
	
		
		    
		 $("[data-toggle=tooltip]").tooltip();

		$('#periodStartDateEdit1').datetimepicker({
		    format: 'MM/DD/YYYY',
		    widgetPositioning: {
		        vertical: 'bottom',
		        horizontal: 'auto'
		    }
		});
		
		$('#periodEndDateEdit1').datetimepicker({
		    format: 'MM/DD/YYYY',
		    widgetPositioning: {
		        vertical: 'bottom',
		        horizontal: 'auto'
		    }
		});
		

		$('#periodStartDateAdd1').datetimepicker({
		    format: 'MM/DD/YYYY',
		    widgetPositioning: {
		        vertical: 'bottom',
		        horizontal: 'auto'
		    }
		});
		
		$('#periodStartDateAddCond1').datetimepicker({
		    format: 'MM/DD/YYYY',
		    widgetPositioning: {
		        vertical: 'bottom',
		        horizontal: 'auto'
		    }
		});		
		
		$('#periodEndDateAdd1').datetimepicker({
		    format: 'MM/DD/YYYY',
		    widgetPositioning: {
		        vertical: 'bottom',
		        horizontal: 'auto'
		    }
		});
		
        $('#periodYearAdd1').datetimepicker({
            viewMode: 'years',
            format: 'YYYY'
        });		
	
		
		$('#addCompanyAccountingPeriod').on('shown.bs.modal', function (e) {
					
		 	document.location.hash = 'accountingperiod';
		});		


		
		$('#editCompanyAccountingPeriod').on('shown.bs.modal', function (e) {
			document.location.hash = 'accountingperiod';
		
		 	var periodYear = $(e.relatedTarget).attr('data-period-year');
		 	var periodNum = $(e.relatedTarget).attr('data-period-num');
		 	var periodStartDate = $(e.relatedTarget).attr('data-period-start');
		 	var periodEndDate = $(e.relatedTarget).attr('data-period-end');
		 	var periodIntRecID = $(e.relatedTarget).attr('data-record-id');
		 	
	    	var $modal = $(this);

			 $modal.find('#periodYearEdit1').empty().append("<strong>Fiscal Year</strong>: " + periodYear);
			 $modal.find('#periodNumEdit1').empty().append("<strong>Period</strong>: " + periodNum);
			 
			 $("#txtPeriodNumEdit1").val(periodNum);
			 $("#txtPeriodYearEdit1").val(periodYear);
			 $("#txtIntRecID1").val(periodIntRecID);
		 	 $("#periodStartDateEdit1").datetimepicker("defaultDate", periodStartDate);
		 	 $("#txtPeriodStartDateEdit1").val(periodStartDate);
		 	 $("#periodEndDateEdit1").datetimepicker("defaultDate", periodEndDate);
		 	 $("#txtPeriodEndDateEdit1").val(periodEndDate);
			 $("#txtIntRecID1").val(periodIntRecID);
			 

				    	$.ajax({
							type:"POST",
							url: "../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
							cache: false,
							data: "action=EditStartDateForAccountingYearAdd&periodYear=" + encodeURIComponent(periodYear) + "&periodNum=" + encodeURIComponent(periodNum),
							success: function(response)
							 {
							//alert(response);
							
							 if(response > 1)
							 {
							 //alert(response1);
								document.getElementById("txtPeriodStartDateEdit1").disabled = true;
							 }
							 else
							 {
								document.getElementById("txtPeriodStartDateEdit1").disabled = false;
							 }
								
				             },
				             failure: function(response)
							 {
							    //$('#selPeriodNumContainerDivAdd2').empty().append("Failed");
				             }
						});
			 
		});		


		$('#deleteCompanyAccountingPeriod').on('show.bs.modal', function(e) {

	    	var $modal = $(this);
			var chkBoxArray = [];
			$(".checkthis:checked").each(function() {
			    chkBoxArray.push(this.id);
			});			
	    	
	    	if (chkBoxArray.length > 0) {
				//alert("Test1: "+chkBoxArray.length);
		    	$.ajax({
					type:"POST",
					url: "../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
					cache: false,
					data: "action=GetAccountingPeriodDeleteInformationForModal&accountingPeriodsArray="+encodeURIComponent(chkBoxArray),
					success: function(response)
					 {
		               	 $modal.find('#deleteAccountingPeriodsInfo').html(response);	               	 
		             },
		             failure: function(response)
					 {
					   $modal.find('#deleteAccountingPeriodsInfo').html("Failed");
		             }
				});	
			}
			else {
				//alert("Test2: "+chkBoxArray.length);
				swal("Please select at least one accounting period to delete.");
				deleteCompanyAccountingPeriod().show=false;
			}    
		});
	
	 
	});	
	
	
</script>


<!-- local custom css !-->
<style type="text/css">
	.form-control{
		overflow-x: hidden;
		}
		
	.post-labels{
 		padding-top: 5px;
 	}
 	.row-margin{
	 	margin-bottom: 20px;
	 	margin-top: 20px;
 	}
 	
 	h3{
	 	margin-top: 0px;
 	}
 	
 	.table-size .category{
	 	width: 35%;
	 	font-weight: normal;
 	}
 	
 	.table-size .group-name{
	 	width: 40%
 	}
 	
 	.table-size .sort-order{
	 	width: 10%;
 	}
 	
 	.table-size .display{
	 	width: 15%;
 	}
 
	 .col-line{
		 margin-bottom: 20px;
	  }

</style>
<!-- eof local custom css !-->

<!--<h1 class="page-header"><i class="fa fa-calendar-o" aria-hidden="true"></i> Company Accounting Periods</h1>-->

		<%
		SQLCheckPeriodsOrMonths = "SELECT PeriodsOrMonths FROM Settings_CompanyID"	

		Set cnnBuildDateDataSource = Server.CreateObject("ADODB.Connection")
		cnnBuildDateDataSource.open (Session("ClientCnnString"))
		Set rsBuildDateDataSource = Server.CreateObject("ADODB.Recordset")
		rsBuildDateDataSource.CursorLocation = 3 
		
		Set rsBuildDateDataSource = cnnBuildDateDataSource.Execute(SQLCheckPeriodsOrMonths)
		If rsBuildDateDataSource.EOF = false Then			
			PeriodsOrMonths = rsBuildDateDataSource("PeriodsOrMonths")
		End If	
		
		
		rsBuildDateDataSource.close
		Set rsBuildDateDataSource = Nothing
		Set rsBuildDateDataSource = Nothing
		
		'Response.Write PeriodsOrMonths
		
		%>
		<div class="row">
			<div class="col-lg-12 col-line">
				<div class="panel panel-default" style="margin:10px;">
				<% If PeriodsOrMonths = "M" Then %>
				<div class="panel-heading" style="color:red;">Your company options are currently set to use Months, not Periods. If you wish to change this it can be
done on the Identity tab. Any entries mode here will not be used by Insight while not in Period mode.</div>
				<% End If %>
					<div class="panel-heading">Build your custom period date ranges for each year's accounting.</div>
					<div class="panel-body">
						<div class="container">
						<div class="row">
					        <div class="col-md-12">
					        <h4>
					        <p data-placement="top" data-toggle="tooltip" title="Add Accounting Period"><a class="btn btn-success btn-large" data-title="Add Accounting Period" data-toggle="modal" data-target="#addCompanyAccountingPeriod"><i class="fa fa-plus-circle" aria-hidden="true"></i> Add New Accounting Period</a></p>					        
				            </h4>  
				            <div class="table-responsive">
				            <table id="mytable" class="table table-bordred table-striped">
				                   <thead>
				                   <th><input type="checkbox" id="checkall" /></th>
										<th>Fiscal Year</th>
										<th>Period</th>
										<th>Begin Date</th>
										<th>End Date</th>
										<th>Edit</th>
										<th>Delete</th>
				                   </thead>
					    <tbody>
						<%
						
						Server.ScriptTimeout = 500
						
						SQLBuildPeriodsDataSource = "SELECT * FROM Settings_AccountingPeriods ORDER BY PeriodYear ASC, Period ASC"
						
						Set cnnBuildPeriodsDataSource = Server.CreateObject("ADODB.Connection")
						cnnBuildPeriodsDataSource.open (Session("ClientCnnString"))
						Set rsBuildPeriodsDataSource = Server.CreateObject("ADODB.Recordset")
						rsBuildPeriodsDataSource.CursorLocation = 3 
						
						Set rsBuildPeriodsDataSource = cnnBuildPeriodsDataSource.Execute(SQLBuildPeriodsDataSource)
						
						If not rsBuildPeriodsDataSource.EOF Then
						
							
							Do While Not rsBuildPeriodsDataSource.EOF
							
								IntRecID = rsBuildPeriodsDataSource("InternalRecordIdentifier")
								PeriodYear = rsBuildPeriodsDataSource("PeriodYear")
								Period = rsBuildPeriodsDataSource("Period")
								PeriodBeginDate = formatDateTime(rsBuildPeriodsDataSource("BeginDate"),2)
								PeriodEndDate = formatDateTime(rsBuildPeriodsDataSource("EndDate"),2)
								
								%>
							    <tr>
							    <td><input type="checkbox" class="checkthis" id="<%= IntRecID %>"></td>
							    <td><%= PeriodYear %></td>
							    <td><%= Period %></td>
							    <td><%= PeriodBeginDate %></td>
							    <td><%= PeriodEndDate %></td>
							    <td><p data-placement="top" data-toggle="tooltip" title="Edit Accounting Period"><a class="btn btn-primary btn-xs" data-period-year="<%= PeriodYear %>" data-period-num="<%= Period %>" data-period-start="<%= PeriodBeginDate %>" data-period-end="<%= PeriodEndDate %>" data-record-id="<%= IntRecID %>" data-title="Edit Accounting Period" data-toggle="modal" data-target="#editCompanyAccountingPeriod"><span class="glyphicon glyphicon-pencil"></span></a></p></td>
							    <td><p data-placement="top" data-toggle="tooltip" title="Delete Accounting Period"><a class="btn btn-danger btn-xs" data-record-id="<%= IntRecID %>" data-title="Delete Accounting Period" data-toggle="modal" data-target="#deleteCompanyAccountingPeriod" ><span class="glyphicon glyphicon-trash"></span></a></p></td>
							    </tr>
							    
								<%
								rsBuildPeriodsDataSource.MoveNext
							Loop
						Else
							%><tr><td colspan="7">No Accounting Periods Have Been Added. Please Click The Green Button Above To Start Building Your Periods.</td></tr><%							
							
						End If
											
						Set rsBuildPeriodsDataSource = Nothing
						cnnBuildPeriodsDataSource.Close
						Set cnnBuildPeriodsDataSource = nothing
						
						%>					    
					    </tbody>
					        
					</table>
					                
					            </div>
					            
					        </div>
						</div>
					</div>

					</div>
				</div>
			</div>
		</div>
		
<!---------------------------------------------------------------------------------------------------------------------------------->
<!---------------------------------------------------------------------------------------------------------------------------------->
<!-- ADD, EDIT AND DELETE MODALS FOR COMPANY ACCOUNTING PERIODS                                                                       -->
<!---------------------------------------------------------------------------------------------------------------------------------->
<!---------------------------------------------------------------------------------------------------------------------------------->


<div class="modal fade" id="addCompanyAccountingPeriod" tabindex="-1" role="dialog" aria-labelledby="edit" aria-hidden="true">
	<div class="modal-dialog">
		<div class="modal-content">

		    <script>
		    
				function validateAddNewAccountingPeriodFields()
			    {
							    				       
				   var selectedPeriodAdd = $("#selPeriodNumAdd1 option:selected").val();
				   var selectedYearAdd = $("#txtPeriodYearAdd1").val();
				   var selectedStartDateAdd = $("#txtPeriodStartDateAdd1").val();
				   var selectedEndDateAdd = $("#txtPeriodEndDateAdd1").val();
				   		    
			       if (selectedPeriodAdd == "") {
			            swal("Please select a accounting period number.");
			            return false;
			       }	
			       if (selectedYearAdd == "") {
			            swal("Please select a accounting year.");
			            return false;
			       }						       			       
				   if (selectedStartDateAdd == "") {
			            swal("Please select a start date for this accounting period.");
			            return false;
			       }	
				   if (selectedEndDateAdd == "") {
			            swal("Please select an end date for this accounting period.");
			            return false;
			       }	
			       
					var d1 = Date.parse(selectedStartDateAdd);
					var d2 = Date.parse(selectedEndDateAdd);
					
					if (d1 > d2) {
			            swal("The end date must occur AFTER the start date.");
			            return false;
			       }	
		       			       	
			       return true;
			    }

		    
				$(document).ready(function(){
		
					$('#periodYearAdd1').on("dp.change", function (e){
					
					    var selectedPeriodAdd = $("#selPeriodNumAdd1").val();
					  	var selectedYearAdd = $("#txtPeriodYearAdd1").val();
						
				  		if (selectedPeriodAdd == "" || selectedPeriodAdd == null) {
				  			selectedPeriodAdd = 1;
				  		}
				  		
				    	$.ajax({
							type:"POST",
							url: "../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
							cache: false,
							data: "action=WritePeriodsInUseDropdownForAccountingYearAdd&periodYear=" + encodeURIComponent(selectedYearAdd) + "&periodNum=" + encodeURIComponent(selectedPeriodAdd),
							success: function(response)
							 {
				               	 $('#selPeriodNumContainerDivAdd1').empty().append(response);               	 
				             },
				             failure: function(response)
							 {
							    $('#selPeriodNumContainerDivAdd1').empty().append("Failed");
				             }
						});
						


				    	$.ajax({
							type:"POST",
							url: "../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
							cache: false,
							data: "action=WriteStartDateForAccountingYearAdd&periodYear=" + encodeURIComponent(selectedYearAdd) + "&periodNum=" + encodeURIComponent(selectedPeriodAdd),
							success: function(response)
							 {
							//alert(response);
							  var str = "+0";
							  var res = response.split("+");						  
							  var response1 = res[0];
							  var response2 = res[1];
							
							 if(response1 != "")
							 {
							 //alert(response1);
							 var date1 = new Date(response1);
							 date1.setDate(date1.getDate() + 1);
							 //alert(date1);
							  var day = date1.getDate();
							  var month = date1.getMonth() + 1;
							  var year = date1.getFullYear();
							  var newDate = month + '/' + day + '/' + year;
								$("#periodStartDateAddCond1").datetimepicker("defaultDate", newDate);
								document.getElementById("txtPeriodStartDateAdd1").disabled = true;
							 }
							 else
							 {
								$('#periodStartDateAddCond1').datetimepicker();
								$('#txtPeriodStartDateAdd1').val('');
								document.getElementById("txtPeriodStartDateAdd1").disabled = false;
							 }
								
				             },
				             failure: function(response)
							 {
							    $('#selPeriodNumContainerDivAdd2').empty().append("Failed");
				             }
						});
						
					});	
				    
				
					$('#btnAddNewPeriod1').on("click", function (e){
					
					    var selectedPeriodAdd = $("#selPeriodNumAdd1 option:selected").val();
					  	var selectedYearAdd = $("#txtPeriodYearAdd1").val();
					  	var selectedStartDateAdd = $("#txtPeriodStartDateAdd1").val();
					  	var selectedEndDateAdd = $("#txtPeriodEndDateAdd1").val();
					    
					    
					    if (validateAddNewAccountingPeriodFields()) {
						    
					    	$.ajax({
								type:"POST",
								url: "../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
								cache: false,
								data: "action=ValidateAndAddAccountingPeriod&periodYear=" + encodeURIComponent(selectedYearAdd) + "&periodNum=" + encodeURIComponent(selectedPeriodAdd) + "&periodStartDate=" + encodeURIComponent(selectedStartDateAdd) + "&periodEndDate=" + encodeURIComponent(selectedEndDateAdd),
								success: function(response)
								 {
					               	 if (response == 'Success') {
					               	 	location.reload();
									 }	 
									 else {
									 	swal(response);
									 }              	 
					             },
					             failure: function(response)
								 {
								    swal("Failed");
					             }
							});
						}
					});
					

				});	//end document.ready() function
				
		    </script>
	    <div class="modal-header">
	        <button type="button" class="close" data-dismiss="modal" aria-hidden="true"><span class="glyphicon glyphicon-remove" aria-hidden="true"></span></button>
	        <h4 class="modal-title custom_align" id="Heading">Add Accounting Period</h4>
	    </div>
	    <div class="modal-body">
    
	        <div class="form-group">
	        	Fiscal Year
	            <div class='input-group date' id='periodYearAdd1'>
	                <input type='text' class="form-control" id="txtPeriodYearAdd1">
	                <span class="input-group-addon">
	                    <span class="glyphicon glyphicon-calendar">
	                    </span>
	                </span>
	            </div>
	        </div>
	
			<div class="form-group" id="selPeriodNumContainerDivAdd1">
				<label for="selPeriodNum1">Period</label>
				<select class="form-control" id="selPeriodNum1" name="selPeriodNumAdd1">				
					<%
					For i = 1 To 100
					  	%><option value="<%= i %>"><%= i %></option><%
					Next
					%>				
				</select>				
				<!--<div class="form-group input-group date" id="periodStartDateAdd">
					<label for="txtPeriodStartDateAdd">Period Start Date</label>
	                <input type='text' class="form-control" id="txtPeriodStartDateAdd" name="txtPeriodStartDateAdd">
	                <span class="input-group-addon">
	                    <span class="glyphicon glyphicon-calendar">
	                    </span>
	                </span>
				</div>-->
			</div>
				
	        <div class="form-group">
			<div id="selPeriodNumContainerDivAdd2">
	            <label for="txtPeriodStartDateAdd1">Period Start Date</label>
				<div class="input-group date" id="periodStartDateAddCond1">	
					<input type="text" class="form-control" id="txtPeriodStartDateAdd1" name="txtPeriodStartDateAdd1" value="hello">		
					<span class="input-group-addon">
						<span class="glyphicon glyphicon-calendar">
						</span>
					</span>
				</div>
				
				<!--<div id="startDateCond1">
				<div class="input-group date" id="periodStartDateAdd">	
	                <input type="text" class="form-control" id="txtPeriodStartDateAdd" name="txtPeriodStartDateAdd" value="hello">
	                <span class="input-group-addon">
	                    <span class="glyphicon glyphicon-calendar">
	                    </span>
	                </span>
	            </div>
				</div>

				<div id="startDateCond2" style="display:none;">
				<div class="input-group date" id="periodStartDateAddCond">	
	                <input type="text" class="form-control" id="txtPeriodStartDateAdd" name="txtPeriodStartDateAdd" value="hello">
	                <span class="input-group-addon">
	                    <span class="glyphicon glyphicon-calendar">
	                    </span>
	                </span>
	            </div>
				</div>-->				
				
				<!--<div class="input-group date" id="periodStartDateAdd">	
	                <input type="text" class="form-control" id="txtPeriodStartDateAdd" name="txtPeriodStartDateAdd" value="hello">
	                <span class="input-group-addon">
	                    <span class="glyphicon glyphicon-calendar">
	                    </span>
	                </span>
	            </div>-->
	        </div>
			</div>
	
	        <div class="form-group">
	        	Period End Date
	            <div class='input-group date' id='periodEndDateAdd1'>
	                <input type='text' class="form-control" id="txtPeriodEndDateAdd1">
	                <span class="input-group-addon">
	                    <span class="glyphicon glyphicon-calendar">
	                    </span>
	                </span>
	            </div>
	        </div>
				
	      </div>
      
		<div class="modal-footer ">
			<button type="button" class="btn btn-success btn-lg" style="width: 100%;" id="btnAddNewPeriod1"><i class="fa fa-plus" aria-hidden="true"></i> Add New Period</button>
		</div>

       </div>
	<!-- /.modal-content --> 
	</div>
<!-- /.modal-dialog --> 
</div>
    



<div class="modal fade" id="editCompanyAccountingPeriod" tabindex="-1" role="dialog" aria-labelledby="edit" aria-hidden="true">
	<div class="modal-dialog">
		<div class="modal-content">
			<form name="frmEditAccountingPeriods" id="frmEditAccountingPeriods" method="post" action="accountingPeriodsUpdateFromModal.asp">
			    <script>
			    
					function validateEditAccountingPeriodFields()
				    {
								    				       
					   var selectedStartDateEdit = $("#txtPeriodStartDateEdit1").val();
					   var selectedEndDateEdit = $("#txtPeriodEndDateEdit1").val();
					   		    						       			       
					   if (selectedStartDateEdit == "") {
				            swal("Please select a start date for this accounting period.");
				            return false;
				       }	
					   if (selectedEndDateEdit == "") {
				            swal("Please select an end date for this accounting period.");
				            return false;
				       }	
				       
						var d1 = Date.parse(selectedStartDateEdit);
						var d2 = Date.parse(selectedEndDateEdit);
						
						if (d1 > d2) {
				            swal("The end date must occur AFTER the start date.");
				            return false;
				       }	
			       			       	
				       return true;
				    }
			    
					$(document).ready(function(){						
						
						$('#btnEditPeriod1').on("click", function (e){
						
						    var selectedPeriodEdit = $("#txtPeriodNumEdit1").val();
						  	var selectedYearEdit = $("#txtPeriodYearEdit1").val();
						  	var selectedStartDateEdit = $("#txtPeriodStartDateEdit1").val();
						  	var selectedEndDateEdit = $("#txtPeriodEndDateEdit1").val();
						    var selectedIntRecIDEdit = $("#txtIntRecID1").val();
						    
						    if (validateEditAccountingPeriodFields()) {
							    
						    	$.ajax({
									type:"POST",
									url: "../../inc/InSightFuncs_AjaxForBizIntelModals.asp",
									cache: false,
									data: "action=UpdateAccountingPeriod&periodIntRecID=" + encodeURIComponent(selectedIntRecIDEdit) + "&periodYear=" + encodeURIComponent(selectedYearEdit) + "&periodNum=" + encodeURIComponent(selectedPeriodEdit) + "&periodStartDate=" + encodeURIComponent(selectedStartDateEdit) + "&periodEndDate=" + encodeURIComponent(selectedEndDateEdit),
									success: function(response)
									 {
						               	 if (response == 'Success') {
						               	 	location.reload();
										 }	 
										 else {
										 	swal(response);
										 }              	 
						             },
						             failure: function(response)
									 {
									    swal("Failed");
						             }
								});
							}
						});
				
	
					});	
			    </script>
		    <div class="modal-header">
		        <button type="button" class="close" data-dismiss="modal" aria-hidden="true"><span class="glyphicon glyphicon-remove" aria-hidden="true"></span></button>
		        <h4 class="modal-title custom_align" id="Heading">Edit Accounting Period</h4>
		    </div>
		    <div class="modal-body">
	    		<input type="hidden" name="txtIntRecID1" id="txtIntRecID1">
	    		<input type="hidden" name="txtPeriodYearEdit1" id="txtPeriodYearEdit1">
	    		<input type="hidden" name="txtPeriodNumEdit1" id="txtPeriodNumEdit1">
	    		
	    		<div class="form-group" id="periodYearEdit1"></div>
	    				
				<div class="form-group" id="periodNumEdit1"></div>
					
		        <div class="form-group">
		        	Period Start Date
		            <div class='input-group date' id='periodStartDateEdit1'>
		                <input type='text' class="form-control" id="txtPeriodStartDateEdit1">
		                <span class="input-group-addon">
		                    <span class="glyphicon glyphicon-calendar">
		                    </span>
		                </span>
		            </div>
		        </div>
		
		        <div class="form-group">
		        	Period End Date
		            <div class='input-group date' id='periodEndDateEdit1'>
		                <input type='text' class="form-control" id="txtPeriodEndDateEdit1">
		                <span class="input-group-addon">
		                    <span class="glyphicon glyphicon-calendar">
		                    </span>
		                </span>
		            </div>
		        </div>
					
		      </div>
	      
				<div class="modal-footer ">
					<button type="button" class="btn btn-primary btn-lg" style="width: 100%;" id="btnEditPeriod1"><span class="fa fa-pencil"></span> Update Period</button>
				</div>
			</form>
       </div>
	<!-- /.modal-content --> 
	</div>
<!-- /.modal-dialog --> 
</div>

    
<div class="modal fade" id="deleteCompanyAccountingPeriod" tabindex="-1" role="dialog" aria-labelledby="edit" aria-hidden="true">
	<div class="modal-dialog">
		<div class="modal-content">
			<form name="frmDeleteAccountingPeriods" id="frmDeleteAccountingPeriods" method="post" action="accountingPeriodsDeleteFromModal.asp">
			
			  	<div class="modal-header">
					<button type="button" class="close" data-dismiss="modal" aria-hidden="true"><span class="fas fa-trash-alt" aria-hidden="true"></span></button>
					<h4 class="modal-title custom_align" id="Heading">Delete Period</h4>
				</div>
				
				<div class="modal-body">
					<div class="alert alert-danger"><span class="glyphicon glyphicon-warning-sign"></span> Are you sure you want to delete the following period(s)?</div>
					<div id="deleteAccountingPeriodsInfo"></div>
				</div>
				
				<div class="modal-footer ">
					<button type="button" class="btn btn-success" onclick="frmDeleteAccountingPeriods.submit()"><i class="fas fa-trash-alt" aria-hidden="true"></i> Yes, Delete</button>
					<button type="button" class="btn btn-default" data-dismiss="modal"><i class="fa fa-ban" aria-hidden="true"></i> Nevermind, Do Not Delete</button>
				</div>
			
			</form>
		</div>
	<!-- /.modal-content --> 
</div>
<!-- /.modal-dialog --> 

<!---------------------------------------------------------------------------------------------------------------------------------->
<!---------------------------------------------------------------------------------------------------------------------------------->
<!-- END ADD, EDIT AND DELETE MODALS FOR COMPANY ACCOUNTING PERIODS                                                                       -->
<!---------------------------------------------------------------------------------------------------------------------------------->
<!---------------------------------------------------------------------------------------------------------------------------------->



</div>				
				</div>
				
			</div>
		</div>
	
</div>
</div>




<!-- eof row !-->    
<!--#include file="../../inc/footer-main.asp"-->
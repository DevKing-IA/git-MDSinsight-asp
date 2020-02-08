<!--#include file="../../inc/header.asp"-->
<script>

	function OnCancel() 
	{ 	
		var url = [location.protocol, '//', location.host, location.pathname].join('');
		document.getElementById('txtProgress').innerHTML = "All File Upload Operations Cancelled";
		window.location.href = url;
	}
		
	
	$(document).ready(function(){
	
		$('#txtFileToUploadCustomers').filestyle({
			buttonName : 'btn-info',
	        buttonText : ' Select a FileTo Upload'
		});   
		$('#txtFileToUploadCustomerShipTos').filestyle({
			buttonName : 'btn-info',
	        buttonText : ' Select a File To Upload'
		});   
		$('#txtFileToUploadCustomerBillTos').filestyle({
			buttonName : 'btn-info',
	        buttonText : ' Select a File To Upload'
		});   
		$('#txtFileToUploadCustomerPOS').filestyle({
			buttonName : 'btn-info',
	        buttonText : ' Select a File To Upload'
		});   
		$('#txtFileToUploadEquipment').filestyle({
			buttonName : 'btn-info',
	        buttonText : ' Select a File To Upload'
		});   

		
		$('#txtFileToUploadCustomers').change(function(){
		    var fileArray = this.files;
		    var validFileExtensions = ["csv", "txt", "doc", "docx", "xls", "xlsx", "pdf", "zip"];
		    $.each(fileArray,function(i,v){
		    
			    var filename = v.name.toLowerCase();
			    var filesize = v.size;
			    var filetype = v.type;
			    
				var period = filename.lastIndexOf('.');
				var pluginName = filename.substring(0, period);
				var fileExtension = filename.substring(period + 1).toLowerCase();	
				
				if ($.inArray(fileExtension, validFileExtensions) == -1) {
	            	swal(fileExtension + " files are not allowed. Please upload one of the following file types: " + validFileExtensions.join(', '));
	            	$('#btnUpload').prop('disabled', true);
	            	$('#btnStop').prop('disabled', true)
	            	return false;
	            }
	        	else {
	        		$('#btnUpload').prop('disabled', false);
	            	$('#btnStop').prop('disabled', false)
	        	}

		      
		    })
		    
		    return true;
		});	
		
		$('#txtFileToUploadCustomerShipTos').change(function(){
		    var fileArray = this.files;
		    var validFileExtensions = ["csv", "txt", "doc", "docx", "xls", "xlsx", "pdf", "zip"];
		    $.each(fileArray,function(i,v){
		    
			    var filename = v.name.toLowerCase();
			    var filesize = v.size;
			    var filetype = v.type;
			    
				var period = filename.lastIndexOf('.');
				var pluginName = filename.substring(0, period);
				var fileExtension = filename.substring(period + 1).toLowerCase();	
				
				if ($.inArray(fileExtension, validFileExtensions) == -1) {
	            	swal(fileExtension + " files are not allowed. Please upload one of the following file types: " + validFileExtensions.join(', '));
	            	$('#btnUpload').prop('disabled', true);
	            	$('#btnStop').prop('disabled', true)
	            	return false;
	            }
	        	else {
	        		$('#btnUpload').prop('disabled', false);
	            	$('#btnStop').prop('disabled', false)
	        	}
		      
		    })
		    
		    return true;
		});	

		$('#txtFileToUploadCustomerBillTos').change(function(){
		    var fileArray = this.files;
		    var validFileExtensions = ["csv", "txt", "doc", "docx", "xls", "xlsx", "pdf", "zip"];
		    $.each(fileArray,function(i,v){
		    
			    var filename = v.name.toLowerCase();
			    var filesize = v.size;
			    var filetype = v.type;
			    
				var period = filename.lastIndexOf('.');
				var pluginName = filename.substring(0, period);
				var fileExtension = filename.substring(period + 1).toLowerCase();	
				
				if ($.inArray(fileExtension, validFileExtensions) == -1) {
	            	swal(fileExtension + " files are not allowed. Please upload one of the following file types: " + validFileExtensions.join(', '));
	            	$('#btnUpload').prop('disabled', true);
	            	$('#btnStop').prop('disabled', true)
	            	return false;
	            }
	        	else {
	        		$('#btnUpload').prop('disabled', false);
	            	$('#btnStop').prop('disabled', false)
	        	}

		      
		    })
		    
		    return true;
		});	

		$('#txtFileToUploadCustomerPOS').change(function(){
		    var fileArray = this.files;
		    var validFileExtensions = ["csv", "txt", "doc", "docx", "xls", "xlsx", "pdf", "zip"];
		    $.each(fileArray,function(i,v){
		    
			    var filename = v.name.toLowerCase();
			    var filesize = v.size;
			    var filetype = v.type;
			    
				var period = filename.lastIndexOf('.');
				var pluginName = filename.substring(0, period);
				var fileExtension = filename.substring(period + 1).toLowerCase();	
				
				if ($.inArray(fileExtension, validFileExtensions) == -1) {
	            	swal(fileExtension + " files are not allowed. Please upload one of the following file types: " + validFileExtensions.join(', '));
	            	$('#btnUpload').prop('disabled', true);
	            	$('#btnStop').prop('disabled', true)
	            	return false;
	            }
	        	else {
	        		$('#btnUpload').prop('disabled', false);
	            	$('#btnStop').prop('disabled', false)
	        	}

		      
		    })
		    
		    return true;
		});	

		$('#txtFileToUploadEquipment').change(function(){
		    var fileArray = this.files;
		    var validFileExtensions = ["csv", "txt", "doc", "docx", "xls", "xlsx", "pdf", "zip"];
		    $.each(fileArray,function(i,v){
		    
			    var filename = v.name.toLowerCase();
			    var filesize = v.size;
			    var filetype = v.type;
			    
				var period = filename.lastIndexOf('.');
				var pluginName = filename.substring(0, period);
				var fileExtension = filename.substring(period + 1).toLowerCase();	
				
				if ($.inArray(fileExtension, validFileExtensions) == -1) {
	            	swal(fileExtension + " files are not allowed. Please upload one of the following file types: " + validFileExtensions.join(', '));
	            	$('#btnUpload').prop('disabled', true);
	            	$('#btnStop').prop('disabled', true)
	            	return false;
	            }
	        	else {
	        		$('#btnUpload').prop('disabled', false);
	            	$('#btnStop').prop('disabled', false)
	        	}
		      
		    })
		    
		    return true;
		});	
		
		
		$('#btnUpload').prop('disabled', false);
		$('#btnStop').prop('disabled', false);
		
	});
	
 </script>   

<script src="progress_ajax.js"></script>
<script type="text/javascript" src="<%= BaseURL %>js/bootstrap-filestyle.min.js"></script>
<%

If (Request.ServerVariables("REQUEST_METHOD") = "POST") Then

	'*******************************************
	' Prepare ASPUpload to receive files
	'*******************************************
	Set UploadProgress = Server.CreateObject("Persits.UploadProgress")
	PID = UploadProgress.CreateProgressID()
	
	Session.CodePage = 65001
	
	Set Upload = Server.CreateObject("Persits.Upload.1")
	Upload.CodePage = 65001

	' This is needed to enable the progress indicator
	Upload.ProgressID = Request.QueryString("PID")
	
	Upload.IgnoreNoPost = True ' to use upload script in the same file as the form.
	
	'Upload.SetMaxSize 52428, false
	Upload.SetMaxSize 1*1024*1024*1024, false
	
	Upload.OverwriteFiles = false

	Upload.Save
	
	If Upload.Files.Count > 0 Then
		UploadMessage = "Success! " & Upload.Files.Count & " file(s) have been uploaded."
	Else
		UploadMessage = "<font color='red'>You have not uploaded any files. Please select files again.<font>"
	End If
			
	'******************************************************************
	' Loop through the File object and upload all files to the
	' virtual client uploaded_data directory with a datetime stamp
	'******************************************************************
	
	' Construct the save path
	
	SQL = "SELECT * FROM tblServerInfo where clientKey='" & trim(MUV_READ("ClientID")) & "'"
	Set Connection = Server.CreateObject("ADODB.Connection")
	Set Recordset = Server.CreateObject("ADODB.Recordset")
	Connection.Open InsightCnnString
	
	'Open the recordset object executing the SQL statement and return records
	Recordset.Open SQL,Connection,3,3
	
	'First lookup the ClientKey in tblServerInfo
	If NOT Recordset.EOF then
		BatchDataFolder = Recordset.Fields("BatchDataFolder")
	End If
	
	
	' Construct the save path
	VirtualPath ="..\..\..\..\BatchData\" & trim(BatchDataFolder) & "\"
	'VirtualPath ="c:\home\BatchData\" & trim(BatchDataFolder) & "\"

	
	'******************************************************************
	' Loop through the File object and upload all files to the
	' virtual client uploaded_data directory with a datetime stamp
	'******************************************************************
	'Response.Write("<br><br><br>VirtualPath : " & VirtualPath)
	
	FileNameForAuditTrail = ""

	For Each File in Upload.Files
	
		'Get date/time stamp:
		N = Now
		DateTimeStamp = Right("0" & Year(N),4) & Right("0" & Month(N),2) & Right("0" & Day(N),2) & "_" & Right("0" & Hour(N),2) & Right("0" & Minute (N),2) & Right("0" & Second(Now),2)	
		
		'Get file name without extension
		FileNameArray = Split(File.FileName, ".")
		FileNameForUpload = FileNameArray(0)

		
		If File.Name = "txtFileToUploadCustomers" Then
			FileNameForUpload = "BatchDataCustomer"
		ElseIf File.Name = "txtFileToUploadCustomerShipTos" Then
			FileNameForUpload = "BatchDataCustomerShipTos"
		ElseIf File.Name = "txtFileToUploadCustomerBillTos" Then
			FileNameForUpload = "BatchDataCustomerBillTos"
		ElseIf File.Name = "txtFileToUploadCustomerPOS" Then
			FileNameForUpload = "BatchDataCustomerPOS"
		ElseIf File.Name = "txtFileToUploadEquipment" Then
			FileNameForUpload = "BatchDataEquipment"
		End If
		
		FileExtentionForUpload = FileNameArray(1)
		'FileNameForAuditTrail = FileNameForAuditTrail & ", " & FileNameForUpload & "_" & DateTimeStamp & File.Ext
		FileNameForAuditTrail = FileNameForAuditTrail & ", " & FileNameForUpload & File.Ext
		'Only allow PDF, XLS, XLSX, DOC, DOCX, TXT, CSV or ZIP Files to be uploaded
		
		If UCASE(FileExtentionForUpload) = "PDF" OR UCASE(FileExtentionForUpload) = "XLS" OR UCASE(FileExtentionForUpload) = "XLSX" OR UCASE(FileExtentionForUpload) = "DOC" OR _
			UCASE(FileExtentionForUpload) = "DOCX" OR UCASE(FileExtentionForUpload) = "TXT" OR UCASE(FileExtentionForUpload) = "CSV" OR UCASE(FileExtentionForUpload) = "ZIP" Then
	    		'Response.Write("<br>" & VirtualPath & FileNameForUpload & "_" & DateTimeStamp & File.Ext & "<br>")
	    		'File.SaveAsVirtual VirtualPath & FileNameForUpload & "_" & DateTimeStamp & File.Ext
	    		File.SaveAsVirtual VirtualPath & FileNameForUpload & File.Ext
	    Else
	    	File.Delete
	    End If
	    
	Next

	CreateAuditLogEntry "Data Exchange File(s) Uploaded", "Data Exchange File(s) Uploaded", "Minor", 1, MUV_Read("DisplayName") & " uploaded the following data exchange file(s): " & FileNameForAuditTrail	

	
End If
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
	  
	.input-group .form-control {
	    position: relative;
	    z-index: 2;
	    float: left;
	    width: 100%;
	    height: 40px;
	    margin-bottom: 0;
	}
		
	.btn-info {
	    color: #fff;
	    background-color: #5bc0de;
	    border-color: #46b8da;
	    height: 40px;
	    vertical-align: middle;
	}	  
	
	.buttonText{
	vertical-align: middle;
	}

	.glyphicon {
	    margin-right: 2px;
	    margin-top:2px;
	}
	
	.progress-text {
		font-size:15px;
		color:green;
		font-weight:bold;
		margin-left:20px;
	}
</style>
<!-- eof local custom css !-->


<h1 class="page-header"><i class="fas fa-file-upload"></i> Data Exchange</h1>


<!--<form method="post" action="main.asp?pid=<% = PID %>" name="frmDataExchange" id="frmDataExchange" ENCTYPE="multipart/form-data" OnSubmit="ShowProgress('<% = PID %>')">-->

	<FORM METHOD="POST" ENCTYPE="multipart/form-data"
			ACTION="main.asp?pid=<% = PID %>"
			OnSubmit="ShowProgress('<% = PID %>')"> 

	<div class="row">
	
		<div class="col-lg-5 col-line">
			<h2>Customers</h2>
			<div class="form-group">
				<div class="col-lg-11">
			        <label style="margin-top:20px;margin-bottom:20px;">Please Choose Customer File to Upload</label>
			        <input type="file" id="txtFileToUploadCustomers" name="txtFileToUploadCustomers" accept = "*">	
			    </div>
			 </div>
		</div>
		
		<div class="col-lg-5 col-line">
			<h2>Customer Ship Tos</h2>
			<div class="form-group">
				<div class="col-lg-11">
			        <label style="margin-top:20px;margin-bottom:20px;">Please Choose Customer Ship To File to Upload</label>
			        <input type="file" id="txtFileToUploadCustomerShipTos" name="txtFileToUploadCustomerShipTos" accept = "*">	
			    </div>
			 </div>
		</div>	
			
	</div>

	<div class="row" style="margin-top:20px;">
	
		<div class="col-lg-5 col-line">
			<h2>Customer Bill Tos</h2>
			<div class="form-group">
				<div class="col-lg-11">
			        <label style="margin-top:20px;margin-bottom:20px;">Please Choose Customer Bill To File to Upload</label>
			        <input type="file" id="txtFileToUploadCustomerBillTos" name="txtFileToUploadCustomerBillTos" accept = "*">	
			    </div>
			 </div>
		</div>
		
		<div class="col-lg-5 col-line">
			<h2>Customer Points of Service</h2>
			<div class="form-group">
				<div class="col-lg-11">
			        <label style="margin-top:20px;margin-bottom:20px;">Please Choose Customer POS File to Upload</label>
			        <input type="file" id="txtFileToUploadCustomerPOS" name="txtFileToUploadCustomerPOS" accept = "*">	
			    </div>
			 </div>
		</div>	
			
	</div>

	<div class="row" style="margin-top:20px;">
	
		<div class="col-lg-5 col-line">
			<h2>Equipment</h2>
			<div class="form-group">
				<div class="col-lg-11">
			        <label style="margin-top:20px;margin-bottom:20px;">Please Choose Eqipment File to Upload</label>
			        <input type="file" id="txtFileToUploadEquipment" name="txtFileToUploadEquipment" accept = "*">	
			    </div>
			 </div>
		</div>	
		
		<div class="col-lg-5 col-line">
			&nbsp;
		</div>	
	
	</div>


	<div class="row" style="margin-top:20px;">
	
		<div class="col-lg-4 col-line pull-right">
			<div id="txtProgress"></div>
			<span class="progress-text"><%= UploadMessage %></span>
		</div>
	</div>
	
	<div class="row" style="margin-top:20px;">
	
		<div class="col-lg-4 col-line pull-right">
			<button type="button" class="btn btn-default" name="btnCancel" id="btnCancel" OnClick="OnCancel()"><i class="fas fa-redo-alt"></i> Clear &amp; Start Over</button>
			<!--<button type="button" class="btn btn-danger" name="btnStop" id="btnStop" OnClick="OnStop()"><i class="fas fa-stop"></i> Stop Upload</button>-->
			<button type="submit" class="btn btn-primary" name="btnUpload" id="btnUpload"><i class="fas fa-upload"></i> Upload All File(s)</button>       
		</div>
	</div>
	
</form>



<!-- eof row !-->    
<!--#include file="../../inc/footer-main.asp"-->
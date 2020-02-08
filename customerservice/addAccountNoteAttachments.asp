<!--#include file="../inc/header.asp"-->

<link rel="stylesheet" type="text/css" href="<%= BaseURL %>css/tabs.css" />


<style type="text/css">
	.alert{
 		padding: 6px 12px;
	}
	
	.form-control{
		margin-bottom: 20px;
	}
	
	a:hover{
		text-decoration: none;
	}
</style>


<h1 class="page-header"><i class="fa fa-file-text"></i> New <%=GetTerm("Account")%> Attachment</h1>

	
<form method="POST" action="addAccountNoteAttachments_submit.asp" name="frmAddAccountNote" ENCTYPE="multipart/form-data">		    
      

<div class="row">
	
     <!--account number !-->
	<div class="col-lg-3 col-md-3 col-sm-12 col-xs-12">
		<div class="alert alert-success" role="alert">  <strong>Account #: <%=Session("ServiceCustID") %><br><%= FormattedCustInfoByCustNum(Session("ServiceCustID"))%></strong>
				<input type="hidden" id="txtAccount" name="txtAccount" value="<%=Session("ServiceCustID") %>"  class="form-control last-run-inputs"> 
		</div>
     </div>					
     <!-- eof account number !-->
		        
	 
		<div class="col-lg-4 col-md-4 col-sm-12 col-xs-12">	  
			<!-- Account Note !-->
			<strong>Enter your notes regarding the attachment here</strong>
			<textarea name="txtAccountNote" spellcheck="True" id="txtAccountNote" rows="5"  class="form-control"></textarea>
			<!-- Account Note !-->
		</div>
		
	    <div class="col-xs-6 col-sm-1 col-md-1 col-lg-3">
		    <strong>Attachment</strong> <INPUT TYPE=FILE SIZE=40 NAME="FILE1">
		</div>

	</div>

	<div class="row">
			
		<div class="col-lg-12">	<br>
		    <a href="<%= BaseURL %>customerservice/main.asp#Attachments">
		    	<button type="button" class="btn btn-default">&lsaquo; Cancel &amp; Go Back To Notes Screen</button>
			</a>
			<button type="submit" class="btn btn-primary"><i class="fa fa-upload"></i> Submit</button>
		</div>
			
			
	</div>
			<!-- eof row !-->    

</form>

   
<!--#include file="../inc/footer-service.asp"-->

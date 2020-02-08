<!--#include file="../inc/header.asp"-->

<link rel="stylesheet" type="text/css" href="<%= BaseURL %>css/tabs.css" />


<!-- date picker !-->
<link rel="stylesheet" href="<%= baseURL %>css/datepicker/BeatPicker.min.css"/>
<script src="<%= baseURL %>js/datepicker/BeatPicker.min.js"></script>
<!-- eof date picker !-->


<style type="text/css">
    
    .beatpicker-clear{
	    display: none;
    }
    
    .sticky-box{
	    min-height: 50px;
    }
    
</style>

          
    

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


<h1 class="page-header"><i class="fa fa-file-text"></i> New <%=GetTerm("Account")%> Note</h1>

	
<form method="POST" action="addAccountNote_submit.asp" name="frmAddAccountNote">		    
      

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
 			<strong>Enter your notes here</strong>
			<textarea name="txtAccountNote" spellcheck="True" id="txtAccountNote" rows="5"  class="form-control"></textarea>
			<!-- Account Note !-->
		</div>
		
		<!-- sticky / calendar !-->
		<div class="col-lg-4 col-md-4 col-sm-12 col-xs-12">
			<div class="row">
				
				<!-- sticky !-->
	    <div class="col-lg-12 sticky-box">
		    <strong>Sticky</strong> <input type="checkbox" unchecked id="chkSticky"  name="chkSticky">
	    </div>
	    <!-- eof sticky !-->
	    
	    <!-- calendar !-->
	    <div class="col-lg-12">
		    <strong>Expires</strong><br>
		    	<input type="text" id="txtExpirationDate" name="txtExpirationDate" value='<%=DateAdd("yyyy",1,Date()) %>'  class="form-control last-run-inputs" data-beatpicker="true" data-beatpicker-format="['MM','DD','YYYY'],separator:'/'">
	    </div>
	    <!-- eof calendar !-->
		
			</div>
		</div>
			<!-- eof sticky / calendar !-->
 	</div>

	<div class="row">
			
		<div class="col-lg-12">	<br>
		    <a href="<%= BaseURL %>customerservice/main.asp#home">
		    	<button type="button" class="btn btn-default">&lsaquo; Cancel &amp; Go Back To Notes Screen</button>
			</a>
			<button type="submit" class="btn btn-primary"><i class="fa fa-upload"></i> Submit</button>
		</div>
			
			
	</div>
			<!-- eof row !-->    

</form>

   
<!--#include file="../inc/footer-service.asp"-->

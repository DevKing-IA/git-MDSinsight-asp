<!--#include file="../inc/header-field-service-mobile.asp"-->

<%
ClientKeyForFileNames = MUV_READ("ClientKeyForFileNames")
%>

<link rel="stylesheet" type="text/css" href="<%= BaseURL %>css/tabs.css" />

<% 

If Session("userNo") = "" Then Response.Redirect (BaseURL) ' Not logged in
If Session("ServiceCustID") <> "" Then 	CurrentService_CustID = Session("ServiceCustID") %>
<!-- select and auto complete !-->
  

<script type="text/javascript">

	$(function () {
		var autocompleteJSONFileURL = "../clientfiles/<%= ClientKeyForFileNames %>/autocomplete/customer_account_list_<%= ClientKeyForFileNames %>.json";
		
		var options = {
		  url: autocompleteJSONFileURL,
		  placeholder: "Search by name or account #",
		  getValue: "name",
		  list: {	
	        onChooseEvent: function() {
	             var custID = $("#txtCustID").getSelectedItemData().code;
				 if (custID!=""){
				 	$.ajax({
						type:"POST",
						url: "../inc/InSightAjaxFuncs.asp",
						data: "action=selectAccountFSNewMemo&custID="+encodeURIComponent(custID),
							success: function(msg){
								window.location = "addServiceMemo.asp";
							}
					}) 
				
				  }

        	},		  
		    match: {
		      enabled: true
			},
			maxNumberOfElements: 15		
		  },
		  theme: "round"
		};
		$("#txtCustID").easyAutocomplete(options);

	})
</script>
 

<style type="text/css">
.input-lg::-webkit-input-placeholder, textarea::-webkit-input-placeholder {
  color: #666;
}
.input-lg:-moz-placeholder, textarea:-moz-placeholder {
  color: #666;
}

.checkboxes label{
	font-weight: normal;
	margin-right: 20px;
}

.close-service-client-output{
	text-align: left;
}

.ticket-details{
	margin-bottom: 15px;
} 

.btn-default{
	margin-bottom: 15px;
} 

</style>


<!-- home & logout header buttons !-->
 <div class="container-fluid fieldservice-heading">
	 <div class="row">
		 
		 <!-- home !-->
		 <div class="col-lg-3">
			 <!-- Standard button -->
			 <a href="<%= BaseURL %>fieldserviceADV/main_menu.asp"  >
				<button type="button" class="btn btn-default pull-left"><i class="fa fa-home"></i> Home</button>
			</a>
		 </div>
		 <!-- eof home !-->
	 		 
		 <!-- logout !-->
		 <div class="col-lg-3">
			 <a href="../logout.asp"> 
				<button type="button" class="btn btn-danger pull-right"><i class="fa fa-sign-out"></i> Logout</button>
			 </a>
		 </div>
		 <!-- eof logout !-->
	 </div>
 </div>
 <!-- eof home & logout header buttons !-->


          
<SCRIPT LANGUAGE="JavaScript">
<!--
    function validateServiceMemoform()
    {

        if (document.frmAddServiceMemo.txtCompany.value == "") {
            swal("Company name can not be blank.");
            return false;
        }
        if (document.frmAddServiceMemo.txtContactName.value == "") {
            swal("Contact name can not be blank.");
            return false;
        }
        if (document.frmAddServiceMemo.txtLocation.value == "") {
            swal("Problem Location can not be blank.");
            return false;
        }
        if (document.frmAddServiceMemo.txtDescription.value == "") {
            swal("Problem description can not be blank.");
            return false;
        }


        return true;

    }
    
   
// -->
</SCRIPT>     




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

<div class="container-fluid">
	
<h1 class="page-header"><i class="fa fa-wrench"></i> New Service Memo</h1>

 

	
	<form method="POST" action="addservicememo_submit.asp" ENCTYPE="multipart/form-data" name="frmAddServiceMemo" onsubmit="return validateServiceMemoform();">		    
      
		
 	        <!-- row !-->		
	        <div class="row ">
		        
		        <!-- select company !-->
		        <div class="col-lg-4 col-md-4 col-sm-12 col-xs-12">
					<input id="txtCustID" name="txtCustID" style="width:100%">
					<i id="searchIcon" class="fa fa-search fa-2x"></i>
		        <!-- eof select company !-->
		        
		        <%If CurrentService_CustID <> "" Then 

					SQL = "Select * from AR_Customer WHERE CustNum = '" & CurrentService_CustID & "'"
					
					Set cnn8 = Server.CreateObject("ADODB.Connection")
					cnn8.open (Session("ClientCnnString"))
					Set rs = Server.CreateObject("ADODB.Recordset")
					rs.CursorLocation = 3	
					
					set rs = cnn8.Execute (SQL)
					
					If not rs.EOF then 
						acctnum = rs("CustNum")
						custAddress1 = rs("Addr1")
						custAddress2 = rs("Addr2")
						custCityStateZip = rs("CityStateZip")
						custContact =  rs("Contact")
						custPhone =  rs("Phone")
						 %>
			
				        <!--account number !-->
				        <div class="col-lg-4 col-md-4 col-sm-12 col-xs-12">
					        <div class="alert alert-success" role="alert">  <strong>Account #: <%=Session("ServiceCustID") %><br><%= FormattedCustInfoByCustNum(Session("ServiceCustID"))%></strong>
							<br>
					        <input type="hidden" id="txtAccount1" name="txtAccount1" value="<%=Session("ServiceCustID")%>"  class="form-control last-run-inputs"></div>
					        
					        </div>
				        <!-- eof account number !-->
			        <% End If %>
		      	
						        
		
		        </div>
 <!-- eof row !-->

 <!-- main row !-->
 <div class="row">
	 
	 <!-- left col !-->
	 <div class="col-lg-7 col-md-7 col-sm-12 col-xs-12">
	
	 
 		        <!-- row !-->			
			    <div class="row">

			    	<!-- Contact Name !-->
			    <div class="col-lg-4">
			        <strong>Contact Name</strong>
			        <input type="text" id="txtContactName" name="txtContactName"   class="form-control" value="<%= custContact %>">
			        </div>
			    	<!-- Contact Name !-->
	
			    	<!-- Contact Phone !-->
			    	  <div class="col-lg-4">
				    	 <strong>Contact Phone</strong>
				    	 <input type="text" id="txtContactPhone" name="txtContactPhone" class="form-control"  value="<%= custPhone %>">
			        </div>
			    	<!-- Contact Phone !-->
			    	
		    	
		    		<!-- Contact Email 
			    	<div class="col-lg-4">
			        <strong>Contact Email</strong>
			        <input type="text" id="txtContactEmail" name="txtContactEmail"   class="form-control">
			        </div>
			    	 Contact Email !-->

		    		<!-- Asset DropDown !-->
	   		    	   <div class="col-lg-4">
				       		<strong>Asset List</strong>
						        <select name="txtAssetID" id="txtAssetID" class="form-control input-lg">
									<option value="">Tap to select from assets assigned to this account</option>
									<option value="noneselected">-- NONE or NOT FOUND, USE THE NUMBER FROM THE BOX BELOW --</option>
									<%	
									strAssets= ""	
									
									Set cnn8 = Server.CreateObject("ADODB.Connection")
									cnn8.open (Session("ClientCnnString"))
									Set rs = Server.CreateObject("ADODB.Recordset")
									rs.CursorLocation = 3 
										
									SQL = "SELECT assetNumber,description,serno FROM " & MUV_Read("SQL_Owner") & ".Assets WHERE CustAcctNum = " & Session("ServiceCustID") &" ORDER BY assetTypeNo, assetNumber"
									
									set rs = cnn8.Execute (SQL)
									If not rs.EOF Then
							
										Do While Not rs.EOF
											tempAssetNum = rs("assetNumber")
											tempAssetDescription = rs("description")
											tempSerialNumber = rs("serno")
											
											If tempAssetNum = CurrentService_AssetNum Then
												strSelect =  "<option selected value='"& tempAssetNum &"'>"& tempAssetNum  & " -- SN:" &  tempSerialNumber & "-- " &  tempAssetDescription & "</option>"
											Else
												strSelect =  "<option value='"& tempAssetNum &"'>"& tempAssetNum  & " -- SN:" & tempSerialNumber & "-- " & tempAssetDescription & "</option>"
											End If
											strAccounts = strAccounts & """" & tempAssetNum  & " -- " & tempAssetNum &" --- SN:"& tempSerialNumber &" --- "& tempAssetDescription & ""","
							
											Response.Write(strSelect)
											rs.MoveNext
										Loop
										
										If Len(strAssets)>0 Then strAssets= Left(strAssets,Len(strAssets)-1)
										
									End If
									Set rs = Nothing
									Set Cnn8 = Nothing
								%>
								</select>
				       </div>
					<!-- eof Asset DropDown !-->

					<!-- asset tag number !-->
					<div class="col-lg-4 selectedhidden" id="noneselected" style="display:none;">
						<strong>If not found enter the asset tag below</strong>
						<input type="text" class="form-control input-lg" name="txtAssetTagNumber" id="txtAssetTagNumber">
					</div>
					<!-- eof asset tag number !-->

			    	</div>
			    <!-- eof row !-->
			    	
			   



				<!-- row !-->			
			    <div class="row">
 
			   				    
			    <!-- Memo Type !-->
			    <input type="hidden" name="selMemoType" value="Open">
			    <!-- Memo Type !-->
			    	   		  
		  			   
			    <!-- Problem Location !-->
			   <div class="col-lg-6">
					<strong>Problem Location</strong> <small>(Please include floor # if applicable)</small>
						<input type="text" id="txtLocation" name="txtLocation"   class="form-control"  value="">
		        </div>
		    	<!-- Problem Location !-->

		    	
		        <!-- end row !-->		
		        </div>
 
	 </div>
		    	 <!-- eof left col !-->
		    	 
		    	 
 		       
  	
  			<!-- right col !-->
  			<div class="col-lg-5 col-md-5 col-sm-12 col-xs-12">	  
	  					    	
			    	<!-- Description of problem !-->
 					<strong>Please enter a description as completely as possible.</strong>
					<textarea name="txtDescription" id="txtDescription"  spellcheck="True" rows="5"  class="form-control"></textarea>
 			    	<!-- Description of problem !-->
			
  			</div>
			<!-- eof right col !-->
		        		         
		</div>
		<!-- eof main row !-->
					<div class="row">

						<div class="col-lg-6">
							<button type="submit" class="btn btn-primary btn-lg btn-block col-lg-12"><i class="fa fa-upload"></i>  Tap Here To Submit Ticket</button><br>
						</div>
					</div>
		<% End if%>

	</form>
</div> 

<!-- show content if NONE or NOT FOUND is selected !-->
<script>
	 $(function() {
        $('#txtAssetID').change(function(){
            $('.selectedhidden').hide();
            $('#' + $(this).val()).show();
        });
    });
	</script>
<!-- eof show content !-->

<!--#include file="../inc/footer-field-service-noTimeout.asp"-->
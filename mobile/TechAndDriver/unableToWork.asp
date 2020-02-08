<!--#include file="inc/header-tech-and-driver.asp"-->


<%SelectedMemoNumber = Request.Form("txtTicketNumber") %>


<SCRIPT LANGUAGE="JavaScript">
    function validateunableToWorkform()
    {
        if (document.frmunableToWork.selAssetID.value != "") {
      	  if (document.frmunableToWork.txtAssetTagNumber.value != "") {
            swal("You selected an asset from the list AND typed in an asset tag number. Only one or the other is permitted. Please clear one entry before submitting.");
            return false;
      	  }
        }
        
        if (document.frmunableToWork.selAssetID.value != "") {
      	  if (document.frmunableToWork.txtAssetLocation.value == "") {
            swal("You selected an asset from the list but did not specify the location. Please enter the location before submitting.");
            return false;
      	  }
        }
        
        if (document.frmunableToWork.txtAssetTagNumber.value != "") {
      	  if (document.frmunableToWork.txtAssetLocation.value == "") {
            swal("You entered an asset tag but did not specify the location. Please enter the location before submitting.");
            return false;
      	  }
        }

                
        return true;
    }
</SCRIPT>  

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

.accordion-box{
margin-bottom:15px;
} 
</style>

 
<h1 class="fieldservice-heading" ><form method="post" action="onSite.asp" name="frmOnSiteH" id="frmOnSiteH">
		<input type='hidden' id='txtTicketNumber' name='txtTicketNumber' value='<%=SelectedMemoNumber%>'>		 
		<button type="button" onclick="document.forms['frmOnSiteH'].submit();" class="btn btn-default btn-home pull-left"><i class="fa fa-home"></i> Back</button>
</form>
Unable To Work Ticket # <%=SelectedMemoNumber%></h1>
 

<div class="container-fluid fieldservice-container">
	<div class="row">
 
		<form action="unableToWork_submit.asp" method="POST" ENCTYPE="multipart/form-data" name="frmunableToWork" id="frmunableToWork" onsubmit="return validateunableToWorkform();"> 

		<input type='hidden' id='txtTicketNumber' name='txtTicketNumber' value='<%=SelectedMemoNumber%>'>		 


			<!-- order notes !-->
			<div class="row row-rest">
				<h4 class="close-service-h4"><i class="fa fa-comments"></i> Service Notes</h4>
				<div class="col-lg-12 close-service-box">
					<textarea class="form-control  input-lg" rows="5" name="ServiceNotes"  spellcheck="True" id="ServiceNotes" ></textarea>
				</div>
			</div>
			<!-- eof order notes !-->

			<!-- Reason for no work accordion !-->
			<div class="container-fluid accordion-box"  id="theaccordion"  >
				<div class="row row-rest">
					<% If SelectedMemoNumber <> "" Then %>		
						<div class="panel-group panel-group-asset" id="accordion" role="tablist" aria-multiselectable="true">
 
	  						<div class="panel panel-default">
		   
								<!-- title !-->
    							<div class="panel-heading" role="tab" id="headingOne">
       								<h4 class="panel-title close-service-h4-panel">
	       								<a role="button" data-toggle="collapse"   data-parent="#accordion" href="#collapseOne"  aria-expanded="false" aria-controls="collapseOne">
	       									<i class="fa fa-plus-square-o"></i> Tap to select the reason you could not perform the work
	       								</a>
      								</h4>
								</div>

								<div id="collapseOne" class="panel-collapse collapse" role="tabpanel" aria-labelledby="headingOne">
									<div class="panel-body">
								        <div class="col-lg-12">

									        <select name="selReason" id="selReason" class="form-control input-lg">
												<option value="noneselected">-- None Selected --</option>
												<option value="COI-Insurance Problem">COI-Insurance Problem</option>
												<option value="Freight Elevator Down">Freight Elevator Down</option>
												<option value="Office or Freight Closed">Office or Freight Closed</option>
												<option value="No Contact">No Contact</option>
												<option value="Data Entry Error">Data Entry Error</option>												
												<option value="Tech Not On List">Tech Not On List</option>																								
												<option value="Other - See Notes">Other - See Notes</option>													
										</select>
									</div> 

								</div>
							</div>
							<!-- eof content !-->
						</div>
					</div>
					<% End If%>
				</div>
			</div>
			<!-- eof Reson for no work accordion !-->

			<!-- assets accordion !-->
			<div class="container-fluid accordion-box"  id="theaccordion"  >
				<div class="row row-rest">
					<% If SelectedMemoNumber <> "" Then %>		
						<div class="panel-group panel-group-asset" id="accordion" role="tablist" aria-multiselectable="true">
 
	  						<div class="panel panel-default">
		   
								<!-- title !-->
    							<div class="panel-heading" role="tab" id="headingOne">
       								<h4 class="panel-title close-service-h4-panel">
	       								<a role="button" data-toggle="collapse"   data-parent="#accordion" href="#collapseOne"  aria-expanded="false" aria-controls="collapseOne">
	       									<i class="fa fa-plus-square-o"></i> Update Asset Location
	       								</a>
      								</h4>
								</div>

								<div id="collapseOne" class="panel-collapse collapse" role="tabpanel" aria-labelledby="headingOne">
									<div class="panel-body">
								        <div class="col-lg-12">

									        <select name="selAssetID" id="selAssetID" class="form-control input-lg">
												<option value="">Tap to select from assets assigned to this account</option>
												<option value="noneselected">-- NONE or NOT FOUND, USE THE NUMBER FROM THE BOX BELOW --</option>
												<%	
												
												Set cnn8 = Server.CreateObject("ADODB.Connection")
												cnn8.open (Session("ClientCnnString"))
												Set rs = Server.CreateObject("ADODB.Recordset")
												rs.CursorLocation = 3 
													
												SQL = "SELECT assetNumber,description,serno FROM " & MUV_Read("SQL_Owner") & ".Assets WHERE CustAcctNum = " & GetServiceTicketCust(SelectedMemoNumber) &" ORDER BY assetTypeNo, assetNumber"
												
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
										
														Response.Write(strSelect)
														rs.MoveNext
													Loop
													
													
												End If
												Set rs = Nothing
												Set Cnn8 = Nothing
											%>
										</select>
									</div> 

									<!-- asset tag number !-->
									<div class="col-lg-12 selectedhidden" id="noneselected" style="display:none;">
										<h4 class="close-service-h4">If not found enter the asset tag below</h4>
										<input type="text" class="form-control input-lg" name="txtAssetTagNumber" id="txtAssetTagNumber">
									</div>
									<!-- eof asset tag number !-->

									<!-- asset location !-->
									<div class="col-lg-12">
										<h4 class="close-service-h4"><i class="fa fa-map-marker"></i> Asset Location </h4>
										<input type="text" class="form-control input-lg" name="txtAssetLocation" id="txtAssetLocation">
									</div>
									<!-- eof asset location !-->
								</div>
							</div>
							<!-- eof content !-->
						</div>
					</div>
				</div>
			</div>
			<!-- eof assets accordion !-->

			<!-- cancel / submit buttons !-->
			<div class="container-fluid ">
				<div class="row accordion-box">

					<!-- left col with pictures !-->
 						<!-- pictures accordion !-->
					    <div class="panel-group panel-group-asset pictures" id="accordion" role="tablist" aria-multiselectable="true">
							<div class="panel panel-default">
								<!-- title !-->
								<div class="panel-heading " role="tab" >
									<h4 class="panel-title close-service-h4-panel">
										<a role="button" data-toggle="collapse"   data-parent="#accordion" href="#uploadPictures"  aria-expanded="false" aria-controls="uploadPictures">
											<i class="fa fa-plus-square-o"></i> Attach Pictures
										</a>
									</h4>
								</div>
								<!-- eof title !-->

								<!-- content !-->
								<div id="uploadPictures" class="panel-collapse collapse " role="tabpanel" aria-labelledby="pictures">
									<div class="panel-body">
										Picture 1:<input type="file" capture="camera" accept="image/*" id="cameraInput1" name="cameraInput1"><br>
										Picture 2:<input type="file" capture="camera" accept="image/*" id="cameraInput2" name="cameraInput2"><br>
										Picture 3:<input type="file" capture="camera" accept="image/*" id="cameraInput3" name="cameraInput3"><br>
										Picture 4:<input type="file" capture="camera" accept="image/*" id="cameraInput4" name="cameraInput4"><br>
										Picture 5:<input type="file" capture="camera" accept="image/*" id="cameraInput5" name="cameraInput5"><br>
										Picture 6:<input type="file" capture="camera" accept="image/*" id="cameraInput6" name="cameraInput6"><br>
									</div>
								</div>
								<!-- eof content !-->
 						</div>
						<!-- eof pictures accordion !-->
					</div>
				</div>

				 <div class="row">
				 <div class="row">
						<button   class="btn btn-info btn-block  btn-lg close-buttons" name="Submit" value="Submit Service Memo" type="submit"   data-action="save" id="btn-download"><i class="fa fa-upload"></i> Submit Service Memo</button>
						</div>
						</div>
						
						</div>
			</div>
			<% End If%>
		</form>

<div class="row">		
  		<form method="post" action="onSite.asp" name="frmOnSite" id="frmOnSite">
			<input type='hidden' id='txtTicketNumber' name='txtTicketNumber' value='<%=SelectedMemoNumber%>'>
<div class="row">			
			<button type="button" onclick="document.forms['frmOnSite'].submit();" class="btn btn-warning btn-block  btn-lg close-buttons"><i class="fa fa-times-circle-o"></i> Cancel</button>
			</div>
		</form>
		</div>
		
		
	</div>
</div>

  

 <!-- show content if NONE or NOT FOUND is selected !-->
<script>
	 $(function() {
        $('#selAssetID').change(function(){
            $('.selectedhidden').hide();
            $('#' + $(this).val()).show();
        });
    });
</script>
<!-- eof show content !-->
 
<!--#include file="inc/footer-tech-and-driver.asp"-->
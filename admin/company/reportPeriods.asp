<!--#include file="../../inc/header.asp"-->

<script src="http://cdnjs.cloudflare.com/ajax/libs/moment.js/2.10.3/moment.js"></script>

<script type="text/javascript" src="<%= BaseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.js"></script>
<link rel="stylesheet" type="text/css" href="<%= BaseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.css" />


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
					
		 	
		});		


		
		$('#editCompanyReportPeriod').on('shown.bs.modal', function (e) {
		
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

<h1 class="page-header"><i class="fa fa-calendar-o" aria-hidden="true"></i> Company Report Periods</h1>


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
<!-- eof row !-->    
<!--#include file="../../inc/footer-main.asp"-->
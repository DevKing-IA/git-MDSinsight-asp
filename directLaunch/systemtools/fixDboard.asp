<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs_Routing.asp"-->

<%

If (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 AND Instr(ucase(ClientKey),"D") <> 0) Then clientIDsArray = Array("1071d","1230d","1128d") 
If (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") = 0 AND Instr(ucase(ClientKey),"D") = 0) Then clientIDsArray = Array("1071","1230","1128") 
If (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"FL.") <> 0 AND Instr(ucase(ClientKey),"D") <> 0) Then clientIDsArray = Array("1071d","1230d","1128d") 
If (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"FL2.") <> 0 AND Instr(ucase(ClientKey),"D") <> 0) Then clientIDsArray = Array("1071d","1230d","1128d") 
%>

<!DOCTYPE html>
<html>
<head>

<title>Fix Delivery Board</title>

<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">

<link href="https://maxcdn.bootstrapcdn.com/font-awesome/4.7.0/css/font-awesome.min.css" rel="stylesheet" type="text/css" />
<link href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" rel="stylesheet" type="text/css" />

<script src="http://code.jquery.com/jquery.min.js"></script>
<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"></script>

<!-- sweet alert jquery modal alerts !-->	
<script src="https://cdn.jsdelivr.net/sweetalert2/6.6.2/sweetalert2.min.js"></script>
<link rel="stylesheet" type="text/css" href="https://cdn.jsdelivr.net/sweetalert2/6.6.2/sweetalert2.css">
<!-- end sweet alert jquery modal alerts !-->	


<script type="text/javascript">

	function confirmUpdateDeliveryBoard(clientKey,clientShortName) {
	
		swal({
		  title: 'Are you sure you want to update ' + clientShortName + '?',
		  text: "You won't be able to revert this!",
		  type: 'warning',
		  showCancelButton: true,
		  confirmButtonColor: '#3085d6',
		  cancelButtonColor: '#d33',
		  confirmButtonText: 'Yes, update ' + clientShortName + ' board!'
  		}).then(function () {
		  
			$.ajax({		
				type:"POST",
				url: "fixDboardCallStoredProcedure.asp",
				data: "clientKey=" + encodeURIComponent(clientKey),
				success: function (data) {
					swal(
						'Success!',
		    			'The Delivery Board has been updated for ' + clientShortName + '.',
		    			'success'
		  			)
					window.location.href = "fixDboard.asp";
				}
			})	
		  
		}, function (dismiss) {
		  if (dismiss === 'cancel') {
		    swal(
		      'Cancelled',
		      'The ' + clientShortName + ' Delivery Board was not updated :)',
		      'error'
		    )
		  }
		})

  }

</script>

</head>

<style type="text/css">

	a.list-group-item-success, button.list-group-item-success {
	    color: #43ac6a;
	    font-weight: bold;
	}
	
	a.list-group-item-info, button.list-group-item-info {
	    color: #5bc0de;
	    font-weight: bold;
	}
	header {
	  padding: 3.5rem 0 0 0;
	  height: 40px;
	  position: absolute;
	  top: 0;
	  left: 0;
	  width: 100%;
	  display: block;
	  z-index: 1;
	}
	
	h1{
		font-size: 19px;
		font-weight:bold;
	}
	h1 span{
		color:#F26D21;
	}
	
	h4{
	line-height:1.3em;
	}
	
	h4 span{
		font-weight:bold;
	}
	
	.label{
	    font-size: 1em;
	}
	
	.list-group-item-heading {
	    margin-top: 0;
	    margin-bottom: 5px;
	}
	
	.list-group-item-text {
	    margin-bottom: 0;
	    line-height: 1.3;
	    /*margin-top:20px;*/
	}	

</style>


<body>  
	<%

	For i = 0 to uBound(clientIDsArray)
	
		SQL = "SELECT * FROM tblServerInfo WHERE Active = 1 AND routingModule='Enabled' AND ClientKey = '" & clientIDsArray(i) & "'"
		
		Set TopConnection = Server.CreateObject("ADODB.Connection")
		Set TopRecordset = Server.CreateObject("ADODB.Recordset")
		TopConnection.Open InsightCnnString
			
		'Open the recordset object executing the SQL statement and return records
		TopRecordset.Open SQL,TopConnection,3,3
		
		ClientCnnString = "Driver={SQL Server};Server=" & TopRecordset.Fields("dbServer")
		ClientCnnString = ClientCnnString & ";Database=" & TopRecordset.Fields("dbCatalog")
		ClientCnnString = ClientCnnString & ";Uid=" & TopRecordset.Fields("dbLogin")
		ClientCnnString = ClientCnnString & ";Pwd=" & TopRecordset.Fields("dbPassword") & ";"
		
		currClientKey = clientIDsArray(i)
		ShortCompanyName = TopRecordset.Fields("ShortCompanyName")
		LongCompanyName = TopRecordset.Fields("CompanyName")

		%>
        
      	<form action="fixDboardCallStoredProcedure.asp" method="POST" name="frmUpdateDeliveryBoard<%= currClientKey %>" id="frmUpdateDeliveryBoard<%= currClientKey %>">
		<input type="hidden" name="txtClientKey<%= currClientKey %>" id="txtClientKey<%= currClientKey %>" value="<%= currClientKey %>">
	
		
		  <div class="container" style="margin-bottom:30px;">
		      <div class="row">
		        <div class="col-md-12">
		          <h1>Delivery Board Status <span><%= LongCompanyName %></span></h1>
		        </div>
		      </div>   
			
			      <div class="row clearfix">
			          <div class="col-md-12 column">
			              <div class="panel panel-primary">
			                <div class="panel-heading">
			                  <h3 class="panel-title">
			                    Today's Date <%= FormatDateTime(Now(),2) %>
			                    <small class="pull-right"></small>
			                  </h3>
			                </div>                
			              </div>
			              
			              <%
			              	'Response.write("ClientCnnString : " & ClientCnnString & "<br>")
			              	
							Set cnnTechInfo = Server.CreateObject("ADODB.Connection")
							cnnTechInfo.open (ClientCnnString)
							Set rsTechInfo = Server.CreateObject("ADODB.Recordset")
							rsTechInfo.CursorLocation = 3 
						
							SQL_TechInfo = "SELECT * FROM SC_TechInfo"
							Set rsTechInfo = cnnTechInfo.Execute(SQL_TechInfo)
								              
							NightlyDBoard_Start = rsTechInfo("NightlyDBoard_Start")
							NightlyDBoard_Finish = rsTechInfo("NightlyDBoard_Finish")
							NightlyDBoard_LastStatus = rsTechInfo("NightlyDBoard_LastStatus")
							NightlyDBoard_LastAction = rsTechInfo("NightlyDBoard_LastAction")


							Set cnnDeliveryBoardSum = Server.CreateObject("ADODB.Connection")
							cnnDeliveryBoardSum.open (ClientCnnString)
							Set rsDeliveryBoardSum = Server.CreateObject("ADODB.Recordset")
							rsDeliveryBoardSum.CursorLocation = 3 
							
							SQLDeliveryBoardSum = "SELECT COUNT(IvsNum) as DeliveriesForToday FROM RT_DeliveryBoard"

							Set rsDeliveryBoardSum = cnnDeliveryBoardSum.Execute(SQLDeliveryBoardSum)
							If NOT rsDeliveryBoardSum.EOF Then
								DeliveriesForToday = rsDeliveryBoardSum("DeliveriesForToday")
							Else
								DeliveriesForToday = 0 
							End If
							
							SQLDeliveryBoardSum2 = "SELECT COUNT(IvsNum) as DeliveredOrSkipped FROM RT_DeliveryBoard WHERE DeliveryStatus IS NOT NULL OR DeliveryInProgress = 1"

							Set rsDeliveryBoardSum = cnnDeliveryBoardSum.Execute(SQLDeliveryBoardSum2)
							If NOT rsDeliveryBoardSum.EOF Then
								DeliveredOrSkipped = rsDeliveryBoardSum("DeliveredOrSkipped")
							Else
								DeliveredOrSkipped = 0 
							End If
							
							SQLDeliveryBoardSum3 = "SELECT COUNT(IvsNum) as UnmarkedDeliveries FROM RT_DeliveryBoard WHERE DeliveryStatus IS NULL"

							Set rsDeliveryBoardSum = cnnDeliveryBoardSum.Execute(SQLDeliveryBoardSum3)
							If NOT rsDeliveryBoardSum.EOF Then
								UnmarkedDeliveries = rsDeliveryBoardSum("UnmarkedDeliveries")
							Else
								UnmarkedDeliveries = 0 
							End If
							
							SQLDeliveryBoardSum4 = "SELECT COUNT (*) AS NumberOfDrivers FROM RT_Truck"

							Set rsDeliveryBoardSum = cnnDeliveryBoardSum.Execute(SQLDeliveryBoardSum4)
							If NOT rsDeliveryBoardSum.EOF Then
								NightlyDBoard_NumberOfDrivers = rsDeliveryBoardSum("NumberOfDrivers")
							Else
								NightlyDBoard_NumberOfDrivers = 0 
							End If
							
							
			              
			              %>
			              <div class="row clearfix">
			                  <div class="col-md-12 column">
			                      <div class="list-group">

			                          <div class="list-group-item">
			                              <h4 class="list-group-item-heading">
			                                  # of Drivers <span><%= NightlyDBoard_NumberOfDrivers %></span>
			                              </h4>
			                              
			                              <% If cInt(NightlyDBoard_NumberOfDrivers) < 1 Then %>
				                              <p class="list-group-item-text">
				                                  <span class="label label-warning">Potential Issue</span>
				                              </p>
			                              <% Else %>
				                              <p class="list-group-item-text">
				                                  <span class="label label-success">Valid Driver Count</span>
				                              </p>
			                              <% End If %>
			                          </div>
			                        
			                          <div class="list-group-item">
			                              <h4 class="list-group-item-heading">
			                                  Nightly Delivery Board Start <span><%= NightlyDBoard_Start %></span>
			                              </h4>
			                              
			                              <% If DateDiff("d",NightlyDBoard_Start,Now()) <> 0 Then %>
				                              <p class="list-group-item-text">
				                                  <span class="label label-warning">Potential Issue</span>
				                              </p>
			                              <% Else %>
				                              <p class="list-group-item-text">
				                                  <span class="label label-success">Valid Date</span>
				                              </p>
			                              <% End If %>
			                          </div>
			                        
			                          <div class="list-group-item">
			                              <h4 class="list-group-item-heading">
			                                  Nightly Delivery Board Finish <span><%= NightlyDBoard_Finish %></span>
			                              </h4>
			                              
			                              <% If DateDiff("d",NightlyDBoard_Finish,Now()) <> 0 Then %>
				                              <p class="list-group-item-text">
				                                  <span class="label label-warning">Potential Issue</span>
				                              </p>
			                              <% Else %>
				                              <p class="list-group-item-text">
				                                  <span class="label label-success">Valid Date</span>
				                              </p>
			                              <% End If %>
			                          </div>
			                        
			                          <div class="list-group-item">
			                              <h4 class="list-group-item-heading">
			                                  Nightly Delivery Board Last Status <span><%= NightlyDBoard_LastStatus %></span>
			                              </h4>
			                              
			                              <% If NightlyDBoard_LastStatus <> "Finished" Then %>
				                              <p class="list-group-item-text">
				                                  <span class="label label-warning">Potential Issue</span>
				                              </p>
			                              <% Else %>
				                              <p class="list-group-item-text">
				                                  <span class="label label-success">Valid Status</span>
				                              </p>
			                              <% End If %>
			                              
			                          </div>
	
			                        
			                          <div class="list-group-item">
			                              <h4 class="list-group-item-heading">
			                                  Nightly Delivery Board Last Action <span><%= NightlyDBoard_LastAction %></span>
			                              </h4>
			                          </div>
			                        
			                          <div class="list-group-item">
			                              <h4 class="list-group-item-heading">
			                                  # Deliveries For Today <span><%= DeliveriesForToday %></span>
			                              </h4>
			                              <% If cInt(DeliveriesForToday) < 1 Then %>
				                              <p class="list-group-item-text">
				                                  <span class="label label-warning">Potential Issue</span>
				                              </p>
			                              <% Else %>
				                              <p class="list-group-item-text">
				                                  <span class="label label-success">Valid Delivery Count</span>
				                              </p>
			                              <% End If %>
			                          </div>


			                          <div class="list-group-item">
			                              <h4 class="list-group-item-heading">
			                                  # Deliveries Marked Delivered or Skipped <span><%= DeliveredOrSkipped %></span>
			                              </h4>
			                              <% If cInt(DeliveredOrSkipped) > 0 Then %>
				                              <p class="list-group-item-text">
				                                  <span class="label label-warning">Potential Issue</span>
				                              </p>
			                              <% Else %>
				                              <p class="list-group-item-text">
				                                  <span class="label label-success">Valid DeliveredOrSkipped Count</span>
				                              </p>
			                              <% End If %>
			                          </div>


			                          <div class="list-group-item">
			                              <h4 class="list-group-item-heading">
			                                  # Deliveries Unmarked <span><%= UnmarkedDeliveries %></span>
			                              </h4>
			                              <% If cInt(UnmarkedDeliveries) <> cInt(DeliveriesForToday) Then %>
				                              <p class="list-group-item-text">
				                                  <span class="label label-warning">Potential Issue</span>
				                              </p>
			                              <% Else %>
				                              <p class="list-group-item-text">
				                                  <span class="label label-success">Valid Unmarked Count</span>
				                              </p>
			                              <% End If %>
			                          </div>
		                          
						                <a class="list-group-item list-group-item-success clearfix" style="font-size:19px;" onclick="location.reload();">
						                    REFRESH PAGE
						                    <span class="pull-right">
						                        <span class="btn btn-xs btn-default" onclick="location.reload();">
						                            <i class="fa fa-refresh fa-2x" aria-hidden="true"></i>
						                        </span>
						                    </span>
						                </a>
						                <a class="list-group-item list-group-item-info clearfix" style="font-size:19px;" onclick="confirmUpdateDeliveryBoard('<%= currClientKey %>','<%= ShortCompanyName %>');">
						                    UPDATE <%= LongCompanyName %>
						                    <span class="pull-right">
						                        <span class="btn btn-xs btn-default" onclick="confirmUpdateDeliveryBoard('<%= currClientKey %>','<%= ShortCompanyName %>');">
						                            <i class="fa fa-cloud-upload fa-2x" aria-hidden="true"></i>
						                        </span>
						                    </span>
						                </a>			                          
  
			                      </div>
			                  </div>
			              </div>
 
			          </div>
			      </div>
			  </div>
	  
	  </form>
	  
	<%	 
	Next
	

	Set rsTechInfo = Nothing
	cnnTechInfo.Close
	Set cnnTechInfo = Nothing

	Set rsDeliveryBoardSum = Nothing
	cnnDeliveryBoardSum.Close
	Set cnnDeliveryBoardSum = Nothing

	
	%>
		
  	<div class="container" style="margin-bottom:30px;">
      <div class="row">
        <div class="col-md-12">
          <a href="setNagMaster.asp"><button type="button" class="btn btn-primary btn-block">Go To Set Nag Master <i class="fa fa-arrow-circle-right" aria-hidden="true"></i></button></a>
        </div>
      </div>   
	</div>



</body>
</html>

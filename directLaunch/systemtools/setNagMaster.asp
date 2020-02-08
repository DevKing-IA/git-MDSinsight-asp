<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs_Routing.asp"-->

<%
clientIDsArray = Array("1071d","1230d") 
%>

<!DOCTYPE html>
<html>
<head>

<title>Set Master Nag Alerts</title>

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

	function confirmUpdateMasterNagAlertON(clientKey,clientShortName) {
	
		swal({
		  title: 'Are you sure you want to update ' + clientShortName + '?',
		  text: "You won't be able to revert this!",
		  type: 'warning',
		  showCancelButton: true,
		  confirmButtonColor: '#3085d6',
		  cancelButtonColor: '#d33',
		  confirmButtonText: 'Yes, turn ON ' + clientShortName + ' Master Nag!'
  		}).then(function () {
		  
			$.ajax({		
				type:"POST",
				url: "setNagMaster_AjaxFuncs.asp",
				cache: false,
				data: "action=TurnOnMasterNagAlertsForClientID&clientKey=" + encodeURIComponent(clientKey),
				success: function (data) {
					swal(
						'Success!',
		    			'The Master Nag Alert Setting has been turned ON for ' + clientShortName + '.',
		    			'success'
		  			)
					window.location.href = "setNagMaster.asp";
				}
			})	
  
		}, function (dismiss) {
		  if (dismiss === 'cancel') {
		    swal(
		      'Cancelled',
		      'The ' + clientShortName + ' Master Nag Alert Setting was not updated :)',
		      'error'
		    )
		  }
		})

  }
  
  


	function confirmUpdateMasterNagAlertOFF(clientKey,clientShortName) {
	
		swal({
		  title: 'Are you sure you want to update ' + clientShortName + '?',
		  text: "You won't be able to revert this!",
		  type: 'warning',
		  showCancelButton: true,
		  confirmButtonColor: '#3085d6',
		  cancelButtonColor: '#d33',
		  confirmButtonText: 'Yes, turn OFF ' + clientShortName + ' Master Nag!'
  		}).then(function () {
		  
			$.ajax({		
				type:"POST",
				url: "setNagMaster_AjaxFuncs.asp",
				cache: false,
				data: "action=TurnOffMasterNagAlertsForClientID&clientKey=" + encodeURIComponent(clientKey),
				success: function (data) {
					swal(
						'Success!',
		    			'The Master Nag Alert Setting has been turned OFF for ' + clientShortName + '.',
		    			'success'
		  			)
					window.location.href = "setNagMaster.asp";
				}
			})	
		  
		}, function (dismiss) {
		  if (dismiss === 'cancel') {
		    swal(
		      'Cancelled',
		      'The ' + clientShortName + ' Master Nag Alert Setting was not updated :)',
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
	
		SQL = "SELECT * FROM tblServerInfo WHERE Active = 1 AND routingModule='Enabled' AND serviceModule ='Enabled' AND ClientKey = '" & clientIDsArray(i) & "'"
		
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
        
		<input type="hidden" name="txtClientKey<%= currClientKey %>" id="txtClientKey<%= currClientKey %>" value="<%= currClientKey %>">

      <%
      	'Response.write("ClientCnnString : " & ClientCnnString & "<br>")
      	
		Set cnnMasterNagAlert = Server.CreateObject("ADODB.Connection")
		cnnMasterNagAlert.open (ClientCnnString)
		Set rsMasterNagAlert = Server.CreateObject("ADODB.Recordset")
		rsMasterNagAlert.CursorLocation = 3 
	
		SQL_MasterNagAlert = "SELECT   MasterNagMessageONOFF FROM Settings_Global"
		Set rsMasterNagAlert = cnnMasterNagAlert.Execute(SQL_MasterNagAlert)
			              
		MasterNagMessageONOFF = rsMasterNagAlert("MasterNagMessageONOFF")
		
		If MasterNagMessageONOFF = 0 Then
			MasterNagMessageONOFFStatus = "OFF"
	    ElseIf MasterNagMessageONOFF = 1 Then
	    	MasterNagMessageONOFFStatus = "ON"
	    Else
	    	MasterNagMessageONOFFStatus = "OFF"
	    End If
      
      %>
			              
		
		  <div class="container" style="margin-bottom:30px;">
		      <div class="row">
		        <div class="col-md-12">
		          <h1><%= ShortCompanyName %> Current Master Nag Status <span><%= MasterNagMessageONOFFStatus %></span></h1>
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
		              
		              
			          <div class="row clearfix">
			              <div class="col-md-12 column">
			                  <div class="list-group">
			
					                <a class="list-group-item list-group-item-success clearfix" style="font-size:19px;" onclick="location.reload();">
					                    REFRESH PAGE
					                    <span class="pull-right">
					                        <span class="btn btn-xs btn-default" onclick="location.reload();">
					                            <i class="fa fa-refresh fa-2x" aria-hidden="true"></i>
					                        </span>
					                    </span>
					                </a>
					                
					                <% If MasterNagMessageONOFF = 1 Then %>
						                <a class="list-group-item list-group-item-danger clearfix" style="font-size:19px;" onclick="confirmUpdateMasterNagAlertOFF('<%= currClientKey %>','<%= ShortCompanyName %>');">
						                    Turn Master Nag OFF for <%= LongCompanyName %>
						                    <span class="pull-right">
						                        <span class="btn btn-xs btn-danger" onclick="confirmUpdateMasterNagAlertOFF('<%= currClientKey %>','<%= ShortCompanyName %>');">
						                            <i class="fa fa-toggle-on fa-2x" aria-hidden="true"></i>
						                        </span>
						                    </span>
						                </a>	
						            <% Else %>
							           <a class="list-group-item list-group-item-info clearfix" style="font-size:19px;" onclick="confirmUpdateMasterNagAlertON('<%= currClientKey %>','<%= ShortCompanyName %>');">
						                    Turn Master Nag ON for <%= LongCompanyName %>
						                    <span class="pull-right">
						                        <span class="btn btn-xs btn-default" onclick="confirmUpdateMasterNagAlertON('<%= currClientKey %>','<%= ShortCompanyName %>');">
						                            <i class="fa fa-toggle-off fa-2x" aria-hidden="true"></i>
						                        </span>
						                    </span>
						                </a>					            

									<% End If %>		                          
			
			                  </div>
			              </div>
			          </div>

		          </div>
		      </div>
		  </div>
	
	<%	 
	Next
	

	Set rsMasterNagAlert = Nothing
	cnnMasterNagAlert.Close
	Set cnnMasterNagAlert = Nothing

	
	%>

  	<div class="container" style="margin-bottom:30px;">
      <div class="row">
        <div class="col-md-12">
          <a href="fixDboard.asp"><button type="button" class="btn btn-primary btn-block">Go To Fix Delivery Board <i class="fa fa-arrow-circle-right" aria-hidden="true"></i></button></a>
        </div>
      </div>   
	</div>


</body>
</html>

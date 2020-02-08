<!DOCTYPE html>
<!--[if lt IE 7 ]> <html class="no-js ie6 oldie" lang="en"> <![endif]-->
<!--[if IE 7 ]>    <html class="no-js ie7 oldie" lang="en"> <![endif]-->
<!--[if IE 8 ]>    <html class="no-js ie8 oldie" lang="en"> <![endif]-->
<!--[if IE 9 ]>    <html class="no-js ie9" lang="en"> <![endif]-->
<!--[if (gte IE 9)|!(IE)]><![endif]--><!-->
<html class="no-js" lang="en">
<!--<![endif]-->
<!--#include file="../../../inc/subsandfuncs.asp"-->
<!--#include file="../../../inc/InsightFuncs.asp"-->
<!--#include file="../../../inc/InSightFuncs_Routing.asp"-->
<%
UsageMessage = "http://www.mdsinsight.com/directLaunch/kiosks/routing/deliveryboardKiosk.asp?pp={Your Passphrase}&cl={Your Client ID}&ri={Refresh Interval In Seconds}&tn={Truck#,Truck#,Trick#} (Max 5 trucks)"
UsageMessage = UsageMessage & "<br>-OR-<br>"
UsageMessage = UsageMessage & "For last parameter use &tn=auto"

'These must be declared here
Dim DelBoardNextStopColor, DelBoardScheduledColor, DelBoardCompletedColor, DelBoardSkippedColor, DelBoardAMColor, DelBoardPriorityColor, DelBoardTitleText, DelBoardTitleTextFontColor, DelBoardTitleGradientColor 
Dim DelBoardRoutesToIgnore, DelBoardUPSRoutes

PassPhrase = Request.QueryString("pp")
ClientKey = Request.QueryString("cl")
If Request.QueryString("ri") <> "" Then Session("RefreshInterval") = Request.QueryString("ri") else Session("RefreshInterval") = 60
TruckNums = ""
TruckNums = Request.QueryString("tn")
If TruckNums = "" or PassPhrase = "" or ClientKey = "" Then
	Response.Write(UsageMessage)
	Respnse.End
End If


SQL = "SELECT * FROM tblServerInfo where clientKey='"& ClientKey &"'"

Set Connection = Server.CreateObject("ADODB.Connection")
Set Recordset = Server.CreateObject("ADODB.Recordset")
Connection.Open InsightCnnString
'Response.Write("InsightCnnString:" & InsightCnnString & "<br>")

'Open the recordset object executing the SQL statement and return records
Recordset.Open SQL,Connection,3,3
'Response.Write("SQL:" & SQL& "<br>")

'First lookup the ClientKey in tblServerInfo
'If there is no record with the entered client key, close connection
'and go back to login with QueryString
If Recordset.recordcount <= 0 then
	Recordset.close
	Connection.close
	set Recordset=nothing
	set Connection=nothing
	%>MDS Insight: Unable to connect to SQL database.<%
	Response.End
Else
	Session("ClientCnnString") = "Driver={SQL Server};Server=" & Recordset.Fields("dbServer")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Database=" & Recordset.Fields("dbCatalog")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Uid=" & Recordset.Fields("dbLogin")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Pwd=" & Recordset.Fields("dbPassword") & ";"
	dummy = MUV_Write("SQL_Owner",Recordset.Fields("dbLogin"))
	dummy = MUV_Write("ClientID",Recordset.Fields("clientkey"))
	If PassPhrase <>  Recordset.Fields("directLaunchPassphrase") Then
		Response.Write("Access Denied")
		Session.Abandon
		Response.End
	End If
	Recordset.close
	Connection.close
End If

Call Read_Settings_Global
%>

  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <meta name="description" content="">
    <meta name="author" content="">

    <title>MDS Insight</title>

    <!-- Bootstrap core CSS -->
    <link href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" rel="stylesheet">
    <!-- End Bootstrap core CSS -->

    <!-- Custom styles for MDS Insight -->
    <link href="<%= BaseURL %>css/dashboard.css" rel="stylesheet">
    <link href="<%= BaseURL %>css/screensize.css" rel="stylesheet">
	

    <!-- HTML5 shim and Respond.js for IE8 support of HTML5 elements and media queries -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/html5shiv/3.7.2/html5shiv.min.js"></script>
      <script src="https://oss.maxcdn.com/respond/1.4.2/respond.min.js"></script>
    <![endif]-->
    
    <!-- icons and notification styles !-->
     <link href="<%= BaseURL %>css/font-awesome/css/font-awesome.min.css" rel="stylesheet">
     <link href="<%= BaseURL %>css/notifications.css" rel="stylesheet">
    <!-- eof icons and notification styles !-->
    
    <!-- fonts !-->
    <link href='http://fonts.googleapis.com/css?family=Coda' rel='stylesheet' type='text/css'>
    <link href='http://fonts.googleapis.com/css?family=Oswald:400,300,700' rel='stylesheet' type='text/css'>
	<link href='http://fonts.googleapis.com/css?family=Indie+Flower' rel='stylesheet' type='text/css'>
    
    <!-- eof fonts !-->
	
	<!-- sort table script !-->
	<script src="<%= BaseURL %>js/sorttable.js"></script>
	<script src="<%= BaseURL %>js/sorttable1.js"></script>
	<!-- eof sort table script !-->

	<!-- *********************************************************************** -->
	<!-- IMPORTANT - USE OLDER VERSION OF JQUERY FOR SORTABLE PLUGIN             -->
	<!-- *********************************************************************** -->
  	<script src="http://code.jquery.com/jquery-1.11.2.min.js"></script>
	<!--<script src="https://code.jquery.com/jquery-3.1.1.js" integrity="sha256-16cdPddA6VdVInumRGo6IbivbERE8p7CQR3HzTBuELA=" crossorigin="anonymous"></script>  -->	
	<!-- *********************************************************************** -->

	
	<!-- Bootstrap core JS - must load after jQuery -->
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"></script>
    <!-- End Bootstrap core JS - must load after jQuery -->
		
	<!-- Including jQuery UI CSS & jQuery Dialog UI Here-->
	<link href="http://ajax.googleapis.com/ajax/libs/jqueryui/1.9.2/themes/ui-darkness/jquery-ui.css" rel="stylesheet">
	<script src="http://ajax.googleapis.com/ajax/libs/jqueryui/1.9.2/jquery-ui.min.js"></script>
	<!-- End Including jQuery UI CSS & jQuery Dialog UI Here-->
       
	<style type="text/css">
	        
	body{
		margin:10px;
		border: 2px solid <%=DelBoardTitleGradientColor%>;
		border-radius:5px;
		padding: 0px;
	}
	
	
	.wrapper{
		margin:0px;
	}
	
	#wrapper-margin{
		margin-top: -20px;
	}
	
	.heading-legend{
		margin-top:15px;
	}
	
	.heading-legend h4{
		font-weight:bold;
		margin:0px;
		padding:0px;
		text-transform:uppercase;
		text-align:center;
   }
   
   .legend-complete{
	   background:#d8f9d1;
	   padding:10px 15px 10px 15px;
   }
   
   .legend-nodelivery{
	   background:#fcb3b3;
	   padding:10px 15px 10px 15px;
   }
   
   .legend-nextstop{
	   background:#ffa500;
	   padding:10px 15px 10px 15px;
   }

	.legend-priority{
		<% Response.Write("-webkit-box-shadow:inset 0px 0px 0px 3px " & DelBoardPriorityColor & ";")%>
		<% Response.Write("-moz-box-shadow:inset 0px 0px 0px 3px " & DelBoardPriorityColor & ";")%>
		<% Response.Write("box-shadow:inset 0px 0px 0px 3px " & DelBoardPriorityColor & ";")%>    
		<% Response.Write("background-color:#FFFFFF;")%> 		
	   padding:10px 15px 10px 15px;
	   color:#000;
	   text-align:center;
	}

	.legend-am{
		<% Response.Write("-webkit-box-shadow:inset 0px 0px 0px 3px " & DelBoardAMColor & ";")%>
		<% Response.Write("-moz-box-shadow:inset 0px 0px 0px 3px " & DelBoardAMColor & ";")%>
		<% Response.Write("box-shadow:inset 0px 0px 0px 3px " & DelBoardAMColor & ";")%>    
		<% Response.Write("background-color:#FFFFFF;")%> 		
	   padding:10px 15px 10px 15px;
	   color:#000;
	   text-align:center;
	}
   
   .navbar-inverse{
	   border: 0px;
	  border-top-left-radius: 5px;
	  border-top-right-radius: 5px;
	  border-bottom-left-radius: 0px;
	  border-bottom-right-radius: 0px;
	  margin-top: -2px;
	  margin-bottom: -40px;
   }
   
   .delivery-status{
	   margin-top: 0px;
	   color: #fff;
    }
   
   .navbar-logo{
	   position: absolute;
	   margin-top: 5px;
	   margin-left: 5px;
	   max-height:40px;
   }
   
   .delivery-status h2{
	   margin:8px 0px 0px 0px;
	   line-height:1;
   }
  
   .pages{
	   float:right;
	   text-align:right;
	   margin: 0px 15px 0px 0px;
   }
    .pause{
	   float:right;
	   margin:10px 30px 0px 0px;
	   color:#337ab7;
   }
   
   .pages #number{
	   font-size:30px;
   }

	 .material-switch > input[type=checkbox] {
	    display: none;   
	}
	
	.material-switch > label {
	    cursor: pointer;
	    height: 0px;
	    position: relative; 
	    width: 40px;  
	}

	.material-switch > label::before {
	    background: rgb(0, 0, 0);
	    box-shadow: inset 0px 0px 10px rgba(0, 0, 0, 0.5);
	    border-radius: 8px;
	    content: '';
	    height: 16px;
	    margin-top: -8px;
	    position:absolute;
	    opacity: 0.3;
	    transition: all 0.4s ease-in-out;
	    width: 40px;
	}
	.material-switch > label::after {
	    background: rgb(255, 255, 255);
	    border-radius: 16px;
	    box-shadow: 0px 0px 5px rgba(0, 0, 0, 0.3);
	    content: '';
	    height: 24px;
	    left: -4px;
	    margin-top: -8px;
	    position: absolute;
	    top: -4px;
	    transition: all 0.3s ease-in-out;
	    width: 24px;
	}
	.material-switch > input[type=checkbox]:checked + label::before {
	    background: inherit;
	    opacity: 0.5;
	}
	.material-switch > input[type=checkbox]:checked + label::after {
	    background: inherit;
	    left: 20px;
	}  
	</style>

	<!-- countdown script !-->
	<script src="<%= BaseURL %>js/countdown/jquery.pietimer.js"></script>
	<script type="text/javascript">

		function Timer(callback, delay) {
		    var timerId, start, remaining = delay;
		
		    this.pause = function() {
		        window.clearTimeout(timerId);
		        remaining -= new Date() - start;
		    };
		
		    this.resume = function() {
		        start = new Date();
		        window.clearTimeout(timerId);
		        timerId = window.setTimeout(callback, remaining);
		    };
		
		    this.resume();
		}

		function hexToRgb(hex) {
		  var arrBuff = new ArrayBuffer(4);
		  var vw = new DataView(arrBuff);
		  vw.setUint32(0,parseInt(hex, 16),false);
		  var arrByte = new Uint8Array(arrBuff);
		
		  return "rgba(" + arrByte[1] + "," + arrByte[2] + "," + arrByte[3] + ",0.8)";
		}
			
		$(function(){
		
	
			var rgbcolor = '<%=Session("DelBoardPieTimerColor")%>';

			var pagetimer = new Timer(function() {
			    location.reload();
			}, <%=Session("RefreshInterval")%>*1000);
					
			$('#timer').pietimer({
				seconds: <%=Session("RefreshInterval")%>,
				color: hexToRgb(rgbcolor),
				height: 35,
				width: 35,
				is_reversed: true
			});
		
			
			$('#timer').pietimer('start');
	
			$('#switchAutomaticRefresh').on('change', function() {
			   if (this.checked) {
			        $('#timer').pietimer('pause');
			        pagetimer.pause();
			        return false;
			   }
			   else {
					$('#timer').pietimer('start');
					pagetimer.resume();
					return false;
			    }
	
			})	
		});
	</script>
	<!-- eof countdown script !-->
	

<%
Response.Write("<style type='text/css'>")
	Response.Write(".navbar-inverse{")
	Response.Write("background: " & DelBoardTitleGradientColor &"; /* For browsers that do not support gradients */" & "<br>")
    Response.Write("background: -webkit-linear-gradient(" & DelBoardTitleGradientColor & ", #fff); /* For Safari 5.1 to 6.0 */" & "<br>")
    Response.Write("background: -o-linear-gradient(" & DelBoardTitleGradientColor & ", #fff); /* For Opera 11.1 to 12.0 */" & "<br>")
    Response.Write("background: -moz-linear-gradient(" & DelBoardTitleGradientColor & ", #fff); /* For Firefox 3.6 to 15 */" & "<br>")
    Response.Write("background: linear-gradient(" & DelBoardTitleGradientColor & ", #fff); /* Standard syntax (must be last) */" & "<br>")
Response.Write("}")
Response.Write("</style>")
%> 

        
<%
If MUV_INSPECT("DBOARD_TOT_TRUCKS") <> True Then ' Only 1st time the page loads

	'This is where we get a total truck count, counting which ones will actually be displayed
	'So we can come up with the paging, etc
	
	TotalNumberOfTrucks = 0
	
	Set cnn_DeliveryBoardSum = Server.CreateObject("ADODB.Connection")
	cnn_DeliveryBoardSum.open (Session("ClientCnnString"))
	Set rs_DeliveryBoardSum = Server.CreateObject("ADODB.Recordset")
	rs_DeliveryBoardSum.CursorLocation = 3 
	
	If TruckNums = "auto" Then
		If DelBoardRoutesToIgnore = "" Or IsNull(DelBoardRoutesToIgnore) Then
			SQL_DeliveryBoardSum = "SELECT Count(DISTINCT TruckNumber) AS Expr1 FROM RT_DeliveryBoard"
		Else
			SQL_DeliveryBoardSum = "SELECT Count(DISTINCT TruckNumber) AS Expr1 FROM RT_DeliveryBoard WHERE LTRIM(RTRIM(TruckNumber)) NOT IN (" & DelBoardRoutesToIgnore & ")"
		End If
	Else
		If DelBoardRoutesToIgnore ="" Then
			'Redo the trucknums var to get rid of any that are in the ignored field
			DelBoardTruckArray = split(TruckNums,",")
			IgNoreTruckArray = split(DelBoardRoutesToIgnore,",")
			
			TruckNums = ""
			
			For i = 0 to Ubound(DelBoardTruckArray)
				FoundInIgnore = False
				For z = 0 to Ubound(IgNoreTruckArray)
					If Trim(DelBoardTruckArray(i)) = Trim(IgNoreTruckArray(z)) Then FoundInIgnore = True
				Next 
				If FoundInIgnore <> True Then TruckNums = TruckNums  & DelBoardTruckArray(i) & ","
			Next
			
			If Right(TruckNums,1) = "," Then TruckNums = Left(TruckNums, Len(TruckNums) -1) 
			
		End IF
		
		temparray = split(TruckNums,",")
		
		SQL_DeliveryBoardSum = "SELECT Count(TruckNumber) AS Expr1 From (SELECT DISTINCT TruckNumber FROM RT_DeliveryBoard "
	
		SQL_DeliveryBoardSum = SQL_DeliveryBoardSum & "WHERE "
		
		For x = 0 to ubound(temparray)
			SQL_DeliveryBoardSum = SQL_DeliveryBoardSum & "ltrim(rtrim(TruckNumber)) = '" & temparray(x) & "' OR "
		next
	
		If right(trim(SQL_DeliveryBoardSum),2)="OR" Then SQL_DeliveryBoardSum = Left(trim(SQL_DeliveryBoardSum),Len(trim(SQL_DeliveryBoardSum))-2) ' strip trailig OR
		
		SQL_DeliveryBoardSum = SQL_DeliveryBoardSum & ")AS derivedtbl_1 "
	End If

	Set rs_DeliveryBoardSum = cnn_DeliveryBoardSum.Execute(SQL_DeliveryBoardSum)
	
	If Not rs_DeliveryBoardSum.Eof Then TotalNumberOfTrucks = rs_DeliveryBoardSum("Expr1") 
	
	Set rs_DeliveryBoardSum = Nothing
	cnn_DeliveryBoardSum.Close
	Set cnn_DeliveryBoardSum = Nothing
	
	dummy = MUV_WRITE("DBOARD_TOT_TRUCKS",TotalNumberOfTrucks)
	'OK, now work out the paging
	If MUV_READ("DBOARD_TOT_TRUCKS") < 1 Then
		'No deliveries for the specified turcks, figure out what to do here
	Elseif MUV_READ("DBOARD_TOT_TRUCKS") < 6 Then
		dummy = MUV_WRITE("DBOARD_NUM_PAGES",1)
		dummy = MUV_WRITE("DBOARD_CURRENT_PAGE",1)
	Elseif MUV_READ("DBOARD_TOT_TRUCKS") > 5 Then
		'Need more than 1 page
		NumPages = MUV_READ("DBOARD_TOT_TRUCKS") / 5
		If NumPages <> Int(NumPages) Then NumPages = Int(NumPages) + 1 'Not a whole number, add a page
		dummy = MUV_WRITE("DBOARD_NUM_PAGES",NumPages)
		dummy = MUV_WRITE("DBOARD_CURRENT_PAGE",1)
	End If		
End If
%>
      
  </head>

<body>
	
  

 <!-- header !-->
<nav class="navbar navbar-inverse">

      <div class="container-fluid">
        <div class="navbar-header">
          
          <!-- row !-->
          <div class="row ">
	          
                     
          <!-- legend !-->
          <div class="col-lg-12 delivery-status">
          
          	<div class="col-lg-3">
	      		<a href="<%= BaseURL %>main/default.asp"><img src="<%= BaseURL %>clientfiles/<%= MUV_Read("ClientID") %>/logos/logo.png" class="navbar-logo"></a>
	      	</div>
            
            <div class="col-lg-6">          
		  		<h2 align="center"><font color="<%=DelBoardTitleTextFontColor%>"><%=DelBoardTitleText%></font></h2>
			</div>
            
        	<div class="col-lg-3" >

			<!-- 	<div class="pause">
			        Pause Automatic Refresh&nbsp;&nbsp;
                    <div class="material-switch pull-right">
                        <input id="switchAutomaticRefresh" name="chkAutomaticRefresh" type="checkbox"/>
                        <label for="switchAutomaticRefresh" class="label-primary"></label>
                    </div>
				</div>!-->
	        	
	        	<%IF MUV_READ("DBOARD_NUM_PAGES") <> "" AND MUV_READ("DBOARD_CURRENT_PAGE") <> "" Then%>
	        		<div class="pages"><font color="black">Page&nbsp;<span id="number"><%=MUV_READ("DBOARD_CURRENT_PAGE")%></span>&nbsp;of&nbsp;<span id="number"><%=MUV_READ("DBOARD_NUM_PAGES")%></span></font></div>
	        	<%End If%>
	 

				<div id="timer" class="pull-right" style="height:30px; margin-right:5px; margin-top:5px; margin-bottom:5px;"></div>
				
			</div>  

		 </div>
		<!-- legend ends here !-->


        </div>
          <!-- eof legend !-->
          
          <!-- welcome !-->
          <div class="col-lg-4">
          
 
 </div>
 <!-- eof row !-->
 
          </div>
          <!-- eof welcome !-->
          
          </div>
          <!-- eof row !-->
          
        </div>
         
      </div>
    </nav>
<!-- eof header !-->    


   
 <!-- eof side bar !-->

        <!-- content area !-->
        <div class="wrapper " >

 
<!--#include file="../../../inc/jquery_table_search.asp"-->

<%
RefreshURL = "DeliveryBoardKiosk.asp"
RefreshURL = RefreshURL & "?pp=" & PassPhrase 
RefreshURL = RefreshURL & "&cl=" & ClientKey 
RefreshURL = RefreshURL & "&ri=" & Session("RefreshInterval")
RefreshURL = RefreshURL & "&tn=" & TruckNums 
%>

<!-- DYNAMIC FORM !-->
<style type="text/css">
	 .ativa-scroll{
	 max-height: 300px
 }
</style>

<style type="text/css">
mark {
    background-color: yellow;
    color: black;
}
</style>

<!-- END DYNAMIC FORM !-->


 
 <style type="text/css">
 
 body{
	 overflow-x:hidden;
 }
 	.email-table{
		width:46%;
	}
	
	.bs-example-modal-lg-customize .row{
	margin-bottom: 10px;
 	width: 100%;
	overflow: hidden;
}

.bs-example-modal-lg-customize .left-column{
	background: #eaeaea;
	padding-bottom: 1000px;
    margin-bottom: -1000px;
}

.bs-example-modal-lg-customize .left-column h4{
	margin-top: 0px;
}

.bs-example-modal-lg-customize .right-column{
	background: #fff;
	padding-bottom: 1000px;
    margin-bottom: -1000px;
}


	table.sortable th:not(.sorttable_sorted):not(.sorttable_sorted_reverse):not(.sorttable_nosort):after { 
    content: " \25B4\25BE" 
}

table thead a{
	color: #000;
}

.tr-even{
	background: #f6f6f6;
  }

.tr-even-with-border{
	background: #f6f6f6;
	border:2px solid #fcd537;
  }


.tr-green{
	background:#D8F9D1;
  }
 
 .tr-green-with-border{
	background:#D8F9D1;
	border:2px solid #fcd537;
 }
 
 .tr-border-line{
	 border-bottom: 1px solid #ccc;
 }
 
 .tr-border-line-red{
	 border-bottom: 1px solid #999;
 }

.tr-red{
	background:#FCB3B3;
  }

.tr-red-with-border{
	background:#FCB3B3;
	border:2px solid #fcd537;
  }


.tr-orange{
	background:orange;
  }
 
.btn-link{
	padding: 0px;
	text-align: left;
}

.date-time-hidden-value{
	display:none;
}

.row{
	font-size:12px;
}

.fa-exclamation-triangle{
 	color:#ddcd1e;
 	cursor:pointer;
}

.legend-title{
	margin: 0px;
	padding: 0px;
}

.legend-row{
	margin-bottom: 10px;
	margin-left: 0px;
	margin-right: 0px;
 }

.legend-box{
 	padding-top: 10px;
	margin-bottom: 15px;
}
 
.high-priority{
	background:#fad5d5;
}

.urgent-priority{
	background:#ACF29C;
}

.alert-priority{
	background:#faf99d;
}

.alert-high-priority{
	background:#fa9090;
}

.yesbtn{
	background: transparent;
	border: 0px;
	color: green;
}

.nobtn{
	background: transparent;
	border: 0px;
	color: red;
}

.table-info{
	padding: 5px;
	border: 1px solid #eaeaea;
}

.table-info .table>tbody>tr>td, .table>tbody>tr>th, .table>tfoot>tr>td, .table>tfoot>tr>th, .table>thead>tr>td, .table>thead>tr>th{
	border: 0px;
	font-weight: normal;
	line-height: 1;
}

.table-condensed>tbody>tr>td, .table-condensed>tbody>tr>th, .table-condensed>tfoot>tr>td, .table-condensed>tfoot>tr>th, .table-condensed>thead>tr>td, .table-condensed>thead>tr>th{
	padding: 2px;
    max-width: 100px;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
}
 
 
 .page-header{
	 border-bottom:0px;
 }

.heading-legend{
	border-bottom:1px solid #eee;
	margin-bottom:20px;
 }

.heading-legend h1{
	margin:0px;
}

.custom-table{
 	font-size: 11px;
}

.btn-dispatch{
	font-size: 11px;
	padding: 5px;
}

.scrollable-table{
 	overflow: hidden;
	border: 1px solid #ccc;
 	font-size: 11px;
 	border-bottom-left-radius: 5px;
 	border-bottom-right-radius: 5px;
}

.row-line{
	margin-bottom: 25px;
}

.scrollable-title{
	border: 1px solid #ccc;
	padding: 3px 10px 3px 10px;
	margin-bottom: -1px;
	background: #DCE6E9;
	font-size: 12px;
	border-top-left-radius: 5px;
	border-top-right-radius: 5px;
}
 
 .tooltip-button{
	 padding: 0px;
	 border: 0px;
	 background: transparent;
	 font-size: 9px;
	 vertical-align: top;
 }
 
  .tooltip-button:hover{
	  background: transparent;
  }
  
  [class^="col-"]{
 	 padding:2px;
   }
   
   .col-lg-cust{
	   width:20%;
	   margin-top:10px;
   }
   

   .table>tbody>tr>td, .table>tbody>tr>th, .table>tfoot>tr>td, .table>tfoot>tr>th, .table>thead>tr>td, .table>thead>tr>th{
	   border:0px;
   }
 
 .tr-client{
	 font-size: 10px;
 }
 
 
 </style>

<!--- eof on/off scripts !-->





<style type="text/css">
.tr-inprogress{
	<% Response.Write("background:" & DelBoardInProgressColor & ";") %>
}

.tr-completed{
	<% Response.Write("background:" & DelBoardCompletedColor & ";") %>
}

.tr-nodelivery{
	<% Response.Write("background:" & DelBoardSkippedColor & ";") %>
}

.tr-nextstop{
	<% Response.Write("background:" & DelBoardNextStopColor & ";") %>
}
	
.tr-scheduled{
	<% Response.Write("background:" & DelBoardScheduledColor & ";") %>
}

.tr-border-line-AMDelivery{
	<% Response.Write("border: 3px solid " & DelBoardAMColor & ";") %>
}

.tr-border-line-PriorityDelivery{
	<% Response.Write("border: 3px solid " & DelBoardPriorityColor & ";") %>
}
 

.fa-star{
	color:blue;
  }

 
 </style>

<%
'Lets get all the trucks

If TruckNums = "auto" Then

	TruckNums = ""

	'If set to auto, we build it oursleves, all trucks
	'but still put it into the old style array
	Set cnn_DeliveryBoardSum = Server.CreateObject("ADODB.Connection")
	cnn_DeliveryBoardSum.open (Session("ClientCnnString"))
	Set rs_DeliveryBoardSum = Server.CreateObject("ADODB.Recordset")
	rs_DeliveryBoardSum.CursorLocation = 3 
	
	If DelBoardRoutesToIgnore <> "" Then 
		SQL_DeliveryBoardSum = "SELECT DISTINCT TruckNumber FROM RT_DeliveryBoard WHERE LTRIM(RTRIM(TruckNumber)) NOT IN (" & DelBoardRoutesToIgnore & ") ORDER BY Trucknumber"
	Else
		SQL_DeliveryBoardSum = "SELECT DISTINCT TruckNumber FROM RT_DeliveryBoard ORDER BY Trucknumber"
	End If
	
	Set rs_DeliveryBoardSum = cnn_DeliveryBoardSum.Execute(SQL_DeliveryBoardSum)
	
	If Not rs_DeliveryBoardSum.Eof Then
		Do While Not rs_DeliveryBoardSum.Eof
		
			TruckNums = TruckNums & rs_DeliveryBoardSum("TruckNumber") & ","
		
			rs_DeliveryBoardSum.MoveNext
		Loop
	End If
	
	TruckNums = Left(TruckNums,Len(TruckNums)-1) ' Strp Last comma
			
	Set rs_DeliveryBoardSum = Nothing
	cnn_DeliveryBoardSum.Close
	Set cnn_DeliveryBoardSum = Nothing

End If

TruckNumsArray = split(TruckNums,",")

'Now if paging is involved, we figure out what trucks to show
'an load up the TruckNumsArray appropriately
If MUV_READ("DBOARD_NUM_PAGES") = 1 AND MUV_READ("DBOARD_CURRENT_PAGE") = 1 Then
	'No real paging
	TruckNumsArray = split(TruckNums,",")
Else

	tempTruckNumsArray = split(TruckNums,",")
	
	StartTruckElement = (MUV_READ("DBOARD_CURRENT_PAGE")-1) * 5

	CurPageNum = MUV_READ("DBOARD_CURRENT_PAGE")
	StartTruckElement = ((CurPageNum-1)*5) 


	'Adjust for arrays starting at 0, except 1st page
	'If StartTruckElement > 0 Then StartTruckElement = StartTruckElement - 1

	EndTruckElement = StartTruckElement + 4	
	If EndTruckElement > Ubound(tempTruckNumsArray) Then EndTruckElement = Ubound(tempTruckNumsArray) 
	
	'response.write("<BR>")
	'response.write("StartTruckElement:" & StartTruckElement  & "<BR>")
	'response.write("EndTruckElement :" & EndTruckElement & "<BR><BR>")

	NewTruckVar = ""
	For x = StartTruckElement to EndTruckElement 
		NewTruckVar = NewTruckVar & tempTruckNumsArray(x) & ","
	Next
	NewTruckVar = Left(NewTruckVar,Len(NewTruckVar)-1) ' Strp Last comma
	
	TruckNumsArray = split(NewTruckVar,",")
	
	If MUV_READ("DBOARD_CURRENT_PAGE") = MUV_READ("DBOARD_NUM_PAGES") Then
		dummy = MUV_WRITE("DBOARD_CURRENT_PAGE",1)
	Else
		SetPg = MUV_READ("DBOARD_CURRENT_PAGE")
		SetPg = SetPg + 1
		dummy = MUV_WRITE("DBOARD_CURRENT_PAGE",SetPg)
	End If
	
End If



Set cnn_DeliveryBoardSum = Server.CreateObject("ADODB.Connection")
cnn_DeliveryBoardSum.open (Session("ClientCnnString"))
Set rs_DeliveryBoardSum = Server.CreateObject("ADODB.Recordset")
rs_DeliveryBoardSum.CursorLocation = 3 
SQL_DeliveryBoardSum = "SELECT TruckNumber From (SELECT DISTINCT TruckNumber FROM RT_DeliveryBoard "

SQL_DeliveryBoardSum = SQL_DeliveryBoardSum & "WHERE "
For x = 0 to ubound(TruckNumsArray)
	SQL_DeliveryBoardSum = SQL_DeliveryBoardSum & "ltrim(rtrim(TruckNumber)) = '" & TruckNumsArray (x) & "' OR "
next

If right(trim(SQL_DeliveryBoardSum),2)="OR" Then SQL_DeliveryBoardSum = Left(trim(SQL_DeliveryBoardSum),Len(trim(SQL_DeliveryBoardSum))-2) ' strip trailig OR

SQL_DeliveryBoardSum = SQL_DeliveryBoardSum & ")AS derivedtbl_1 Order By Case TruckNumber "
For x = 0 to ubound(TruckNumsArray)
	SQL_DeliveryBoardSum = SQL_DeliveryBoardSum & "WHEN '" & TruckNumsArray (x) & "' THEN " & x
next
SQL_DeliveryBoardSum = SQL_DeliveryBoardSum & " END "

'response.write(SQL_DeliveryBoardSum)
Set rs_DeliveryBoardSum = cnn_DeliveryBoardSum.Execute(SQL_DeliveryBoardSum)


GridColumn = 1

If not rs_DeliveryBoardSum.EOF Then
	Response.Write("<div class='row row-line' id='wrapper-margin'>")
	Do While Not rs_DeliveryBoardSum.Eof

			DriverUserNo = Trim(GetUserNumberByTruckNumber(rs_DeliveryBoardSum("TruckNumber")))
			
			If userIsArchived(DriverUserNo) = False AND userIsEnabled(DriverUserNo) = True Then
				
				Response.Write("<div class='col-lg-1 col-lg-cust'>")
				Response.Write("<div class='scrollable-title'><font size='4'>" & rs_DeliveryBoardSum("TruckNumber") & " - " & GetUserDisplayNameByUserNo(Trim(GetUserNumberByTruckNumber(rs_DeliveryBoardSum("TruckNumber")))) & "</font></div>")
				%> 
			        <div class='table-responsive scrollable-table'>
				        <% Response.Write("<table id='truck" & rs_DeliveryBoardSum("TruckNumber") & "' name='truck" & rs_DeliveryBoardSum("TruckNumber") & "' class='food_planner table table-condensed sortable '>") %>
					        <tbody class='searchable'>
					        	<%'Get all the tickets for this truck
								Set cnn_Tickets = Server.CreateObject("ADODB.Connection")
								cnn_Tickets.open (Session("ClientCnnString"))
								Set rs_DeliveryBoardDet = Server.CreateObject("ADODB.Recordset")
								rs_DeliveryBoardDet.CursorLocation = 3 
								SQL_Tickets = "SELECT * FROM RT_DeliveryBoard "
								SQL_Tickets = SQL_Tickets & "WHERE TruckNumber = '" & rs_DeliveryBoardSum("TruckNumber")  & "' "
								If DelBoardDontUseStopSequencing() = False Then
			                        SQL_Tickets = SQL_Tickets & "Order By SequenceNumber, CustNum" 
			                    Else
		   	                        SQL_Tickets = SQL_Tickets & "Order By CustNum" 
			                    End If
		                        Set rs_DeliveryBoardDet = cnn_Tickets.Execute(SQL_Tickets)
								If not rs_DeliveryBoardDet.Eof Then
		
									NumLines = 0
									Do While not rs_DeliveryBoardDet.Eof
									
										PriorityDelivery = rs_DeliveryBoardDet("Priority")
		
										If rs_DeliveryBoardDet("CustNum") = GetNextCustomerStopByTruck(rs_DeliveryBoardSum("TruckNumber")) Then
										
											If rs_DeliveryBoardDet("AMorPM") = "AM" Then
											
												If rs_DeliveryBoardDet("DeliveryInProgress") = 1 Then
												
													trclass = "<tr class='tr-inprogress tr-border-line-AMDelivery'>"
													
												Else
												
													If PriorityDelivery = 1 Then
														trclass = "<tr class='tr-nextstop tr-border-line-PriorityDelivery'>"
													Else
														trclass = "<tr class='tr-nextstop'>"
													End If
													
												End If
											
											Else
												If rs_DeliveryBoardDet("DeliveryInProgress") = 1 Then

													If PriorityDelivery = 1 Then
														trclass = "<tr class='tr-inprogress tr-border-line-PriorityDelivery'>"
													Else
														trclass = "<tr class='tr-inprogress'>"
													End If
										
												Else

													If PriorityDelivery = 1 Then
														trclass = "<tr class='tr-nextstop tr-border-line-PriorityDelivery'>"
													Else
														trclass = "<tr class='tr-nextstop'>"
													End If
													
												End If
											End If	
										Else
											If rs_DeliveryBoardDet("DeliveryStatus") = "Delivered" Then 
											
												If rs_DeliveryBoardDet("AMorPM") = "AM" Then 
													
													trclass = "<tr class='tr-completed tr-border-line-AMDelivery'>"
													
												Else 
													
													If PriorityDelivery = 1 Then
														trclass = "<tr class='tr-completed tr-border-line-PriorityDelivery'>"
													Else
														trclass = "<tr class='tr-completed'>"
													End If
													
												End If
												
											ElseIf rs_DeliveryBoardDet("DeliveryStatus") = "No Delivery" Then 
											
												If rs_DeliveryBoardDet("AMorPM") = "AM" Then 
													
													trclass = "<tr class='tr-nodelivery tr-border-line-AMDelivery'>"
													
												Else 

													If PriorityDelivery = 1 Then
														trclass = "<tr class='tr-nodelivery tr-border-line-PriorityDelivery'>"
													Else
														trclass = "<tr class='tr-nodelivery'>"
													End If
													
												End If
												
											Else
											
												If rs_DeliveryBoardDet("AMorPM") = "AM" Then
												
													If rs_DeliveryBoardDet("DeliveryInProgress") = 1 Then
														
														trclass = "<tr class='tr-inprogress tr-border-line-AMDelivery'>"
														
													Else
	
														If PriorityDelivery = 1 Then
															trclass = "<tr class='tr-scheduled tr-border-line-PriorityDelivery'>"
														Else
															trclass = "<tr class='tr-scheduled tr-border-line-AMDelivery'>"
														End If
														
													End If
												Else
													If rs_DeliveryBoardDet("DeliveryInProgress") = 1 Then
													
														If PriorityDelivery = 1 Then
															trclass = "<tr class='tr-inprogress tr-border-line-PriorityDelivery'>"
														Else
															trclass = "<tr class='tr-inprogress'>"
														End If
													
													Else
														
														If PriorityDelivery = 1 Then
															trclass = "<tr class='tr-scheduled tr-border-line-PriorityDelivery'>"
														Else
															trclass = "<tr class='tr-scheduled'>"
														End If
														
											
													End If
													
												End If
												
											End If
										End If
										
										Response.Write(trclass)
										If GetLastInvoiceMarkedByTruckNumber(rs_DeliveryBoardSum("TruckNumber")) = rs_DeliveryBoardDet("IvsNum") Then
											Response.Write("<td width='12%'><font size='3' >" & rs_DeliveryBoardDet("IvsNum") & "</font> </td>")
											Response.Write("<td width='10%' align='left'><i class='fa fa-star' aria-hidden='true'></i></td>")
											zzz=16
										Else
											Response.Write("<td><font size='3'>" & rs_DeliveryBoardDet("IvsNum") & "</font></td>")
											Response.Write("<td>&nbsp;</td>")
											zzz=16
										End If
		
										'Handle display of customer name
										CustDisplayName = rs_DeliveryBoardDet("CustName")
												
										Response.Write("<td><font size='3'>" & CustDisplayName & "</font></td></tr>")
										
										
										rs_DeliveryBoardDet.movenext
										NumLines = NumLines + 1
									
									Loop
									
									'Make all boxes even
									If NumLines < MaxNumberOfDeliveries() Then
										For x = 1 to MaxNumberOfDeliveries() - NumLines
											'Response.Write("<tr><td>&nbsp;</td></tr>")
											'Response.Write("<tr><td>&nbsp;</td></tr>")
										Next
									End IF
									
								End IF%>
					        </tbody>
				        </table>
			        </div>
		        <%Response.Write("</div>")
		        
		     End If

		rs_DeliveryBoardSum.Movenext
	Loop
	Response.Write("</div>")
End If



Set rs_DeliveryBoardSum = Nothing
cnn_DeliveryBoardSum.Close
Set cnn_DeliveryBoardSum = Nothing
%>	




 

<!-- tooltip JS !-->
<script type="text/javascript">
$(function () {
  $('[data-toggle="tooltip"]').tooltip()
})

 </script>
<!-- eof tooltip JS !-->


<script src="https://cdnjs.cloudflare.com/ajax/libs/highlight.js/8.3/highlight.min.js"></script>

<%'Subs and Funcs here

Sub Read_Settings_Global
	'Read delivery board settings
	SQL = "SELECT * FROM Settings_Global"
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
	If not rs.EOF Then
			DelBoardNextStopColor = rs("DelBoardNextStopColor")
			DelBoardScheduledColor = rs("DelBoardScheduledColor")	
			DelBoardCompletedColor = rs("DelBoardCompletedColor")				
			DelBoardSkippedColor = rs("DelBoardSkippedColor")	
			DelBoardAMColor = rs("DelBoardAMColor")
			DelBoardPriorityColor = rs("DelBoardPriorityColor")
			DelBoardTitleText = rs("DelBoardTitleText")	
			DelBoardTitleTextFontColor = rs("DelBoardTitleTextFontColor")
			DelBoardTitleGradientColor = rs("DelBoardTitleGradientColor")
			DelBoardRoutesToIgnore = rs("DelBoardRoutesToIgnore")
			DelBoardUPSRoutes = rs("DelBoardUPSRoutes")
			DelBoardPieTimerColor = rs("DelBoardPieTimerColor")
	End If
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	If DelBoardNextStopColor = "" Then DelBoardNextStopColor = "#FFA500"
	If IsNull(DelBoardNextStopColor) Then DelBoardNextStopColor = "#FFA500"
	If DelBoardScheduledColor = "" Then DelBoardScheduledColor = "#F6F6F6"
	If IsNull(DelBoardScheduledColor) Then DelBoardScheduledColor = "#F6F6F6"
	If DelBoardCompletedColor = "" Then DelBoardCompletedColor = "#D8F9D1"
	If IsNull(DelBoardCompletedColor) Then DelBoardCompletedColor = "#D8F9D1"
	If DelBoardSkippedColor = "" Then DelBoardSkippedColor = "#FCB3B3"
	If IsNull(DelBoardSkippedColor) Then DelBoardSkippedColor = "#FCB3B3"
	If DelBoardAMColor = "" Then DelBoardAMColor = "#000000"
	If DelBoardPriorityColor = "" Then DelBoardPriorityColor = "#000000"
	If IsNull(DelBoardAMColor) Then DelBoardAMColor = "#000000"
	If IsNull(DelBoardPriorityColor) Then DelBoardPriorityColor = "#000000"
	If DelBoardTitleTextFontColor = "" Then DelBoardTitleTextFontColor = "#000000"
	If IsNull(DelBoardTitleTextFontColor) Then DelBoardTitleTextFontColor = "#000000"
	If DelBoardTitleText = "" Then DelBoardTitleText = "Delivery Status"
	If IsNull(DelBoardTitleText ) Then DelBoardTitleText = "Delivery Status"
	DelBoardTitleText = Replace(DelBoardTitleText,"'","")
	DelBoardTitleText = Replace(DelBoardTitleText,"~today~",FormatDateTime(Now(),2))
	DelBoardTitleText = Replace(DelBoardTitleText,"~dow~",WeekDayName(Datepart("w",Now())))
	If DelBoardTitleGradientColor = "" Then DelBoardTitleGradientColor = "#80B8FF"
	If IsNull(DelBoardTitleGradientColor) Then DelBoardTitleGradientColor = "#80B8FF"
	If DelBoardPieTimerColor = "" Then DelBoardPieTimerColor = "000000"
	If IsNull(DelBoardPieTimerColor ) Then DelBoardPieTimerColor = "000000"
	Session("DelBoardPieTimerColor") = Replace(DelBoardPieTimerColor,"#","") ' Just this one for Javascript
End Sub

%><!--#include file="../../../inc/footer-main.asp"-->
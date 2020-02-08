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

UsageMessage = "http://www.mdsinsight.com/directLaunch/kiosks/routing/deliveryboardKiosk.asp?pp={Your Passphrase}&cl={Your Client ID}&ri={Refresh Interval In Seconds}&tn={Truck#,Truck#,Trick#}&mx=(Max # trucks, default 12)"
UsageMessage = UsageMessage & "<br>-OR-<br>"
UsageMessage = UsageMessage & "For tn parameter use &tn=auto"

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
MaxTrucksPerPage = Request.QueryString("mx")
If MaxTrucksPerPage = "" Then MaxTrucksPerPage = 12
MaxTrucksPerPage = cint(MaxTrucksPerPage)

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
	Elseif MUV_READ("DBOARD_TOT_TRUCKS") < MaxTrucksPerPage + 1 Then
		dummy = MUV_WRITE("DBOARD_NUM_PAGES",1)
		dummy = MUV_WRITE("DBOARD_CURRENT_PAGE",1)
	Elseif MUV_READ("DBOARD_TOT_TRUCKS") > MaxTrucksPerPage Then
		'Need more than 1 page
		NumPages = MUV_READ("DBOARD_TOT_TRUCKS") / MaxTrucksPerPage 
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

.findElement { 
		color:red !important; 
		font-weight:bolder !important;.
		background-color: yellow !important;
	}
	
	.progress{
	    position: relative;
		height: 16px;
		margin-top: 5px;
    	margin-bottom: 2px;
	}
	.progress > .progress-type {
		position: absolute;
		left: 0px;
		font-weight: 800;
	    font-size:11px;
		padding: 0px 30px 0px 5px;
		color: #FFF;
		background-color: rgba(25, 25, 25, 0.2);
	}
	.progress > .progress-completed {
		position: absolute;
		right: 0px;
		font-weight: 800;
	    color: #000;
		padding: 0px 10px 1px;
	}
	
	.progress-type { top: 1px } 
	.progress-completed { font-size: 12px } 


		 .ativa-scroll{
		 max-height: 300px
	 }
	
 

.fa-star{
	color:blue;
  }

 
 </style>

<style type="text/css">

	/*.findElement { background-color: yellow !important;}*/
	
	.findElement { 
		color:red !important; 
		font-weight:bolder !important;.
		background-color: yellow !important;
	}
	
	.progress{
	    position: relative;
		height: 16px;
		margin-top: 5px;
    	margin-bottom: 2px;
	}
	.progress > .progress-type {
		position: absolute;
		left: 0px;
		font-weight: 800;
	    font-size:11px;
		padding: 0px 30px 0px 5px;
		color: #FFF;
		background-color: rgba(25, 25, 25, 0.2);
	}
	.progress > .progress-completed {
		position: absolute;
		right: 0px;
		font-weight: 800;
	    color: #000;
		padding: 0px 10px 1px;
	}
	
	.progress-type { top: 1px } 
	.progress-completed { font-size: 12px } 


		 .ativa-scroll{
		 max-height: 300px
	 }
	
	 .alarm-bell{
		 position:absolute;
		 /*right:5px;*/
	 }
	 
 
 .alarm-bell .alert-pop-up{
	 display: none;
	 background: #000;
	 color: #fff;
	 position: absolute;
	 padding: 5px 10px 5px 10px;
	 z-index: 900;
	 margin:-17px 0px 0px 20px;
	 font-weight: bold;
  }
 
	 .alarm-bell:hover .alert-pop-up{
		 display: block;
	 }
	 
	 .alarm-bell2{
		 position:absolute;
		 left:5px;
	 }
	 
	 
	 .alarm-bell2 .alert-pop-up{
		 display: none;
		 background: #000;
		 color: #fff;
		 position: absolute;
		 padding: 5px 10px 5px 10px;
		 z-index: 900;
		 margin:-17px 0px 0px 20px;
		 font-weight: bold;
	  }
	 
	 .alarm-bell2:hover .alert-pop-up{
		 display: block;
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

	.modal-body{
		font-size:14px;
	}
	
	.modal-body label{
		font-weight:bold;
		padding-top:10px;
	}
	
	.modal-body .row-line{
		width:100%;
		display:inline-block;
		margin:0px 0px 10px 0px;
	}
	
	.modal-body .row-line .multiselect,.textarea{
		min-height:110px;
		max-height:110px;
		margin-bottom:5px;
	}
	
	.modal-body .row-line .right{
		text-align:right;
	}

	.bottom-alert{
		text-align:center;
		font-weight:bold;
		color:red;
	}
	
	
	 body{
		 overflow-x:hidden;
		/*overflow: auto; */
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
	
	
	
	
	table thead a{
		color: #000;
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

	/*.tr-scheduled-top{
		<% Response.Write("border-top: 1px solid #000000;") %>
		<% Response.Write("border-left: 1px solid #000000;") %>
		<% Response.Write("border-right: 1px solid #000000;") %>
	}
	
	.tr-scheduled-bottom{
		<% Response.Write("border-bottom: 1px solid #000000;") %>
		<% Response.Write("border-left: 1px solid #000000;") %>
		<% Response.Write("border-right: 1px solid #000000;") %>
	}*/
	
	#AM-border-top{
		<% 'Response.Write("border: 3px solid " & DelBoardAMColor & ";") %>
			-webkit-box-shadow:inset 0px 0px 0px 3px #f00;
			-moz-box-shadow:inset 0px 0px 0px 3px #f00;
			box-shadow:inset 0px 0px 0px 3px #f00;
  	}
	
	
	#Priority-border-top{
		<% 'Response.Write("border: 3px solid " & DelBoardPriorityColor & ";") %>
			-webkit-box-shadow:inset 0px 0px 0px 3px #f00;
			-moz-box-shadow:inset 0px 0px 0px 3px #f00;
			box-shadow:inset 0px 0px 0px 3px #f00;
  	}
	
	/*.AM-border-bottom{
		<% Response.Write("border-bottom: 3px solid " & DelBoardAMColor & ";") %>
		<% Response.Write("border-left: 3px solid " & DelBoardAMColor & ";") %>
		<% Response.Write("border-right: 3px solid " & DelBoardAMColor & ";") %>
	}*/
	
	/*.Priority-border-bottom{
		<% Response.Write("border-bottom: 3px solid " & DelBoardPriorityColor & ";") %>
		<% Response.Write("border-left: 3px solid " & DelBoardPriorityColor & ";") %>
		<% Response.Write("border-right: 3px solid " & DelBoardPriorityColor & ";") %>
	}*/
	
	
	.tr-user-alert{
		<% Response.Write("background:" & DelBoardUserAlertColor & ";") %>
	}
	
	.tr-user-alert-top{
		<% Response.Write("border: 1px solid #000000;") %>
 	}
	
	/*.tr-user-alert-bottom{
		<% Response.Write("border-bottom: 1px solid #000000;") %>
		<% Response.Write("border-left: 1px solid #000000;") %>
		<% Response.Write("border-right: 1px solid #000000;") %>
	}*/
 
	 .tr-border-line{
		 border-bottom: 1px solid #ccc;
	 }
	 
	 .tr-border-line-red{
		 border-bottom: 1px solid #999;
	 }
	
	 
	.btn-link{
		padding: 0px;
		text-align: left;
	}
	
	.date-time-hidden-value{
		display:none;
	}
	
	.row{
		/*font-size:12px;*/
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
	 	font-size: 9px;
	 	border-bottom-left-radius: 5px;
	 	border-bottom-right-radius: 5px;
	 	
	}

	.row-line{
		margin-bottom: 25px;
		margin-top: 30px;
	}
	
	.scrollable-title{
		border: 1px solid #ccc;
		padding: 10px;
		margin-bottom: -1px;
		background: #DCE6E9;
		font-size: 12px;
		border-top-left-radius: 5px;
		border-top-right-radius: 5px;
	}
	
	.scrollable-title strong{
		width:100%;
		display:block;
		white-space:normal;
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
	   width:7%;
	   display:inline-block;
	   vertical-align:top;
   }
   
 	.ui-state-highlight.item{height: 100px;}
 
	.list-boxes{
	/*margin-left: 230px;*/
	}
	
	.timer-countdown{
	 position:absolute;
	 top:10px;
	 right:20px;
	}
	  
	  .fa-star{
	color:blue;
	}
	
	.container{
		max-width:800px;
		margin:0 auto;
		padding-top:40px;
	}
	
	.container-fluid-trucks{
		padding-top:40px;
	}
	
	.btn-truck{
		width:100%;
		padding:10px 15px 10px 15px;
		font-weight:bold;
		border:0px;
		border-radius:0px 0px 0px;
		background:#f5f5f5;
		color:#000;
		display:block;
		text-align:left;
		outline:none;
 	}
	
	.btn-truck:hover{
		text-decoration:none;
 	}
	
	.well{
		background-color:#fff;
		border:1px solid #f5f5f5;
		box-shadow:0px;	
		border-radius:0px;
		width:100%;
		float:left;
	}
	
	.fa-star{
		color:blue;
	}
	
	.tr-completed{
		background: #dbefd2;
	}
	
	.tr-nextstop{
    	background: #fce5cd;
	}
	
	.tr-AM-border{
		border:2px solid red;
	}
 	
	button[aria-expanded=true]{
	  	background-color: #80B8FF;
	  	color:#fff;
	}
	
	.col-lg-12{
		margin-bottom:3px;
	}
	
	.col-lg-2-border{
		border:1px solid #666;
		padding-top:5px;
		padding-bottom:5px;
		padding-right:0px;
		padding-left:5px;
		margin-right:-1px;
		margin-bottom:-1px;
		font-size:11px;
		min-height: 85px;
	}

	.input-group .icon-addon .form-control {
	    border-radius: 0;
	}
	
	.icon-addon {
	    position: relative;
	    color: #555;
	    display: block;
	}
	
	.icon-addon:after,
	.icon-addon:before {
	    display: table;
	    content: " ";
	}
	
	.icon-addon:after {
	    clear: both;
	}
	
	.icon-addon.addon-md .glyphicon,
	.icon-addon .glyphicon, 
	.icon-addon.addon-md .fa,
	.icon-addon .fa {
	    position: absolute;
	    z-index: 2;
	    left: 10px;
	    font-size: 14px;
	    width: 20px;
	    margin-left: -2.5px;
	    text-align: center;
	    padding: 10px 0;
	    top: 1px
	}
	
	.icon-addon.addon-lg .form-control {
	    line-height: 1.33;
	    height: 46px;
	    font-size: 18px;
	    padding: 10px 16px 10px 40px;
	}
	
	.icon-addon.addon-sm .form-control {
	    height: 30px;
	    padding: 5px 10px 5px 28px;
	    font-size: 12px;
	    line-height: 1.5;
	}

	.icon-addon.addon-lg .fa,
	.icon-addon.addon-lg .glyphicon {
	    font-size: 18px;
	    margin-left: 0;
	    left: 11px;
	    top: 4px;
	}
	
	.icon-addon.addon-md .form-control,
	.icon-addon .form-control {
	    padding-left: 30px;
	    float: left;
	    font-weight: normal;
	}
	
	.icon-addon.addon-sm .fa,
	.icon-addon.addon-sm .glyphicon {
	    margin-left: 0;
	    font-size: 12px;
	    left: 5px;
	    top: -1px
	}
	
	.icon-addon .form-control:focus + .glyphicon,
	.icon-addon:hover .glyphicon,
	.icon-addon .form-control:focus + .fa,
	.icon-addon:hover .fa {
	    color: #2580db;
	}
	
	.stopinfo {
	float:right !important;
	font-size:11px;
	font-weight:normal;
	
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
	
	StartTruckElement = (MUV_READ("DBOARD_CURRENT_PAGE")-1) * MaxTrucksPerPage 

	CurPageNum = MUV_READ("DBOARD_CURRENT_PAGE")
	StartTruckElement = ((CurPageNum-1)*MaxTrucksPerPage ) 


	'Adjust for arrays starting at 0, except 1st page
	'If StartTruckElement > 0 Then StartTruckElement = StartTruckElement - 1

	EndTruckElement = StartTruckElement + (MaxTrucksPerPage -1 )	
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
	
	'Write ordered trucks
	For Each TruckNumber In TruckNumsArray 
		SQL_DeliveryBoardSum = "SELECT DISTINCT TruckNumber FROM RT_DeliveryBoard WHERE TruckNumber = " & TruckNumber
		Set rs_DeliveryBoardSum = cnn_DeliveryBoardSum.Execute(SQL_DeliveryBoardSum)
		If not rs_DeliveryBoardSum.EOF Then
		
			CurrentColumn = 0
		
			Do While Not rs_DeliveryBoardSum.Eof
			
				If CurrentColumn = 0 Then 
					%><div class="row" style="margin-right:0px !important; margin-left :0px !important; font-size:14px;"><%
				End If
				
				If DelBoardIgnoreThisRoute(rs_DeliveryBoardSum("TruckNumber")) <> True Then 
				
					DriverUserNo = Trim(GetUserNumberByTruckNumber(rs_DeliveryBoardSum("TruckNumber")))
				
					If userIsArchived(DriverUserNo) = False AND userIsEnabled(DriverUserNo) = True Then
				
						%><div class="col-lg-6"><%
						Call TruckNumberWrite(rs_DeliveryBoardSum("TruckNumber"), GridColumn)
						%></div><% 
						CurrentColumn = CurrentColumn + 1
						
					End If
					
				End If

				If CurrentColumn = 2 Then 
					%></div><%
				End If
				
				rs_DeliveryBoardSum.Movenext
			Loop
		End If
	Next
	
%> 

<%
Sub TruckNumberWrite(TruckNumber, GridColumn)

	TotalStops = GetTotalStopsByTruckNumber(TruckNumber)
	RemainingStops = GetRemainingStopsByTruckNumber(TruckNumber)
	CurrentStop = Abs(RemainingStops - TotalStops)
	
	If TotalStops > 0 Then
		PercentComplete = Round(((TotalStops - RemainingStops) / TotalStops) * 100)
	Else
		PercentComplete = 0
	End If
		
	Response.Write("<div class='col-lg-12 col-lg-hide' TruckNumber='" & TruckNumber & "'>") %>

	<button class="btn-truck" role="button" data-toggle="collapse" href="#<%=TruckNumber%>" aria-expanded="false" aria-controls="<%=TruckNumber%>">
         
		<div class="col-lg-12">
			<div class="col-lg-6">Route: <%= TruckNumber %>:  <%= GetUserDisplayNameByUserNo(Trim(GetUserNumberByTruckNumber(TruckNumber))) %></div>
			<div class="col-lg-6"><span class="stopinfo">Total Stops: <%= TotalStops %>, Remaining: <%= RemainingStops %></span></div>
		</div>
		
		<% If TotalStops > 0 Then %>
			<div class="col-lg-12">
				<div class="progress">
					<div class="progress-bar progress-bar-success" role="progressbar" aria-valuenow="40" aria-valuemin="0" aria-valuemax="100" style="width: <%= PercentComplete %>%">
						<span class="sr-only"><%= PercentComplete %>% Complete (success)</span>
					</div>
					<span class="progress-completed"><%= PercentComplete %>%</span>
				</div>				
			</div>	
		<% End If %>
		
	</button>
       
	<%
	Response.Write("</div>")
 	GridColumn = GridColumn +1
 	
End Sub 

Set rs_DeliveryBoardSum = Nothing
cnn_DeliveryBoardSum.Close
Set cnn_DeliveryBoardSum = Nothing
%>	


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
	If IsNull(DelBoardAMColor) Then DelBoardAMColor = "#000000"
	If DelBoardPriorityColor = "" Then DelBoardPriorityColor = "#000000"
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

%><!--#include file="../../../inc/footer-noTimeount.asp"-->
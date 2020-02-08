<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/jquery_table_search.asp"-->
<!--#include file="../../inc/InSightFuncs_BizIntel.asp"--> 
<!--#include file="../../inc/InSightFuncs_Equipment.asp"--> 
<!--#include file="../../inc/InsightFuncs_AR_AP.asp"-->
<%


'**********************************************************************************
'HELPER FUNCTION TO GET LAST DAY OF MONTH
'**********************************************************************************
Function GetLastDayofMonth(aDate)
    dim intMonth
    dim dteFirstDayNextMonth

    dtefirstdaynextmonth = dateserial(year(adate),month(adate) + 1, 1)
    GetLastDayofMonth = Day(DateAdd ("d", -1, dteFirstDayNextMonth))
End Function

'**********************************************************************************

Server.ScriptTimeout = 900000 'Default value


Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs2 = Server.CreateObject("ADODB.Recordset")
rs2.CursorLocation = 3 

%>
<style>

	.well {
	    box-shadow: 0 0 4px 0 rgba(0,0,0,.08),0 2px 4px 0 rgba(0,0,0,.12);
	    border-radius: 4px;
	
	}
	.container {
	    width: 100%;
	}

	#chartdiv,
	#chartdivSls2,
	#chartdivSls1,
	#chartdivCustType,
	#chartdivRef{
		width: 100%;
		height: 100%;
	}	
	
	.filter-search-width{
		max-width: 36%;
	}
	
	table.sortable th:not(.sorttable_sorted):not(.sorttable_sorted_reverse):not(.sorttable_nosort):after { 
	    content: " \25B4\25BE" 
	    
	}
	
	table.sortable thead {
	    color:#222;
	    font-weight: bold;
	    cursor: pointer;
	}

	#PleaseWaitPanel{
		position: fixed;
		left: 470px;
		top: 275px;
		width: 975px;
		height: 300px;
		z-index: 9999;
		background-color: #fff;
		opacity:1.0;
		text-align:center;
	}    

	.referral-color{
		background: #D43F3A;
		color:#fff;
		text-align:center;
	}

	.referral-header{
		background: #D43F3A;
		color:#fff;
		text-align:center;
		width:20%;
	}
	
	.cust-type-header{
		background: #F0AD4E;
		color:#fff;
		text-align:center;
		font-weight:bold;
		width:20%;
	}

	.cust-type-color{
		background: #F0AD4E;
		color:#fff;
		text-align:center;
		font-weight:bold;
	}
	
	.primary-slsmn-header{
		background: #337AB7;
		color:#fff;
		text-align:center;
		font-weight:bold;
		width:20%;
	}

	.primary-slsmn-color{
		background: #337AB7;
		color:#fff;
		text-align:center;
		font-weight:bold;
	}

	.secondary-slsmn-header{
		background: #5CB85C;
		color:#fff;
		text-align:center;
		font-weight:bold;
		width:20%;
	}

	.gen-info-header{
		background: #3B579D;
		color:#fff;
		text-align:center;
		font-weight:bold;
		width:60%;
	}
	
	.dollar-amount-header{
		background: #808080;
		color:#fff;
		text-align:center;
		width:16%;
		font-size: 0.8em;
	}


	.dollar-amount-footerNum{
		background: #808080;
		color:#fff;
		text-align:right;
		width:16%;
		font-size: 0.8em;
	}


	.percent-header{
		background: #808080;
		color:#fff;
		text-align:center;
		width:16%;
		font-size: 0.8em;
	}
		
	.dataTables_wrapper .dataTables_filter input {
	    margin-left: 0.5em;
	    box-shadow: none;
  		border-radius:6px;
		-webkit-border-radius: 6px;
		-moz-border-radius: 6px; 
	    padding: 3px;
	    border: solid 1px #E4E4E4;
	    background-color: #fff;		
	    margin-top:10px;
	    margin-bottom:10px;
	}	
	
	.negative{
		font-weight:normal;
		color:red;	
	}

	.neutral{
		font-weight:bold;
		color:black;
	}


	.smaller-detail-line{
		font-size: 0.8em;
	}	

	.fixed-col-header{
		color:#000;
		font-size: 1.1em;
		text-align:left;
		font-weight:bold;
		margin-bottom:10px;
		margin-left: -10px;
    	margin-right: 63px;
    	/*width: 350px;*/
	}

	.fixed-col {
	    height: 100%;
	    background-color: #fff;
	    text-align: center;
	    margin-right: 20px;
	    /*overflow-y: scroll;*/
	    border: solid 1px #000;
	    /*width: 430px;*/
	}
		
	.headerText {
		display: inline-block;
		text-align:left;
		vertical-align:middle;
		margin-left:10px;
	}

	.smaller-header{
		font-size: 0.8em;
		vertical-align: top !important;
		text-align: center;
	}	
	
	
	.red{
		font-weight:bold;
		color:red;	
	}

	.blue{
		font-weight:bold;
		color:blue;	
	}
	
	.pct14 {
	  width: 14%;
	  max-width: 14%;
	  word-wrap: break-word;
	}  

	.pct10 {
	  width: 10%;
	  max-width: 10%;
	  word-wrap: break-word;
	}  

	.pct9 {
	  width: 9%;
	  max-width: 9%;
	  word-wrap: break-word;
	}  

	.pct8 {
	  width: 8%;
	  max-width: 8%;
	  word-wrap: break-word;
	}  

	.pct7 {
	  width: 7%;
	  max-width: 7%;
	  word-wrap: break-word;
	}  
	.pct6 {
	  width: 6%;
	  max-width: 6%;
	  word-wrap: break-word;
	}
</style>

<%
	Response.Write("<div id=""PleaseWaitPanel"" class=""container"">")
	Response.Write("<br><br>Business Intelligence Dashboard Loading<br><br>Please wait...<br><br>")
	Response.Write("<img src='" & baseURL & "/img/loading.gif'/>")
	Response.Write("</div>")
	Response.Flush()
%>

<!-- Begin Cotainer Top Row -->
<div class="container">



 <!--#include file="../../bizintel/dashboard/dashboard_screen_header.asp"--> 

	<!-- Begin Top Row -->
	<div class="row">
	    
	    <!-- Nav tabs -->
	    <ul class="nav nav-tabs" role="tablist">
	      <li class="active"><a href="#tab1" role="tab" data-toggle="tab">Segments</a></li>
	      <li><a href="#tab2" role="tab" data-toggle="tab">Graphs</a></li>
	    </ul>

		<div class="tab-content">
		
			<div class="tab-pane fade active in" id="tab1" style="padding-top:20px;">

				<div class="col-md-12">
		    
			  	  <div class="panel panel-default">
			          <div class="panel-heading">
			            <div class="row">
				              <div class="col-xs-12">
				
								    <!-- Nav tabs -->
								    <ul class="nav nav-tabs" role="tablist">
								      <li><a href="#referral" role="tab" data-toggle="tab">Referral</a></li>
								      <li><a href="#custtype" role="tab" data-toggle="tab">Cust Type</a></li>
								      <li><a href="#sls1" role="tab" data-toggle="tab"><%= GetTerm("Primary Salesman") %></a></li>
								      <li class="active"><a href="#sls2" role="tab" data-toggle="tab"><%= GetTerm("Secondary Salesman") %></a></li>
								    </ul>

								    <!-- Tab panes -->
								    <div class="tab-content">
								      <div class="tab-pane fade" id="referral">
								      		<!--#include file="table_LCPSales_overview_referralnew.asp"-->
								      </div>
								      <div class="tab-pane fade" id="custtype">
								      		<!--#include file="table_LCPSales_overview_custtypenew.asp"-->
								      </div>
								      <div class="tab-pane fade" id="sls1">
											<!--#include file="table_LCPSales_overview_sls1new.asp"-->
								      </div>
								      <div class="tab-pane fade active in" id="sls2">
								     	 	<!--#include file="table_LCPSales_overview_sls2new.asp"--> 
								      </div>
								
									</div>
				              </div>
	    	        </div>
	    	      </div>
	    	   </div>
			</div>
		

			</div>		


			<div class="tab-pane fade  in" id="tab2" style="padding-top:20px;">

				<div class="col-md-2">
				    <div class="panel panel-info">
			          <div class="panel-heading">
			            <div class="row">
			              <div class="col-xs-12">
			               <%'   <!--#include file="graphs/graph_leakage_overview_referral.asp"--> %>
			              </div>
			            </div>
			          </div>
			        </div>
				</div>
		
				<div class="col-md-2">
				    <div class="panel panel-info">
			          <div class="panel-heading">
			            <div class="row">
			              <div class="col-xs-12">
			                  <!--#include file="graphs/graph_overview_custtype.asp"-->
			              </div>
			            </div>
			          </div>
			        </div>
				</div>
		
				<div class="col-md-2">
				    <div class="panel panel-info">
			          <div class="panel-heading">
			            <div class="row">
			              <div class="col-xs-12">
			                  <%'  <!--#include file="graphs/graph_leakage_overview_sls1.asp"--> %>
			              </div>
			            </div>
			          </div>
			        </div>
				</div>

				<div class="col-md-2">
				    <div class="panel panel-info">
			          <div class="panel-heading">
			            <div class="row">
			              <div class="col-xs-12">
			                 <%'   <!--#include file="graphs/graph_leakage_overview_sls2.asp"--> %>
			              </div>
			            </div>
			          </div>
			        </div>
				</div>
		
				<div class="col-md-4">
				    <div class="panel panel-warning">
			          <div class="panel-heading">
			            <div class="row">
			              <div class="col-xs-12">
			                 <%'   <!--#include file="graphs/graph_InvoiceHistoryCounts.asp"--> %>
			              </div>
			            </div>
			          </div>
			        </div>
				</div>


				<!----------------------------------------->
				<!----------------------------------------->				
				<!-- Second set of pie charts start here -->
				<!----------------------------------------->
				<!----------------------------------------->								
	
				<div class="row">
					<div class="col-md-2">
								    <div class="panel panel-info">
							          <div class="panel-heading">
							            <div class="row">
							              <div class="col-xs-12">
							                  <%'  <!--#include file="graphs/graph_SALES_overview_referral.asp"--> %>
							              </div>
							            </div>
							          </div>
							        </div>
								</div>
						
								<div class="col-md-2">
								    <div class="panel panel-info">
							          <div class="panel-heading">
							            <div class="row">
							              <div class="col-xs-12">
							                <%'    <!--#include file="graphs/graph_SALES_overview_custtype.asp"--> %>
							              </div>
							            </div>
							          </div>
							        </div>
								</div>
						
								<div class="col-md-2">
								    <div class="panel panel-info">
							          <div class="panel-heading">
							            <div class="row">
							              <div class="col-xs-12">
							                  <%'  <!--#include file="graphs/graph_SALES_overview_sls1.asp"--> %>
							              </div>
							            </div>
							          </div>
							        </div>
								</div>
				
								<div class="col-md-2">
								    <div class="panel panel-info">
							          <div class="panel-heading">
							            <div class="row">
							              <div class="col-xs-12">
							                <%'    <!--#include file="graphs/graph_SALES_overview_sls2.asp"--> %>
							              </div>
							            </div>
							          </div>
							        </div>
								</div>
						
								<div class="col-md-4">
								    <div class="panel panel-warning">
							          <div class="panel-heading">
							            <div class="row">
							              <div class="col-xs-12">
							               <%'   <!--#include file="graphs/graph_InvoiceHistoryDollars.asp"--> %>
							              </div>
							            </div>
							          </div>
							        </div>
								</div>

						</div>			


				<!----------------------------------------->
				<!----------------------------------------->				
				<!-- Second set of pie charts start here -->
				<!----------------------------------------->
				<!----------------------------------------->								
	
				<div class="row">
				
					<div class="col-md-2">
					    &nbsp;
					</div>
			
					<div class="col-md-2">
					    &nbsp;
					</div>
			
					<div class="col-md-2">
					    &nbsp;
					</div>
	
					<div class="col-md-2">
					    &nbsp;
					</div>
			
					<div class="col-md-4">
					    <div class="panel panel-warning">
				          <div class="panel-heading">
				            <div class="row">
				              <div class="col-xs-12">
				                 <%' <!--#include file="graphs/graph_sales_and_leakage.asp"--> %>
				              </div>
				            </div>
				          </div>
				        </div>
					</div>

			</div>					
		</div>
		
	</div><!-- End Top Row -->

	
</div><!-- End Container -->
	

<!--#include file="table_functions.asp"-->
<!--#include file="graphs/graph_dashboard_footer_js_code.asp"-->
<!--#include file="graphs/graph_dashboard_footer_LCPSales_js_code.asp"-->
<!--#include file="../../inc/footer-main.asp"-->
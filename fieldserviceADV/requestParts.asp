<!--#include file="../inc/header-field-service-mobile.asp"-->
<!--#include file="../inc/InsightFuncs_Service.asp"-->

<style type="text/css">
	.fieldservice-heading{
		line-height: 1;
		padding-top: 5px;
		padding-bottom: 5px;
	}
	
	.btn-home{
		color: #fff;
		margin-top: -2px;
		margin-left: 5px;
		float: left;
 	}
	
 ul{
	 color: #666;
	 font-size: 13px;
	 text-transform: uppercase;
	 list-style-type: none;
	     -webkit-margin-before: 0px;
    -webkit-margin-after: 0px;
    -webkit-margin-start: 0px;
    -webkit-margin-end: 0px;
    -webkit-padding-start: 0px;
 }
 
 .enroute{
	 color: green;
 }
 
 .btn-spacing{
	 margin-bottom: 40px;
 }
 
 
 
.btn-block {
    width: auto;
    display: inline-block;
    margin-right:2em;
}
 
 .row{
 	/* flex-wrap: nowrap !important; */
 }

 
 @media (max-width: 767px) {
 	.mob-col{
 		/* width: auto !important;  */
 	}
 }
 
 .driver-menu{
 	text-align:center;
 	margin-bottom:10px;
 	margin-top:10px;
 }
 
.badge-pill-icon-letter {
    padding-right: .3em;
    padding-left: .3em;
    border-radius: 8rem;
}	
	
</style>       

<style type="text/css">
 	.email-table{
		width:46%;
	}
	
	table.sortable th:not(.sorttable_sorted):not(.sorttable_sorted_reverse):not(.sorttable_nosort):after { 
    content: " \25B4\25BE" 
}

.nav-tabs>li>a{
	background: #f5f5f5;
	border: 1px solid #ccc;
	color: #000;
}

.nav-tabs>li>a:hover{
	border: 1px solid #ccc;
}

.nav-tabs>li.active>a, .nav-tabs>li.active>a:focus, .nav-tabs>li.active>a:hover{
	color: #000;
	border: 1px solid #ccc;
}

.container{
	max-width:600px;
	margin:0 auto;
}

.narrow-results{
	margin:0px 0px 20px 0px;
}

#filter{
	width:40%;
}

.modal-link{
	cursor:pointer;
}

.modal-content{
	max-height:360px;
	overflow-y:auto;
}

 .modal-content .row{
	 padding-bottom:20px;
 }

 .modal-content p{
	 margin-bottom:20px;
	 white-space:normal;
 }
 
table.dataTable thead tr {
  background-color: #007AFF;
} 
 </style>
 

<script src="http://cdn.datatables.net/1.10.19/js/jquery.dataTables.min.js"></script>
<link href="http://cdn.datatables.net/1.10.19/css/jquery.dataTables.min.css" rel="stylesheet">

<script type="text/javascript">
$(document).ready( function () {
    //$('#myTable').DataTable();
	
	
$('#myTable').dataTable( {
		"paging":   false,
        "columnDefs": [ 
		{"targets": [2,3], "orderable": false,} ]
} );	
	
	
	$(".dataTables_length").hide();
	$(".dataTables_info").hide();
	$(".dataTables_paginate").hide();
} );
</script> 
<h1 class="fieldservice-heading"><a class="btn-home" href="main_menu.asp" role="button"><i class="fa fa-bars"></i></a> Parts Request</h1>

<div class="container-fluid">

<form method="POST" action="requestParts_submit.asp" name="frmRequestpart" id="frmRequestpart">

<div class="table-responsive">
<div id="example_wrapper" class="dataTables_wrapper">
            <table  id="myTable"  class="table table-striped table-condensed table-hover table-bordered sortable">
              <thead>
                <tr>
                  <th class="sorttable_nosort">Part#</th>
				  <th class="sorttable_nosort">Description</th>
				  <th class="sorttable_nosort" style="display:none;">Search Keywords</th>
                  <th class="sorttable_nosort" style="width:20px;">Qty</th>
                  <th class="sorttable_nosort">Notes</th>
                </tr>
              </thead>
               <tbody class='searchable'>
              
				<%
			
				SQL = "SELECT * FROM FS_Parts WHERE InternalRecordIdentifier > 0 order by PartNumber"
		
				Set cnnparts = Server.CreateObject("ADODB.Connection")
				cnnparts.open (Session("ClientCnnString"))
				Set rsparts = Server.CreateObject("ADODB.Recordset")
				rsparts.CursorLocation = 3 
				Set rsparts = cnnparts.Execute(SQL)
		
				If not rsparts.EOF Then

					Do While Not rsparts.EOF
				
			        %>
						<!-- table line !-->
						<tr>							
							<td><%= rsparts.Fields("PartNumber")%>
							<input type="hidden" class="form-control" id="txtPartNumber[]" name="txtPartNumber[]" value="<%= rsparts.Fields("PartNumber")%>"></td>
							<td><%= rsparts.Fields("PartDescription")%>
							<input type="hidden" class="form-control" id="txtPartDescription[]" name="txtPartDescription[]" value="<%= rsparts.Fields("PartDescription")%>"></td>
							<td style="display:none;"><%= rsparts.Fields("SearchKeywords")%>
							<input type="hidden" class="form-control" id="txtSearchKeywords[]" name="txtSearchKeywords[]" value="<%= rsparts.Fields("SearchKeywords")%>">
							</td>							
							<!--<td><input type="text" class="form-control" id="txtPartQty[]" name="txtPartQty[]" placeholder="0"></td>-->
							<td>
							<select id="txtPartQty[]" name="txtPartQty[]">
								<option value="0">0</option>
								<option value="1">1</option>
								<option value="2">2</option>
								<option value="3">3</option>
								<option value="4">4</option>
								<option value="5">5</option>
								<option value="6">6</option>
								<option value="7">7</option>
								<option value="8">8</option>
								<option value="9">9</option>
								<option value="10">10</option>
								<option value="11">11</option>
								<option value="12">12</option>
								<option value="13">13</option>
								<option value="14">14</option>
								<option value="15">15</option>
								<option value="16">16</option>
								<option value="17">17</option>
								<option value="18">18</option>
								<option value="19">19</option>
								<option value="20">20</option>
								<option value="21">21</option>
								<option value="22">22</option>
								<option value="23">23</option>
								<option value="24">24</option>
								<option value="25">25</option>								
							<select>
							<td><input type="text" class="form-control" id="txtPartNotes[]" name="txtPartNotes[]" ></td>
					   	</tr>
					<%
						rsparts.movenext
					loop
				End If
				set rsparts = Nothing
				cnnparts.close
				set cnnparts = Nothing
	            %>
			<!--<tr>
				<td>Others</td>
				<td colspan="2"><input type="text" class="form-control" id="txtPartDesc" name="txtPartDesc" ></td>
				<td style="display:none;"></td>
			</tr>-->
				
			</tbody>
			
		</table>
		</div>
	</div>
		<table class="table table-striped table-condensed table-hover table-bordered sortable">
			<tr>
				<td>Others notes or parts not listed here:</td>
			</tr>
			<tr>	
				<td colspan="2">
				<!--<input type="text" class="form-control" id="txtPartDesc" name="txtPartDesc" >-->
				<textarea class="form-control" id="txtPartDesc" name="txtPartDesc" rows="4"></textarea>
				</td>
				<td style="display:none;"></td>
			</tr>	
		</table>
	<button type="submit" class="btn btn-primary" align="center"><i class="far fa-save"></i> Request Parts</button>
</form>	
</div>






<!--#include file="../inc/footer-field-service-noTimeout.asp"-->
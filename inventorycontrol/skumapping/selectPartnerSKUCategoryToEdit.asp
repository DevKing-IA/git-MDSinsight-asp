<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/jquery_table_search.asp"-->
<!--#include file="../../inc/InsightFuncs_InventoryControl.asp"-->



 
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
	max-width:1100px;
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
 .command-panel {
    margin: 45px 0 20px;
   }
 tr.for-select {cursor:pointer;}
 tr.for-select.row-selected,tr.for-select.row-selected:hover, .table-striped>tbody>tr.row-selected:nth-of-type(odd) {background-color:#337ab7;color:#ffffff;}

 </style>

<script type="text/javascript">
    $(window).on("load", function () {
        $("tr.for-select").on("click", function () {
            $("tr.for-select").removeClass("row-selected");

            if ($(this).find("td.SKU-qty").html() != "0") {
                if ($(".csv-sku").hasClass("disabled")) {
                    $(".csv-sku").removeClass("disabled");
                    $(".csv-sku").on("click", function () {

                        location.href = "exportSKUCategoryToCSV.asp?i=" + $("tr.for-select.row-selected").attr("data-i") + "&c=" + $("tr.for-select.row-selected").attr("data-c");
                    });
                }
            }
            else {
                $(".csv-sku").addClass("disabled");
                $(".csv-sku").off("click");

            }
            $(this).addClass("row-selected");
            if ($(".selected-sku").hasClass("disabled")) {
                $(".selected-sku").removeClass("disabled");
                $(".selected-sku").on("click", function () {
                    location.href = "editPartnerSKUCategoryToEdit.asp?i=" + $("tr.for-select.row-selected").attr("data-i") + "&c=" + $("tr.for-select.row-selected").attr("data-c");
                });
            }
        });
        $(".all-sku").on("click", function () {
            location.href = "editPartnerSKUCategoryToEdit.asp?i=" + $("tr.for-select.all").attr("data-i") + "&c=" + $("tr.for-select.all").attr("data-c");
        });
    });
</script>
<!--- eof on/off scripts !-->
<div class="container-fluid">
    <div class="row ">
        <div class="col-lg-6 col-md-6 col-sm-6 col-xs-6">
            <h1 class="page-header"><i class="fa fa-map-marker" aria-hidden="true"></i> SKU Mapping Tool - Category Selection</h1>
        </div>
        <div class="col-lg-6 col-md-6 col-sm-6 col-xs-6 command-panel">
            
            <button type="button" class="btn btn-success selected-sku disabled">MAP SKUS FOR SELECTED <i class="fa fa-arrow-circle-o-right" aria-hidden="true"></i></button>
            <button type="button" class="btn btn-warning csv-sku disabled">DOWNLOAD AS .csv&nbsp;<span class="glyphicon glyphicon-save-file" aria-hidden="true"></span></button>
        </div>
    </div>
</div>


	<!-- tabs start here !-->
	<div class="container">

	<div class="table-responsive">
            <table    class="table table-striped table-condensed table-hover table-bordered sortable for-datatable">
              <thead>
                <tr>
                  <th>Category</th>
                  <th>SKUs in Category</th>
                  <th>Equivalent SKUs in Category</th>
                  
                </tr>
              </thead>
              <tbody class='searchable'>
              
				<%
				
				InternalRecordIdentifier = Request.QueryString("i")

						SQL9 = "SELECT COUNT(PartNo) as TotalSKUCount FROM Product" 
						Set cnn9 = Server.CreateObject("ADODB.Connection")
						cnn9.open (Session("ClientCnnString"))
						Set rs9 = Server.CreateObject("ADODB.Recordset")
						rs9.CursorLocation = 3 
						Set rs9 = cnn9.Execute(SQL9)
						If not rs9.EOF Then
							TotalSKUCount = rs9("TotalSKUCount")
						Else
							TotalSKUCount = 0
						End If
						
						SQL9 = "SELECT COUNT(SKU) as TotalEquivalentSKUCount FROM IC_ProductMapping WHERE partnerIntRecID = " & InternalRecordIdentifier
						Set rs9 = cnn9.Execute(SQL9)
						If not rs9.EOF Then
							TotalEquivalentSKUCount = rs9("TotalEquivalentSKUCount")
						Else
							TotalEquivalentSKUCount = 0
						End If
						set rs9 = Nothing
						cnn9.close
						set cnn9 = Nothing
						

				%>              
				<tr class="for-select all" data-i="<%= InternalRecordIdentifier %>" data-c="all">
					<td>ALL PRODUCT CATEGORIES</td>
					<td class="SKU-qty"><%= TotalSKUCount %></td>
					<td><%= TotalEquivalentSKUCount %></td>	
					
			   	</tr>
             
				<%
				
				
			
				SQL = "SELECT * FROM tblCategories ORDER BY CategoryID ASC"
		
				Set cnn8 = Server.CreateObject("ADODB.Connection")
				cnn8.open (Session("ClientCnnString"))
				Set rs = Server.CreateObject("ADODB.Recordset")
				rs.CursorLocation = 3 
				Set rs = cnn8.Execute(SQL)
		
				If NOT rs.EOF Then

					Do While Not rs.EOF
					
						CategoryID = rs.Fields("CategoryID")
						CategoryName = rs.Fields("CategoryName")	
					
						SQL9 = "SELECT COUNT(PartNo) as TotalSKUCount FROM Product WHERE Category = " & CategoryID
						Set cnn9 = Server.CreateObject("ADODB.Connection")
						cnn9.open (Session("ClientCnnString"))
						Set rs9 = Server.CreateObject("ADODB.Recordset")
						rs9.CursorLocation = 3 

						Set rs9 = cnn9.Execute(SQL9)
						If not rs9.EOF Then
							TotalSKUCount = rs9("TotalSKUCount")
						Else
							TotalSKUCount = 0
						End If
						
						SQL9 = "SELECT COUNT(SKU) as TotalEquivalentSKUCount FROM IC_ProductMapping WHERE CategoryID = " & CategoryID & " AND partnerIntRecID = " & InternalRecordIdentifier
						Set rs9 = cnn9.Execute(SQL9)
						If not rs9.EOF Then
							TotalEquivalentSKUCount = rs9("TotalEquivalentSKUCount")
						Else
							TotalEquivalentSKUCount = 0
						End If
						set rs9 = Nothing
						cnn9.close
						set cnn9 = Nothing
						

				
			        %>
						<!-- table line !-->
						<tr class="for-select" data-i="<%= InternalRecordIdentifier %>" data-c="<%= CategoryID %>">
							<td><%= CategoryName %></td>
							<td class="SKU-qty"><%= TotalSKUCount %></td>
							<td><%= TotalEquivalentSKUCount %></td>	
							
					   	</tr>
					<%
						rs.movenext
					Loop
				End If
				set rs = Nothing
				cnn8.close
				set cnn8 = Nothing
	            %>
			</tbody>
		</table>
	</div>
 
		</div>

</div>
<!-- eof row !-->
								

<!--#include file="../../inc/footer-main.asp"-->
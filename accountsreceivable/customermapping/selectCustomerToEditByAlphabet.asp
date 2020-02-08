<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/jquery_table_search.asp"-->
<!--#include file="../../inc/InsightFuncs_AR_AP.asp"-->

<%
				
				InternalRecordIdentifier = Request.QueryString("i")

				SQL9 = "SELECT COUNT(CustNum) as TotalCustomerCount FROM AR_Customer WHERE AcctStatus = 'A'" 
				Set cnn9 = Server.CreateObject("ADODB.Connection")
				cnn9.open (Session("ClientCnnString"))
				Set rs9 = Server.CreateObject("ADODB.Recordset")
				rs9.CursorLocation = 3 
				Set rs9 = cnn9.Execute(SQL9)
				If not rs9.EOF Then
					TotalCustomerCount = rs9("TotalCustomerCount")
				Else
					TotalCustomerCount = 0
				End If
				
				SQL9 = "SELECT COUNT(partnerCustID) as TotalEquivalentPartnerCustCount FROM AR_CustomerMapping WHERE partnerRecID = " & InternalRecordIdentifier
				Set rs9 = cnn9.Execute(SQL9)
				If not rs9.EOF Then
					TotalEquivalentPartnerCustCount = rs9("TotalEquivalentPartnerCustCount")
				Else
					TotalEquivalentPartnerCustCount = 0
				End If
				set rs9 = Nothing
				cnn9.close
				set cnn9 = Nothing
						

				%>              

 
 <style type="text/css">
 
	
	
 	.email-table{
		width:46%;
	}
	
	
	
	.container{
		max-width:1100px;
		margin:0 auto;
	}
     .command-panel {
    margin: 45px 0 20px;
   }
     .table-striped>tbody>tr.for-select.selected,.table-striped>tbody>tr:nth-of-type(odd).selected {
         color: #fff;
    cursor: default;
    background-color: #337ab7;
    border-color: #337ab7;
     }
      .table-striped>tbody>tr.for-select.selected input {color:#000000;}
 </style>
<script type="text/javascript">
    $(window).on("load", function () {
        $('a[data-toggle="tab"]').on('shown.bs.tab', function (e) {

           
            if ($("#" + $(e.target).attr("aria-controls")).html().trim().length == 0) loadDataTabs($(e.target).attr("aria-controls"), 1);
            else {
                if ($("#" + $(e.target).attr("aria-controls")).find(".total-customer").html() == "0") {
                    $(".csv-map").addClass("disabled");
                    $(".csv-map").removeAttr("onclick");
                }
                else {
                    $(".csv-map").removeClass("disabled");
                    $(".csv-map").attr("onclick","javascript:toCSV();");
                }
            }
        });
        loadDataTabs("all",1);
    });

    function loadDataTabs(letter, page) {
        $(".waitdiv").css("display", "block");
        $("div.tab-pane.active").load(
            "tablebyletter.asp",
            "letter="+letter+"&i=<%=InternalRecordIdentifier%>&page="+page+"&pagesize="+$(".rows-per-page").val()+"&filterdata="+$("#filter").val(),
            function (responseTxt, statusTxt, xhr) {
                if (statusTxt == "success") {
                    $(".waitdiv").css("display", "none");
                    if ($("div.tab-pane.active .total-customer").html() == "0") {
                        $(".csv-map").addClass("disabled");
                        $(".csv-map").removeAttr("onclick");
                    }
                    else {
                    $(".csv-map").removeClass("disabled");
                    $(".csv-map").attr("onclick","javascript:toCSV();");
                }
                }
                if (statusTxt == "error") {
                    $(".csv-map").addClass("disabled");
                    $(".csv-map").removeAttr("onclick");
                    $(".waitdiv").css("display", "none");
                    alert("Error: " + xhr.status + ": " + xhr.statusText);
                }
            }
        );
    }
    function clearContent() {
        $(".tab-letter>li:not(.active)").each(function () {
            var href = $(this).children("a").attr("href");
            $(href).html("");
        });
    }
    function clearfind() {

    }
    function gotoPage(pageNo) {

        
        loadDataTabs($(".tab-letter li.active").attr("data-letter"),pageNo)
    }
    function selectRow(obj) {
        
        $(".for-select").removeClass("selected");
        $(obj).addClass("selected");
    }
    function editRow(obj) {
        $(obj).find("td:last").attr("data-value", $(obj).find("td:last").html().trim());
        $(obj).find("td:last").html("<input type='text' class='col-lg-8 input-sm' value='" + $(obj).find("td:last").attr("data-value") + "'/><button class='btn btn-default btn-sm' type='button' onclick='javascript:saveCustomerID(this);'><span class='glyphicon glyphicon-ok'></span></button><button class='btn btn-default btn-sm' type='button' onclick='javascript:cancelEdit(this);'><span class='glyphicon glyphicon-remove'></span></button>");
        $(obj).closest("table").find("tr.for-select").removeAttr("onclick").removeAttr("ondblclick").removeAttr("data-toggle");
        $(".pagination").addClass("invisible");
        $(".tab-letter>li:not(.active)").each(function () {
            $(this).children("a").attr("data-href", $(this).children("a").attr("href")).attr("href", "#");
            $(this).children("a").removeAttr("data-toggle");
            $(this).addClass("disabled");

        });
    }
    function cancelEdit(obj) {
        var editedObj = $(obj).parent();

        $(editedObj).html($(editedObj).attr("data-value"));
        $(editedObj).closest("table").find("tr.for-select").attr("onclick", "javascript:selectRow(this);").attr("ondblclick", "javascript:editRow(this);").attr("data-toggle","tooltip");
        $(".pagination").removeClass("invisible");
        $(".tab-letter>li:not(.active)").each(function () {
            $(this).children("a").attr("href", $(this).children("a").attr("data-href"));
            $(this).removeClass("disabled");
            $(this).children("a").attr("data-toggle","tab");

        });
    }
    function saveCustomerID(obj) {
        var inputEnteredValue = $(obj).parent().find("input").val();
        var inputID = $(obj).closest("tr").attr("data-id");
        $(".waitdiv").css("display", "block");
        $.ajax({
				type:"POST",
				url: "../../inc/InSightFuncs_AjaxForARAP.asp",
				cache: false,
				data: "action=SaveEquivalentCustomerAccount&id=" + encodeURIComponent(inputID) + "&equivID=" + encodeURIComponent(inputEnteredValue),
				success: function(response)
				 {
				 	if (response == "Success") {
				 		//If successfully saved, change the style of the input box, so the user gets a visual cue that the SKU was saved   
				 		//alert("#" + inputID);
                        var editedObj = $(obj).parent();
                        $(editedObj).html(inputEnteredValue);

                        $(editedObj).closest("table").find("tr.for-select").attr("onclick", "javascript:selectRow(this);").attr("ondblclick", "javascript:editRow(this);");
                        $(".pagination").removeClass("invisible");
                        $(".tab-letter>li:not(.active)").each(function () {
                            $(this).children("a").attr("href", $(this).children("a").attr("data-href"));
                            $(this).removeClass("disabled");
                            $(this).children("a").attr("data-toggle","tab");

                          });
                          $(".waitdiv").css("display", "none");
				 	}
                      else {
                          $(".waitdiv").css("display", "none");
				 	    alert(response);
					}       	 
            },

            complete: function () {
                $(".waitdiv").css("display", "none");
            }
			});	//end ajax post to data: "action=saveCustomerIDToEquivalentsTable"
		
    }
    function toCSV() {
        var hrefdata = "exportUserMapToCSV.asp?letter=" + $(".tab-letter li.active").attr("data-letter") + "&i=<%=InternalRecordIdentifier%>";
        location.href = hrefdata;
    }
    function gotoFilter() {
        var href = $(".tab-letter>li.active").attr("data-letter");
        var activePage = $("#" + href + " ul.pagination>li.active a").html();
        clearContent();
        gotoPage("1");
        
    }
</script>
<!--- eof on/off scripts !-->

<!-- <h1 class="page-header"><i class="fa fa-map-marker" aria-hidden="true"></i> Customer Mapping Tool - Account By First Letter Selection</h1> -->
<div class="container-fluid">
    <div class="row ">
        <div class="col-lg-4 col-md-4 col-sm-4 col-xs-4">
            <h1 class="page-header"><i class="fa fa-map-marker" aria-hidden="true"></i>Customer Mapping Table</h1>
        </div>
        <div class="col-lg-3 command-panel">
		    <div class="input-group"> <span class="input-group-addon">Find Customer</span>
		        <input id="filter" type="text" class="form-control " placeholder="Type here...">
                <span class="input-group-btn">
                    <button class="btn btn-default" type="button"><span class="glyphicon glyphicon-search" onclick="javascript:gotoFilter();"></span></button>
                </span>
                <span class="input-group-btn">
                    <button class="btn btn-default" type="button"><span class="glyphicon glyphicon-remove" onclick="javascript:$('#filter').val('');gotoFilter();"></span></button>
                </span>
		    </div>
	    </div>
        <div class="col-lg-2 col-md-2 col-sm-2 col-xs-2 command-panel text-right"> Rows Per Page :</div>
        <div class="col-lg-1 col-md-1 col-sm-1 col-xs-1 command-panel">
           
            <select class="form-control rows-per-page col-lg-5" onchange="javascript:clearContent();gotoPage(1);">
                <option value="10" selected="selected">10</option>
                <option value="50">50</option>
                <option value="100">100</option>
                <option value="500">500</option>
    
            </select>
        </div>
        <div class="col-lg-2 col-md-2 col-sm-2 col-xs-2 command-panel">
            
            
            <button onclick="javascript:toCSV();" type="button" class="btn btn-warning csv-map disabled">DOWNLOAD AS .csv&nbsp;<span class="glyphicon glyphicon-save-file" aria-hidden="true"></span></button>
        </div>
    </div>
</div>


<div class="container-fluid">
    <div class="row">
        <div class="col-lg-12 col-md-12 col-sm-12 col-xs-12">
            <ul class="nav nav-tabs nav-justified tab-letter" role="tablist">
                <li role="presentation" class="active" data-i="<%= Request.QueryString("i")%>" data-letter="all"><a href="#all" aria-controls="all" role="tab" data-toggle="tab">All</a></li>
                <% for i = asc("A") to asc("Z") %>
                <li role="presentation" data-letter="<%= chr(i) %>" data-i="<%= Request.QueryString("i")%>"><a href="#<%= chr(i) %>" aria-controls="<%= chr(i) %>" role="tab" data-toggle="tab" ><%= chr(i) %></a></li>
	            <% next %>
            </ul>
            <!-- Tab panes -->
            <div class="tab-content">
                <div role="tabpanel" class="tab-pane fade in active" id="all">
                     
                </div>
                <% for i = asc("A") to asc("Z") %>
                    <div role="tabpanel" class="tab-pane fade" id="<%= chr(i) %>">
                       
                    </div>
                <% next %>
               
            </div>
        </div>
    </div>
</div>

 <div class="waitdiv" style="display:none;position: fixed;z-index: 999999999; top: 0px; left: 0px; width: 100%; height:80%; background-color:transparent; text-align: center; padding-top: 20%; filter: alpha(opacity=0); opacity:0; "></div>
    <div class="waitdiv" style="padding-bottom: 90px;text-align: center; vertical-align:middle;padding-top:50px;background-color:#ebebeb;width:300px;height:100px;margin: 0 auto; top:40%; left:40%;position:fixed;display:none;-webkit-box-shadow: 0 5px 10px rgba(0, 0, 0, 0.2); -moz-box-shadow: 0 5px 10px rgba(0, 0, 0, 0.2); box-shadow: 0 5px 10px rgba(0, 0, 0, 0.2); z-index:999999999;">
        <img src="/img/loading_wait.gif" alt="" /><br />Request to Server. Please wait ...
    </div>
</div>
<!-- eof row !-->						

<!--#include file="../../inc/footer-main.asp"-->
<!--#include file="../../../../inc/header.asp"-->
<!--#include file="../../../../inc/insightfuncs.asp"-->

<!-- function that gets the value of the tab when it is clicked and then
updates the value of a hidden form field so when the page posts, it returns
back to the tab that was previously opened -->

<script type="text/javascript">
	$(function () {
		$('a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
		var target = $(e.target).attr("href");
		$('input[name="txtTab"]').val(target);
		//alert(target);
		});
	})
</script>

<%
'Check to see if there is a querystring value for 's'
'Any value at all indicates a failure to read the 
'seetings from UNIX
If Request.QueryString("s") <> "" Then %>
	<div class="col-lg-6 days-hours">
		<div class="table-responsive"> 
			<br><font color="red"><strong>
			Insight was	unable to read the quote information from <%=GetTerm("Backend")%>.<br>
			Please try again and contact techincal support if<br> the problem continues.<br><br>
			</strong></font>
		</div>	
	</div>
<%End If%> 

<%
ClientKeyForFileNames = MUV_READ("ClientKeyForFileNames")
%>

<script type="text/javascript">

	$(function () {
		var autocompleteJSONFileURL = "../../../../clientfiles/<%= ClientKeyForFileNames %>/autocomplete/customer_account_list_CSZ_<%= ClientKeyForFileNames %>.json";
		
		var options = {
		  url: autocompleteJSONFileURL,
		  placeholder: "Search for a customer by name, account, city, state, zip",
		  getValue: "name",
		  list: {	
	        onChooseEvent: function() {
	            var custID = $("#txtCustID").getSelectedItemData().code;
	            window.location = "readQuotesloading.asp?custID="+custID;
        	},		  
		    match: {
		      enabled: true
			},
			maxNumberOfElements: 150		
		  },
		  theme: "round"
		};
		$("#txtCustID").easyAutocomplete(options);

	})
</script>


<style> 
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
</style>
	
<script type="text/javascript">
	$(document).ready(function() {
	    $("#PleaseWaitPanel").hide();
	});
</script>
  
  
<%
Response.Write("<div id=""PleaseWaitPanel"">")
Response.Write("<br><br>Processing " & GetTerm("Customer") & " Center, please wait...<br><br>")
Response.Write("<img src=""../../../../img/loading.gif"" />")
Response.Write("</div>")
Response.Flush()
%>

<% SelectedCustomer = Session("ServiceCustID") %>


<h1 class="page-header"><i class="fa fa-file-text-o"></i> <%=GetTerm("Customer")%> Quotes</h1>

 

<!-- row !-->
<div class="row row-line">
    <div class="col-lg-8">
    		<div class="row">
        		<!-- select company !-->
		        <div class="col-lg-6 col-md-3 col-sm-12 col-xs-12">
					<input id="txtCustID" name="txtCustID">
					<i id="searchIcon" class="fa fa-search fa-2x"></i>
				</div>
				<!-- eof select company !-->
		</div>
	</div>
</div>
<!-- eof row !-->
   


<!--#include file="../../../../inc/footer-service.asp"-->
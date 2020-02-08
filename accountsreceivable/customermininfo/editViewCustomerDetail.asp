<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/InsightFuncs_AR_AP.asp"-->
<!--#include file="../../inc/InsightFuncs_Prospecting.asp"-->
<!--#include file="../../inc/InsightFuncs_Service.ASP"-->
<!--#include file="../../inc/InsightFuncs_AR_AP.asp"-->
<!--#include file="../../css/fa_animation_styles.css"-->
<!--#include file="editViewCustomerDetailStylesheet.asp"-->

<script src="<%= baseURL %>js/moment.min.js" type="text/javascript"></script>
<script type="text/javascript" src="<%= BaseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.min.js"></script>
<link href="<%= BaseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.css" rel="stylesheet">

<% 

InternalRecordIdentifier = Request.QueryString("i") 

customerID = Request.QueryString("cid")

If customerID = "" Then
	customerID = REQUEST("customerID")
End If

InternalRecordIdentifier = customerID

If InternalRecordIdentifier = "" Then Response.Redirect("main.asp")


OpenTabNum = 1 'Dfault open tab #
OpenTabNum = Request.QueryString("t") 


SQL = "SELECT * FROM AR_Customer WHERE CustNum = '" & customerID & "'"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)

If NOT rs.EOF Then
	InternalRecordIdentifier = rs("InternalRecordIdentifier")
	CustomerName = rs("Name")
	CustomerContactFirstName = rs("ContactFirstName")
	CustomerContactLastName = rs("ContactLastName")	
	CustomerAddr1 = rs("Addr1")
	CustomerAddr2 = rs("Addr2")
	CustomerCity = rs("City")
	CustomerState = rs("State")
	CustomerZip = rs("Zip")
	CustomerCountry = rs("Country")
	CustomerPhone = rs("Phone")
	CustomerFax = rs("Fax")
	CustomerAcctStatus = rs("AcctStatus")
	CustomerFullName = CustomerContactFirstName & " " & CustomerContactLastName 
	CustomerFullAddress = CustomerAddr1 & " " & CustomerAddr2 & ", " & CustomerCity & ", " & CustomerState & ", " & CustomerZip
	CustomerLastPriceChangeDate = rs("LastPriceChangeDate")
End If

'****************************************************************
'See if we will be using popup message for service tickets
'****************************************************************
ShowPopup = 0
SQL = "SELECT * FROM " & MUV_Read("SQL_Owner") & ".Settings_Global"
Set rs = cnn8.Execute(SQL)
If not rs.eof then ShowPopup = rs("NotesScreenShowPopup")
If ShowPopup = 0 Then ShowPopup = False
If ShowPopup = 1 Then ShowPopup = True

If MUV_READ("ShowServiceTicketAlertSwalShown") <> "true" Then
	dummy = MUV_Write("ShowServiceTicketAlertSwalShown","false")
End If

'Give them a message if there are open tickets

If ShowPopup = True Then ' From global settings table

	OPTick=NumberOfServiceTicketsOpenForCust(customerID)
	HLDTick=NumberOfServiceTicketsHOLDForCust(customerID)
	
	If OPTick <> 0 AND HLDTick <> 0 AND MUV_READ("ShowServiceTicketAlertSwalShown") = "false" Then
		Response.write("<script type=""text/javascript"">swal(""" & GetTerm("Account") & " has " & OPTick & " open service ticket(s) and " & HLDTick & " service ticket(s) on hold"");</script>")
		dummy = MUV_Write("ShowServiceTicketAlertSwalShown","true")
	ElseIF OPTick <> 0 AND MUV_READ("ShowServiceTicketAlertSwalShown") = "false" Then
		Response.write("<script type=""text/javascript"">swal(""" & GetTerm("Account") & " has " & OPTick & " open service ticket(s)"");</script>")
		dummy = MUV_Write("ShowServiceTicketAlertSwalShown","true")
	ElseIF HLDTick<> 0 AND MUV_READ("ShowServiceTicketAlertSwalShown") = "false" Then
		Response.write("<script type=""text/javascript"">swal(""" & GetTerm("Account") & " has " & HLDTick& " service ticket(s) on hold"");</script>")	
		dummy = MUV_Write("ShowServiceTicketAlertSwalShown","true")	
	End If
	
End If


'*******************************************************************************************
' CHECK AR_CUSTOMERBILLTO TO SEE IF THERE IS AT LEAST ONE DEFAULT BILL TO LOCATION DEFINED
' IF THERE IS NOT, CREATE ONE FROM AR_CUSTOMER DATA
'******************************************************************************************

SQL = "SELECT * FROM AR_CustomerBillTo WHERE CustNum = '" & customerID & "'"

Set rs = cnn8.Execute(SQL)

If rs.EOF Then

	SQLCreateDefaultBillTo = "INSERT INTO AR_CustomerBillTo (CustNum, BillName, ContactFirstName, ContactLastName, Contact, Addr1, Addr2, City, [State], Zip, Country, Phone, DefaultBillTo)"
	SQLCreateDefaultBillTo = SQLCreateDefaultBillTo &  " VALUES (" 
	SQLCreateDefaultBillTo = SQLCreateDefaultBillTo & "'" & customerID & "','"  & CustomerName & "','" & CustomerContactFirstName & "','" & CustomerContactLastName & "', "
	SQLCreateDefaultBillTo = SQLCreateDefaultBillTo & "'" & CustomerFullName & "','" & CustomerAddr1 & "','"  & CustomerAddr2 & "','"  & CustomerCity & "','"  & CustomerState & "' ,"
	SQLCreateDefaultBillTo = SQLCreateDefaultBillTo & "'" & CustomerZip & "','" & CustomerCountry & "','" & CustomerPhone & "',1)"
		
	Set cnnCreateDefaultBillTo = Server.CreateObject("ADODB.Connection")
	cnnCreateDefaultBillTo.open (Session("ClientCnnString"))
	Set rsCreateDefaultBillTo = Server.CreateObject("ADODB.Recordset")
	rsCreateDefaultBillTo.CursorLocation = 3 
	Set rsCreateDefaultBillTo = cnnCreateDefaultBillTo.Execute(SQLCreateDefaultBillTo)
	
End If

'******************************************************************************************



'*******************************************************************************************
' CHECK AR_CUSTOMERSHIPTO TO SEE IF THERE IS AT LEAST ONE DEFAULT SHIP TO LOCATION DEFINED
' IF THERE IS NOT, CREATE ONE FROM AR_CUSTOMER DATA
'******************************************************************************************

SQL = "SELECT * FROM AR_CustomerShipTo WHERE CustNum = '" & customerID & "'"

Set rs = cnn8.Execute(SQL)

If rs.EOF Then

	SQLCreateDefaultShipTo = "INSERT INTO AR_CustomerShipTo (CustNum, ShipName, ContactFirstName, ContactLastName, Contact, Addr1, Addr2, City, [State], Zip, Country, Phone, DefaultShipTo)"
	SQLCreateDefaultShipTo = SQLCreateDefaultShipTo &  " VALUES (" 
	SQLCreateDefaultShipTo = SQLCreateDefaultShipTo & "'" & customerID & "','" & CustomerName & "','" & CustomerContactFirstName & "','" & CustomerContactLastName & "', "
	SQLCreateDefaultShipTo = SQLCreateDefaultShipTo & "'" & CustomerFullName & "','" & CustomerAddr1 & "','"  & CustomerAddr2 & "','"  & CustomerCity & "','"  & CustomerState & "',"
	SQLCreateDefaultShipTo = SQLCreateDefaultShipTo & "'" & CustomerZip & "','" & CustomerCountry & "','" & CustomerPhone & "',1)"
	
	Set cnnCreateDefaultShipTo = Server.CreateObject("ADODB.Connection")
	cnnCreateDefaultShipTo.open (Session("ClientCnnString"))
	Set rsCreateDefaultShipTo = Server.CreateObject("ADODB.Recordset")
	rsCreateDefaultShipTo.CursorLocation = 3 
	Set rsCreateDefaultShipTo = cnnCreateDefaultShipTo.Execute(SQLCreateDefaultShipTo)
	
End If



'******************************************************************************************

%>


<script type="text/javascript">

	
	$(document).ready(function() {
		
		$('a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
			var target = $(e.target).attr("href");
			$('input[name="txtTab"]').val(target);
		});
	
		$('#filter-billtos').keyup(function() {
			//alert('Handler for .keyup() called.');
		});

		$('#filter-shiptos').keyup(function() {
			//alert('Handler for .keyup() called.');
		});

		$('#filter-notes').keyup(function() {
			//alert('Handler for .keyup() called.');
		});
				       
		$('#modalEditARCustomer').on('show.bs.modal', function(e) {
		
		    //get data-id attribute of the clicked prospect
		    var CustID = $(e.relatedTarget).data('custid');	
	
	    	var $modal = $(this);
	    	
	    	$.ajax({
				type:"POST",
				url: "../../inc/InSightFuncs_AjaxForARAP.asp",
				cache: false,
				data: "action=GetCustomerAccountInformationForModal&CustID=" + encodeURIComponent(CustID),
				success: function(response)
				 {
	               	 $modal.find('#customerInfo').html(response);
				 }		
			});
			
		});
		
		$('#modalEditLastPriceChangeDate').on('show.bs.modal', function(e) {
		
		    //get data-id attribute of the clicked prospect
		    var CustID = $(e.relatedTarget).data('custid');	
		    var IntRecID = $(e.relatedTarget).data('int-rec-id');
	
	    	var $modal = $(this);
	    	
	    	$.ajax({
				type:"POST",
				url: "../../inc/InSightFuncs_AjaxForARAP.asp",
				cache: false,
				data: "action=GetCustomerPricingInformationForModal&CustID=" + encodeURIComponent(CustID) + "&IntRecID=" + encodeURIComponent(IntRecID),
				success: function(response)
				 {
	               	 $modal.find('#customerPricingInfo').html(response);
				 }		
			});
			
		});
		
		$('#modalEditCustomerNotes').on('show.bs.modal', function(e) {
	
		    //get data-id attribute of the clicked order
		    var CustID = $(e.relatedTarget).data('cust-id');
		    var CategoryID = $(e.relatedTarget).data('category-id');
		    
		    //populate the textbox with the id of the clicked order
		    $(e.currentTarget).find('input[name="txtCustIDToPassMainSearch"]').val(CustID);
		    $(e.currentTarget).find('input[name="txtCustIDToPassToGenerateNotes"]').val(CustID);
		    $(e.currentTarget).find('input[name="txtCategoryID"]').val(CategoryID);
		    	    
		    var $modal = $(this);
	
	    	$.ajax({
				type:"POST",
				url: "../../../../inc/InSightFuncs_AjaxForARAP.asp",
				cache: false,
				data: "action=GetContentForCustomerNotesModal&CustID="+encodeURIComponent(CustID),
				success: function(response)
				 {
	               	 $modal.find('#modalEditCustomerNotesContent').html(response);               	 
	             },
	             failure: function(response)
				 {
				  	$modal.find('#modalEditCustomerNotesContent').html("Failed");
		            //var height = $(window).height() - 600;
		            //$(this).find(".modal-body").css("max-height", height);
	             }
			});
		});
	
		
	
	});
	

	
	
	function ajaxRowMode(type, id, mode) {
	

		$('#ajaxRow'+type+'-'+id).attr("class", "ajaxRow"+mode);
		if(id==0){
			$('#ajaxRow'+type+'-' + 0 + '').remove();
		}	
	
		 $(".ajaxRowEdit").find("input[disabled='true']").each(function () {
		     $(this).removeAttr("disabled");
		 });
		 
		 if (mode == "View"){
			 $(".visibleRowView").find("input[type=checkbox]").each(function () {
			     $(this).attr("disabled", "disabled");
			 });
		}
		 

	   var phonesTab = [{ "mask": "(###) ###-####"}];
	    $('#txtPhoneTab' + id).inputmask({ 
	        mask: phonesTab, 
	        greedy: false, 
	        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
	    $('#txtFaxTab' + id).inputmask({ 
	        mask: phonesTab, 
	        greedy: false, 
	        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
		
	}
	
		
</script>




<!-- title / lead owner !-->
<div class="row">
	<div class="page-header">

		<div class="col-lg-3">

            <div class="business-card">
            
				<a style="position:absolute; right:20px" class="pull-right" data-toggle="modal" data-show="true" href="#" data-custid="<%= customerID %>" data-int-rec-id="<%= InternalRecordIdentifier %>" data-target="#modalEditARCustomer" data-tooltip="true" data-title="Edit Customer"><button class="btn btn-success" role="button" type="button"><i class="fas fa-lg fa-pencil-alt"></i></button></a>					
								   
                <div class="media">
                    <div class="media-left">
                        <img class="media-object img-circle profile-img" src="http://s3.amazonaws.com/37assets/svn/765-default-avatar.png">
                        <small style="margin-left:5px;">(<%= InternalRecordIdentifier %>/<%= CustomerAcctStatus %>)</small><br>
                    </div>
                    <div class="media-body">
                    	<h2 class="custid">Acct. <%= customerID %></h2>
                    	<h2 class="company"><%= CustomerName %></h2>
                        <h2 class="name"><%= CustomerContactFirstName %>&nbsp;<%= CustomerContactLastName %></h2>
                                                
                        <div class="address">
                        	<nobr>
                        		<a href="https://maps.google.com/?q=<%= CustomerFullAddress %>" target="_blank"><%= CustomerAddr1 %>
                        		<% If CustomerAddr2 <> "" then %>
                        			<br><%= CustomerAddr2 %>
                        		<% End If %>
                        		</a>
                        	</nobr>
                        </div>
                        
                        <% If CustomerCity <> "" AND CustomerState <> "" AND CustomerZip <> "" Then %>
                        	<div class="address"><a href="https://maps.google.com/?q=<%= CustomerFullAddress %>" target="_blank"><%= CustomerCity %>, <%= CustomerState %>&nbsp; <%= CustomerZip %></a></div>
                        <% End If %>
                        
                        <% If CustomerPhone <> "" Then %>
                        	<div class="phone"><i class="fa fa-phone" aria-hidden="true"></i>&nbsp;&nbsp;<%= CustomerPhone %></div>
                        <% End If %>
                                            
                         <% If CustomerFax <> "" Then %>
                        	<div class="fax"><i class="fa fa-printer" aria-hidden="true"></i>&nbsp;&nbsp;<%= CustomerFax %></div>
                        <% End If %>
                    
                    </div>
                    <div class="media-footer">
                   		&nbsp;
                    </div>
                </div>
            </div>

			<a class="btn btn-primary btn-lg btn-block" href="<%= BaseURL %>accountsreceivable/customermininfo/main.asp" role="button" style="margin-top:15px;"><i class="fa fa-arrow-left"></i> &nbsp;Back To <%= GetTerm("Customer") %> List</a>

			<% If UserHasAnyUnviewedNotes(customerID) = True Then %>
				<button type="button" class="btn btn-warning btn-lg btn-block yes-unread-notes-button" data-toggle="modal" data-target="#modalEditCustomerNotes" data-category-id="-1" data-cust-id="<%= customerID %>"><span class="fa fa-file-text-o fa-2x" aria-hidden="true"></span> <%= GetTerm("Customer") %> Notes</button>
			<% Else %>
				<button type="button" class="btn btn-warning btn-lg btn-block no-unread-notes-button" data-toggle="modal" data-target="#modalEditCustomerNotes" data-category-id="-1" data-cust-id="<%= customerID %>"><span class="fa fa-file-text-o" aria-hidden="true"></span> <%= GetTerm("Customer") %> Notes</button>
			<% End If %>
			
		</div>

		<div class="col-lg-3">

            <div class="quick-info-block quick-info-block-green">
                <h2 class="heading-md"><i class="icon-2x color-light fa fa-usd"></i>&nbsp;<%= GetTerm("Account Pricing") %>
					<a class="pull-right" data-toggle="modal" data-show="true" href="#" data-custid="<%= customerID %>" data-int-rec-id="<%= InternalRecordIdentifier %>" data-target="#modalEditLastPriceChangeDate" data-tooltip="true" data-title="Edit Customer Pricing"><button class="btn btn-success" role="button" type="button"><i class="fas fa-lg fa-pencil-alt"></i></button></a>					
                </h2>
                
                <hr class="tile">   
 
                <% If IsNull(CustomerLastPriceChangeDate) OR CustomerLastPriceChangeDate="1/1/1900" Then %>
                	<p>Last Price Change Date (Not Entered)</p> 
                <% Else %>
                	<p>Last Price Change Date <%= CustomerLastPriceChangeDate %></p>
				<% End If %>     
				        
            </div>

		</div>

		<div class="col-lg-3">
			Col Here
		</div>

		<div class="col-lg-3">
			Col Here
		</div>
		
		
	</div>
</div>
<!-- eof title / lead owner !-->

		 
<!-- tabs start here !-->
<div class="bottom-table">
	<div class="row">
		<div class="col-lg-12">
			<div class="bottom-tabs-section">

				<!-- tab navigation !-->
				<ul class="bottom-tabs nav nav-tabs" role="tablist">
					<li role='presentation'><a href='#billtos' class='tabBillToColor' aria-controls='billtos' role='tab' data-toggle='tab'>Bill To Locations (<%= NumberOfARCustAccountBillToLocationsByCustID(customerID) %>)</a></li>
					<li role='presentation'><a href='#shiptos' class='tabShipToColor' aria-controls='shiptos' role='tab' data-toggle='tab'>Ship To Locations (<%= NumberOfARCustAccountShipToLocationsByCustID(customerID) %>)</a></li>
					<li role='presentation'><a href='#contacts' class='tabContactsColor' aria-controls='contacts' role='tab' data-toggle='tab'><%= GetTerm("Contacts") %> (<%= NumberOfARCustContactsByCustID(customerID) %>)</a></li>
					<li role='presentation'><a href='#service' class='tabServiceTicketsColor' aria-controls='service' role='tab' data-toggle='tab'><%= GetTerm("Service") %> <%= GetTerm("Tickets") %> (<%= NumberOfServiceTicketsEver(customerID) %>)</a></li>
				</ul>
				<!-- eof tab navigation -->
			
				<div class="bottom-tabs tab-content">
					<!--#include file="editViewCustomerDetail_billto_tab.asp"-->
					<!--#include file="editViewCustomerDetail_shipto_tab.asp"-->
					<!--#include file="editViewCustomerDetail_contacts_tab.asp"-->
					<!--#include file="editViewCustomerDetail_service_tab.asp"-->
				</div>	
										
			</div>
		</div>
	</div>
</div>


<!-- tabs js  !-->
<script type="text/javascript">

	 $('a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
	  	e.target // newly activated tab
	  	e.relatedTarget // previous active tab
	})
	
	$(document).ready(function(){
	
		$("#demo").on("hide.bs.collapse", function(){
			$(".btn-custom").html('<span class="glyphicon glyphicon-collapse-down"></span> Click to Expand');
		});
		$("#demo").on("show.bs.collapse", function(){
			$(".btn-custom").html('<span class="glyphicon glyphicon-collapse-up"></span> Click to Collapse');
		});
	  
        $('#filter-billtolocations').keyup(function () {

           var rex = new RegExp($(this).val(), 'i');
           $('.searchable-billtos tr').hide();		   
           $('.searchable-billtos tr').filter(function () {
               return rex.test($(this).text());
            }).show();
			
        })

        $('#filter-shiptolocations').keyup(function () {

           var rex = new RegExp($(this).val(), 'i');
           $('.searchable-shiptos tr').hide();		   
           $('.searchable-shiptos tr').filter(function () {
               return rex.test($(this).text());
            }).show();
			
        })

        $('#filter-contacts').keyup(function () {

           var rex = new RegExp($(this).val(), 'i');
           $('.searchable-contacts tr').hide();		   
           $('.searchable-contacts tr').filter(function () {
               return rex.test($(this).text());
            }).show();
			
        })
        
	  
	});

</script>
<!-- eof custom table search !-->



<!-- checkboxes JS !-->
<script type="text/javascript">
    function changeState(el) {
        if (el.readOnly) el.checked=el.readOnly=false;
        else if (!el.checked) el.readOnly=el.indeterminate=true;
    }
</script>
<!-- eof checkboxes JS !-->

 
 
 
<!-- ******************************************************************************************************************************** -->
<!-- MODAL WINDOW DESIGN AND DEFINITIONS -->
<!-- ******************************************************************************************************************************** -->

	<!--#include file="editViewCustomerDetailModals.asp"-->    

	<!-- modal window contacts tab -->
	<!--#include file="onthefly_contacttitle_forcontactstab.asp"--> 
	<!-- end modal contacts tab -->   

<!-- ******************************************************************************************************************************** -->
<!-- END MODAL WINDOW DESIGN AND DEFINITIONS -->
<!-- ******************************************************************************************************************************** -->



<!--#include file="../../inc/footer-main.asp"-->

<!--#include file="../inc/header-prospecting.asp"-->
<!--#include file="../inc/InSightFuncs_Prospecting.asp"-->

<% InternalRecordIdentifier = Request.QueryString("i") 
If InternalRecordIdentifier = "" Then Response.Redirect("mainWonPool.asp")
OpenTabNum = 1 'Dfault open tab #
OpenTabNum = Request.QueryString("t") 
%>


<% 'Read edit prospect tab color settings
SQL = "SELECT * FROM Settings_Global"
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
If not rs.EOF Then
	CRMTabLogColor = rs("CRMTabLogColor")
	CRMTabProductsColor = rs("CRMTabProductsColor")
	CRMTabEquipmentColor = rs("CRMTabEquipmentColor")
	CRMTabDocumentsColor = rs("CRMTabDocumentsColor")
	CRMTabLocationColor = rs("CRMTabLocationColor")
	CRMTabContactsColor = rs("CRMTabContactsColor")
	CRMTabCompetitorsColor = rs("CRMTabCompetitorsColor")
	CRMTabOpportunityColor = rs("CRMTabOpportunityColor")
	CRMTabAuditTrailColor	 = rs("CRMTabAuditTrailColor")
	CRMTileOfferingColor = rs("CRMTileOfferingColor")
	CRMTileCompetitorColor = rs("CRMTileCompetitorColor")
	CRMTileDollarsColor = rs("CRMTileDollarsColor")
	CRMTileActivityColor = rs("CRMTileActivityColor")
	CRMTileStageColor = rs("CRMTileStageColor")
	CRMTileOwnerColor = rs("CRMTileOwnerColor")
	CRMTileCommentsColor = rs("CRMTileCommentsColor")
	CRMHideLocationTab = rs("CRMHideLocationTab")
	CRMHideProductsTab = rs("CRMHideProductsTab")
	CRMHideEquipmentTab = rs("CRMHideEquipmentTab")	
End If

SQL = "SELECT * FROM Settings_Prospecting"
Set rs = cnn8.Execute(SQL)
If not rs.EOF Then
	TabSocialMediaColor = rs("TabSocialMediaColor")
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing

If CRMTabLogColor = "" Then CRMTabLogColor = "#D8F9D1"
If IsNull(CRMTabLogColor) Then CRMTabLogColor = "#D8F9D1"

If CRMTabProductsColor = "" Then CRMTabProductsColor = "#D8F9D1"
If IsNull(CRMTabProductsColor) Then CRMTabProductsColor = "#D8F9D1"

If CRMTabEquipmentColor = "" Then CRMTabEquipmentColor = "#FFA500"
If IsNull(CRMTabEquipmentColor) Then CRMTabEquipmentColor = "#FFA500"

If CRMTabDocumentsColor = "" Then CRMTabDocumentsColor = "#F6F6F6"
If IsNull(CRMTabDocumentsColor) Then CRMTabDocumentsColor = "#F6F6F6"

If CRMTabLocationColor = "" Then CRMTabLocationColor = "#D8F9D1"
If IsNull(CRMTabLocationColor) Then CRMTabLocationColor = "#D8F9D1"

If CRMTabContactsColor = "" Then CRMTabContactsColor = "#FCB3B3"
If IsNull(CRMTabContactsColor) Then CRMTabContactsColor = "#FCB3B3"

If CRMTabCompetitorsColor = "" Then CRMTabCompetitorsColor = "#FCB3B3"
If IsNull(CRMTabCompetitorsColor) Then CRMTabCompetitorsColor = "#FCB3B3"

If CRMTabOpportunityColor = "" Then CRMTabOpportunityColor = "#FFA500"
If IsNull(CRMTabOpportunityColor) Then CRMTabOpportunityColor = "#FFA500"

If CRMTabAuditTrailColor = "" Then CRMTabAuditTrailColor = "#FFA500"
If IsNull(CRMTabAuditTrailColor) Then CRMTabAuditTrailColor = "#FFA500"

If CRMTileOfferingColor = "" Then CRMTileOfferingColor = "#3498db"
If IsNull(CRMTileOfferingColor) Then CRMTileOfferingColor = "#3498db"

If CRMTileCompetitorColor = "" Then CRMTileCompetitorColor = "#9b6bcc"
If IsNull(CRMTileCompetitorColor) Then CRMTileCompetitorColor = "#9b6bcc"

If CRMTileDollarsColor = "" Then CRMTileDollarsColor = "#2ecc71"
If IsNull(CRMTileDollarsColor) Then CRMTileDollarsColor = "#2ecc71"

If CRMTileActivityColor = "" Then CRMTileActivityColor = "#f1c40f"
If IsNull(CRMTileActivityColor) Then CRMTileActivityColor = "#f1c40f"

If CRMTileStageColor = "" Then CRMTileStageColor = "#e67e22"
If IsNull(CRMTileStageColor) Then CRMTileStageColor = "#e67e22"

If CRMTileOwnerColor = "" Then CRMTileOwnerColor = "#95a5a6"
If IsNull(CRMTileOwnerColor) Then CRMTileOwnerColor = "#95a5a6"

If CRMTileCommentsColor = "" Then CRMTileCommentsColor = "#d43f3a"
If IsNull(CRMTileCommentsColor) Then CRMTileCommentsColor = "#d43f3a"

%>

<!-- function that gets the value of the tab when it is clicked and then
updates the value of a hidden form field so when the page posts, it returns
back to the tab that was previously opened -->

<script type="text/javascript">
	$(function () {
		$('a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
		var target = $(e.target).attr("href");
		$('input[name="txtTab"]').val(target);
		});
	})
	$(document).ready(function() {
	  $('#filter-contacts').keyup(function() {
	    //alert('Handler for .keyup() called.');
	  });
	  
  
	});	
	function ajaxRowMode(type, id, mode) {
	
		$('#ajaxRow'+type+'-'+id).attr("class", "ajaxRow"+mode);
		if(id==0){
			$('#ajaxRow'+type+'-' + 0 + '').remove();
		}	
	
		 $(".ajaxRowEdit").find('input[disabled="true"]').each(function () {
		     $(this).removeAttr("disabled");
	});	 
			
}
</script>

<style>

.ajax-loading {
    position: relative;
}
.ajax-loading::after {
    background-image: url("/img/loading.gif");
    background-position: center top;
    background-repeat: no-repeat;
    content: "";
    display: block;
    height: 100%;
    min-height: 100px;
    position: absolute;
    top: 0;
    width: 100%;
}
.ajaxRowView .visibleRowEdit, .ajaxRowEdit .visibleRowView { display: none; }


.styled{
 	cursor:pointer;
}

.plus-button{
	cursor:pointer;
}

.beatpicker-clear{
	display: none;
}
.nav-tabs{
	font-size: 12px;
}

.the-tabs .nav>li>a{
	padding: 5px 10px;
	font-weight: bold;
}


.tab-content{
	margin-top:20px;
	font-size:12px;
}

   
.tab-content .split-arrows{
	 text-align:left;
	 margin-top:10px;
	 margin-bottom: 10px;
 }
 
 .tab-content .split-arrows a{
	 display:inline-block;
	 background:#f5f5f5;
	 padding:5px;
 }
 
  .tab-content .split-arrows a:hover{
	  background:#ccc;
	  text-decoration:none;
  }


.row-line{
	margin-bottom:15px;
}
 

.first-name{
	width: 54%;
	display: inline-block;
}

.first-name-mr{
	width: 30%;
	display: inline-block;
	margin-right: 2px;
}

.page-header{
	padding-top:0px;
	margin-top: 0px;
	float:left;
	width:100%;
	margin-top:20px;
}

.standard-font{
	display: block;
    padding: 3px 20px;
    font-weight: normal;
    font-size:13px;
    line-height: 25px;
    color: #333333;
    white-space: normal;
 }
 
.nav-tabs {
    font-size: 13px;
}
table th{
	font-weight:normal
}

 
.label-col{
	width:26%;
}

 .input-col{
	 width:74%;
 }
 
 .projected-monthly-spend{
	 max-width:35%;
  }

.proposal-meeting-date{
	max-width:35%;
}
 

.table-responsive{
	font-size:11px;
}

.form-control{
	width:90%;
	height:auto;
	border-radius:4px;
	border:1px solid #ccc;
	display:inline;
} 

.form-control-modal {
    width: 90%;
    height: auto;
    border-radius: 4px;
    border: 1px solid #ccc;
    display: inline;
}
.top-section-tables   .table>tbody>tr>td, .table>tbody>tr>th, .table>tfoot>tr>td, .table>tfoot>tr>th, .table>thead>tr>td, .table>thead>tr>th{
	border:0px;
	padding:1px;
   }


.bottom-tabs-section{
	border:1px solid #ccc;
	padding:10px;
	margin-top:20px;
	float:left;
	width:100%;
}

#txtWebsite{
	width:80%;
	float:left;
}

.fa-globe{
	float:left;
	margin:4px 0px 0px 5px;
}

.btn-custom{
	width:100%;
	text-align:left;
	color: #333;
    background-color: #f5f5f5;
	border:1px solid #ddd;
	outline:none;
	border-top-left-radius:5px;
	border-top-right-radius:5px;
	font-size:16px;
	padding:10px;
}

.btn-custom:hover{
	background:#ccc;
}

#demo{
	border:1px solid #ddd;
	border-bottom-left-radius:5px;
	border-bottom-right-radius:5px;
	padding:15px;
	margin-top:-1px;
}
 

 

.bottom-table table thead th{
	padding:6px;
	font-weight:bold;
	border:1px solid #ddd;
	vertical-align:top;
}

.bottom-table table>tbody>tr>td{
	padding:6px;
	font-weight:normal;
	border:1px solid #ddd;
}

.narrow-results{
	margin-bottom:15px;
}

#filter-notes{
	width:40%;
	padding:10px;
	height:34px;
}
 
#filter-documents{
	width:40%;
	padding:10px;
	height:34px;
}
 

#filter-contacts{
	width:40%;
	padding:10px;
	height:34px;
}
 
#filter-competitors{
	width:40%;
	padding:10px;
	height:34px;
}
 

 #filter-audit{
	width:40%;
	padding:10px;
	height:34px;
}
 
 
 #filter-opportunity{
	width:40%;
	padding:10px;
	height:34px;
}


.nav-tabs > li.active > a,
.nav-tabs > li.active > a:hover,
.nav-tabs > li.active > a:focus{
    color: #fff;
    /*background-color: red;*/
} 


.nav-tabs>li>a {
color: #fff;
font-size:16px;
}


.return {
color: #333;
font-size:16px;
}

.CRMTabLogColor{
	<% Response.Write("background:" & CRMTabLogColor & " !important;") %>
}
.CRMTabProductsColor{
	<% Response.Write("background:" & CRMTabProductsColor & " !important;") %>
}
.CRMTabEquipmentColor{
	<% Response.Write("background:" & CRMTabEquipmentColor & " !important;") %>
}
.CRMTabDocumentsColor{
	<% Response.Write("background:" & CRMTabDocumentsColor & " !important;") %>
}
.CRMTabLocationColor{
	<% Response.Write("background:" & CRMTabLocationColor & " !important;") %>
}
.CRMTabContactsColor{
	<% Response.Write("background:" & CRMTabContactsColor & " !important;") %>
}
.CRMTabCompetitorsColor{
	<% Response.Write("background:" & CRMTabCompetitorsColor & " !important;") %>
}
.CRMTabOpportunityColor{
	<% Response.Write("background:" & CRMTabOpportunityColor & " !important;") %>
}
.CRMTabAuditTrailColor{
	<% Response.Write("background:" & CRMTabAuditTrailColor & " !important;") %>
}


.CRMTileOfferingColor {
	<% Response.Write("background:" & CRMTileOfferingColor & " !important;") %>
}
.CRMTileCompetitorColor {
	<% Response.Write("background:" & CRMTileCompetitorColor & " !important;") %>
}
.CRMTileDollarsColor {
	<% Response.Write("background:" & CRMTileDollarsColor & " !important;") %>
}
.CRMTileActivityColor {
	<% Response.Write("background:" & CRMTileActivityColor & " !important;") %>
}
.CRMTileStageColor {
	<% Response.Write("background:" & CRMTileStageColor & " !important;") %>
}

.CRMTileOwnerColor {
	<% Response.Write("background:" & CRMTileOwnerColor & " !important;") %>
}
.CRMTileCommentsColor {
	<% Response.Write("background:" & CRMTileCommentsColor & " !important;") %>
}

.nav-tabs > li.active > a,
.nav-tabs > li.active > a:hover,
.nav-tabs > li.active > a:focus{

	color: #fff;
	font-weight:normal;
	font-size:24px;
    /*background-color: #111 !important;*/
    border-color: #2e6da4 !important;
    margin-bottom:0px;
    margin-top:0px;
     
} 

.nav-tabs > li > a,
.nav-tabs > li > a:hover,
.nav-tabs > li > a:focus{

	color: #fff;
    margin-bottom:20px;
    margin-top:0px;
     
} 

/*Business Card Css */
.business-card {
  border: 1px solid #cccccc;
  background: #f8f8f8;
  padding: 10px;
  border-radius: 7px;
  margin-bottom: 10px;
}
.profile-img {
  height: 120px;
  background: white;
}
.company {
    font-size: 25px;
    margin-top:0px;
    margin-bottom:5px;
}
.name {
  font-size: 20px;
  margin-top:0px;
  margin-bottom:0px;
  color:#337ab7;
 }

.job {
  color: #449d44;
  font-size: 14px;
  margin-bottom:8px;
}
.address {
  color: #666;
  font-size: 14px;
  margin-bottom:1px;
}

.mail {
  font-size: 15px;
  color: #666;
  margin-bottom:5px;
  margin-top:15px;
 }
 .phone{
  color: #666;
  font-size: 15px;
  margin-bottom:5px;
  margin-top:8px;

}
 .cell{
  color: #666;
  font-size: 15px;
  margin-bottom:5px;

}

 .fax{
  color: #666;
  font-size: 15px;
  margin-bottom:5px;

}

 .website{
  color: #666;
  font-size: 15px;
  margin-bottom:5px;
  margin-left:-5px;

}

i.icon-2x {
  font-size: 30px;
}

.color-light{
    color:#FFFFFF;
}

/*Colored Content Boxes
------------------------------------*/
.quick-info-block {
  padding: 3px 20px;
  text-align: center;
  margin-bottom: 20px;
  border-radius: 7px;
}

.quick-info-block p{
  color: #fff;
  font-size:16px;
}
.quick-info-block h2 {
  color: #fff;
  font-size:20px;
}

.quick-info-block h2 a:hover{
  text-decoration: none;
}

.quick-info-block-light,
.quick-info-block-default {
  background: #fafafa;
  border: solid 1px #eee; 
}

.quick-info-block-default:hover {
  box-shadow: 0 0 8px #eee;
}

.quick-info-block-light p,
.quick-info-block-light h2,
.quick-info-block-default p,
.quick-info-block-default h2 {
  color: #555;
}

.overdue
{
	color:#FF0000 !important;
}

.quick-info-block-u {
  background: #72c02c;
}
.quick-info-block-blue {
  background: #3498db;
}
.quick-info-block-red {
  background: #e74c3c;
}
.quick-info-block-sea {
  background: #1abc9c;
}
.quick-info-block-grey {
  background: #95a5a6;
}
.quick-info-block-yellow {
  background: #f1c40f;
}
.quick-info-block-orange {
  background: #e67e22;
}
.quick-info-block-green {
  background: #2ecc71;
}
.quick-info-block-purple {
  background: #9b6bcc;
}
.quick-info-block-aqua {
  background: #27d7e7;
}
.quick-info-block-brown {
  background: #9c8061;
}
.quick-info-block-dark-blue {
  background: #4765a0;
}
.quick-info-block-light-green {
  background: #79d5b3;
}
.quick-info-block-dark {
  background: #555;
}
.quick-info-block-light {
  background: #ecf0f1;
}

.note {
color: #4cae4c;
}
.activity {
color: #888;
}
.stagechange {
color: #e67e22;
}
.email {
color: #3d85c6;
}

.email a.address{
	color:#168bf4;
}

.email a.address:hover{
	color:#5cb85c;
}

.fileicon {
width:40%;
}

hr.tile {
    border: 0;
    height: 3px;
    background-image: linear-gradient(to right, rgba(0, 0, 0, 0), rgba(255, 255, 255, 0.95), rgba(0, 0, 0, 0));
}

</style>
<!-- eof css !-->

<%

SQL = "SELECT * FROM PR_Prospects where InternalRecordIdentifier = " & InternalRecordIdentifier 

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)

If not rs.EOF Then
	InternalRecordIdentifier = rs("InternalRecordIdentifier")
	Company = rs("Company")
	Street= rs("Street")
	City= rs("City")
	State= rs("State")
	PostalCode = rs("PostalCode")
	Country= rs("Country")
	Suite= rs("Floor_Suite_Room__c")								
	Website= rs("Website")								
	LeadSourceNumber = rs("LeadSourceNumber")
	LeadSource = GetLeadSourceByNum(LeadSourceNumber)				
	StageNumber = GetProspectCurrentStageByProspectNumber(InternalRecordIdentifier)
	IndustryNumber = rs("IndustryNumber")	
	Industry = GetIndustryByNum(IndustryNumber)											
	OwnerUserNo = rs("OwnerUserNo")				
	CreatedDate= rs("CreatedDate")
	CreatedByUserNo= rs("CreatedByUserNo")				
	TelemarketerUserNo = rs("TelemarketerUserNo")
	Telemarketer = GetUserDisplayNameByUserNo(TelemarketerUserNo)
	ProjectedGPSpend= rs("ProjectedGPSpend")
	NumberOfPantries = rs("NumberOfPantries")
	EmployeeRangeNumber = rs("EmployeeRangeNumber")
	NumEmployees = GetEmployeeRangeByNum(EmployeeRangeNumber)
	CreatedDate = rs("CreatedDate")
	FormerCustNum = rs("FormerCustNum")
	CancelDate = rs("CancelDate")
	LeaseExpirationDate = rs("LeaseExpirationDate")	
	Comments = rs("Comments")
	CurrentOffering = rs("CurrentOffering")			
End If
%>

<!-- title / lead owner !-->
<div class="row">
	<div class="page-header">

		<div class="col-lg-3">
			<%
		
			SQLContacts1 = "SELECT * FROM PR_ProspectContacts WHERE ProspectIntRecID = " & InternalRecordIdentifier & " AND PrimaryContact = 1"
			
			Set cnnContacts1 = Server.CreateObject("ADODB.Connection")
			cnnContacts1.open (Session("ClientCnnString"))
			Set rsContacts1 = Server.CreateObject("ADODB.Recordset")
			rsContacts1.CursorLocation = 3 
			Set rsContacts1 = cnnContacts1.Execute(SQLContacts1)
			
			If not rsContacts1.EOF Then
			
			  	primarySuffix = rsContacts1("Suffix")
			  	primaryFirstName = rsContacts1("FirstName")
				primaryLastName = rsContacts1("LastName")	
				primaryTitleNumber = rsContacts1("ContactTitleNumber")
				primaryTitle = GetContactTitleByNum(primaryTitleNumber)
				primaryEmail = rsContacts1("Email") 
				primaryPhone = rsContacts1("Phone")
				primaryPhoneExt = rsContacts1("PhoneExt")
				primaryCell = rsContacts1("Cell")
				primaryFax = rsContacts1("Fax")

								
			End If
			Set rsContacts1 = Nothing
			cnnContacts1.Close
			Set cnnContacts1 = Nothing
				
			%>
            <div class="business-card">
                <div class="media">
                	
                    <div class="media-left">
                        <img class="media-object img-circle profile-img" src="http://s3.amazonaws.com/37assets/svn/765-default-avatar.png">
                        <small style="margin-left:5px;">(<%=InternalRecordIdentifier%>)</small><br>
                    </div>
                    <div class="media-body">
                    	<h2 class="company"><%= Company %></h2>
                        <h2 class="name"><%= primarySuffix %>&nbsp;<%= primaryFirstName %>&nbsp;<%= primaryLastName %></h2>
                        
                        <% If primaryTitle <> "0" Then %>
                        	<div class="job"><%= primaryTitle %></div>
                        <% End If %>
                        
                        <div class="address"><%= Street %></div>
                        <div class="address"><%= Suite %></div>
                        
                        <% If State <> "" AND City <> "" AND PostalCode <> "" Then %>
                        	<div class="address"><%= City %>, <%= State %>&nbsp;<%= PostalCode %></div>
                        <% End If %>
                                                
                        <% If primaryPhone <> "" Then %>
                        	<div class="phone"><i class="fa fa-phone" aria-hidden="true"></i>&nbsp;&nbsp;<%= primaryPhone %></div>
                        <% End If %>

                        <% If primaryPhoneExt <> "" Then %>
                        	<div class="phoneext"><i class="fa fa-phone" aria-hidden="true"></i>&nbsp;&nbsp;<%= primaryPhoneExt %></div>
                        <% End If %>
                         
                        <% If primaryCell <> "" Then %>
                        	<div class="cell"><i class="fa fa-mobile fa-lg" aria-hidden="true"></i>&nbsp;&nbsp;&nbsp;<%= primaryCell %> (cell)</div>
                        <% End If %>
                        
                        <% If primaryFax <> "" Then %>
                        	<div class="fax"><i class="fa fa-fax" aria-hidden="true"></i>&nbsp;&nbsp;<%= primaryFax %> (fax)</div>
                        <% End If %>                        
                        
                        <% If primaryEmail <> "" Then %>
                        	<div class="mail"><i class="fa fa-envelope" aria-hidden="true"></i>&nbsp;&nbsp;<a href="mailto:<%= primaryEmail %>"><%= primaryEmail %></a></div>
                        <% End If %>
                     
                        <% If Industry <> "" Then %>
                        	<div class="address"><%= Industry %></div>
                        <% End If %>
                        
                        
                        <% If Website <> "" Then %>
                        	<div class="website"><i class="fa fa-globe" aria-hidden="true"></i>&nbsp;&nbsp;<a href="http://<%= Website %>" target="_blank"><%= Website %></a></div>
                        <% End If %>

                    </div>
                </div>
            </div>
             <div class="quick-info-block CRMTileOwnerColor">
                <h2 class="heading-md"><i class="icon-2x color-light fa fa-user"></i>&nbsp;<%= GetTerm("Owner") %></h2>
                <hr class="tile">
                <p><% If OwnerUserNo <> 0 Then Response.Write(GetUserDisplayNameByUserNo(OwnerUserNo)) %></p>                        
            </div>

			<a class="btn btn-primary btn-lg btn-block" href="mainWonPool.asp" role="button" style="margin-top:15px;"><i class="fa fa-arrow-left"></i> &nbsp;Back To <%= GetTerm("New Customer Pool") %></a>

		</div>

		<div class="col-lg-3">
             <div class="quick-info-block CRMTileCommentsColor">
                <h2 class="heading-md"><i class="icon-2x color-light fa fa-comment"></i>&nbsp;<%= GetTerm("Comments") %></h2>
                <hr class="tile">
                <p><%= Comments %></p>                        
            </div>

            <div class="quick-info-block CRMTileDollarsColor">
                <h2 class="heading-md"><i class="icon-2x color-light fa fa-usd"></i>&nbsp;<%= GetTerm("Opportunity") %></h2>
                <hr class="tile">
                <p>Projected GP Spend <%= FormatCurrency(ProjectedGPSpend,2) %></p>     
                <p># Employees <%= NumEmployees %></p>
                <p># Pantries <%= NumberOfPantries %></p>    
                <p>Lease Expiration Date <%= LeaseExpirationDate %></p>             
            </div>

		</div>
            

		<div class="col-lg-3">

            <div class="quick-info-block CRMTileOfferingColor">
                <h2 class="heading-md"><i class="icon-2x color-light fa fa-clock-o"></i>&nbsp;<%= GetTerm("Current Supplier Info") %></h2>
                <hr class="tile">
                <p><%= CurrentOffering %></p>                        
            </div>

				<%
					PrimaryCompetitorID = GetPrimaryCompetitorIDByProspectNumber(InternalRecordIdentifier)
					
					If PrimaryCompetitorID <> "" Then
						PrimaryCompetitorName = GetCompetitorByNum(PrimaryCompetitorID)
					
						SQLCompetitors1 = "SELECT * FROM PR_ProspectCompetitors WHERE CompetitorRecID = " & PrimaryCompetitorID & " AND ProspectRecID = " &  InternalRecordIdentifier
						
						Set cnnCompetitors1 = Server.CreateObject("ADODB.Connection")
						cnnCompetitors1.open (Session("ClientCnnString"))
						Set rsCompetitors1 = Server.CreateObject("ADODB.Recordset")
						rsCompetitors1.CursorLocation = 3 
						Set rsCompetitors1 = cnnCompetitors1.Execute(SQLCompetitors1)
						
						If not rsCompetitors1.EOF Then
						
							BottledWater = rsCompetitors1 ("BottledWater")
							FilteredWater = rsCompetitors1 ("FilteredWater")
							OCS = rsCompetitors1 ("OCS")
							OCS_Supply = rsCompetitors1 ("OCS_Supply")
							OfficeSupplies = rsCompetitors1 ("OfficeSupplies")
							Vending = rsCompetitors1 ("Vending")
							Micromarket = rsCompetitors1 ("Micromarket")
							Pantry = rsCompetitors1 ("Pantry")
											
						End If
						Set rsCompetitors1 = Nothing
						cnnCompetitors1.Close
						Set cnnCompetitors1 = Nothing
						
						
						If BottledWater = vbTrue Then BottledWater = "Bottled Water" Else BottledWater = ""
						If FilteredWater = vbTrue Then FilteredWater = "Filtered Water" Else FilteredWater = ""
						If OCS = vbTrue Then OCS = "OCS" Else OCS = ""
						If OCS_Supply = vbTrue Then OCS_Supply = "OCS Supply" Else OCS_Supply = ""
						If OfficeSupplies = vbTrue Then OfficeSupplies = "Office Supplies " Else OfficeSupplies = ""
						If Vending = vbTrue Then Vending = "Vending" Else Vending = ""
						If Micromarket = vbTrue Then Micromarket = "Micromarkets" Else Micromarket = ""
						If Pantry = vbTrue Then Pantry = "Pantry" Else Pantry = ""
					Else

						PrimaryCompetitorName = "None Selected"
						BottledWater = ""
						FilteredWater = ""
						OCS = ""
						OCS_Supply = ""
						OfficeSupplies = ""
						Vending = ""
						Micromarket = ""
						Pantry = ""
					
					End If
						
					%>
			
		            <div class="quick-info-block CRMTileCompetitorColor">
		                <h2 class="heading-md"><i class="icon-2x color-light fa fa-user-circle-o"></i>&nbsp;<%= GetTerm("Primary Competitor") %></h2>
		                <hr class="tile">
		                <p><%= PrimaryCompetitorName %></p>
		                <p>
		                	<% If BottledWater <> "" Then Response.Write(BottledWater) %>
		                	
		                	<% If BottledWater <> "" Then %>
		                		<% If FilteredWater <> "" Then Response.Write(", " & FilteredWater) %>
		                	<% Else %>
		                		<% If FilteredWater <> "" Then Response.Write(FilteredWater) %>
							<% End If %>
							
							<% If FilteredWater <> "" Then %>
		                		<% If OCS <> "" Then Response.Write(", " & OCS) %>
		                	<% Else %>
		                		<% If OCS <> "" Then Response.Write(OCS) %>
		                	<% End If %>
		                	
		                	<% If OCS <> "" Then %>
		                		<% If OCS_Supply <> "" Then Response.Write(", " & OCS_Supply) %>
		                	<% Else %>	
		                		<% If OCS_Supply <> "" Then Response.Write(OCS_Supply) %>
		                	<% End If %>	
		                	
		                	<% If OCS_Supply <> "" Then %>
		                		<% If OfficeSupplies <> "" Then Response.Write(", " & OfficeSupplies) %>
		                	<% Else %>
		                		<% If OfficeSupplies <> "" Then Response.Write(OfficeSupplies) %>
		                	<% End If %>
		                	
		                	<% If OfficeSupplies <> "" Then %>
		                		<% If Vending <> "" Then Response.Write(", " & Vending) %>
		                	<% Else %>
		                		<% If Vending <> "" Then Response.Write(Vending) %>
		                	<%  End If %>
		                	
		                	<% If Vending <> "" Then %>
		                		<% If Micromarket <> "" Then Response.Write(", " & Micromarket) %>
		                	<% Else %>	
		                		<% If Micromarket <> "" Then Response.Write(Micromarket) %>
							<%  End If %>
							
							<% If Micromarket <> "" Then %>
		                		<% If Pantry <> "" Then Response.Write(", " & Pantry) %>
		                	<% Else %>
		                		<% If Pantry <> "" Then Response.Write(Pantry) %>
		                	<% End If %>
		                </p>
		                <p>Source (<%= LeadSource %>)</p>  
		                <hr class="tile">
		                <p>Former Customer #: <%= FormerCustNum %></p>
		                <p>Cancel Date: <%= CancelDate %></p>                      
		            </div>
		</div>

		<div class="col-lg-3">
	
			<%
		
			SQLContacts1 = "SELECT * FROM PR_ProspectActivities where ProspectRecID = " & InternalRecordIdentifier & " AND Status IS NULL"
			
			Set cnnContacts1 = Server.CreateObject("ADODB.Connection")
			cnnContacts1.open (Session("ClientCnnString"))
			Set rsContacts1 = Server.CreateObject("ADODB.Recordset")
			rsContacts1.CursorLocation = 3 
			Set rsContacts1 = cnnContacts1.Execute(SQLContacts1)
			
			If not rsContacts1.EOF Then
				ActivityRecID = rsContacts1("ActivityRecID")
			  	nextActivity = GetActivityByNum(rsContacts1("ActivityRecID"))
				nextActivityDueDate = FormatDateTime(rsContacts1("ActivityDueDate"),2) & " " & FormatDateTime(rsContacts1("ActivityDueDate"),3)
				daysOld = DateDiff("d",rsContacts1("RecordCreationDateTime"),Now())
				daysOverdue = DateDiff("d",rsContacts1("ActivityDueDate"),Now())	
							
			End If
			Set rsContacts1 = Nothing
			cnnContacts1.Close
			Set cnnContacts1 = Nothing
				
			%>
			
            <div class="quick-info-block CRMTileActivityColor">
                <h2 class="heading-md"><i class="icon-2x color-light fa fa-arrow-right"></i>&nbsp;<%= GetTerm("Next Activity") %></h2> 
                <hr class="tile">
                <p><%= GetTerm("Prospects") %> do not have a next activity in the <%= GetTerm("New Customer Pool") %>.</p>
            </div>
		
			
            <div class="quick-info-block CRMTileStageColor">
                <h2 class="heading-md"><i class="icon-2x color-light fa fa-tasks"></i>&nbsp;<%= GetTerm("Stage") %></h2> 
                <hr class="tile">
                <p><%= GetStageByNum(StageNumber) %></p>
                <hr class="tile">
                <p><i class="fa fa-pencil-square-o" aria-hidden="true"></i>&nbsp;Last Change Date:&nbsp;&nbsp;<%= GetProspectLastStageChangeDateByProspectNumber(InternalRecordIdentifier) %></p>
                <p><i class="fa fa-calendar-o" aria-hidden="true"></i>&nbsp;Days Since Qualified:&nbsp;&nbsp;XYZ</p>
                <div class="progressbarsone" progress="<%= GetPercentForStage(StageNumber)%>%"></div>       
            </div>
            

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
				<ul class="nav nav-tabs" role="tablist">
					<li role='presentation' class="active"><a href='#log' class='CRMTabLogColor' aria-controls='notes' role='tab' data-toggle='tab'><%= GetTerm("Journal") %> (<%=NumberOfLogItemsByProspectNumber(InternalRecordIdentifier)%>)</a></li>
					<% If CRMHideProductsTab = 0 Then %>
						<li role='presentation'><a href='#products' class='CRMTabProductsColor' aria-controls='documents' role='tab' data-toggle='tab'><%= GetTerm("Products") %></a></li>
					<% End If %>
					<% If CRMHideEquipmentTab = 0 Then %>
						<li role='presentation'><a href='#equipment' class='CRMTabEquipmentColor' aria-controls='documents' role='tab' data-toggle='tab'><%= GetTerm("Equipment") %></a></li>
					<% End If %>
					<li role='presentation'><a href='#documents' class='CRMTabDocumentsColor' aria-controls='documents' role='tab' data-toggle='tab'><%= GetTerm("Documents") %> (<%=NumberOfDocumentsByProspectNumber(InternalRecordIdentifier)%>)</a></li>  
					<li role='presentation'><a href='#contacts' class='CRMTabContactsColor' aria-controls='contacts' role='tab' data-toggle='tab'><%= GetTerm("Contacts") %> (<%=NumberOfContactsByProspectNumber(InternalRecordIdentifier)%>)</a></li>
					<li role='presentation'><a href='#competitors' class='CRMTabCompetitorsColor' aria-controls='general' role='tab' data-toggle='tab'><%= GetTerm("Competitors") %> (<%=NumberOfCompetitorsByProspectNumber(InternalRecordIdentifier)%>)</a></li>
					<% If CRMHideLocationTab = 0 Then %>
						<li role='presentation'><a href='#location' class='CRMTabLocationColor' aria-controls='general' role='tab' data-toggle='tab'><%= GetTerm("Location") %></a></li>
					<% End If %>
					<li role='presentation'><a href='#audit' class='CRMTabAuditTrailColor' aria-controls='audit' role='tab' data-toggle='tab'><%= GetTerm("Audit Trail") %></a></li>
					<!--<li role='presentation'><a href="mainWonPool.asp" class="btn btn-secondary active" role="button" aria-pressed="true"><span class="return">Back To <%= GetTerm("Prospecting") %> Main</span></a></li>-->
				</ul>
				<!-- eof tab navigation -->
				
				<div class="tab-content">
					<!--#include file="viewProspectReadOnly_log_tab.asp"-->
					<% If CRMHideProductsTab = 0 Then %>
						<!--#include file="viewProspectReadOnly_products_tab.asp"-->
					<% End If %>
					<% If CRMHideEquipmentTab = 0 Then %>
						<!--#include file="viewProspectReadOnly_equipment_tab.asp"-->
					<% End If %>
					<!--#include file="viewProspectReadOnly_documents_tab.asp"-->
					<!--#include file="viewProspectReadOnly_contacts_tab.asp"-->
					<!--#include file="viewProspectReadOnly_competitors_tab.asp"-->
					<% If CRMHideLocationTab = 0 Then %>
						<!--#include file="viewProspectReadOnly_location_tab.asp"-->
					<% End If %>
					<!--#include file="viewProspectReadOnly_audit_tab.asp"-->
				</div>
			</div>
		</div>
	</div>
</div>

<%
set rs = Nothing
cnn8.close
set cnn8 = Nothing
%>
 
 <!-- tabs js  !-->
 <script type="text/javascript">
 $('a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
  e.target // newly activated tab
  e.relatedTarget // previous active tab
})
   </script>

   <script>
$(document).ready(function(){
  $("#demo").on("hide.bs.collapse", function(){
    $(".btn-custom").html('<span class="glyphicon glyphicon-collapse-down"></span> Click to Expand');
  });
  $("#demo").on("show.bs.collapse", function(){
    $(".btn-custom").html('<span class="glyphicon glyphicon-collapse-up"></span> Click to Collapse');
  });
});
</script>
 <!-- eof tabs js !-->

 <!-- custom table search !-->

<script>

$(document).ready(function () {

    (function ($) {

        $('#filter-audit').keyup(function () {

            var rex = new RegExp($(this).val(), 'i');
            $('.searchable-audit tr').hide();
            $('.searchable-audit tr').filter(function () {
                return rex.test($(this).text());
            }).show();
        })
        
        $('#filter-notes').keyup(function () {

            var rex = new RegExp($(this).val(), 'i');
            $('.searchable-notes tr').hide();
            $('.searchable-notes tr').filter(function () {
                return rex.test($(this).text());
            }).show();
        })
        
        $('#filter-documents').keyup(function () {

            var rex = new RegExp($(this).val(), 'i');
            $('.searchable-documents tr').hide();
            $('.searchable-documents tr').filter(function () {
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
 
        $('#filter-competitors').keyup(function () {

           var rex = new RegExp($(this).val(), 'i');
           $('.searchable-competitors tr').hide();
           $('.searchable-competitors tr').filter(function () {
               return rex.test($(this).text());
            }).show();
        })
        
       
        $('#filter-opportunity').keyup(function () {

            var rex = new RegExp($(this).val(), 'i');
            $('.searchable-opportunity tr').hide();
            $('.searchable-opportunity tr').filter(function () {
                return rex.test($(this).text());
            }).show();
        })
        

    }(jQuery));

});
</script>
<!-- eof custom table search !-->

<!-- progress bar !-->
<link rel="stylesheet" href="<%= BaseURL %>js/jprogress/jprogress.css">
<script src="<%= BaseURL %>js/jprogress/jprogress.js" type="text/javascript"></script>

<script>
    // activate jprogress
    $(".progressbars").jprogress();
    $(".progressbarsone").jprogress({
        background: "url(../js/jprogress/progress_bar_tiles.png)"
     });
</script>
<!-- eof progress bar !-->


<!-- checkboxes JS !-->
<script type="text/javascript">
    function changeState(el) {
        if (el.readOnly) el.checked=el.readOnly=false;
        else if (!el.checked) el.readOnly=el.indeterminate=true;
    }
</script>
<!-- eof checkboxes JS !-->




<!--#include file="../inc/footer-main.asp"-->
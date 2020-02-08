<% 
'Read edit prospect tab color settings
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

If TabSocialMediaColor = "" Then TabSocialMediaColor = "#fa92e3"
If IsNull(TabSocialMediaColor) Then TabSocialMediaColor = "#fa92e3"

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
    height: 35px;
    padding-left:10px;
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

/*
.modal-content{
	max-height:650px;
	overflow-y:auto;
	width:750px;
}

 .modal-content .row{
	 padding-bottom:20px;
 }

 .modal-content p{
	 margin-bottom:20px;
	 white-space:normal;
 }

*/
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
.CRMTabSocialMediaColor{
	<% Response.Write("background:" & TabSocialMediaColor & " !important;") %>
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

.nextActivityDate {
	font-size:26px;
	font-weight:400;
}

.nextActivityTime {
	font-size:12px;
	font-weight:normal;	
}

	.red-line{
		border-left:3px solid red;
	}   


</style>

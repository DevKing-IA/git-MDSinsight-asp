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
.bottom-tabs .nav-tabs{
	font-size: 12px;
}

.bottom-tabs .nav-tabs .nav>li>a{
	padding: 5px 10px;
	font-weight: bold;
}


.bottom-tabs .tab-content{
	margin-top:20px;
	font-size:12px;
}

   
.bottom-tabs .tab-content .split-arrows{
	 text-align:left;
	 margin-top:10px;
	 margin-bottom: 10px;
 }
 
.bottom-tabs .tab-content .split-arrows a{
	 display:inline-block;
	 background:#f5f5f5;
	 padding:5px;
 }
 
.bottom-tabs .tab-content .split-arrows a:hover{
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

.bottom-tabs .standard-font{
	display: block;
    padding: 3px 20px;
    font-weight: normal;
    font-size:13px;
    line-height: 25px;
    color: #333333;
    white-space: normal;
 }

.standard-font{
    font-weight: normal;
    font-size:13px;
    line-height: 25px;
    color: #333333;
    white-space: normal;
 }

.bottom-tabs .nav-tabs {
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

table thead th{
	padding:6px;
	font-weight:bold;
	border:1px solid #ddd;
	vertical-align:top;
}

table>tbody>tr>td{
	padding:6px;
	font-weight:normal;
	border:1px solid #ddd;
}

.narrow-results{
	margin-bottom:15px;
	margin-top:15px;
}


#filter-billtolocations{
	width:40%;
	padding:10px;
	height:34px;
}

#filter-shiptolocations{
	width:40%;
	padding:10px;
	height:34px;
}


#filter-contacts{
	width:40%;
	padding:10px;
	height:34px;
}


.bottom-tabs .nav-tabs > li.active > a,
.bottom-tabs .nav-tabs > li.active > a:hover,
.bottom-tabs .nav-tabs > li.active > a:focus{
    color: #fff;
    /*background-color: red;*/
} 


.bottom-tabs .nav-tabs>li>a {
color: #fff;
font-size:16px;
}


.return {
color: #333;
font-size:16px;
}

.tabBillToColor {
	background: #4cae4c !important;
    margin-right: 2px!important;
    line-height: 1.42857143!important;
    border: 1px solid transparent!important;
    border-radius: 4px 4px 0 0!important;
    font-size: 16px !important;
    color: #FFF!important;
}

.tabShipToColor {
	background: #555 !important;
    margin-right: 2px!important;
    line-height: 1.42857143!important;
    border: 1px solid transparent!important;
    border-radius: 4px 4px 0 0!important;
    font-size: 16px !important;
    color: #FFF!important;	
}


.tabContactsColor {
	background: #007bff !important;
    margin-right: 2px!important;
    line-height: 1.42857143!important;
    border: 1px solid transparent!important;
    border-radius: 4px 4px 0 0!important;
    font-size: 16px !important;
    color: #FFF!important;	
}

.tabServiceTicketsColor {
	background: #6A1B9A !important;
    margin-right: 2px!important;
    line-height: 1.42857143!important;
    border: 1px solid transparent!important;
    border-radius: 4px 4px 0 0!important;
    font-size: 16px !important;
    color: #FFF!important;		
}

.bottom-tabs .nav-tabs > li.active > a,
.bottom-tabs .nav-tabs > li.active > a:hover,
.bottom-tabs .nav-tabs > li.active > a:focus{

	color: #fff;
	font-weight:normal;
	font-size:24px;
    /*background-color: #111 !important;*/
    /*border-color: #2e6da4 !important;*/
    margin-bottom:0px;
    margin-top:0px;
     
} 

.bottom-tabs .nav-tabs > li > a,
.bottom-tabs .nav-tabs > li > a:hover,
.bottom-tabs .nav-tabs > li > a:focus{

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
.custid {
    font-size: 22px;
    margin-top:0px;
    margin-bottom:5px;
    color:#4cae4c;
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
/*color: #4cae4c;*/
color: #000;
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

	.yes-unread-notes-button {
	  text-align: center;
	  color: white;
	  border: none;
	  font-size: 18px;
	  background: #ffc12c;
	  cursor: pointer;
	  box-shadow: 0 0 0 0 rgba(#ffc12c, .5);
	  -webkit-animation: pulse 2.5s infinite;
	}
	.yes-unread-notes-button:hover {
	  -webkit-animation: none;
	}
	
	/*@-webkit-keyframes pulse {
	  0% {
	    @include transform(scale(.9));
	  }
	  70% {
	    @include transform(scale(1));
	    box-shadow: 0 0 0 25px rgba(#5a99d4, 0);
	  }
	    100% {
	    @include transform(scale(.9));
	    box-shadow: 0 0 0 0 rgba(#5a99d4, 0);
	  }
	}	*/

	@keyframes pulse{
	    0% { transform: scale(1); }
	    30% { transform: scale(1); }
	    40% { transform: scale(1.08); }
	    50% { transform: scale(1); }
	    60% { transform: scale(1); }
	    70% { transform: scale(1.05); }
	    80% { transform: scale(1); }
	    100% { transform: scale(1); }
	}	
	.no-unread-notes-button {
	  text-align: center;
	  color: white;
	  border: none;
	  font-size:1.1em;
	  background: #ffc12c;
	  cursor: pointer;
	  box-shadow: 0 0 0 0 rgba(#ffc12c, .5);
	  -webkit-animation: none;
	}
	.no-unread-notes-button:hover {
	  -webkit-animation: none;
	}

	.modal.modal-wide .modal-dialog {
	  width: 75%;
	}
	.modal-wide .modal-body {
	  overflow-y: auto;
	}
	
	.modal.modal-xwide .modal-dialog {
	  width: 70%;
	}
	.modal-xwide .modal-body {
	  overflow-y: auto;
	  max-height:600px;
	}

</style>

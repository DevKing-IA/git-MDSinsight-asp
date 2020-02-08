<!--#include file="../inc/header.asp"-->
<!--#include file="../inc/insightfuncs.ASP"-->
<!--#include file="../inc/insightfuncs_service.ASP"-->
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
ClientKeyForFileNames = MUV_READ("ClientKeyForFileNames")
%>

<script type="text/javascript">

	$(function () {
		var autocompleteJSONFileURL = "../clientfiles/<%= ClientKeyForFileNames %>/autocomplete/customer_account_list_CSZ_<%= ClientKeyForFileNames %>.json";
		
		var options = {
		  url: autocompleteJSONFileURL,
		  placeholder: "Search for a customer by name, account, city, state, zip",
		  getValue: "name",
		  list: {	
	        onChooseEvent: function() {
	            var custID = $("#txtCustID").getSelectedItemData().code;
	            $("#txtCustIDToPass").val(custID);
			
				 if (custID!=""){
				 	$.ajax({
						type:"POST",
						url: "../inc/InSightAjaxFuncs.asp",
						data: "action=selectAccount_AccountNotes&custID="+encodeURIComponent(custID),
							success: function(msg){
								window.location = "main.asp";
							}
					}) 
				
				  }

        	},		  
		    match: {
		      enabled: true
			},
			maxNumberOfElements: 15		
		  },
		  theme: "round"
		};
		$("#txtCustID").easyAutocomplete(options);

	})
</script>


<script>
  function myFunction(num)
	  {   

		  var  notnum=num;
				
		   if(num!='')
		   {
		    $.ajax({
		   type:'post',
		      url:'toggleSticky.asp',
		          data:{notnum: notnum},
					success: function(msg){
						window.location = "main.asp";
					}
		 });
		  }
	}
</script>

<script>

  function ManualreturnDateFunc(cnum)
	  {   

		  var  cstnum=cnum;
		  var rtdte = document.getElementById("txtCloseCancelDate").value;
		  				
		   if(cnum!='')
		   {
		    $.ajax({
		   type:'post',
		      url:'setManualReturnDate.asp',
		          data:{cstnum: cstnum,rtdte: rtdte},
					success: function(msg){
						window.location = "main.asp";
					}
		 });
		  }
	}
</script>

<script>

  function AutoreturnDateFunc(cnum)
	  {   

		  var  cstnum=cnum;
		  var rsn = document.getElementById("selReturnDateReason").value;
		  				
		   if(cnum!='')
		   {
		    $.ajax({
		   type:'post',
		      url:'setAutoReturnDate.asp',
		          data:{cstnum: cstnum,rsn: rsn},
					success: function(msg){
						window.location = "main.asp";
					}
		 });
		  }
	}
</script>

<script>
  function expirDateFunc(num)
	  {   

		  var  notnum=num;
  		  var ExpirDate= document.getElementById("txtExpirDate").value;

				
		   if(num!='')
		   {
		    $.ajax({
		   type:'post',
		      url:'SetExpirDate.asp',
		          data:{notnum: notnum,ExpirDate: ExpirDate},
		          	success: function(msg){
						window.location = "main.asp";
					}
		 });
		  }
	}
</script>


  <!-- date picker !-->
  <link rel="stylesheet" href="<%= baseURL %>css/datepicker/BeatPicker.min.css"/>
 <script src="<%= baseURL %>js/datepicker/BeatPicker.min.js"></script>
   <!-- eof date picker !-->

  
  <style type="text/css">
	  .alert{
 		padding: 6px 12px;
 		font-weight: normal;
 		font-size: 13px;
	}
	
	  .row-line{
		  margin-bottom: 25px;
	  }
	  
	  .status-date{
		  text-align: right;
		  font-size: 13px;
	  }
	  

.table-responsive{
	font-size: 12px;
}

.table-responsive .arrows{
	font-size: 16px;
	margin-right: 3px;
	color: green;
}

.table-responsive .delete{
	font-size: 16px;
	margin-right: 3px;
	color: red;
}

 .notes-col{
	width: 50%;
	 
}

.date-col{
	width: 10%;
}	

.createdby-col{
	width: 10%;
}

.sticky-col{
	width: 4%;
}

.sticky-table-line{
	background: #fcfcf0;
}

.table-line-new-for-user{
	background: #fff200;
}

 .table-info{
 	 border: 1px solid #ccc;
 	 margin-left: 10px;
 	 padding: 10px;
 }

.table-info .table>tbody>tr>td, .table>tbody>tr>th, .table>tfoot>tr>td, .table>tfoot>tr>th, .table>thead>tr>td, .table>thead>tr>th{
	border: 0px;
	font-weight: bold;
	line-height: 0.8;
}
 

    .input-parent{
	    display: block;
	    float: left;
	    margin-right: 5px;
    }
    
    .last-run-inputs{
	    max-width: 110px;
    }
    
    .setbtn{
	    font-size: 12px;
	    padding: 5px 10px 5px 10px;
    }
    
table.sortable th:not(.sorttable_sorted):not(.sorttable_sorted_reverse):not(.sorttable_nosort):after { 
    content: " \25B4\25BE" 
}

table{
	font-size: 12px;
}

.beatpicker-clear{
	display: none;
}

.tr-even{
	background: #f6f6f6;
}

.tr-odd{
	background: #fff;
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
	border: 1px solid #eaeaea;
	padding-top: 10px;
	margin-bottom: 15px;
}

.high-priority{
	background:#fad5d5;
}

.alert-priority{
	background:#faf99d;
}

.alert-high-priority{
	background:#fa9090;
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
</style>

  
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
  
  
<!-- eof select and auto complete !-->
<%
Response.Write("<div id=""PleaseWaitPanel"">")
Response.Write("<br><br>Processing " & GetTerm("Customer") & " Center, please wait...<br><br>")
Response.Write("<img src=""../img/loading.gif"" />")
Response.Write("</div>")
Response.Flush()



'See if we will be using popup message for service tickets
ShowPopup = 0
SQL = "SELECT * FROM " & MUV_Read("SQL_Owner") & ".Settings_Global"
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
If not rs.eof then ShowPopup = rs("NotesScreenShowPopup")
set rs = nothing
cnn8.close
set cnn8=nothing
If ShowPopup = 0 Then ShowPopup = False
If ShowPopup = 1 Then ShowPopup = True
%>

<% SelectedCustomer = Session("ServiceCustID") %>

<h1 class="page-header"><i class="fa fa-file-text-o"></i> <%=GetTerm("Account")%> Center</h1>

<!-- row !-->
<div class="row row-line">
    <div class="col-lg-8">
    		<div class="row">
	        		
	        		<!-- select company !-->
				    <div class="col-lg-4 col-md-3 col-sm-12 col-xs-12">
				    
				    	<input id="txtCustID" name="txtCustID">
				    	<input type="hidden" id="txtCustIDToPass" name="txtCustIDToPass">
						<i id="searchIcon" class="fa fa-search fa-2x"></i>
				
					</div>
					<!-- eof select company !-->
		</div><!-- eof row -->
	</div><!-- eof col -->
</div><!-- eof row row-line -->


<!-- row !-->
<div class="row row-line">
    <div class="col-lg-8">
    		<div class="row">
								
<%
If SelectedCustomer <> "" Then 

	'Give them a message if there are open tickets
	If ShowPopup = True Then ' From global settings table
		OPTick=NumberOfServiceTicketsOpenForCust(SelectedCustomer)
		HLDTick=NumberOfServiceTicketsHOLDForCust(SelectedCustomer)
		If OPTick <> 0 AND HLDTick <> 0 Then
			Response.write("<script type=""text/javascript"">swal(""" & GetTerm("Account") & " has " & OPTick & " open service ticket(s) and " & HLDTick & " service ticket(s) on hold"");</script>")
		ElseIF OPTick <> 0 Then
			Response.write("<script type=""text/javascript"">swal(""" & GetTerm("Account") & " has " & OPTick & " open service ticket(s)"");</script>")
		ElseIF HLDTick<> 0 Then
			Response.write("<script type=""text/javascript"">swal(""" & GetTerm("Account") & " has " & HLDTick& " service ticket(s) on hold"");</script>")		
		End If
	End If%>

	
    	<!-- common customer display !-->
<div class="col-lg-8">
	<!--#include file="..\inc\commonCustomerDisplay.asp"-->
			        		</div>
	<!-- eof common customer display !-->
	
	</div>
    </div>
    
    

    
    <!-- info !-->
    <div class="col-lg-4">
	    <div class="table-info">
	    <div class="table-responsive">
			<table class="table">
						
						<tbody>
							<tr><td align="right">Return Type:</td><td>&nbsp;&nbsp;&nbsp;</td><td align="left"><%=tmpRetType %></td></tr>
							<tr><td align="right">Return Time:</td><td>&nbsp;&nbsp;&nbsp;</td><td align="left"><%=tmpRetTime %></td></tr>
							<tr><td align="right">Return Date:</td><td>&nbsp;&nbsp;&nbsp;</td><td align="left">
							<%'Code for datepicker start date
							sdate = year(tmpReturn) & "/ " & month(tmpReturn) & "/ " & day(tmpReturn)
							If IsNull(tmpReturn) Then 
								sdate = year(Now()) & "/ " & month(now()) & "/ " & day(now())
							Else	
								If DateDiff("d",cdate(tmpReturn),cdate("7/7/2077")) = 0 Then sdate  = year(Now()) & "/ " & month(now()) & "/ " & day(now())
							End If
							sdate =  replace(sdate,"/",",") 
							'Response.Write(sdate)%>
							<input type="text" id="txtCloseCancelDate" name="txtCloseCancelDate" value="<%=tmpReturn %>"  class="form-control last-run-inputs" data-beatpicker="true"     data-beatpicker-extra="customOptions"  data-beatpicker-format="['MM','DD','YYYY'],separator:'/'"> 
							<%  If Not IsNull(tmpReturn) Then
									If DateDiff("d",cdate(tmpReturn),cdate("7/7/2077")) <> 0 Then %>
										<input type="button" id="dteAuto" class="btn btn-primary setbtn" value="AUTO ADVANCE" onclick="AutoreturnDateFunc(<%=SelectedCustomer %>)"> 
									<% End If
							End If %>
							<input type="button" id="dteMan" class="btn btn-primary setbtn" value="MANUALLY SET" onclick="ManualreturnDateFunc(<%=SelectedCustomer %>)"> </td></tr>
						</tbody>	
													
  			</table>
		</div>
		       		
		<select class="form-control" name="selReturnDateReason" id="selReturnDateReason">
			<option value="0"selected>-- Select option if changing return date --</option>
			<option value="119">No order needed</option>
			<option value="120">Left message</option>
			<option value="121">Will call back</option>
			<option value="122">Email sent - last option</option>
			<option value="123">Charge rent - no order</option>
			<option value="124">Received order</option>
			<option value="125">Updated return date</option>
			<option value="126">Change return date to 7/7/77</option>
			<option value="127">Referred to sales department</option>
		</select>

		</div>
    </div>
    <!-- eof info !-->
    

    </div>
   <!-- eof row !-->
   


<!-- tabs start here !-->
<div class="row row-line">
	
	<!--tabs navigation -->
  <ul class="nav nav-tabs" role="tablist">
    <li role="presentation" class="active"><a href="#home" aria-controls="home" role="tab" data-toggle="tab">Current&nbsp;(<%=NumberOfCurrentNotes(SelectedCustomer)%>)</a></li>
    <li role="presentation"><a href="#Archived" aria-controls="Archived" role="tab" data-toggle="tab">Archived Notes&nbsp;(<%=NumberOfArchivedNotes(SelectedCustomer)%>)</a></li>
    <li role="presentation"><a href="#Attachments" aria-controls="Attachments" role="tab" data-toggle="tab">Attachments&nbsp;(<%=NumberOfAttachmentsNotes(SelectedCustomer)%>)</a></li>
    <li role="presentation"><a href="#ServiceTickets" aria-controls="Attachments" role="tab" data-toggle="tab">Service Tickets&nbsp;(<%=NumberOfServiceTicketsEver(SelectedCustomer)%>)</a></li>

   </ul>
  <!-- eof tabs navigation

  <!-- the tabs -->
  <div class="tab-content">
	  
	  <!-- table tab !-->
    <div role="tabpanel" class="tab-pane fade in active" id="home"> 

   <div class="row row-line">
	  
	   <div class="col-lg-12">
		    <a href="addAccountNote.asp">
		    <br><br>
			    <button type="button" class="btn btn-success">New Note</button>
		   </a>
	   </div>
	   
   </div>

<% ' S T I C K Y  F I R S T
SQL = "SELECT * FROM tblCustomerNotes Where CustNum ='" & SelectedCustomer & "' AND Sticky = 1 AND Archived <> 1 order by EntryDateTime Desc"
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
If not rs.EOF Then%>
	<!-- row !-->
	<div class="row-line">
		<div class="table-responsive">
 			<table class="table">
				<thead class="table-titles">
					<tr>
						<th class="date-col">Date/Time</th>
						<th class="notes-col">Notes</th>
						<th class="createdby-col">Created By</th>
						<th class="sticky-col">Sticky</th>
						<th class="sticky-col">Archive</th>
						<th class="sticky-col">&nbsp;</th>
					</tr>
				</thead>
				<tbody><%
					Do While Not rs.EOF
						If NoteNewForUser(SelectedCustomer,rs("EntryDateTime")) = True then%>
				   			<tr class="table-line-new-for-user">
						<%Else%> 
				   			<tr class="sticky-table-line">
			   			<%End If%>
						   <%
								Response.Write("<td class='date-col'>" & FormatDateTime(rs("EntryDateTime")) & "</td>")
							   	Response.Write("<td class='notes-col'>" & Replace(rs("Note"),"&nbsp;"," ") & "</td>")
							   	Response.Write("<td class='createdby-col'>" & GetUserDisplayNameByUserNo(rs("Userno")) & "</td>")
								Response.Write("<td><a href='#'><input type='checkbox' checked name='chk" & rs("InternalNoteNumber") & "' id='chk" & rs("InternalNoteNumber") & "' onclick='myFunction(" & rs("InternalNoteNumber") & ")')></a></td>")
						   		Response.Write("<td><a href='archiveAccountNoteQues.asp?nt=" & rs.Fields("InternalNoteNumber") & "'><i class='fa fa-archive' ></i></a></td>")
								Response.Write("<td class='date-col'>&nbsp;</td>")
						   %>
			   			</tr>
		   			<%
		   				rs.movenext
		   			Loop%>
					</tbody>
				</table>
			</div>
		</div>
<%
	cnn8.close
	set rs = nothing
	set cnn8 = nothing		
End IF%>


<% ' N O W      N O T      S T I C K Y 

SQL = "SELECT * FROM tblCustomerNotes Where CustNum ='" & SelectedCustomer & "' AND (Sticky = 0 or Sticky Is Null) AND Archived <> 1 order by EntryDateTime DESC"
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
If not rs.EOF Then%>
	<!-- row !-->
	<div class="row-line">
		<div class="table-responsive">
 			<table class="table  sortable">
				<thead class="table-titles">
					<tr>
						<th class="sorttable_numeric date-col">Date/Time</th>
						<th class="sorttable notes-col">Notes</th>
						<th class="sorttable createdby-col">Created By</th>
						<th class="sorttable_nosort sticky-col">Sticky</th>
						<th class="sorttable_nosort sticky-col">Archive</th>
						<th class="sorttable_numeric">Expires</th>
					</tr>
				</thead>
				<tbody><%
					LineX=1
					Do While Not rs.EOF 
						If NoteNewForUser(SelectedCustomer,rs("EntryDateTime")) = True then%>
				   			<tr class="table-line-new-for-user">
						<%Else
							If LineX Mod 2 = 0 then
								'THESE ARE EVEN LINES
								%><tr class="tr-even"><%
							Else
								'THESE ARE ODD LINE
								%><tr class="tr-odd"><%
							End If				   			
			   			End If
								Response.Write("<td class='date-col'>" & FormatDateTime(rs("EntryDateTime")) & "</td>")
							   	Response.Write("<td class='notes-col'>" & Replace(rs("Note"),"&nbsp;"," ") & "</td>")
							   	Response.Write("<td class='createdby-col'>" & GetUserDisplayNameByUserNo(rs("Userno")) & "</td>")
								Response.Write("<td><a href='#'><input type='checkbox' name='chk" & rs("InternalNoteNumber") & "' id='chk" & rs("InternalNoteNumber") & "' onclick='myFunction(" & rs("InternalNoteNumber") & ")')></a></td>")
						   		Response.Write("<td><a href='archiveAccountNoteQues.asp?nt=" & rs.Fields("InternalNoteNumber") & "'><i class='fa fa-archive' ></i></a></td>")
						   		'Response.Write("<td><input type='text' onchange='expirDateFunc(" & rs("InternalNoteNumber") & ")' id='txtExpirDate' name='txtExpirDate' value='"& FormatDateTime(rs("ExpirationDate")) &"'  class='form-control last-run-inputs' data-beatpicker='true'></td>")
						   		Response.Write("<td class='date-col'>" & FormatDateTime(rs("ExpirationDate")) & "</td>")
						   %>
			   			</tr>
		   			<%
		   				LineX = LineX + 1
		   				rs.movenext
		   			Loop%>
				</tbody>
			</table>
		</div>
	</div>
<%End IF%>

<%End If %>
</div>
    <!-- eof table tab !-->
    
<!-- Archive tab !-->
<div role="tabpanel" class="tab-pane fade" id="Archived">
    
<% ' A R C H I V E D
SQL = "SELECT * FROM tblCustomerNotes Where CustNum ='" & SelectedCustomer & "' AND Archived =1 order by EntryDateTime Desc"
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
If not rs.EOF Then%>
	<!-- row !-->
	<div class="row-line">
		<div class="table-responsive">
 			<table class="table">
				<thead class="table-titles">
					<tr>
						<th>Date/Time</th>
						<th class="notes-col">Notes</th>
						<th>Created By</th>
						<th>Archived By</th>
						<th>Move To<br>Current</th>
					</tr>
				</thead>
				<tbody><%
					Do While Not rs.EOF%> 
			   			<tr class="sticky-table-line">
						   <%
								Response.Write("<td>" & FormatDateTime(rs("EntryDateTime")) & "</td>")
							   	Response.Write("<td class='notes-col'>" & Replace(rs("Note"),"&nbsp;"," ") & "</td>")
							   	Response.Write("<td>" & GetUserDisplayNameByUserNo(rs("Userno")) & "</td>")
							   	Response.Write("<td>" & GetUserDisplayNameByUserNo(rs("ArchivedByUserno")) & "</td>")
						   		Response.Write("<td><a href='currentAccountNoteQues.asp?nt=" & rs.Fields("InternalNoteNumber") & "'><i class='fa fa-undo' ></i></a></td>")
						   %>
			   			</tr>
		   			<%
		   				rs.movenext
		   			Loop%>
					</tbody>
				</table>
			</div>
		</div>
<%
	cnn8.close
	set rs = nothing
	set cnn8 = nothing		
End IF%>

    
</div>
<!-- eof Archive tab !-->
    
    
<!-- Attachment tab !-->
<div role="tabpanel" class="tab-pane fade" id="Attachments">

   <div class="row row-line">
	  
	   <div class="col-lg-12">
		    <a href="addAccountNoteAttachments.asp">
		    <br><br>
			    <button type="button" class="btn btn-success">New Attachment</button>
		   </a>
	   </div>
	   
   </div>
    
<% ' A T T A C H M E N T S 
SQL = "SELECT * FROM tblCustomerNotesAttachments Where CustNum ='" & SelectedCustomer & "' order by EntryDateTime Desc"
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
If not rs.EOF Then%>
	<!-- row !-->
	<div class="row-line">
		<div class="table-responsive">
 			<table class="table sortable">
				<thead class="table-titles">
					<tr>
						<th class="sorttable_numeric date-col">Date/Time</th>
						<th class="sorttable notes-col">Notes</th>
						<th class="sorttable createdby-col">Created By</th>
						<th class="sorttable_nosort">Attachment</th>
					</tr>
				</thead>
				<tbody><%
					Do While Not rs.EOF%> 
			   			<tr>
						   <%
								Response.Write("<td class='date-col'>" & FormatDateTime(rs("EntryDateTime")) & "</td>")
							   	Response.Write("<td class='notes-col'>" & Replace(rs("Note"),"&nbsp;"," ") & "</td>")
							   	Response.Write("<td class='createdby-col'>" & GetUserDisplayNameByUserNo(rs("Userno")) & "</td>")
							   	Set fs = CreateObject("Scripting.FileSystemObject")
							   	Pth =  "../clientfiles/" & trim(MUV_Read("ClientID")) & "/attachments/" & rs("AttachmentFilename")
							   	If fs.FileExists(Server.MapPath(Pth)) Then
									Response.Write("<td><a href='" & baseURL & "clientfiles/" & trim(MUV_Read("ClientID")) & "/attachments/" & rs("AttachmentFilename") & "' target='_blank'>" & right(rs("AttachmentFilename"),len(rs("AttachmentFilename"))-Instr(rs("AttachmentFilename"),"-")) & "</a></td>")
								Else
									Response.Write("<td>&nbsp;</td>")
								End If
								Set fs = Nothing
						   %>
			   			</tr>
		   			<%
		   				rs.movenext
		   			Loop%>
				</tbody>
			</table>
		</div>
	</div>
	
<%End IF
cnn8.close
set rs = nothing
set cnn8 = nothing	%>
 
</div>
<!-- eof Attachment tab !-->


<!-- Service tickets tab !-->
<div role="tabpanel" class="tab-pane fade" id="ServiceTickets">

   <div class="row row-line">
	   
   </div>
<% ' S E R V I C E  T I C K E T S
SQL = "SELECT * FROM FS_ServiceMemos "
SQL = SQL & " where AccountNumber ='" & SelectedCustomer & "' "
SQL = SQL & " order by submissionDateTime desc"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)

If not rs.EOF Then
%>
	<!-- row !-->
	<div class="row-line">
		<div class="table-responsive">
 			<table class="table sortable table-striped">
              <thead>
                <tr>
	              <th class="sorttable_numeric">Date</th>
	              <th>Ticket #</th>	              
                  <th>Status</th>
                  <th class="sorttable_nosort">&nbsp;</th>
                  <th class="sorttable_nosort">Description</th>
                  <% If advancedDispatchIsOn() Then %>
		              <th>Stage</th>
	              <% Else %>
	                  <th>Dispatched</th>
	              <% End If %>
                  <th class="sorttable_numeric">Elapsed<br>Time</th>
                  <th class="sorttable_nosort">PIC</th>
                  <th class="sorttable_nosort">SIG</th>
                  <th>Submitted Via</th>
                </tr>
              </thead>
              
              <tbody class='searchable'>
<%
				Do While Not rs.EOF
						If rs.Fields("CurrentStatus") = rs.Fields("RecordSubType") Then ' Show only 1 line per memo, the most current status
				        %>
							<!-- table line !-->
							<tr class="low-priority">
							<%Response.write("<td sorttable_customkey=" & FormatAsSortableDateTime(rs("submissionDateTime")) & ">" & FormatDateTime(rs("submissionDateTime")) & "</td>")%>
							<%If rs.Fields("CurrentStatus")="OPEN" Then %>
								<td><a href='../service/editServiceMemo.asp?memo=<%= rs.Fields("MemoNumber")%>'><%= rs.Fields("MemoNumber")%></a></td>
							<% Else %>
								<td><a href='../service/viewServiceMemo.asp?memo=<%= rs.Fields("MemoNumber")%>'><%= rs.Fields("MemoNumber")%></a></td>
							<% End If %>
							<td><%= rs.Fields("RecordSubType") %></td>
							<td>							</td>
							<td>
							<%
								CompressLen = 27
								'See if there are linefeeds in there that need to come out
								If Instr(rs.Fields("ProblemDescription"),"<br>") <> 0 Then CompressLen = Instr(rs.Fields("ProblemDescription"),"<br>")
								If CompressLen > 27 Then CompressLen = 27
								If len(rs.Fields("ProblemDescription")) > CompressLen Then Response.Write(Left(rs.Fields("ProblemDescription"),CompressLen)) Else Response.Write(rs.Fields("ProblemDescription"))%>
							</td>
							<%
								If rs.Fields("CurrentStatus") <> "CLOSE" and rs.Fields("CurrentStatus") <> "CANCEL" Then ' dont show a stage if they are closed or cancelled
									Response.Write("<td><b>"& GetServiceTicketCurrentStage(rs.Fields("MemoNumber")) & "</b><br>")
									Response.Write(GetServiceTicketSTAGEUser(rs.Fields("MemoNumber"),GetServiceTicketCurrentStage(rs.Fields("MemoNumber"))) & "<br>")
									Response.Write(GetServiceTicketSTAGEDateTime(rs.Fields("MemoNumber"),GetServiceTicketCurrentStage(rs.Fields("MemoNumber")))& "</td>")
								Else
									Response.Write("<td>&nbsp;</td>")
								End If
							If ElapsedTimeCalcMethod() = "Actual" Then
								If rs.Fields("CurrentStatus") = "CLOSE" or rs.Fields("CurrentStatus") = "CANCEL" Then
									elapsedMinutes = datediff("n",GetServiceTicketOpenDateTime(rs.Fields("MemoNumber")),GetServiceTicketCloseDateTime(rs.Fields("MemoNumber")))
									elapsedMinutesForSorting = elapsedMinutes
									elapsedString = ""
									minutesInServiceDay = GetNumberOfMinutesInServiceDay()
									elapsedDays = 	elapsedMinutes \ minutesInServiceDay
									If int(elapsedDays) > 0 Then
										elapsedMinutes = elapsedMinutes - (int(elapsedDays) * minutesInServiceDay)
										elapsedString = elapsedDays & "d "
									End If
									elapsedHours = 	elapsedMinutes \ 60
									If int(elapsedHours) > 0 Then 
										elapsedMinutes = elapsedMinutes - (int(elapsedHours) * 60)
										elapsedString = elapsedString  & elapsedHours & "h "
									End IF	
									If int(elapsedMinutes) > 0 Then
										elapsedString = elapsedString  & elapsedMinutes & "m"
									End If
									elapsedMinutes = datediff("n",GetServiceTicketOpenDateTime(rs.Fields("MemoNumber")),GetServiceTicketCloseDateTime(rs.Fields("MemoNumber")))
									Response.Write("<td sorttable_customkey='" & elapsedMinutesForSorting & "'>" & elapsedString & "<br>") 
									'response.Write(elapsedMinutes)
								ElseIf rs.Fields("CurrentStatus") = "OPEN" Then 
									elapsedMinutes = datediff("n",GetServiceTicketOpenDateTime(rs.Fields("MemoNumber")),Now())
									elapsedMinutesForSorting = elapsedMinutes
									elapsedString = ""
									minutesInServiceDay = GetNumberOfMinutesInServiceDay()
									elapsedDays = 	elapsedMinutes \ minutesInServiceDay
									If int(elapsedDays) > 0 Then
										elapsedMinutes = elapsedMinutes - (int(elapsedDays) * minutesInServiceDay)
										elapsedString = elapsedDays & "d "
									End If
									elapsedHours = 	elapsedMinutes \ 60
									If int(elapsedHours) > 0 Then 
										elapsedMinutes = elapsedMinutes - (int(elapsedHours) * 60)
										elapsedString = elapsedString  & elapsedHours & "h "
									End IF	
									If int(elapsedMinutes) > 0 Then
										elapsedString = elapsedString  & elapsedMinutes & "m"
									End If
									elapsedMinutes = datediff("n",GetServiceTicketOpenDateTime(rs.Fields("MemoNumber")),Now())
									Response.Write("<td sorttable_customkey='" & elapsedMinutesForSorting & "'>" & elapsedString & "<br>") 
									'response.Write(elapsedMinutes)
								Elseif rs.Fields("CurrentStatus") = "HOLD" Then
									'Response.Write("<td sorttable_customkey='" & 0 & "'>" & "Hold<br>") 
									elapsedMinutes = datediff("n",GetServiceTicketOpenDateTime(rs.Fields("MemoNumber")),Now())
									elapsedMinutesForSorting = elapsedMinutes
									elapsedString = ""
									minutesInServiceDay = GetNumberOfMinutesInServiceDay()
									elapsedDays = 	elapsedMinutes \ minutesInServiceDay
									If int(elapsedDays) > 0 Then
										elapsedMinutes = elapsedMinutes - (int(elapsedDays) * minutesInServiceDay)
										elapsedString = elapsedDays & "d "
									End If
									elapsedHours = 	elapsedMinutes \ 60
									If int(elapsedHours) > 0 Then 
										elapsedMinutes = elapsedMinutes - (int(elapsedHours) * 60)
										elapsedString = elapsedString  & elapsedHours & "h "
									End IF	
									If int(elapsedMinutes) > 0 Then
										elapsedString = elapsedString  & elapsedMinutes & "m"
									End If
									elapsedMinutes = datediff("n",GetServiceTicketOpenDateTime(rs.Fields("MemoNumber")),Now())
									Response.Write("<td sorttable_customkey='" & elapsedMinutesForSorting & "'>" & elapsedString & "<br>") 
								End If
							Else
								If rs.Fields("CurrentStatus") = "CLOSE" or rs.Fields("CurrentStatus") = "CANCEL" Then
									elapsedMinutes = ServiceCallElapsedMinutes(rs.Fields("MemoNumber"))
									elapsedMinutesForSorting = elapsedMinutes 
									elapsedString = ""
									elapsedHours = 	elapsedMinutes \ 60
									If int(elapsedHours) > 0 Then 
										elapsedMinutes = elapsedMinutes - (int(elapsedHours) * 60)
										elapsedString = elapsedString  & elapsedHours & "h "
									End IF	
									If int(elapsedMinutes) > 0 Then
										elapsedString = elapsedString  & elapsedMinutes & "m"
									End If
									If elapsedMinutes = 0 Then elapsedString = "0"
									Response.Write("<td sorttable_customkey='" & elapsedMinutesForSorting & "'>" & elapsedString & "<br>") 
									'response.Write(elapsedMinutes)
								ElseIf rs.Fields("CurrentStatus") = "OPEN" Then 
									elapsedMinutes = ServiceCallElapsedMinutes(rs.Fields("MemoNumber"))
									elapsedMinutesForSorting = elapsedMinutes
									elapsedString = ""
									elapsedHours = 	elapsedMinutes \ 60
									If int(elapsedHours) > 0 Then 
										elapsedMinutes = elapsedMinutes - (int(elapsedHours) * 60)
										elapsedString = elapsedString  & elapsedHours & "h "
									End IF	
									If int(elapsedMinutes) > 0 Then
										elapsedString = elapsedString  & elapsedMinutes & "m"
									End If
									Response.Write("<td sorttable_customkey='" & elapsedMinutesForSorting & "'>" & elapsedString & "<br>") 
									'response.Write(elapsedMinutes)
								Elseif rs.Fields("CurrentStatus") = "HOLD" Then
									'Response.Write("<td sorttable_customkey='" & 0 & "'>" & "Hold<br>") 
									elapsedMinutes = ServiceCallElapsedMinutes(rs.Fields("MemoNumber"))
									elapsedMinutesForSorting = elapsedMinutes
									elapsedString = ""
									elapsedHours = 	elapsedMinutes \ 60
									If int(elapsedHours) > 0 Then 
										elapsedMinutes = elapsedMinutes - (int(elapsedHours) * 60)
										elapsedString = elapsedString  & elapsedHours & "h "
									End IF	
									If int(elapsedMinutes) > 0 Then
										elapsedString = elapsedString  & elapsedMinutes & "m"
									End If
									Response.Write("<td sorttable_customkey='" & elapsedMinutesForSorting & "'>" & elapsedString & "<br>") 
								End If
							End If
							%>
							</td>
							<td>
							<%
							set fs = CreateObject("Scripting.FileSystemObject")
							Pth =  "../clientfiles/" & trim(MUV_Read("ClientID")) & "/SvcMemoPics/" & rs("MemoNumber") & "-1.jpg"
							Pth2 =  "../clientfiles/" & trim(MUV_Read("ClientID")) & "/SvcMemoPics/" & rs("MemoNumber") & "-1.jpeg"
							If fs.FileExists(Server.MapPath(Pth)) or fs.FileExists(Server.MapPath(Pth2)) Then
								%>X<%
							End If
							%>
							</td>

							<% If rs.Fields("RecordSubType") = "CLOSE" Then 
								
								'----------------------------
								'Service Signature Check
								'----------------------------
								set fs = CreateObject("Scripting.FileSystemObject")
								Pth =  "../clientfiles/" & trim(MUV_Read("ClientID")) & "/signaturesave/TicketID-" & rs("MemoNumber") & ".png"
								
								If fs.FileExists(Server.MapPath(Pth)) Then
									hasServiceSignature = True
								Else
									hasServiceSignature = False
								End If
													
								'Response.Write(Pth)
								
								'***************************************************************************************************
								'Display signature file, if any exist in the signaturesave directory
								''Check for the existance of a thumbnail image in the directory, otherwise, size the image with CSS
								'***************************************************************************************************
				
								Pth =  "../clientfiles/" & trim(MUV_Read("ClientID")) & "/signaturesave/TicketID-" & rs("MemoNumber") & ".png"
								PthThumb =  "../clientfiles/" & trim(MUV_Read("ClientID")) & "/signaturesave/TicketID-" & rs("MemoNumber") & "-thumb.png"
			
								SignaturePathNameFull = BaseURL & "clientfiles/" & trim(MUV_Read("ClientID")) & "/signaturesave/TicketID-" & rs("MemoNumber") & ".png"
								SignaturePathNameThumb = BaseURL & "clientfiles/" & trim(MUV_Read("ClientID")) & "/signaturesave/TicketID-" & rs("MemoNumber") & "-thumb.png"
								
								If hasServiceSignature = True Then
									
									If fs.FileExists(Server.MapPath(PthThumb)) Then
								    	%><td align="left"><a href="<%= SignaturePathNameFull %>" target="_blank" style="border:0px;"><img src="<%= SignaturePathNameThumb %>" alt="Ticket <%= rs("MemoNumber") %> Signature"></a></td><%
								    Else
								    	%><td align="left"><a href="<%= SignaturePathNameFull %>" target="_blank" style="border:0px;"><img src="<%= SignaturePathNameFull %>" alt="Ticket <%= rs("MemoNumber") %> Signature" style="width:200px;"></a></td><%
								    End If
								    
								Else
									 %><td align="left">No Signature</td><%
								End If
								
								
							End If
							set fs=nothing
							%>
						
							<td><%= rs.Fields("SubmissionSource") %></td>
							</tr>
							<!-- eof table line !-->
						<%
				
						End If
						
						rs.movenext	
		   			Loop %>
				</tbody>
			</table>
		</div>
	</div>
	
<%End IF
cnn8.close
set rs = nothing
set cnn8 = nothing	%>
 
</div>
<!-- eof Attachment tab !-->

    
<% dummy = MARKNoteNewForUser(SelectedCustomer) 'Now mark all the notes as having been viewed %>
   
    
    <!-- third tab !-->
    <div role="tabpanel" class="tab-pane fade" id="tab3">...</div>
    <!-- eof third tab !-->
    
   </div>
  <!-- eof the tabs !-->
	
</div>
<!-- tabs end here !-->


<!-- date picker custom options !-->
<script type="text/javascript">
	customOptions = {
 
   startDate: new Date([<%=sdate%>]),
   currentDate: new Date([<%=sdate%>]),
 
   }

	</script>
<!-- eof date picker custom options !-->

<!--#include file="../inc/footer-service.asp"-->
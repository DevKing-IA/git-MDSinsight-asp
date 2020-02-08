<!--#include file="../inc/header-field-service-mobile.asp"-->
<!--#include file="../inc/InsightFuncs.asp"-->
<!--#include file="../inc/InsightFuncs_Service.asp"-->
<% 
	SelectedMemoNumber = Request.QueryString("t")  
	SourceTab = Request.QueryString("tab")
	CustID = GetServiceTicketCust(SelectedMemoNumber)
	CustomerName = GetCustNameByCustNum(CustID)
	UserNo = Session("UserNo")
	UserName = GetUserDisplayNameByUserNo(UserNo)
%>


<SCRIPT LANGUAGE="JavaScript">
    function validateAddServiceMemoNote()
    {

  	  if (document.frmAddServiceMemoNote.txtNewServiceNote.innerText!= "") {
        swal("Please enter a note before submitting.");
        return false;
  	  }
  	  
      return true;
      
    }
</SCRIPT>  

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
		background: transparent;
		border: 0px;
		cursor: pointer;
 	}

	.input-lg::-webkit-input-placeholder, textarea::-webkit-input-placeholder {
	  color: #666;
	}
	.input-lg:-moz-placeholder, textarea:-moz-placeholder {
	  color: #666;
	}
	.checkboxes label{
		font-weight: normal;
		margin-right: 20px;
	}
	.close-service-client-output{
		text-align: left;
	}
	.ticket-details{
		margin-bottom: 15px;
	}

	.btn-link {
	    font-weight: 500;
	    font-size: 1.2em;
	    color: #343173;
	    background-color: transparent;
	    white-space: normal;
	    padding: 0px;
	}

	.btn-link:hover {
	    color: #007bff;
	    text-decoration:none;
	    background-color: transparent;
	    border-color: transparent;
	}	
	.accordion-box{
		margin-bottom:15px;
	}
	
	.close-service-h4{
		text-align:center;
	}
	
	 .fa-stack  { font-size: 0.7em; }
	  i { vertical-align: middle; }
	  
  
	
</style>

<h1 class="fieldservice-heading">
	<form method="post" action="addViewServiceMemoNotes_PassThru.asp" name="frmServiceNoteReturnBack" id="frmServiceNoteReturnBack">
		<input type="hidden" id="txtTicketNumber" name="txtTicketNumber" value="<%= SelectedMemoNumber %>">	
		<input type="hidden" id="txtReturnTab" name="txtReturnTab" value="<%= SourceTab %>">	 
		<button type="button" onclick="document.forms['frmServiceNoteReturnBack'].submit();" class="btn-home"><i class="fa fa-arrow-left"></i></button>
	</form>

	Add Service Note For Ticket #<%= SelectedMemoNumber %>
</h1>

 
<div class="container-fluid">
 
 	<h4>New Note For Ticket #<%= SelectedMemoNumber %></h4>
	<h5><%= GetTerm("Account") %> #<%= CustID %> (<%= CustomerName %>)</h5>

	<form action="addServiceMemoNote_submit.asp" method="POST" name="frmAddServiceMemoNote" id="frmAddServiceMemoNote" onsubmit="return validateAddServiceMemoNote();"> 

		<input type="hidden" id="txtTicketNumber" name="txtTicketNumber" value="<%= SelectedMemoNumber %>">
		<input type="hidden" id="txtReturnTab" name="txtReturnTab" value="<%= SourceTab %>">		
		
		<!-- order notes !-->
		<div class="row row-rest">
			<h4 class="close-service-h4">										
				<span class="fa-stack" style="vertical-align: top;">
					<i class="fas fa-sticky-note fa-stack-2x" style="color:#28a745;"></i>
					<i class="fas fa-plus fa-stack-1x fa-inverse"></i>
				</span>	
 				New Service Note:</h4>
 				
			<div class="col-lg-12 close-service-box">
				<textarea class="form-control input-lg" rows="5" name="txtNewServiceNote" spellcheck="True" id="txtNewServiceNote"></textarea>
			</div>
		</div>
		
		<button class="btn btn-info btn-block btn-lg close-buttons mt-4" name="Submit" value="Add Service Ticket Note" type="submit"><i class="fas fa-save"></i> Save Service Ticket Note</button>
			

		<!-- eof order notes !-->
	</form>
	
	

	<form method="post" action="addViewServiceMemoNotes_PassThru.asp" name="frmServiceNoteReturnBack2" id="frmServiceNoteReturnBack2">
		<input type="hidden" id="txtTicketNumber" name="txtTicketNumber" value="<%= SelectedMemoNumber %>">
		<input type="hidden" id="txtReturnTab" name="txtReturnTab" value="<%= SourceTab %>">
		<button type="button" onclick="document.forms['frmServiceNoteReturnBack2'].submit();" class="btn btn-warning btn-block btn-lg close-buttons"><i class="fa fa-times-circle-o"></i> Cancel</button>
	</form>
	
	
	
	<div class="list-group">

		<%
		
		Set cnn8 = Server.CreateObject("ADODB.Connection")
		cnn8.open (Session("ClientCnnString"))
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = 3 

		SQL = "SELECT InternalRecordIdentifier, RecordCreationDateTime, ServiceTicketID, EnteredByUserNo, Note "
		SQL = SQL & "FROM FS_ServiceMemosNotes WHERE ServiceTicketID = '" & SelectedMemoNumber & "' "
		SQL = SQL & " ORDER BY RecordCreationDateTime DESC"
		
		set rs = cnn8.Execute (SQL)
		
		If NOT rs.EOF Then
		
			%><h3>Previous Notes:</h3><%
	
			Do While NOT rs.EOF
			
				InternalRecordIdentifier = rs("InternalRecordIdentifier")
				RecordCreationDateTime = rs("RecordCreationDateTime")
				EnteredByUserNo = rs("EnteredByUserNo")
				Note = rs("Note")
									
				%>
				<span class="list-group-item list-group-item-action flex-column align-items-start">
					<div class="d-flex w-100 justify-content-between">
						<h5 class="mb-1 font-weight-bold" style="font-size:1.1em;"><%= FormatDateTime(RecordCreationDateTime,2) %></h5>
					</div>
					<h6 class="mb-1 mt-1"><%= FormatDateTime(RecordCreationDateTime,3) %></h6>						
					<p class="mb-1">Entered By <%= GetUserDisplayNameByUserNo(EnteredByUserNo) %></p>
					<h6 class="mb-1 mt-1"><%= Note %></h6>
				</span>
				<%
				
				'********************************************************
				'CODE HERE TO MARK NOTE AS BEING READ
				 Call MarkNoteNewForUserServiceTicket(SelectedMemoNumber)
				'********************************************************
				
				rs.MoveNext
				
			Loop
			
		End If
		
		Set rs = Nothing
		Set Cnn8 = Nothing

	%>	
	
	</div>
		
</div>

 
<!--#include file="../inc/footer-field-service-noTimeout.asp"-->
<!--#include file="../inc/header-accounts-receivable.asp"-->


<% MemoNumber = Request.QueryString("memo") 
If MemoNumber = "" Then Response.Redirect(BaseURL)

SQL = "SELECT * FROM FS_ServiceMemos where MemoNumber = '" & MemoNumber  & "'"
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	OpenServiceMemoRecNumber = rs("ServiceMemoRecNumber")
	OpenCurrentStatus = rs("CurrentStatus")
	OpenRecordSubType = rs("RecordSubType")
	OpenSubmittedByName = rs("SubmittedByName")
	OpenAccountNumber = rs("AccountNumber")
	OpenCompany = rs("Company")
	OpenProblemLocation = rs("ProblemLocation")
	OpenSubmittedByPhone = rs("SubmittedByPhone")
	OpenSubmittedByEmail = rs("SubmittedByEmail")
	OpenSubmissionDateTime = rs("SubmissionDateTime")
	OpenProblemDescription = rs("ProblemDescription")
	OpenMode = rs("Mode")
	OpenSubmissionSource = rs("SubmissionSource")
	OpenUserNoOfServiceTech = rs("UserNoOfServiceTech")
End If
set rs = Nothing
cnn8.close
set cnn8 = Nothing
If OpenSubmittedByName = "" Then OpenSubmittedByName = "Not provided"
If OpenSubmittedByPhone = "" Then OpenSubmittedByPhone = "Not provided"
If OpenSubmittedByEmail = "" Then OpenSubmittedByEmail = "Not provided"
If OpenProblemLocation = "" Then OpenProblemLocation = "Not provided"
If OpenProblemDescription = "" Then OpenProblemDescription = "Not provided"
If CCSubmittedByName = "" Then CCSubmittedByName = "Not provided"
If CCSubmittedByPhone = "" Then CCSubmittedByPhone = "Not provided"
If CCSubmittedByEmail = "" Then CCSubmittedByEmail = "Not provided"
If CCProblemLocation = "" Then CCProblemLocation = "Not provided"
If CCProblemDescription = "" Then CCProblemDescription = "Not provided"

'********************
'Advanced dispatching
'*********************
'If advanced dispatching is on, when they come in here we must change the stage to under review
'it does not matter if the operator actually takes action or not. As soon as they click it,
'it becomes under review with their name.

Set Connection = Server.CreateObject("ADODB.Connection")
Set Recordset = Server.CreateObject("ADODB.Recordset")
Recordset.CursorLocation = 3 
Connection.Open InsightCnnString


' If it was only received, this would be the first deail record
SQL = "INSERT INTO " & MUV_Read("SQL_Owner") & ".FS_ServiceMemosDetail (MemoNumber, CustNum, MemoStage, "
SQL = SQL & "SubmissionDateTime, USerNoSubmittingRecord,Remarks)"
SQL = SQL &  " VALUES (" 
SQL = SQL & "'"  & MemoNumber & "'"
SQL = SQL & ",'"  & OpenAccountNumber & "'"
SQL = SQL & ",'Under Review'"
SQL = SQL & ",getdate() "
SQL = SQL & ","  & Session("UserNo") & ",'Viewed by " & GetUserDisplayNameByUserNo(Session("UserNo"))  & "')"
'Response.Write(SQL)
'Response.end
Set Connection2 = Server.CreateObject("ADODB.Connection")
Set Recordset2 = Server.CreateObject("ADODB.Recordset")
Recordset2.CursorLocation = 3 
Connection2.Open Session("ClientCnnString")
Set Recordset2 = Connection2.Execute(SQL)
Connection2.Close
Set Recordset2 = Nothing
Set Connection2 = Nothing


Connection.Close
Set Recordset = Nothing
Set Connection = Nothing
'Write audit trail for under review
'*******************************
Description = "Service ticket number " & ServiceTicketNumber & " changed from received to under review by " & GetUserDisplayNameByUserNo(Session("UserNo")) & " at " & NOW()
CreateAuditLogEntry "Service Ticket System","Under review","Minor",0,Description 

%>

<link rel="stylesheet" type="text/css" href="<%= BaseURL %>css/tabs.css" />

 
<style>
	
	.thumbnail{
		max-width: 100px;
		max-height: 100px;
	}
	
	#lightbox .modal-content {
    display: inline-block;
    text-align: center;   
}

#lightbox .close {
    opacity: 1;
    color: rgb(255, 255, 255);
    background-color: rgb(25, 25, 25);
    padding: 5px 8px;
    border-radius: 30px;
    border: 2px solid rgb(255, 255, 255);
    position: absolute;
    top: -15px;
    right: -55px;
    
    z-index:1032;
}


.beatpicker-clear{
	display: block;
	text-indent:-9999em;
	line-height: 0;
	visibility: hidden;
}    

 	.alert{
 		padding: 6px 12px;
 		margin-bottom: 0px;
	}
	
	.form-control{
		margin-bottom: 20px;
	}
	
	a:hover{
		text-decoration: none;
	}
	
	[class^="col-"]{
	 margin-bottom:25px;
  } 
  
  .custom-hr{
height: 3px;
margin-left: auto;
margin-right: auto;
background-color:#183049;
color:#183049;
border: 0 none;
}

.control-label{
	padding-top: 5px;
}

.table-info .table>tbody>tr>td, .table>tbody>tr>th, .table>tfoot>tr>td, .table>tfoot>tr>th, .table>thead>tr>td, .table>thead>tr>th{
 	font-weight: bold;
	line-height: 0.8;
}

.date-time-col{
	width: 10%;
}

.stage-col{
	width: 10%;
}

.notes-col-{
	width: 60%;
}  

.user-col{
	width: 20%;
}
	</style>


<h1 class="page-header"><i class="fa fa-wrench"></i> View Service Ticket On Hold</h1>

	<form method="POST" action="releaseservicememo_submit.asp" name="frmReleaseServiceMemo">
      

        <input type="hidden" id="txtMemoNumber" name="txtMemoNumber" value="<%=MemoNumber%>"  class="form-control last-run-inputs">
        <input type="hidden" id="txtCustID" name="txtCustID" value="<%=OpenAccountNumber%>"  class="form-control last-run-inputs">


        
 	        <!-- row !-->		
	        <div class="row ">
		        

		        <!--account number !-->
		        <div class="col-lg-6 col-md-4 col-sm-12 col-xs-12">
		        	<%SelectedCustomer = OpenAccountNumber %>
					<!--#include file="../inc/commonCustomerDisplay.asp"-->
			    </div>
		        <!-- eof account number !-->
		        
		        <!-- company name !-->
		        <div class="col-lg-3 col-md-4 col-sm-12 col-xs-12">
			        <div class="alert alert-info" role="alert"><strong>Ticket#: <%= MemoNumber %></strong></div>
			        <br>
			        <!-- row !-->			
			    <div class="row">

			    	<!-- Contact Name !-->
			    <div class="col-lg-12">
			        <strong>Contact Name</strong><br>
			        <% =OpenSubmittedByName %>
			        </div>
			    	<!-- Contact Name !-->
	
			    	<!-- Contact Phone !-->
			    	  <div class="col-lg-12">
				    	 <strong>Contact Phone</strong><br>
				    	 <% =OpenSubmittedByPhone %>
			        </div>
			    	<!-- Contact Phone !-->
			    	
			    	<!-- Problem Location !-->
			   <div class="col-lg-12">
			        <strong>Problem Location</strong><br>
					<% =OpenProblemLocation %>

			        </div>
			    	<!-- Problem Location !-->
	  					    	
			    	<!-- Description of problem !-->
			    	 <div class="col-lg-12">
 					<strong>Problem Description</strong><br>
					<% =OpenProblemDescription %>
			    	 </div>
 			    	<!-- Description of problem !-->
			    	
			    	
			    	</div>
			    <!-- eof row !-->


 		        </div>
		        <!-- eof company name !-->

		        <!-- company name !-->
		        <div class="col-lg-3 col-md-4 col-sm-12 col-xs-12">
			        <div class="row">
					
								
  		        </div>
		        <!-- eof company name !-->

						        		
		        </div>
 <!-- eof row !-->

 <!-- main row !-->
 <div class="row">
 
		    	     			 
			
			<!-- rightmost col !-->
			<div class="col-lg-4 col-md-4 col-sm-12 col-xs-12">  
				
					
					
					<!-- info !-->
    <div class="col-lg-12 table-info">
	    <div class="table-responsive">
	    
		<% ' Lookup the customer record to get the other stuff we need

		SQL = "SELECT * FROM " & MUV_Read("SQL_Owner") & ".AR_Customer WHERE CustNum = '" & OpenAccountNumber & "'"
								
		Set cnn8 = Server.CreateObject("ADODB.Connection")
		cnn8.open (Session("ClientCnnString"))
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = 3 
		Set rs = cnn8.Execute(SQL)
		If not rs.Eof Then
			tmpStatus = rs("AcctStatus")
			tmpChain = rs("ChainNum")
			tmpAssociatedNumber = rs("ArOldAcctNum")
			tmpSalesman = rs("Salesman")
			tmpSalesman2 = rs("SecondarySalesman")	
			tmpReferral = rs("ReferalCode")	
			tmpARrep = rs("ArRep")		
			tmpCustType = rs("CustType")
		End IF
		rs.close
		cnn8.Close
		%>
	    	
		</div>
    </div>
    <!-- eof info !-->

				</div>
			</div>
			<!-- eof rightmost col !-->
	         
	         
		</div>
		<!-- eof main row !-->
		
		
		
		<div class="row">
			<div class="col-lg-12">
				<hr class="custom-hr">
			</div>
		</div>
		
 			    	<!-- service notes !-->
 			    	<div class="col-lg-12">
	 			    	<label>Notes relating to this ticket being released from hold</label>
	 			    	<textarea name="releasenotes"  spellcheck="True" id="releasenotes" class="form-control" rows="6"></textarea>
 			    	</div>
 			    	<!-- eof service notes !-->


			<div class="row">
			
			<div class="col-lg-12">	
			    <a href="<%= BaseURL %>accountsreceivable/TicketsOnHold.asp">
			    	<button type="button" class="btn btn-default">&lsaquo; Go Back To Service Ticket List</button>
				</a>
				
				<button type="submit" id="btnSaveOrRelease" name="btnSaveOrRelease"class="btn btn-primary" value="Save"><i class="far fa-save"></i> Save Notes Only, Don't Release</button>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<button type="submit" id="btnSaveOrRelease" name="btnSaveOrRelease" class="btn btn-primary" value="Release"><i class="fa fa-upload"></i> Release From Hold</button>
			</div>
			
 			
			</div>
			<!-- eof row !-->    
			
			<div class="row">
			<div class="col-lg-12">
				<hr class="custom-hr">
			</div>
		</div>
		
<% MDG_MemoNumber = MemoNumber %>		
<!--#include file="../service/memo_details_grid.asp"-->

   </form>
<!--#include file="../inc/footer-main.asp"-->

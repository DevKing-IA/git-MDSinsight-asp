﻿<!--#include file="../../../inc/header-accounts-receivable.asp"-->
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
	
	.title{
		color: #337ab7;
    	text-decoration: none;
    	font-size:1.3em;
	}
</style>

<div id="PleaseWaitPanel">
	<br><br>Processing, please wait...<br><br>
	<img src="../../../img/loading.gif"/>
</div>

<script type="text/javascript">

	$(document).ready(function() {
	
	    $("#PleaseWaitPanel").hide();

		$('#consolidatedInvoiceEmailModal').on('show.bs.modal', function(e) {
		
		    //get data attributes of invoice to email
		    
		    var CustID = $("#txtCustID").val();
		    var CustName = $("#txtCustName").val();
		    var InvoiceNumber= $("#txtConsInvNumber").val();
		    var EndDate = $("#txtEndDate").val();
		    
		    	    
		    var $modal = $(this);
	
    		$modal.find('#myEmailConsolidatedInvoiceModalAccountLabel').html("Email Consolidated Invoice for " + CustName + " - Invoice #" + InvoiceNumber);
    		
	    	$.ajax({
				type:"POST",
				url: "../../../inc/InSightFuncs_AjaxForARAP.asp",
				cache: false,
				data: "action=GetContentForEmailConsolidatedInvoiceModalAccount&consInvoiceNumber=" + encodeURIComponent(InvoiceNumber) + "&custID=" + encodeURIComponent(CustID) + "&endDate=" + encodeURIComponent(EndDate) + "&paidOrUnpaid=PAID",
				success: function(response)
				 {
	               	 $modal.find('#EmailConsolidatedInvoiceModalAccountContent').html(response);
	             }
	    	});
		    
		});
		
		

		$('#btnSaveConsolidatedInvoice').on('click', function(e) {
		
		    //get data-id attribute of the clicked alert
		    var custID = $("#txtCustID").val();
		    var endDate = $("#txtEndDate").val();
						    							    		    		
	    	$.ajax({
				type:"POST",
				url:"consolidated_InvDetail_frm_acct_save_only.asp",
				data: "e=" + encodeURIComponent(endDate) + "&c=" + encodeURIComponent(custID),
				success: function(response)
				 {
				 	swal("Consolidated Invoice Saved Successfully");
	             }
			});
    	});	
    	
	

		$('#btnSaveAndPostConsolidatedInvoice').on('click', function(e) {
		
		    //get data-id attribute of the clicked alert
		    var consInvoiceNum = $("#txtConsInvNumber").val();
		    var custID = $("#txtCustID").val();
		    var endDate = $("#txtEndDate").val();
		    var dueDateDays = $("#txtDueDateDays").val();
			var dueDateSingle = $("#txtDueDateSingleDate").val();
		    							    		    		
	    	$.ajax({
				type:"POST",
				url:"consolidated_InvDetail_frm_acct_save_and_post.asp",
				data: "e=" + encodeURIComponent(endDate) + "&c=" + encodeURIComponent(custID) + "&ddd=" + encodeURIComponent(dueDateDays) + "&dds=" + encodeURIComponent(dueDateSingle),				
				success: function(response)
				 {
				 	swal("Consolidated Invoice Saved and Posted to Metroplex Successfully");
	             }
			});
    	});	
    	
		
   
	});
</script>

<script type="text/javascript">

	function HideIt()
	{
		$("#PleaseWaitPanel").hide();
	}
	
</script>


<%
'baseURL should always have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)
sURL = Request.ServerVariables("SERVER_NAME")

'Generate a unique number to be used for all pdfs throughout this page
Randomize
UniqueNum = int((9999999-1111111+1)*rnd+1111111)


StartDate = Request.QueryString("s")
EndDate = Request.QueryString("e")
Account = Request.QueryString("c")
StartDate = Replace(StartDate, "~","/")
EndDate = Replace(EndDate, "~","/")
DueDateDays = Request.QueryString("ddd")
DueDateSingleDate = Request.QueryString("dds")
DoNotShowDueDate = Request.QueryString("dnsdd")

Set Pdf = Server.CreateObject("Persits.Pdf")
Set Doc = Pdf.CreateDocument

ImpVar = baseURL & "accountsreceivable/reports/consInvDetail/consolidated_InvDetail_frm_acct_PDFgen.asp?s=" & Replace(StartDate,"/","~") & "&e=" & Replace(EndDate,"/","~")& "&c=" & Account
ImpVar = ImpVar & "&un=" & Session("UserNo")
ImpVar = ImpVar & "&ddd=" & DueDateDays
ImpVar = ImpVar & "&dds=" & DueDateSingleDate
ImpVar = ImpVar & "&dnsdd=" & DoNotShowDueDate

If Left(MUV_Read("ClientID"),4)="1106" Then ImpVar = ImpVar & "&cl=" & MUV_Read("ClientID") & "&u=" & "cdcinsightdev" & "&p=" & "2oobr04dw4Y"
If Left(MUV_Read("ClientID"),4)="1071" Then ImpVar = ImpVar & "&cl=" & MUV_Read("ClientID") & "&u=" & MUV_Read("SQL_Owner") & "&p=" & "5um47AS"
If Left(MUV_Read("ClientID"),4)="1128" Then ImpVar = ImpVar & "&cl=" & MUV_Read("ClientID") & "&u=" & MUV_Read("SQL_Owner") & "&p=" & "04kv!1133SS"
'Response.Write("ImpVar :" & ImpVar & "<br>")

Doc.ImportFromUrl ImpVar, "scale=0.6; hyperlinks=false; drawbackground=true;"

fn = "\clientfiles\" & Left(MUV_Read("ClientID"),4) & "\z_pdfs\ConsolidatedStatement_" & Trim(UniqueNum) & "_Main.pdf"
fn = Replace(fn,"/","-")
fn = Replace(fn,":","-")
'response.write(fn & "<br>")
fn2 = Left(baseURL,Len(baseURL)-1) & fn
fn2 = Replace(fn2,"\","/")
'response.write(fn2 & "<br>")
Main_PDF_Filename = fn
If DebugMessages = True Then response.write("Main_PDF_Filename:" & Main_PDF_Filename & "<br>")
Filename = Doc.Save(Server.MapPath(fn), False)

'Now wait until the file exists on the server before we try to mail it
TimeoutSecs = 60
TimeoutCounter=0
FOundFile = False
Do While TimeoutCounter < TimeoutSecs 
	If CheckRemoteURL(fn2) = True Then
		FoundFile = True
		Exit Do ' The file is there
	End If
	DelayResponse(1) ' wait 1 sec & try again
	TimeoutCounter = TimeoutCounter + 1
Loop

If FoundFile <> True Then 
	Response.Write ("NO FILE FOUND")
	Response.End ' Could not fine the pdf, so just bail
End If

'*******************************************************************************************************************************
'CODE TO COPY PROPER MASTER PDF FOR SAVING
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'*******************************************************************************************************************************
'Now change the name of the file
Orig_Name = "\clientfiles\" & Left(MUV_Read("ClientID"),4) & "\z_pdfs\ConsolidatedStatement_" & Trim(UniqueNum) & "_Main.pdf" 
New_Name =  "\clientfiles\" & Left(MUV_Read("ClientID"),4) & "\z_pdfs\ConsolidatedStatement_Account_" & Trim(Account) & "_" & Trim(Account) & Trim(Replace(EndDate,"/","")) & ".pdf"

Set fso = CreateObject("Scripting.FileSystemObject")

'Kill it first in case an old one is there
On error resume next
fso.DeleteFile Server.MapPath(New_Name)
On error goto 0

fso.CopyFile Server.MapPath(Orig_Name), Server.MapPath(New_Name)

Set fso = Nothing

If DebugMessages = True Then Response.Write(New_Name)
'*******************************************************************************************************************************
'*******************************************************************************************************************************


'Now open the PDF in a new window
%>
<SCRIPT language='javascript'>window.open(' <%=fn2%> ');</SCRIPT>


<h1 class="page-header"><i class="fa fa-file-text"></i>  Detailed Consolidated Invoice (Account)</h1>
<!-- row !-->
<div class="row" style="margin-top:20px;">

    <!-- START !-->
    	
    	<p><a href="#" class="title">Save, Email or Post Detailed Consolidated Invoice</a></p>
    	
    	<div class="row" style="margin-top:15px;">
	    	
	    	<div class="col-md-6">
	    		Saves a copy of the consolidated invoice PDF (<strong>ConsolidatedStatement_Account_<%= Trim(Account) %>_<%= Trim(Account) & Trim(Replace(EndDate,"/","")) %>.pdf</strong>), which was just generated. Updates A/R in metroplex.
	    	</div>
	    	
	    	<div class="col-md-2 pull-left">
	    		<button type="button" class="btn btn-primary" id="btnSaveAndPostConsolidatedInvoice"><i class="fa fa-floppy-o" aria-hidden="true"></i>&nbsp;Save &amp; <i class="fa fa-upload" aria-hidden="true"></i>&nbsp;Post Invoice</button>
	    	</div>
    	
    	</div>
    	

    	<div class="row" style="margin-top:20px;">
	    	
	    	<div class="col-md-6">
	    		Saves a copy of the consolidated invoice PDF (<strong>ConsolidatedStatement_Account_<%= Trim(Account) %>_<%= Trim(Account) & Trim(Replace(EndDate,"/","")) %>.pdf</strong>), which was just generated. DOES NOT update A/R in metroplex.
	    	</div>
	    	
	    	<div class="col-md-2 pull-left">
	    		<button type="button" class="btn btn-primary" id="btnSaveConsolidatedInvoice"><i class="fa fa-floppy-o" aria-hidden="true"></i>&nbsp;Save Invoice Only</button>
	    	</div>
    	
    	</div>


    	<div class="row" style="margin-top:20px;">
	    	
	    	<div class="col-md-6">
	    		Emails a copy of the consolidated invoice PDF (<strong>ConsolidatedStatement_Account_<%= Trim(Account) %>_<%= Trim(Account) & Trim(Replace(EndDate,"/","")) %>.pdf</strong>), to users/emails that you specify. DOES NOT update A/R in metroplex.
	    	</div>
	    	
	    	<div class="col-md-2 pull-left">
	    		<button type="button" data-target="#consolidatedInvoiceEmailModal" data-toggle="modal" class="btn btn-primary"><i class="fa fa-envelope" aria-hidden="true"></i>&nbsp;Email Invoice</button>	    	
	    	</div>
    	
    	</div>

    <!-- END !-->


</div>
<!-- eof row !-->    


<!-- row !-->
<div class="row" style="margin-top:75px;">
    <!-- START !-->
   	<div class="col-md-6 reports-box">
        &nbsp;
    </div> 
	<div class="col-md-2 pull-left">
		<p align="right"><a href="consolidatedInvDetail.asp"><button type="button" class="btn btn-default">&lsaquo; Back To Consolidated Invoice Generator</button></a></p>
	</div>
    
    <!-- END !-->
</div>
<!-- eof row !-->    



<%
'*******************************************************************************************************************************************************************
'*******************************
' SUBs and FUNCTIONs Start Here
'*******************************
Sub DelayResponse(numberOfseconds)
 Dim WshShell
 Set WshShell=Server.CreateObject("WScript.Shell")
 WshShell.Run "waitfor /T " & numberOfSecond & "SignalThatWontHappen", , True
End Sub

Function CheckRemoteURL(fileURL)
    ON ERROR RESUME NEXT
    Dim xmlhttp

    Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")

    xmlhttp.open "GET", fileURL, False
    xmlhttp.send
    If(Err.Number<>0) then
        Response.Write "Could not connect to remote server"
    else
        Select Case Cint(xmlhttp.status)
            Case 200, 202, 302
                Set xmlhttp = Nothing
                CheckRemoteURL = True
            Case Else
                Set xmlhttp = Nothing
                CheckRemoteURL = False
        End Select
    end if
    ON ERROR GOTO 0
End Function

%>

<div class="modal fade" id="consolidatedInvoiceEmailModal" tabindex="-1" role="dialog" aria-labelledby="myEmailConsolidatedInvoiceModalAccountLabel">

	<div class="modal-dialog" role="document">
						
		<div class="modal-content">
	    
			<!-- modal header !-->
			<div class="modal-header">
				<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
				<h4 class="modal-title" id="myEmailConsolidatedInvoiceModalAccountLabel"></h4>
			</div>
			<!-- eof modal header !-->
	  
			<!-- modal body !-->
			<div class="modal-body">
			
				<input type="hidden" name="txtCustID" id="txtCustID" value="<%= Trim(Account) %>">
				<input type="hidden" name="txtCustName" id="txtCustName" value="<%= GetCustNameByCustNum(Trim(Account)) %>">
				<input type="hidden" name="txtStartDate" id="txtStartDate" value="<%= Trim(StartDate) %>">
				<input type="hidden" name="txtEndDate" id="txtEndDate" value="<%= Trim(EndDate) %>">
				<input type="hidden" name="txtConsInvNumber" id="txtConsInvNumber" value="<%= Trim(Account) %>_<%= Trim(Account) & Trim(Replace(EndDate,"/","")) %>">
				<input type="hidden" name="txtDueDateDays" id="txtDueDateDays" value="<%= DueDateDays %>">
				<input type="hidden" name="txtDueDateSingleDate" id="txtDueDateSingleDate" value="<%= DueDateSingleDate %>">
				<input type="hidden" name="txtPaidOrUnpaid" id="txtPaidOrUnpaid" value="PAID">
				
				<div id="EmailConsolidatedInvoiceModalAccountContent">
					<!-- Content for the modal will be generated and written here -->
					<!-- Content generated by Sub GetContentForEmailConsolidatedInvoiceModalAccount() in InsightFuncs_AjaxForARAP.asp -->
				</div>
				
			</div>

		</div>
		<!-- eof modal content !-->
</div>
<!-- eof modal dialog !-->
</div>
<!-- eof modal !-->


<!--#include file="../../../inc/footer-main.asp"-->
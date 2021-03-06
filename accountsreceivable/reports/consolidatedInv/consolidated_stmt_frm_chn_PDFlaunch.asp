<!--#include file="../../../inc/header-accounts-receivable.asp"-->

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
		    
		    var ChainID = $("#txtChainID").val();
		    var ChainName = $("#txtChainName").val();
		    var InvoiceNumber = $("#txtConsInvNumber").val();
		    var EndDate = $("#txtEndDate").val();
		    	    
		    var $modal = $(this);
	
    		$modal.find('#myEmailConsolidatedInvoiceModalChainLabel').html("Email Consolidated Invoice for Chain " + ChainID + " (" + ChainName + ") - Invoice #" + InvoiceNumber);
    		
	    	$.ajax({
				type:"POST",
				url: "../../../inc/InSightFuncs_AjaxForARAP.asp",
				cache: false,
				data: "action=GetContentForEmailConsolidatedInvoiceModalChain&consInvoiceNumber=" + encodeURIComponent(InvoiceNumber) + "&ChainID=" + encodeURIComponent(ChainID) + "&endDate=" + encodeURIComponent(EndDate) + "&paidOrUnpaid=PAID",
				success: function(response)
				 {
	               	 $modal.find('#emailConsolidatedInvoiceModalAccountContent').html(response);
	             }
	    	});
		    
		});
		

		

		$('#btnSaveConsolidatedInvoice').on('click', function(e) {
		
		    //get data-id attribute of the clicked alert
		    var chainID = $("#txtChainID").val();
		    var endDate = $("#txtEndDate").val();
		    					    		    		
	    	$.ajax({
				type:"POST",
				url:"consolidated_stmt_frm_chn_save_only.asp",
				data: "e=" + encodeURIComponent(endDate) + "&c=" + encodeURIComponent(chainID),
				success: function(response)
				 {
				 	swal("Consolidated Invoice Saved Successfully");
	             }
			});
    	});	
		

		$('#btnSaveAndPostConsolidatedInvoice').on('click', function(e) {
		
		    //get data-id attribute of the clicked alert
		    var consInvoiceNum = $("#txtConsInvNumber").val();
	    	var chainID = $("#txtChainID").val();
		    var endDate = $("#txtEndDate").val();
		    var dueDateDays = $("#txtDueDateDays").val();
			var dueDateSingle = $("#txtDueDateSingleDate").val();
							    		    		
	    	$.ajax({
				type:"POST",
				url:"consolidated_stmt_frm_chn_save_and_post.asp",
				data: "e=" + encodeURIComponent(endDate) + "&c=" + encodeURIComponent(chainID) + "&ddd=" + encodeURIComponent(dueDateDays) + "&dds=" + encodeURIComponent(dueDateSingle),
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

Server.ScriptTimeout = 600 ' Ten Minutes
MUV_Remove("ConStmt-StartDate") 

'baseURL should always have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)
sURL = Request.ServerVariables("SERVER_NAME")

DebugMessages = False ' Set to true to turn om Response.Writes

'Generate a unique number to be used for all pdfs throughout this page
Randomize
UniqueNum = int((9999999-1111111+1)*rnd+1111111)

StartDate = Request.QueryString("s")
EndDate = Request.QueryString("e")
Account = Request.QueryString("c")
IncludeIndividuals = Request.QueryString("ind")
StartDate = Replace(StartDate, "~","/")
EndDate = Replace(EndDate, "~","/")
SkipZeroDollar = Request.QueryString("z")
SkipLessThenZero = Request.QueryString("lz")
IncludedType = Request.QueryString("ty")
DueDateDays = Request.QueryString("ddd")
DueDateSingleDate = Request.QueryString("dds")
DoNotShowDueDate = Request.QueryString("dnsdd")


Set Pdf = Server.CreateObject("Persits.Pdf")
Set Doc = Pdf.CreateDocument

ImpVar = baseURL & "accountsreceivable/reports/consolidatedInv/consolidated_stmt_frm_chn_PDFgen.asp?s=" & Replace(StartDate,"/","~") & "&e=" & Replace(EndDate,"/","~")& "&c=" & Account
ImpVar = ImpVar & "&un=" & Session("UserNo")
ImpVar = ImpVar & "&z=" & SkipZeroDollar
ImpVar = ImpVar & "&lz=" & SkipLessThenZero
ImpVar = ImpVar & "&ty=" & IncludedType
ImpVar = ImpVar & "&cl=" & MUV_Read("ClientID")
ImpVar = ImpVar & "&u=" & MUV_Read("SQL_Owner")
ImpVar = ImpVar & "&ddd=" & DueDateDays
ImpVar = ImpVar & "&dds=" & DueDateSingleDate
ImpVar = ImpVar & "&dnsdd=" & DoNotShowDueDate

If DebugMessages = True Then Response.Write("<br><br><br><br>" & ImpVar & "<br>")

Doc.ImportFromUrl ImpVar, "scale=0.75; hyperlinks=false; drawbackground=true"

fn = "\clientfiles\" & Left(MUV_Read("ClientID"),4) & "\z_pdfs\ConsolidatedStatement_" & Trim(UniqueNum) & "_Main.pdf"
fn = Replace(fn,"/","-")
fn = Replace(fn,":","-")
'response.write(fn & "<br>")
fn2 = Left(baseURL,Len(baseURL)-1) & fn
fn2 = Replace(fn2,"\","/")
'response.write(fn2 & "<br>")
Main_PDF_Filename = fn
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

'**********************************************************
'**********************************************************
'**********************************************************
'Now do the individual invoices if that option is turned on
'**********************************************************
'**********************************************************
'**********************************************************
If IncludeIndividuals ="T" Then

	Set cnnIncludeIndividuals = Server.CreateObject("ADODB.Connection")
	cnnIncludeIndividuals.open Session("ClientCnnString")
	
	SQLIncludeIndividuals = "SELECT * FROM " & MUV_Read("SQL_Owner") & ".InvoiceHistory Where "
	SQLIncludeIndividuals = SQLIncludeIndividuals & "IvsHistSequence IN (Select IvsHistSequence from zReportConsolidatedInvoiceInclude_" & Trim(Session("UserNo")) & ") "
	SQLIncludeIndividuals = SQLIncludeIndividuals & " order by CustNum, IvsNum"
	 
	Set rsIncludeIndividuals = Server.CreateObject("ADODB.Recordset")
	rsIncludeIndividuals.CursorLocation = 3 
	Set rsIncludeIndividuals = cnnIncludeIndividuals.Execute(SQLIncludeIndividuals)
	
	If not rsIncludeIndividuals.eof then
	
		DocPartCounter = 1
		
		Do while not rsIncludeIndividuals.EOF
		
			Set Pdf_Individuals = Server.CreateObject("Persits.Pdf")
			Set Doc_Individuals = Pdf_Individuals.CreateDocument
			
			IvsSeq = rsIncludeIndividuals("IvsHistSequence")

			ImpVar = baseURL & "accountsreceivable/reports/consolidatedInv/consolidated_stmt_indv_color_PDFgen.asp?i=" & IvsSeq 
			ImpVar = ImpVar & "&un=" & Session("UserNo")
			ImpVar = ImpVar & "&cl=" & MUV_Read("ClientID")
			ImpVar = ImpVar & "&u=" & MUV_Read("SQL_Owner")

			If DebugMessages = True Then Response.Write(ImpVar & "<br>")

			Doc_Individuals.ImportFromUrl ImpVar, "scale=0.75; hyperlinks=false; drawbackground=true"
			
			fn = "\clientfiles\" & Left(MUV_Read("ClientID"),4) & "\z_pdfs\ConsStmt_" & Trim(UniqueNum) & "_Part" & Trim(DocPartCounter) & ".pdf"
			Individual_File_RootPart = "\clientfiles\" & Left(MUV_Read("ClientID"),4) & "\z_pdfs\ConsStmt_" & Trim(UniqueNum) & "_Part" 
			fn = Replace(fn,"/","-")
			fn = Replace(fn,":","-")
			'response.write(fn & "<br>")
			fn2 = Left(baseURL,Len(baseURL)-1) & fn
			fn2 = Replace(fn2,"\","/")
			'response.write(fn2 & "<br>")
			Filename = Doc_Individuals.Save(Server.MapPath(fn), False)

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

			DocPartCounter = DocPartCounter + 1
			
			Set Doc_Individuals = Nothing
			Set Pdf_Individuals = Nothing
			
			rsIncludeIndividuals.MoveNext
			
		Loop	

	End If
	
	set rsIncludeIndividuals = Nothing
	cnnIncludeIndividuals.Close
	set cnnIncludeIndividuals = Nothing
	
	
	'All the files are generated so now start stitching the file
	
	Set Pdf = Server.CreateObject("Persits.Pdf")
	Set Doc1 = Pdf.OpenDocument(Server.MapPath(Main_PDF_Filename))
	
	ArrayTop = cInt(DocPartCounter)
	ReDim objDocs(ArrayTop)
	
	For x = 1 to DocPartCounter - 1

		IndividualFile = Individual_File_RootPart & Trim(X) & ".pdf"
		
		' Open document 2
		Set objDocs(x) = Pdf.OpenDocument(Server.MapPath(IndividualFile))
		
		' Append doc2 to doc1
		Doc1.AppendDocument objDocs(x)
	
	Next
	
	' Save document, the Save method returns generated file name
	Filename = Doc1.Save(Server.MapPath(Main_PDF_Filename), False)

	Set Pdf = Nothing
	
	'Now change the name of the file
	Orig_Name = "\clientfiles\" & Left(MUV_Read("ClientID"),4) & "\z_pdfs\" & Filename 
	New_Name =  "\clientfiles\" & Left(MUV_Read("ClientID"),4) & "\z_pdfs\ConsolidatedStatement_Chain_" & Trim(Account) & ".pdf"
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	'Kill it first in case an old one is there
	On error resume next
	fso.DeleteFile Server.MapPath(New_Name)
	On error goto 0
	
	fso.MoveFile Server.MapPath(Orig_Name), Server.MapPath(New_Name)
	
	Set fso = Nothing
	
	If DebugMessages = True Then Response.Write(New_Name)
	
End If

'*******************************************************************************************************************************
'CODE TO COPY PROPER MASTER PDF FOR SAVING
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IF INCLUDE INDIVIDUALS IS TRUE, THEN THE MASTER PDF CREATED IS CALLED ConsolidatedStatement_Chain_ACCOUNT.pdf
' IF INCLUDE INDIVIDUALS IS FALSE, THEN THE MASTER PDF CREATED IS CALLED ConsolidatedStatement_UNIQUENUMBER_Main.pdf
' BECAUSE THE INDIVIDUALS WERE NEVER STITCHED TOGETHER INTO AN AGGREGATED MAIN FILE 
'*******************************************************************************************************************************
If IncludeIndividuals = "T" Then
	'Now change the name of the file
	Orig_Name = "\clientfiles\" & Left(MUV_Read("ClientID"),4) & "\z_pdfs\ConsolidatedStatement_Chain_" & Trim(Account) & ".pdf" 
	New_Name =  "\clientfiles\" & Left(MUV_Read("ClientID"),4) & "\z_pdfs\ConsolidatedStatement_Chain_" & Trim(Account) & "_" & Trim(Account) & Trim(Replace(EndDate,"/","")) & ".pdf"
Else
	'Now change the name of the file
	Orig_Name = "\clientfiles\" & Left(MUV_Read("ClientID"),4) & "\z_pdfs\ConsolidatedStatement_" & Trim(UniqueNum) & "_Main.pdf" 
	New_Name =  "\clientfiles\" & Left(MUV_Read("ClientID"),4) & "\z_pdfs\ConsolidatedStatement_Chain_" & Trim(Account) & "_" & Trim(Account) & Trim(Replace(EndDate,"/","")) & ".pdf"
End If

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
If IncludeIndividuals ="T" Then
	Response.Write("<SCRIPT language='javascript'>window.open('" & Replace(New_Name,"\","/") & "');</SCRIPT>")
Else
	Response.Write("<SCRIPT language='javascript'>window.open('" & fn2 & "');</SCRIPT>")
End If	

Response.Write("<script language=javascript>HideIt();</script>")
%>	
<h1 class="page-header"><i class="fa fa-file-text"></i>  Consolidated Invoice (Chain)</h1>
<!-- row !-->
<div class="row" style="margin-top:20px;">

    <!-- START !-->
    	
    	<p><a href="#" class="title">Save, Email or Post Consolidated Invoice</a></p>
    	
    	<div class="row" style="margin-top:15px;">
	    	
	    	<div class="col-md-6">
	    		Saves a copy of the consolidated invoice PDF (<strong>ConsolidatedStatement_Chain_<%= Trim(Account) %>_<%= Trim(Account) & Trim(Replace(EndDate,"/","")) %>.pdf</strong>), which was just generated. Updates A/R in metroplex.
	    	</div>
	    	
	    	<div class="col-md-2 pull-left">
	    		<button type="button" class="btn btn-primary" id="btnSaveAndPostConsolidatedInvoice"><i class="fa fa-floppy-o" aria-hidden="true"></i>&nbsp;Save &amp; <i class="fa fa-upload" aria-hidden="true"></i>&nbsp;Post Invoice</button>
	    	</div>
    	
    	</div>
    	

    	<div class="row" style="margin-top:20px;">
	    	
	    	<div class="col-md-6">
	    		Saves a copy of the consolidated invoice PDF (<strong>ConsolidatedStatement_Chain_<%= Trim(Account) %>_<%= Trim(Account) & Trim(Replace(EndDate,"/","")) %>.pdf</strong>), which was just generated. DOES NOT update A/R in metroplex.
	    	</div>
	    	
	    	<div class="col-md-2 pull-left">
	    		<button type="button" class="btn btn-primary" id="btnSaveConsolidatedInvoice"><i class="fa fa-floppy-o" aria-hidden="true"></i>&nbsp;Save Invoice Only</button>
	    	</div>
    	
    	</div>


    	<div class="row" style="margin-top:20px;">
	    	
	    	<div class="col-md-6">
	    		Emails a copy of the consolidated invoice PDF (<strong>ConsolidatedStatement_Chain_<%= Trim(Account) %>_<%= Trim(Account) & Trim(Replace(EndDate,"/","")) %>.pdf</strong>), to users/emails that you specify. DOES NOT update A/R in metroplex.
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
		<p align="right"><a href="consolidatedStatement.asp"><button type="button" class="btn btn-default">&lsaquo; Back To Consolidated Invoice Generator</button></a></p>
	</div>
    
    <!-- END !-->
</div>
<!-- eof row !-->   <%
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
				<h4 class="modal-title" id="myEmailConsolidatedInvoiceModalChainLabel"></h4>
			</div>
			<!-- eof modal header !-->
	  
			<!-- modal body !-->
			<div class="modal-body">
			
				<input type="hidden" name="txtChainID" id="txtChainID" value="<%= Trim(Account) %>">
				<input type="hidden" name="txtChainName" id="txtChainName" value="<%= GetChainDescByChainNum(Trim(Account)) %>">
				<input type="hidden" name="txtStartDate" id="txtStartDate" value="<%= Trim(StartDate) %>">
				<input type="hidden" name="txtEndDate" id="txtEndDate" value="<%= Trim(EndDate) %>">
				<input type="hidden" name="txtConsInvNumber" id="txtConsInvNumber" value="<%= Trim(Account) %>_<%= Trim(Account) & Trim(Replace(EndDate,"/","")) %>">
				<input type="hidden" name="txtDueDateDays" id="txtDueDateDays" value="<%= DueDateDays %>">
				<input type="hidden" name="txtDueDateSingleDate" id="txtDueDateSingleDate" value="<%= DueDateSingleDate %>">
				<input type="hidden" name="txtPaidOrUnpaid" id="txtPaidOrUnpaid" value="PAID">
				
				<div id="emailConsolidatedInvoiceModalAccountContent">
					<!-- Content for the modal will be generated and written here -->
					<!-- Content generated by Sub GetContentForEmailConsolidatedInvoiceModalChain() in InsightFuncs_AjaxForARAP.asp -->
				</div>
				
			</div>

		</div>
		<!-- eof modal content !-->
</div>
<!-- eof modal dialog !-->
</div>
<!-- eof modal !-->


<!--#include file="../../../inc/footer-main.asp"-->
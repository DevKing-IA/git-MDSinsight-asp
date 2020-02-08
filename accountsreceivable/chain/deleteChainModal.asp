<!--#include file="../../inc/InsightFuncs_AR_AP.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<% 

InternalRecordIdentifier = Request.QueryString("i")
'ClassCode = GetClassCodeByIntRecID(InternalRecordIdentifier)
NumAccountsByChainCode = NumberOfCustomersWithChain(InternalRecordIdentifier)
CustChainToBeDeletedDescription = GetCustChainByIntRecID(InternalRecordIdentifier)

%>
<style type="text/css">
.col-lg-12{
	margin-bottom:20px;
}

.modal-footer{
	margin-top:15px;
}
</style>


<div class="col-lg-12">
	<div class="modal-header">
		<button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
		<h4 class="modal-title">Set Chain to blank Before Deletion</h4>
	</div>
</div>

<form method="post" action="deleteChainFromModal.asp" name="frmDeleteChainFromModal" id="frmDeleteChainFromModal">

	<input type="hidden" name="txtChainCodeToBeDeletedIntRecID" id="txtChainCodeToBeDeletedIntRecID" value="<%=InternalRecordIdentifier %>">

	<div class="col-lg-12">
		<p>There are <%= NumAccountsByChainCode %> customers in the <strong><%= CustChainToBeDeletedDescription %></strong> Chain you are trying to delete. In order to delete this chain, you must first set the chain to blank for all the customers currently associated with this chain.</p>
		
		<p>Would you like to set the Chain to blank for these customers and continue with the deletion of the chain record?</p>
	</div>

	<div class="col-lg-12">
		<div class="modal-footer">
			<a href="main.asp"><button type="button" class="btn btn-default">Cancel Deletion</button></a>
			<button type="submit" class="btn btn-primary">Blank Chain & Delete</button>
		</div>
	</div>
</form>
<!--#include file="../../../inc/subsandfuncs.asp"-->
<!--#include file="../../../inc/Insightfuncs.asp"-->
<%
	
	'baseURL should always have a trailing /slash, just in case, handle either way
	If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)
	sURL = Request.ServerVariables("SERVER_NAME")
	
	'**************************************************************************************************
	'CODE TO SAVE CONSOLIDATED INVOICE TO CLIENTFILES\CLIENTID\CUSTOMER\ACCOUNTSRECEIVABLE
	'**************************************************************************************************
	
	EndDate = Request.Form("e")
	EndDate = Replace(EndDate, "~","/")
	ChainID = Request.Form("c")
	
	Orig_Name =  "\clientfiles\" & Left(MUV_Read("ClientID"),4) & "\z_pdfs\ConsolidatedStatement_Chain_" & Trim(ChainID) & "_" & Trim(ChainID) & Trim(Replace(EndDate,"/","")) & ".pdf"
	New_Name =  "\clientfiles\" & Left(MUV_Read("ClientID"),4) & "\customer\accountsreceivable\ConsolidatedStatement_Chain_" & Trim(ChainID) & "_" & Trim(ChainID) & Trim(Replace(EndDate,"/","")) & ".pdf"
	
	Response.Write("Orig_Name " & Orig_Name & "<br>")
	Response.Write("New_Name " & New_Name & "<br>")
	
	Dim fso
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	'Kill it first in case an old one is there
	On error resume next
	fso.DeleteFile Server.MapPath(New_Name)
	On error goto 0
	
	fso.CopyFile Server.MapPath(Orig_Name), Server.MapPath(New_Name)
	
	Set fso = Nothing
	
	'**************************************************************************************************
	'END CODE TO SAVE CONSOLIDATED INVOICE TO CLIENTFILES\CLIENTID\CUSTOMER\ACCOUNTSRECEIVABLE
	'**************************************************************************************************
%>
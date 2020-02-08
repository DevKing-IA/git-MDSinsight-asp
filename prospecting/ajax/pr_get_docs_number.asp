<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<!--#include file="../../inc/InSightFuncs_Prospecting.asp"-->
<% 
If Session("Userno") = "" Then Response.End() 

ProspectIntRecID = Request.QueryString("i") 
If ProspectIntRecID = "" Then Response.End()

Response.Write("(" & NumberOfDocumentsByProspectNumber(ProspectIntRecID) & ")")
%> 

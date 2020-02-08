<!--#include file="../../../inc/InSightFuncs.asp"-->
<!--#include file="../../../inc/InSightFuncs_Users.asp"-->
<!--#include file="../../../inc/InSightFuncs_Equipment.asp"-->
<% 
If Session("Userno") = "" Then Response.End() 

ModelIntRecID = Request.QueryString("i") 
If ModelIntRecID = "" Then Response.End()

Response.Write("(" & NumberOfLinksByModelIntRecID(ModelIntRecID) & ")")
%> 

<!--#include file="../inc/InSightFuncs.asp"-->
<!--#include file="../inc/InSightFuncs_Users.asp"-->
<%

SelectedMemoNumber = Request.Form("txtTicketNumber")

If SelectedMemoNumber = "" Then
	SelectedMemoNumber = Request.QueryString("t")
End If

SourceTab = Request.Form("txtReturnTab")

If SourceTab = "" Then
	SourceTab = Request.QueryString("tab")
End If


If SourceTab = "unacknowledged" Then
	Response.redirect("main_UnacknowledgedTickets.asp")
ElseIf SourceTab = "open" Then
	Response.redirect("main_OpenTickets.asp")
Else
	Response.redirect("main_ClosedRedoTickets.asp")
End If


%>

<!--#include file="SubsAndFuncs.asp"-->

<%

Dim action 
action = Request("action")

Select Case action
	Case "selectAccount"
		selectAccount()
	Case "selectAccountFS"
		selectAccountFS()
	Case "selectAccountFSNewMemo"
		selectAccountFSNewMemo()		
	Case "selectAssetFSClose"
		selectAssetFSClose()
	Case "selectAccount_AccountNotes"
		selectAccount_AccountNotes()
End Select 

Sub selectAccount_AccountNotes()

	custID = Request.Form("custID") 

	SQL = "SELECT * FROM " & Session("SQL_Owner") & ".AR_Customer WHERE CustNum='" & CustID & "'"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)

	If not rs.EOF then 
		Session("ServiceCustName") = rs("Name")
		Session("ServiceCustID") = rs("CustNum")
	Else
		Session("ServiceCustName") = ""
		Session("ServiceCustID") = ""
	End If
	
	rs.close
	cnn8.close

End Sub


Sub selectAccount()

	custID = Request.Form("custID") 

	SQL = "SELECT * FROM " &  Session("SQL_Owner") & ".AR_Customer WHERE CustNum='" & CustID & "'"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)

	If not rs.EOF then 
		Session("ServiceCustName") = rs("Name")
		Session("ServiceCustID") = rs("CustNum")
	Else
		Session("ServiceCustName") = ""
		Session("ServiceCustID") = ""
	End If
	
	rs.close
	cnn8.close

End Sub



Sub selectAccountFSNewMemo()

	custID = Request.Form("custID") 

	SQL = "SELECT * FROM " &  Session("SQL_Owner") & ".AR_Customer WHERE CustNum='" & CustID & "'"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)

	If not rs.EOF then 
		Session("ServiceCustName") = rs("Name")
		Session("ServiceCustID") = rs("CustNum")

	Else
		Session("ServiceCustName") = ""
		Session("ServiceCustID") = ""

	End If
	
	rs.close
	cnn8.close

End Sub



Sub selectAccountFS()

	MemoNumber = Request.Form("MemoNumber") 

	SQL = "SELECT * FROM " &  Session("SQL_Owner") & ".FS_ServiceMemos WHERE MemoNumber='" & MemoNumber & "'"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)

	If not rs.EOF then 
		Session("MemoNumber") = rs("MemoNumber")
	Else
		Session("MemoNumber") = ""
	End If
	
	rs.close
	cnn8.close

End Sub


%>
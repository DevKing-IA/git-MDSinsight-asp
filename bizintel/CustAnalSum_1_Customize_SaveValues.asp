<%
FilterSlsmn1 = Request.Form("selFilterSlsmn1")
FilterSlsmn2 = Request.Form("selFilterSlsmn2")
FilterReferral = Request.Form("selFilterReferral")
FilterSalesDollars = Request.Form("selDollars")
FilterPercentage = Request.Form("selPercent")

SQL = "SELECT * from Settings_Reports where ReportNumber = 2100 AND UserNo = " & Session("userNo")

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs= cnn8.Execute(SQL)

'Rec does not exist yet, make it quick but empty, update it later
If rs.EOF Then
	SQL = "Insert into Settings_Reports (ReportNumber, UserNo) Values (2100 , " & Session("userNo") & ")"
	rs.Close
	Set rs= cnn8.Execute(SQL)
End If

'Now update the table with the values
SQL = "Update Settings_Reports Set ReportSpecificData1 = '" & FilterSlsmn1 & "', "
SQL = SQL & "ReportSpecificData2 = '" & FilterSlsmn2 & "', " 
SQL = SQL & "ReportSpecificData3 = '" & FilterReferral & "', " 
SQL = SQL & "ReportSpecificData5 = '" & FilterSalesDollars & "', " 
SQL = SQL & "ReportSpecificData6 = '" & FilterPercentage & "'"
SQL = SQL & " WHERE ReportNumber = 2100 AND UserNo = " & Session("userNo")
Set rs= cnn8.Execute(SQL)

cnn8.Close

Set rs = Nothing
Set cnn8 = Nothing

Response.Redirect ("CustAnalSum_1.asp")
%>

 

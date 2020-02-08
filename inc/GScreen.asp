<!-- custom !-->
<style type="text/css">
	.period-end{
		width:8%;
	}
	</style>
<!-- eof custom !-->

<%

GScreenTopPeriodSeq=GetLastClosedPeriodSeqNum()


Response.Write("<div class='col-lg-9'><a class='btn btn-primary' role='button' data-toggle='collapse' href='#collapseExample' aria-expanded='false' aria-controls='collapseExample'>Toggle Period Sales Screen</a> <div class='collapse' id='collapseExample'><div class='well'><div class='table-striped table-condensed table-hover table-responsive'><table class='table table-striped table-condensed table-hover large-table' >")
Response.Write("<tr>")
Response.Write("<th scope='col' class='period-end'>Period End</th>")
Response.Write("<th scope='col'>Period</th>")
Response.Write("<th scope='col'>Period Sales</th>")


'Figure out the group names so we can get the column heading
SQL = "SELECT GroupName, MIN(SortOrder) AS Expr1 FROM Settings_CatGroups "
SQL = SQL & "WHERE (GroupName IN (SELECT DISTINCT GroupName FROM Settings_CatGroups AS Settings_CatGroups_1)) AND (ShowOnGScreen = 1) "
SQL = SQL & "GROUP BY GroupName "
SQL = SQL & "ORDER BY expr1 "
'response.write(SQL)
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs= cnn8.Execute(SQL)

If not rs.eof Then
	Do
		Response.Write("<th scope='col'>" & rs("GroupName") & "</th>")
		rs.movenext
	loop until rs.eof
End If
Response.Write("</tr>")


'Now do the Rows
Set rs= cnn8.Execute(SQL)

'Response.Write("HEREHEREHEREHERE" & GScreenTopPeriodSeq & "<br><br>")

For x = GScreenTopPeriodSeq to (GScreenTopPeriodSeq - 15) Step -1
	rs.movefirst
	Response.Write("<tr>")
	Response.Write("<td class='period-end'>" & GetPeriodEndDateBySeq(x) & "</td>")
	Response.Write("<td>" & Left(GetPeriodAndYearBySeq(x),Instr(GetPeriodAndYearBySeq(x),"-")-2) & "</td>")

	'Now get the totals for each of the category groups

	Set rsGScreen = Server.CreateObject("ADODB.Recordset")
	Set rsCatNums = Server.CreateObject("ADODB.Recordset")
	SQL_G="Select Sum(TotalSales) as PeriodTotSales from CustCatPeriodSales where ThisPeriodSequenceNumber = " & x & " AND CustNum = " & Session("GScreenCust")
	Set rsGScreen= cnn8.Execute(SQL_G)
	'Response.write(SQL_G)
	ZeroHolder = 0 ' duh
	If IsNull(rsGScreen("PeriodTotSales")) Then
		Response.Write("<td>" & FormatCurrency(ZeroHolder) & "</td>")
	Else
		Response.Write("<td>" & FormatCurrency(rsGScreen("PeriodTotSales")) & "</td>")
	End if

	Do
		'Get the lsit of category numbers that natch the name we are working on
		CatNumList=" AND (Category = "
		SQL_CatNums="Select * from Settings_CatGroups where GroupName = '" & rs("GroupName") & "'"
		Set rsCatNums = cnn8.Execute(SQL_CatNums)
		Do
			CatNumList = CatNumList & rsCatNums("Category") & " OR Category="
			rsCatNums.movenext
		Loop until rsCatNums.Eof					
		'Strip the last OR
		CatNumList = Left(CatNumList, len(CatNumList) - 12)
		CatNumList = CatNumList & ")"		
			
		SQL_G="Select Sum(TotalSales) as GroupTot from CustCatPeriodSales where (ThisPeriodSequenceNumber = " & x & " AND CustNum = " & Session("GScreenCust") & ") "
		SQL_G = SQL_G & CatNumList
		'Response.Write(SQL_G)
		Set rsGScreen= cnn8.Execute(SQL_G)
		If IsNull(rsGScreen("GroupTot")) Then GroupTot = 0 Else GroupTot = rsGScreen("GroupTot")
		Response.Write("<td>" & FormatCurrency(GroupTot) & "</td>")
		
		rs.movenext
	Loop until rs.eof	
		
	Response.Write("</tr>")
Next 

Response.Write("</table></div></div></div></div>")

%>

<%

	Set cnnCustCatPeriodSales = Server.CreateObject("ADODB.Connection")
	cnnCustCatPeriodSales.open (Session("ClientCnnString"))
	Set rsCustCatPeriodSales = Server.CreateObject("ADODB.Recordset")
	rsCustCatPeriodSales.CursorLocation = 3 

	
	
	'Drop fields that had the old names
	SQL_CustCatPeriodSales = "SELECT COL_LENGTH('CustCatPeriodSales', '3PriorPeriodsAeverage_Adjusted') AS IsItThere"
	Set rsCustCatPeriodSales = cnnCustCatPeriodSales.Execute(SQL_CustCatPeriodSales)
	If NOT IsNull(rsCustCatPeriodSales("IsItThere")) Then
		SQL_CustCatPeriodSales = "ALTER TABLE CustCatPeriodSales DROP COLUMN 3PriorPeriodsAeverage_Adjusted"
		Set rsCustCatPeriodSales = cnnCustCatPeriodSales.Execute(SQL_CustCatPeriodSales)
	End If

	SQL_CustCatPeriodSales = "SELECT COL_LENGTH('CustCatPeriodSales', 'DiifThisPeriodVSLast3Dollars_Adjusted') AS IsItThere"
	Set rsCustCatPeriodSales = cnnCustCatPeriodSales.Execute(SQL_CustCatPeriodSales)
	If NOT IsNull(rsCustCatPeriodSales("IsItThere")) Then
		SQL_CustCatPeriodSales = "ALTER TABLE CustCatPeriodSales DROP COLUMN DiifThisPeriodVSLast3Dollars_Adjusted"
		Set rsCustCatPeriodSales = cnnCustCatPeriodSales.Execute(SQL_CustCatPeriodSales)
	End If

	SQL_CustCatPeriodSales = "SELECT COL_LENGTH('CustCatPeriodSales', 'DiifThisPeriodVSLast3Percent_Adjusted') AS IsItThere"
	Set rsCustCatPeriodSales = cnnCustCatPeriodSales.Execute(SQL_CustCatPeriodSales)
	If NOT IsNull(rsCustCatPeriodSales("IsItThere")) Then
		SQL_CustCatPeriodSales = "ALTER TABLE CustCatPeriodSales DROP COLUMN DiifThisPeriodVSLast3Percent_Adjusted"
		Set rsCustCatPeriodSales = cnnCustCatPeriodSales.Execute(SQL_CustCatPeriodSales)
	End If

	SQL_CustCatPeriodSales = "SELECT COL_LENGTH('CustCatPeriodSales', 'DiifThisPeriodVSLast12Dollars_Adjusted') AS IsItThere"
	Set rsCustCatPeriodSales = cnnCustCatPeriodSales.Execute(SQL_CustCatPeriodSales)
	If NOT IsNull(rsCustCatPeriodSales("IsItThere")) Then
		SQL_CustCatPeriodSales = "ALTER TABLE CustCatPeriodSales DROP COLUMN DiifThisPeriodVSLast12Dollars_Adjusted"
		Set rsCustCatPeriodSales = cnnCustCatPeriodSales.Execute(SQL_CustCatPeriodSales)
	End If

	SQL_CustCatPeriodSales = "SELECT COL_LENGTH('CustCatPeriodSales', 'DiifThisPeriodVSLast12Percent_Adjusted') AS IsItThere"
	Set rsCustCatPeriodSales = cnnCustCatPeriodSales.Execute(SQL_CustCatPeriodSales)
	If NOT IsNull(rsCustCatPeriodSales("IsItThere")) Then
		SQL_CustCatPeriodSales = "ALTER TABLE CustCatPeriodSales DROP COLUMN DiifThisPeriodVSLast12Percent_Adjusted"
		Set rsCustCatPeriodSales = cnnCustCatPeriodSales.Execute(SQL_CustCatPeriodSales)
	End If


	set rsCustCatPeriodSales = nothing
	cnnCustCatPeriodSales.close
	set cnnCustCatPeriodSales = nothing
			
%>
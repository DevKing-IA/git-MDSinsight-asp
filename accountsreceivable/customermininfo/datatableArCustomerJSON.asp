
<% Server.ScriptTimeout = 360 %>
 
<!--#include file="../../inc/InsightFuncs_AR_AP.asp"--> 
<!--#include file="../../inc/InSightFuncs.asp"--> 
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<%		
	DIM JSONdata
	JSONdata=""
	IF Request.QueryString("length")="-1" THEN
		PageSize=0
		ELSE
		
			PageSize=CLng(Request.QueryString("length"))
	END IF
	IF Request.QueryString("start")="NaN" THEN
		rowStart=1
		ELSE
			rowStart=CLng(Request.QueryString("start"))
	END IF
	IF PageSize=0 THEN
		nPage=1
		ELSE
			nPage=1+rowStart/PageSize
			
	END IF
	
	searchValue=Request.QueryString("search[value]")
	orderValue=Request.QueryString("order[0][column]")
	orderType=Request.QueryString("order[0][dir]")
	
	SQL = "SELECT * FROM AR_Customer "
	
	IF Request.QueryString("maxID")="1" THEN
		SQL = SQL & " WHERE AcctStatus = 'A'"
	ELSE
		SQL = SQL & " WHERE AcctStatus = 'I'"
	END IF
	
	IF LEN(searchValue) > 0 THEN
	
		SQL = SQL & " AND (AR_Customer.CustNum LIKE '%" & searchValue & "%' OR " _
		
		& " AR_Customer.LastPriceChangeDate LIKE '%" & searchValue & "%' OR " _
		& " AR_Customer.Name LIKE '%" & searchValue & "%' OR " _
		& " AR_Customer.City LIKE '%" & searchValue & "%' OR " _
		& " AR_Customer.State LIKE '%" & searchValue & "%' OR " _
		& " AR_Customer.Zip LIKE '%" & searchValue & "%')"
		
	END IF
	
	SELECT CASE orderValue
		CASE "0"
			SQL = SQL & " ORDER BY AR_Customer.CustNum " & orderType
		CASE "1"
			SQL = SQL & " ORDER BY AR_Customer.Name " & orderType
		CASE "2"
			SQL= SQL & " ORDER BY AR_Customer.City " & orderType	
		CASE "3"
			SQL= SQL & " ORDER BY AR_Customer.State " & orderType
		CASE "4"
			SQL= SQL & " ORDER BY AR_Customer.Zip " & orderType	
		CASE "5"
			SQL= SQL & " ORDER BY AR_Customer.LastPriceChangeDate " & orderType	
		CASE ELSE	
			SQL = SQL & " ORDER BY AR_Customer.CustNum ASC"				
	END SELECT
	
	Set cnnARCustomers = Server.CreateObject("ADODB.Connection")
	cnnARCustomers.open (Session("ClientCnnString"))
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	
	'Response.Write(SQL)
	
	rs.open SQL,cnnARCustomers,1
	IF PageSize=0 THEN
		rs.PageSize=rs.recordCount
		
		ELSE
			rs.PageSize = PageSize
	END IF
	nPageCount = rs.PageCount
	nRecordCount=rs.recordCount
	
	If Not rs.Eof Then
		rs.AbsolutePage = nPage
		Do While Not ( rs.Eof Or rs.AbsolutePage <> nPage )
		
			IF LEN(JSONdata)>0 THEN
				JSONdata=JSONdata & ","
			END IF
			
			CustomerID = rs("CustNum")
			CustomerName = rs("Name")
			CustomerCity = rs("City")
			CustomerState = rs("State")
			CustomerZip = rs("Zip")
			CustomerLastPriceChangeDate  = rs("LastPriceChangeDate")
			
			If IsNull(CustomerName) OR IsEmpty(CustomerName) Then CustomerName = " "
			If IsNull(CustomerCity) OR IsEmpty(CustomerCity) Then CustomerCity = " "
			If IsNull(CustomerState) OR IsEmpty(CustomerState) Then CustomerState = " "
			If IsNull(CustomerZip) OR IsEmpty(CustomerZip) Then CustomerZip = " "
			If IsNull(CustomerLastPriceChangeDate) OR IsEmpty(CustomerLastPriceChangeDate) Then CustomerLastPriceChangeDate = " "
			
			JSONdata=JSONdata & "{"
			
			JSONdata=JSONdata & """id"":""" & CustomerID & """"
			JSONdata=JSONdata & ","
			JSONdata=JSONdata & """CustName"":""" & removeUnusualForJSON(CustomerName) & """"
			JSONdata=JSONdata & ","
			JSONdata=JSONdata & """City"":""" & removeUnusualForJSON(CustomerCity) & """"
			JSONdata=JSONdata & ","
			JSONdata=JSONdata & """State"":""" & removeUnusualForJSON(CustomerState) & """"
			JSONdata=JSONdata & ","
			JSONdata=JSONdata & """Zip"":""" & removeUnusualForJSON(CustomerZip) & """"
			JSONdata=JSONdata & ","	
			JSONdata=JSONdata & """LastPriceChangeDate"":""" & removeUnusualForJSON(CustomerLastPriceChangeDate) & """"
			JSONdata=JSONdata & ","				
			JSONdata=JSONdata & """Action"":"""""

			JSONdata=JSONdata & "}"		
			
			rs.movenext
				
		Loop
		
			
End If

retData="{""orderby"":""" & orderValue & """,""draw"": " & CLng(Request.QueryString("draw")) & ",""recordsTotal"": " & nRecordCount & ",""recordsFiltered"": " & nRecordCount & ",""data"": [" & JSONdata & "]}"

 
  
Response.AddHeader "Content-Type", "application/json"
Response.Write retData

function removeUnusualForJSON(value)

removeUnusualForJSON=REPLACE(value,"""","&quot;")
	
END FUNCTION
%>


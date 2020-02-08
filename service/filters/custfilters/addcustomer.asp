<!--#include file="json.asp"-->


<%

jsonstring=Request("data")

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))



				


Set oJSON = New aspJSON


oJSON.loadJSON(jsonstring)
DIM SQL
DIM customerID

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

customerID= oJSON.data("customerID")
SQl="DELETE FROM FS_CustomerFilters WHERE CustID='" & customerID & "'"
cnn8.Execute(SQL)

For Each data In oJSON.data("datatodb")
    Set this = oJSON.data("datatodb").item(data)
	
	selectedDate = this.item("LastChangedData")
	
	filterData = this.item("FilterData")
	location = this.item("location")
	FrequencyTime = this.item("FrequencyTime")
	FrequencyType = this.item("FrequencyType")
	Price = this.item("Price")
	IF LEN(Price)=0 THEN
		Price="0"
	END IF
	qty = this.item("qty")
    SQL="INSERT INTO FS_CustomerFilters (CustID,FilterIntRecID,FrequencyType,FrequencyTime,Price,notes,LastChangeDateTime,qty) VALUES('" & customerID & "'," & filterData & ",'" & FrequencyType & "','" & FrequencyTime & "'," & Price & ",'" & location & "','" & selectedDate & "'," & qty & ")"

	Set rs = Server.CreateObject("ADODB.Recordset")

	Set rs= cnn8.Execute(SQL)
Next
cnn8.Close()

response.write("<br>")
response.write(jsonstring)


%>

<%
'********************************
'List of all the functions & subs
'********************************
'Func NumberOfAPIOrderLines (passedOrderId,passedThread)
'Func GetAPIOrderHighestThread (passedOrderId)
'Func GetAPIOrderLastUpdatedDateTime (passedOrderID)
'Func GetAPIOrderFirstReceivedDateTime (passedOrderID)
'Func GetAPIRepostURL()
'Func APIRepostOrders()
'************************************
'End List of all the functions & subs
'************************************


Function NumberOfAPIOrderLines (passedOrderID, passedOrderThread)

	resultNumberOfAPIOrderLines = ""

	Set cnnNumberOfAPIOrderLines = Server.CreateObject("ADODB.Connection")
	cnnNumberOfAPIOrderLines.open Session("ClientCnnString")

	Set rsNumberOfAPIOrderLines = Server.CreateObject("ADODB.Recordset")		
	rsNumberOfAPIOrderLines.CursorLocation = 3 
	
	'Get header rec id
	SQLNumberOfAPIOrderLines = "SELECT InternalRecordIdentifier FROM API_OR_OrderHeader Where OrderID = '" & passedOrderID & "' AND OrderThread = " & passedOrderThread
	Set rsNumberOfAPIOrderLines = cnnNumberOfAPIOrderLines.Execute(SQLNumberOfAPIOrderLines)
	
	HeaderRecID = rsNumberOfAPIOrderLines("InternalRecordIdentifier") 
		
	SQLNumberOfAPIOrderLines = "Select Count(*) as Expr1 from API_OR_OrderDetail Where OrderHeaderRecID = " & HeaderRecID 
	Set rsNumberOfAPIOrderLines = cnnNumberOfAPIOrderLines.Execute(SQLNumberOfAPIOrderLines)
			 
	If not rsNumberOfAPIOrderLines.EOF Then resultNumberOfAPIOrderLines =  rsNumberOfAPIOrderLines("Expr1")
	
	set rsNumberOfAPIOrderLines= Nothing
	cnnNumberOfAPIOrderLines.Close	
	set cnnNumberOfAPIOrderLines= Nothing
	
	NumberOfAPIOrderLines = resultNumberOfAPIOrderLines
	
End Function


Function GetAPIOrderHighestThread (passedOrderID)

	resultGetAPIOrderHighestThread = ""

	Set cnnGetAPIOrderHighestThread = Server.CreateObject("ADODB.Connection")
	cnnGetAPIOrderHighestThread.open Session("ClientCnnString")
		
	SQLGetAPIOrderHighestThread = "Select Max(OrderThread) as Expr1 from API_OR_OrderHeader WHERE OrderID = '" & passedOrderID & "'"
 
	Set rsGetAPIOrderHighestThread = Server.CreateObject("ADODB.Recordset")
	rsGetAPIOrderHighestThread.CursorLocation = 3 
	Set rsGetAPIOrderHighestThread = cnnGetAPIOrderHighestThread.Execute(SQLGetAPIOrderHighestThread)
			 
	If not rsGetAPIOrderHighestThread.EOF Then resultGetAPIOrderHighestThread =  rsGetAPIOrderHighestThread("Expr1")
	
	set rsGetAPIOrderHighestThread= Nothing
	cnnGetAPIOrderHighestThread.Close	
	set cnnGetAPIOrderHighestThread= Nothing
	
	GetAPIOrderHighestThread = resultGetAPIOrderHighestThread
	
End Function

Function GetAPIOrderLastUpdatedDateTime (passedOrderID)

	resultGetAPIOrderLastUpdatedDateTime = ""

	Set cnnGetAPIOrderLastUpdatedDateTime = Server.CreateObject("ADODB.Connection")
	cnnGetAPIOrderLastUpdatedDateTime.open Session("ClientCnnString")
		
	SQLGetAPIOrderLastUpdatedDateTime = "Select Max(RecordCreationDateTime) as Expr1 from API_OR_OrderHeader WHERE OrderID = '" & passedOrderID & "'"
 
	Set rsGetAPIOrderLastUpdatedDateTime = Server.CreateObject("ADODB.Recordset")
	rsGetAPIOrderLastUpdatedDateTime.CursorLocation = 3 
	Set rsGetAPIOrderLastUpdatedDateTime = cnnGetAPIOrderLastUpdatedDateTime.Execute(SQLGetAPIOrderLastUpdatedDateTime)
			 
	If not rsGetAPIOrderLastUpdatedDateTime.EOF Then resultGetAPIOrderLastUpdatedDateTime =  rsGetAPIOrderLastUpdatedDateTime("Expr1")
	
	set rsGetAPIOrderLastUpdatedDateTime= Nothing
	cnnGetAPIOrderLastUpdatedDateTime.Close	
	set cnnGetAPIOrderLastUpdatedDateTime= Nothing
	
	GetAPIOrderLastUpdatedDateTime = resultGetAPIOrderLastUpdatedDateTime
	
End Function

Function GetAPIOrderFirstReceivedDateTime (passedOrderID)

	resultGetAPIOrderFirstReceivedDateTime = ""

	Set cnnGetAPIOrderFirstReceivedDateTime = Server.CreateObject("ADODB.Connection")
	cnnGetAPIOrderFirstReceivedDateTime.open Session("ClientCnnString")
		
	SQLGetAPIOrderFirstReceivedDateTime = "Select Min(RecordCreationDateTime) as Expr1 from API_OR_OrderHeader WHERE OrderID = '" & passedOrderID & "'"
 
	Set rsGetAPIOrderFirstReceivedDateTime = Server.CreateObject("ADODB.Recordset")
	rsGetAPIOrderFirstReceivedDateTime.CursorLocation = 3 
	Set rsGetAPIOrderFirstReceivedDateTime = cnnGetAPIOrderFirstReceivedDateTime.Execute(SQLGetAPIOrderFirstReceivedDateTime)
			 
	If not rsGetAPIOrderFirstReceivedDateTime.EOF Then resultGetAPIOrderFirstReceivedDateTime =  rsGetAPIOrderFirstReceivedDateTime("Expr1")
	
	set rsGetAPIOrderFirstReceivedDateTime= Nothing
	cnnGetAPIOrderFirstReceivedDateTime.Close	
	set cnnGetAPIOrderFirstReceivedDateTime= Nothing
	
	GetAPIOrderFirstReceivedDateTime = resultGetAPIOrderFirstReceivedDateTime
	
End Function

Function APIOrderIsVoided (passedOrderID)

	resultAPIOrderIsVoided = ""

	Set cnnAPIOrderIsVoided = Server.CreateObject("ADODB.Connection")
	cnnAPIOrderIsVoided.open Session("ClientCnnString")
		
	SQLAPIOrderIsVoided = "Select Top 1 * from API_OR_OrderHeader WHERE OrderID = '" & passedOrderID & "' ORDER BY RecordCreationDateTime DESC"
 
	Set rsAPIOrderIsVoided = Server.CreateObject("ADODB.Recordset")
	rsAPIOrderIsVoided.CursorLocation = 3 
	Set rsAPIOrderIsVoided = cnnAPIOrderIsVoided.Execute(SQLAPIOrderIsVoided)
			 
	If not rsAPIOrderIsVoided.EOF Then
		If rsAPIOrderIsVoided("Voided") = vbTrue Then resultAPIOrderIsVoided = True Else resultAPIOrderIsVoided = False
	End If
	
	set rsAPIOrderIsVoided= Nothing
	cnnAPIOrderIsVoided.Close	
	set cnnAPIOrderIsVoided= Nothing
	
	APIOrderIsVoided = resultAPIOrderIsVoided
	
End Function

Function GetAPIRepostURL()

	resultGetAPIRepostURL = ""

	Set cnnGetAPIRepostURL = Server.CreateObject("ADODB.Connection")
	cnnGetAPIRepostURL.open Session("ClientCnnString")
		
	SQLGetAPIRepostURL = "SELECT * FROM Settings_Global"
 
	Set rsGetAPIRepostURL = Server.CreateObject("ADODB.Recordset")
	rsGetAPIRepostURL.CursorLocation = 3 
	Set rsGetAPIRepostURL = cnnGetAPIRepostURL.Execute(SQLGetAPIRepostURL)
			 
	If not rsGetAPIRepostURL.EOF Then  
		If Not IsNull(rsGetAPIRepostURL("OrderAPIRepostURL")) Then resultGetAPIRepostURL = rsGetAPIRepostURL("OrderAPIRepostURL")
	End If
	
	set rsGetAPIRepostURL= Nothing
	cnnGetAPIRepostURL.Close	
	set cnnGetAPIRepostURL= Nothing
	
	GetAPIRepostURL = resultGetAPIRepostURL
	
End Function

Function APIRepostOrders()

	resultAPIRepostOrders = False ' To err on the side of caution

	Set cnnAPIRepostOrders = Server.CreateObject("ADODB.Connection")
	cnnAPIRepostOrders.open Session("ClientCnnString")
		
	SQLAPIRepostOrders = "SELECT * FROM Settings_Global"
 
	Set rsAPIRepostOrders = Server.CreateObject("ADODB.Recordset")
	rsAPIRepostOrders.CursorLocation = 3 
	Set rsAPIRepostOrders = cnnAPIRepostOrders.Execute(SQLAPIRepostOrders)
			 
	If not rsAPIRepostOrders.EOF Then  
		If IsNull(rsAPIRepostOrders("OrderAPIRepostONOFF")) Then
			resultAPIRepostOrders = False
		Else
			If rsAPIRepostOrders("OrderAPIRepostONOFF") = 1 Then
				resultAPIRepostOrders = True
			Else
				resultAPIRepostOrders = False
			End If
		End If
	End If
	
	set rsAPIRepostOrders= Nothing
	cnnAPIRepostOrders.Close	
	set cnnAPIRepostOrders= Nothing
	
	APIRepostOrders = resultAPIRepostOrders
	
End Function


%>


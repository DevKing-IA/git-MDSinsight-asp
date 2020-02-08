<%
'Func getUserCellNumberModal(passedUserNo)
'Func UserInList(UserToFind,UserList)
'Func MaxNumberOfDeliveries()
'Func GetDriverNameByTruckID(passedTruckID)
'Func GetTruckNumberByUser(passedUserNo)
'Func GetUserNumberByTruckNumber(passedTruckID)
'Func GetDeliveryStatusByCust(passedCustNum)
'Func GetDeliveryStatusByInvoice(passedIvsNum)
'Func GetTotalStopsByUserNo(passedUserNo)
'Func GetTotalPriorityStopsByUserNo(passedUserNo)
'Func GetTotalAMStopsByUserNo(passedUserNo)
'Func GetRemainingStopsByUserNo(passedUserNo)
'Func GetTotalStopsByTruckNumber(passedTruckID)
'Func GetRemainingStopsByTruckNumber(passedTruckID)
'Func GetRemainingPriorityStopsByUserNo(passedUserNo)
'Func CustHasANYPriorityDelivery(passedCustomer,passedUserNo)
'Func GetRemainingAMStopsByUserNo(passedUserNo)
'Func CustHasANYAMDelivery(passedCustomer,passedUserNo)
'Func CustHasANYDelivery(passedCustomer,passedUserNo)
'Func CustHasANYDeliveryByCustAndTruck(passedCustomer,passedTruck)
'Func GetNextCustomerStopByTruck(passedTruck)
'Func DeliveryAlertSet(passedInvoice,passedUserNo)
'Func DeliveryAlertCondition(passedInvoice,passedUserNo)
'Func GetCustNumberByInvoiceNumDelBoard(passedInvoiceNumber)
'Func GetCustNumberByInvoiceNumDelBoardHistory(passedInvoiceNumber)
'Func GetTruckByInvoiceNumDelBoard(passedInvoiceNumber)
'Func GetLastDeliveryStatusChangeBYInvoiceNumDelBoard(passedInvoiceNumber)
'Func GetNumberOfInvoicesByTruckNumber(passsedTruckNumber)
'Func GetNumberOfCustomersByTruckNumber(passsedTruckNumber)
'Func GetValueOfDeliveriesByTruckNumber(passsedTruckNumber)
'Func GetNumberOfInvoicesByTruckNumberHistorical(passedTruckNumber,passedDeliveryDate)
'Func GetNumberOfCustomersByTruckNumberHistorical(passedTruckNumber,passedDeliveryDate)
'Func GetValueOfDeliveriesByTruckNumberHistorical(passedTruckNumber,passedDeliveryDate)
'Func DelBoardHistMostRecentDate()
'Func GetNumberOutOfSequenceByTruckNumber(passsedTruckNumber)
'Func GetNumberOfInvoicesByTruckNumberHistorical(passsedTruckNumber)
'Func AutoPromptNextStopOn()
'Func AutoForceSelectNextStopON()
'Func GetLastInvoiceMarkedByTruckNumber(passedTruckNumber)
'Func DelBoardDontUseStopSequencing()
'Func DelBoardDontShowDeliveryLineItems()
'Func DelBoardIgnoreThisRoute(passedTruckNumber)
'Func DriverNumberHasNagAlerts(passedDriverUserNo)
'Func GetDriverCommentsByInvoiceNumber(passedInvoiceNumber)
'Func InvoiceIsNextStop(passedInvoiceNumber)
'Func GetLastInvoiceMarkedDATETIMEByTruckNumber(passedTruckNumber)
'Func GetNumberOfNagMessagesSent(passedUserno,passedNagType,passedDate)
'Func GetLastDeliveryStatusChangeBYTruck(passedTruckNumber)
'Func GetNumberOfNagMessagesSentSinceDateTime(passedUserno,passedNagType,passedDate)
'Func DriverInNagSkipTable (passedDriverUserNo,passedNagType)
'Func DeliveryInProgress(passedIvsNum)
'Func DeliveryInProgressByCust(passedCustID)
'Func DeliveryIsPriority(passedIvsNum)
'Func DeliveryIsAM(passedIvsNum)

Function getUserCellNumberModal(passedUserNo)

	result = ""
	
	Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.open Session("ClientCnnString")

	SQL = "Select * from tblUsers where UserNo = " & passedUserNo
	 
	Set rsBoost1 = Server.CreateObject("ADODB.Recordset")
	rsBoost1.CursorLocation = 3 
	Set rsBoost1= cnn.Execute(SQL)
	
	If not rsBoost1.eof then 
		If rsBoost1("userCellNumber") <> "" Then result = rsBoost1("userCellNumber")
	End IF	
	set rsBoost1= Nothing
	set cnn= Nothing
	
	getUserCellNumberModal= result

End Function


Function UserInList(UserToFind,UserList)

	result = False
	
	UserNoList = Split(UserList,",")
	For x = 0 To UBound(UserNoList)
		
		If UserNoList(x) <> "" AND UserNoList(x) <> "*Not Found*" AND UserToFind <> "" AND UserToFind <> "*Not Found*" Then
			If cint(UserToFind) = cint(UserNoList(x)) Then
				result = True
				Exit For
			End If
		End If
	Next
	
	UserInList = result
	
End Function


Function MaxNumberOfDeliveries()

	resultMaxNumberOfDeliveries = 0 
	RoutesToIgnore = ""
	
	SQLMaxNumberOfDeliveries = "Select * from Settings_Global"
	 
	Set cnnMaxNumberOfDeliveries = Server.CreateObject("ADODB.Connection")
	cnnMaxNumberOfDeliveries.open (Session("ClientCnnString"))
	Set rMaxNumberOfDeliveries = Server.CreateObject("ADODB.Recordset")
	rMaxNumberOfDeliveries.CursorLocation = 3 
	Set rMaxNumberOfDeliveries = cnnMaxNumberOfDeliveries.Execute(SQLMaxNumberOfDeliveries)
	
	RoutesToIgnore = rMaxNumberOfDeliveries("DelBoardRoutesToIgnore")
	RoutesToIgnore = Replace(RoutesToIgnore,",","','")
	
	If RoutesToIgnore = "" Then
		SQLMaxNumberOfDeliveries = "SELECT MAX(NumberOfDeliveries) AS Expr1 FROM "
		SQLMaxNumberOfDeliveries = SQLMaxNumberOfDeliveries & "(SELECT COUNT(IvsNum) AS NumberOfDeliveries "
		SQLMaxNumberOfDeliveries = SQLMaxNumberOfDeliveries & "FROM RT_DeliveryBoard GROUP BY TruckNumber) AS derivedtbl_1 "
	Else
		SQLMaxNumberOfDeliveries = "SELECT MAX(NumberOfDeliveries) AS Expr1 FROM "
		SQLMaxNumberOfDeliveries = SQLMaxNumberOfDeliveries & "(SELECT COUNT(IvsNum) AS NumberOfDeliveries "
		SQLMaxNumberOfDeliveries = SQLMaxNumberOfDeliveries & "FROM RT_DeliveryBoard WHERE TruckNumber NOT IN ('" & RoutesToIgnore & "') GROUP BY TruckNumber) AS derivedtbl_1 "
	End If


	Set rMaxNumberOfDeliveries = cnnMaxNumberOfDeliveries.Execute(SQLMaxNumberOfDeliveries)

	If not rMaxNumberOfDeliveries.EOF Then resultMaxNumberOfDeliveries = rMaxNumberOfDeliveries ("Expr1")
	
	cnnMaxNumberOfDeliveries.close
	set rMaxNumberOfDeliveries = nothing
	set cnnMaxNumberOfDeliveries= nothing	
	
	MaxNumberOfDeliveries = resultMaxNumberOfDeliveries 
	
End Function

Function GetDriverNameByTruckID(passedTruckID)

	Set cnnGetDriverNameByTruckNumber = Server.CreateObject("ADODB.Connection")
	cnnGetDriverNameByTruckNumber.open Session("ClientCnnString")

	resultGetDriverNameByTruckNumber="*Not Found*"
		
	SQLGetDriverNameByTruckNumber = "Select * from RT_Truck where TruckID = '" & passedTruckID & "'"
	 
	Set rsGetDriverNameByTruckNumber = Server.CreateObject("ADODB.Recordset")
	rsGetDriverNameByTruckNumber.CursorLocation = 3 
	Set rsGetDriverNameByTruckNumber= cnnGetDriverNameByTruckNumber.Execute(SQLGetDriverNameByTruckNumber)
		
	If not rsGetDriverNameByTruckNumber.eof then resultGetDriverNameByTruckNumber = rsGetDriverNameByTruckNumber("driverName")
	
	rsGetDriverNameByTruckNumber.Close
	set rsGetDriverNameByTruckNumber= Nothing
	cnnGetDriverNameByTruckNumber.Close	
	set cnnGetDriverNameByTruckNumber = Nothing
	
	GetDriverNameByTruckID = resultGetDriverNameByTruckNumber 
	
End Function




Function GetUserNumberByTruckNumber(passedTruckID)

	Set cnnGetDriverNameByTruckNumber = Server.CreateObject("ADODB.Connection")
	cnnGetDriverNameByTruckNumber.open Session("ClientCnnString")

	resultGetDriverNameByTruckNumber="*Not Found*"
		
	SQLGetDriverNameByTruckNumber = "Select * from tblUsers where userTruckNumber = '" & passedTruckID & "'"
	 
	Set rsGetDriverNameByTruckNumber = Server.CreateObject("ADODB.Recordset")
	rsGetDriverNameByTruckNumber.CursorLocation = 3 
	Set rsGetDriverNameByTruckNumber= cnnGetDriverNameByTruckNumber.Execute(SQLGetDriverNameByTruckNumber)
		
	If not rsGetDriverNameByTruckNumber.eof then resultGetDriverNameByTruckNumber = rsGetDriverNameByTruckNumber("UserNo")
	
	rsGetDriverNameByTruckNumber.Close
	set rsGetDriverNameByTruckNumber= Nothing
	cnnGetDriverNameByTruckNumber.Close	
	set cnnGetDriverNameByTruckNumber = Nothing
	
	GetUserNumberByTruckNumber = resultGetDriverNameByTruckNumber 
	
End Function


Function GetTruckNumberByUser(passedUserNo)

	Set cnnGetTruckNumberByUser = Server.CreateObject("ADODB.Connection")
	cnnGetTruckNumberByUser.open Session("ClientCnnString")

	resultGetTruckNumberByUser = 0
		
	SQLGetTruckNumberByUser = "Select * from tblUsers where UserNo = " & passedUserNo
	
	'Response.write("<br><br>" & SQLGetTruckNumberByUser & "<br><br>")
	 
	
	Set rsGetTruckNumberByUser = Server.CreateObject("ADODB.Recordset")
	rsGetTruckNumberByUser.CursorLocation = 3 
	Set rsGetTruckNumberByUser= cnnGetTruckNumberByUser.Execute(SQLGetTruckNumberByUser)
		
	If not rsGetTruckNumberByUser.eof then resultGetTruckNumberByUser = rsGetTruckNumberByUser("userTruckNumber")
	
	resultGetTruckNumberByUser = Trim(resultGetTruckNumberByUser )
	'If Len(resultGetTruckNumberByUser) = 1 Then
	'	resultGetTruckNumberByUser = "00" & resultGetTruckNumberByUser 
	'ElseIf Len(resultGetTruckNumberByUser) = 2 Then
	'	resultGetTruckNumberByUser = "0" & resultGetTruckNumberByUser 
	'End If
	
	rsGetTruckNumberByUser.Close
	set rsGetTruckNumberByUser= Nothing
	cnnGetTruckNumberByUser.Close	
	set cnnGetTruckNumberByUser = Nothing
	
	GetTruckNumberByUser = resultGetTruckNumberByUser 
	
End Function

Function GetDeliveryStatusByCust(passedCustNum)

	Set cnnGetDeliveryStatusByCust = Server.CreateObject("ADODB.Connection")
	cnnGetDeliveryStatusByCust.open Session("ClientCnnString")

	resultGetDeliveryStatusByCust=""
		
	SQLGetDeliveryStatusByCust = "Select * from RT_DeliveryBoard WHERE CustNum = '" & passedCustNum & "'"
	 
	Set rsGetDeliveryStatusByCust = Server.CreateObject("ADODB.Recordset")
	rsGetDeliveryStatusByCust.CursorLocation = 3 
	Set rsGetDeliveryStatusByCust= cnnGetDeliveryStatusByCust.Execute(SQLGetDeliveryStatusByCust)
	
	DelCount = 0
	NoDelCount = 0
	TotalCount = 0
			
	If not rsGetDeliveryStatusByCust.eof then 
		Do While Not rsGetDeliveryStatusByCust.eof
			TotalCount = TotalCount + 1
			If rsGetDeliveryStatusByCust("DeliveryStatus") = "Delivered" Then DelCount = DelCount + 1
			If rsGetDeliveryStatusByCust("DeliveryStatus") = "No Delivery" Then NoDelCount = NoDelCount + 1
			rsGetDeliveryStatusByCust.MoveNext
		Loop
	End If
	
	rsGetDeliveryStatusByCust.Close
	set rsGetDeliveryStatusByCust= Nothing
	cnnGetDeliveryStatusByCust.Close	
	set cnnGetDeliveryStatusByCust = Nothing

	'Now figure out what the status will be
	If DelCount = TotalCount Then 
		resultGetDeliveryStatusByCust = "Delivered"
	ElseIf NoDelCount = TotalCount Then
		resultGetDeliveryStatusByCust = "Not Delivered"
	ElseIf (DelCount + NoDelCount) <> 0 Then
		resultGetDeliveryStatusByCust = "Partial Delivery"
	End If
	
	GetDeliveryStatusByCust = resultGetDeliveryStatusByCust 
	
End Function

Function GetDeliveryStatusByInvoice(passedIvsNum)

	Set cnnGetDeliveryStatusByInvoice = Server.CreateObject("ADODB.Connection")
	cnnGetDeliveryStatusByInvoice.open Session("ClientCnnString")

	resultGetDeliveryStatusByInvoice=""
		
	SQLGetDeliveryStatusByInvoice = "Select * from RT_DeliveryBoard WHERE IvsNum = " & passedIvsNum
	 
	Set rsGetDeliveryStatusByInvoice = Server.CreateObject("ADODB.Recordset")
	rsGetDeliveryStatusByInvoice.CursorLocation = 3 
	Set rsGetDeliveryStatusByInvoice= cnnGetDeliveryStatusByInvoice.Execute(SQLGetDeliveryStatusByInvoice)
	
	If not rsGetDeliveryStatusByInvoice.eof then resultGetDeliveryStatusByInvoice =  rsGetDeliveryStatusByInvoice("DeliveryStatus")
		
	rsGetDeliveryStatusByInvoice.Close
	set rsGetDeliveryStatusByInvoice= Nothing
	cnnGetDeliveryStatusByInvoice.Close	
	set cnnGetDeliveryStatusByInvoice = Nothing
	
	GetDeliveryStatusByInvoice = resultGetDeliveryStatusByInvoice 
	
End Function

Function GetTotalStopsByUserNo(passedUserNo)

	Set cnnGetTotalStopsByUserNo = Server.CreateObject("ADODB.Connection")
	cnnGetTotalStopsByUserNo.open Session("ClientCnnString")

	resultGetTotalStopsByUserNo = 0 
		
	SQLGetTotalStopsByUserNo = "SELECT Count(Distinct CustNum) As Expr1 FROM RT_DeliveryBoard "
	SQLGetTotalStopsByUserNo = SQLGetTotalStopsByUserNo & "WHERE TruckNumber = '" & GetTruckNumberByUser(passedUserNo) & "'"
	 
	Set rsGetTotalStopsByUserNo = Server.CreateObject("ADODB.Recordset")
	rsGetTotalStopsByUserNo.CursorLocation = 3 
	Set rsGetTotalStopsByUserNo= cnnGetTotalStopsByUserNo.Execute(SQLGetTotalStopsByUserNo)
	
	If not rsGetTotalStopsByUserNo.eof then resultGetTotalStopsByUserNo =  rsGetTotalStopsByUserNo("Expr1")
		
	rsGetTotalStopsByUserNo.Close
	set rsGetTotalStopsByUserNo= Nothing
	cnnGetTotalStopsByUserNo.Close	
	set cnnGetTotalStopsByUserNo = Nothing
	
	GetTotalStopsByUserNo = resultGetTotalStopsByUserNo 
	
End Function

Function GetRemainingStopsByUserNo(passedUserNo)

	Set cnnGetRemainingStopsByUserNo = Server.CreateObject("ADODB.Connection")
	cnnGetRemainingStopsByUserNo.open Session("ClientCnnString")

	resultGetRemainingStopsByUserNo = 0 
		
	SQLGetRemainingStopsByUserNo = "SELECT Distinct CustNum FROM RT_DeliveryBoard WHERE TruckNumber = '" & GetTruckNumberByUser(passedUserNo) & "'"
 
	Set rsGetRemainingStopsByUserNo = Server.CreateObject("ADODB.Recordset")
	rsGetRemainingStopsByUserNo.CursorLocation = 3 
	Set rsGetRemainingStopsByUserNo= cnnGetRemainingStopsByUserNo.Execute(SQLGetRemainingStopsByUserNo)
	
	If not rsGetRemainingStopsByUserNo.eof then
		Do While Not rsGetRemainingStopsByUserNo.eof
			If CustHasANYDelivery(rsGetRemainingStopsByUserNo("CustNum"),passedUserNo) = False Then resultGetRemainingStopsByUserNo = resultGetRemainingStopsByUserNo + 1
	 		rsGetRemainingStopsByUserNo.Movenext
 		Loop
 	End If
		
	rsGetRemainingStopsByUserNo.Close
	set rsGetRemainingStopsByUserNo= Nothing
	cnnGetRemainingStopsByUserNo.Close	
	set cnnGetRemainingStopsByUserNo = Nothing
	
	GetRemainingStopsByUserNo = resultGetRemainingStopsByUserNo 
	
End Function


Function GetTotalAMStopsByUserNo(passedUserNo)

	Set cnnGetTotalAMStopsByUserNo = Server.CreateObject("ADODB.Connection")
	cnnGetTotalAMStopsByUserNo.open Session("ClientCnnString")

	resultGetTotalAMStopsByUserNo = 0 
		
	SQLGetTotalAMStopsByUserNo = "SELECT Count(Distinct CustNum) As Expr1 FROM RT_DeliveryBoard "
	SQLGetTotalAMStopsByUserNo = SQLGetTotalAMStopsByUserNo & "WHERE TruckNumber = '" & GetTruckNumberByUser(passedUserNo) & "' AND AMorPM = 'AM'"
	 
	Set rsGetTotalAMStopsByUserNo = Server.CreateObject("ADODB.Recordset")
	rsGetTotalAMStopsByUserNo.CursorLocation = 3 
	Set rsGetTotalAMStopsByUserNo= cnnGetTotalAMStopsByUserNo.Execute(SQLGetTotalAMStopsByUserNo)
	
	If not rsGetTotalAMStopsByUserNo.eof then resultGetTotalAMStopsByUserNo =  rsGetTotalAMStopsByUserNo("Expr1")
		
	rsGetTotalAMStopsByUserNo.Close
	set rsGetTotalAMStopsByUserNo= Nothing
	cnnGetTotalAMStopsByUserNo.Close	
	set cnnGetTotalAMStopsByUserNo = Nothing
	
	GetTotalAMStopsByUserNo = resultGetTotalAMStopsByUserNo 
	
End Function


Function GetTotalPriorityStopsByUserNo(passedUserNo)

	Set cnnGetTotalPriorityStopsByUserNo = Server.CreateObject("ADODB.Connection")
	cnnGetTotalPriorityStopsByUserNo.open Session("ClientCnnString")

	resultGetTotalPriorityStopsByUserNo = 0 
		
	SQLGetTotalPriorityStopsByUserNo = "SELECT Count(Distinct CustNum) As Expr1 FROM RT_DeliveryBoard "
	SQLGetTotalPriorityStopsByUserNo = SQLGetTotalPriorityStopsByUserNo & "WHERE TruckNumber = '" & GetTruckNumberByUser(passedUserNo) & "' AND Priority=1"
	 
	Set rsGetTotalPriorityStopsByUserNo = Server.CreateObject("ADODB.Recordset")
	rsGetTotalPriorityStopsByUserNo.CursorLocation = 3 
	Set rsGetTotalPriorityStopsByUserNo= cnnGetTotalPriorityStopsByUserNo.Execute(SQLGetTotalPriorityStopsByUserNo)
	
	If not rsGetTotalPriorityStopsByUserNo.eof then resultGetTotalPriorityStopsByUserNo =  rsGetTotalPriorityStopsByUserNo("Expr1")
		
	rsGetTotalPriorityStopsByUserNo.Close
	set rsGetTotalPriorityStopsByUserNo= Nothing
	cnnGetTotalPriorityStopsByUserNo.Close	
	set cnnGetTotalPriorityStopsByUserNo = Nothing
	
	GetTotalPriorityStopsByUserNo = resultGetTotalPriorityStopsByUserNo 
	
End Function


Function GetTotalStopsByTruckNumber(passedTruckID)

	Set cnnGetTotalStopsByTruckNumber = Server.CreateObject("ADODB.Connection")
	cnnGetTotalStopsByTruckNumber.open Session("ClientCnnString")

	resultGetTotalStopsByTruckNumber = 0 
		
	SQLGetTotalStopsByTruckNumber = "SELECT Count(Distinct CustNum) As Expr1 FROM RT_DeliveryBoard "
	SQLGetTotalStopsByTruckNumber = SQLGetTotalStopsByTruckNumber & "WHERE TruckNumber = '" & passedTruckID & "'"
	 
	Set rsGetTotalStopsByTruckNumber = Server.CreateObject("ADODB.Recordset")
	rsGetTotalStopsByTruckNumber.CursorLocation = 3 
	Set rsGetTotalStopsByTruckNumber= cnnGetTotalStopsByTruckNumber.Execute(SQLGetTotalStopsByTruckNumber)
	
	If not rsGetTotalStopsByTruckNumber.eof then resultGetTotalStopsByTruckNumber =  rsGetTotalStopsByTruckNumber("Expr1")
		
	rsGetTotalStopsByTruckNumber.Close
	set rsGetTotalStopsByTruckNumber= Nothing
	cnnGetTotalStopsByTruckNumber.Close	
	set cnnGetTotalStopsByTruckNumber = Nothing
	
	GetTotalStopsByTruckNumber = resultGetTotalStopsByTruckNumber 
	
End Function


Function GetRemainingStopsByTruckNumber(passedTruckID)

	Set cnnGetRemainingStopsByTruckNumber = Server.CreateObject("ADODB.Connection")
	cnnGetRemainingStopsByTruckNumber.open Session("ClientCnnString")

	resultGetRemainingStopsByTruckNumber = 0 
		
	SQLGetRemainingStopsByTruckNumber = "SELECT Distinct CustNum FROM RT_DeliveryBoard WHERE TruckNumber = '" & passedTruckID & "'"
 
	Set rsGetRemainingStopsByTruckNumber = Server.CreateObject("ADODB.Recordset")
	rsGetRemainingStopsByTruckNumber.CursorLocation = 3 
	Set rsGetRemainingStopsByTruckNumber= cnnGetRemainingStopsByTruckNumber.Execute(SQLGetRemainingStopsByTruckNumber)
	
	If not rsGetRemainingStopsByTruckNumber.eof then
		Do While Not rsGetRemainingStopsByTruckNumber.eof
			If CustHasANYDeliveryByCustAndTruck(rsGetRemainingStopsByTruckNumber("CustNum"),passedTruckID) = False Then resultGetRemainingStopsByTruckNumber = resultGetRemainingStopsByTruckNumber + 1
	 		rsGetRemainingStopsByTruckNumber.Movenext
 		Loop
 	End If
		
	rsGetRemainingStopsByTruckNumber.Close
	set rsGetRemainingStopsByTruckNumber= Nothing
	cnnGetRemainingStopsByTruckNumber.Close	
	set cnnGetRemainingStopsByTruckNumber = Nothing
	
	GetRemainingStopsByTruckNumber = resultGetRemainingStopsByTruckNumber 
	
End Function


Function GetRemainingPriorityStopsByUserNo(passedUserNo)

	Set cnnGetRemainingPriorityStopsByUserNo = Server.CreateObject("ADODB.Connection")
	cnnGetRemainingPriorityStopsByUserNo.open Session("ClientCnnString")

	resultGetRemainingPriorityStopsByUserNo = 0 
		
	SQLGetRemainingPriorityStopsByUserNo = "SELECT Distinct CustNum FROM RT_DeliveryBoard WHERE TruckNumber = '" & GetTruckNumberByUser(passedUserNo) & "' AND Priority=1"
 
	Set rsGetRemainingPriorityStopsByUserNo = Server.CreateObject("ADODB.Recordset")
	rsGetRemainingPriorityStopsByUserNo.CursorLocation = 3 
	Set rsGetRemainingPriorityStopsByUserNo= cnnGetRemainingPriorityStopsByUserNo.Execute(SQLGetRemainingPriorityStopsByUserNo)
	
	If not rsGetRemainingPriorityStopsByUserNo.eof then
		Do While Not rsGetRemainingPriorityStopsByUserNo.eof
			If CustHasANYPriorityDelivery(rsGetRemainingPriorityStopsByUserNo("CustNum"),passedUserNo) = False Then resultGetRemainingPriorityStopsByUserNo = resultGetRemainingPriorityStopsByUserNo + 1
	 		rsGetRemainingPriorityStopsByUserNo.Movenext
 		Loop
 	End If
		
	rsGetRemainingPriorityStopsByUserNo.Close
	set rsGetRemainingPriorityStopsByUserNo= Nothing
	cnnGetRemainingPriorityStopsByUserNo.Close	
	set cnnGetRemainingPriorityStopsByUserNo = Nothing
	
	GetRemainingPriorityStopsByUserNo = resultGetRemainingPriorityStopsByUserNo 
	
End Function


Function CustHasANYPriorityDelivery(passedCustomer,passedUserNo)

	Set cnnCustHasANYPriorityDelivery = Server.CreateObject("ADODB.Connection")
	cnnCustHasANYPriorityDelivery.open Session("ClientCnnString")

	resultCustHasANYPriorityDelivery = False
		
	SQLCustHasANYPriorityDelivery = "SELECT * FROM RT_DeliveryBoard WHERE TruckNumber = '" & GetTruckNumberByUser(passedUserNo) & "'"
	'SQLCustHasANYPriorityDelivery = SQLCustHasANYPriorityDelivery & " AND CustNum = '" & passedCustomer & "' AND Priority = 1 AND (DeliveryStatus IS NOT NULL OR DeliveryInProgress = 1)"
	SQLCustHasANYPriorityDelivery = SQLCustHasANYPriorityDelivery & " AND CustNum = '" & passedCustomer & "' AND Priority = 1"

 
	Set rsCustHasANYPriorityDelivery = Server.CreateObject("ADODB.Recordset")
	rsCustHasANYPriorityDelivery.CursorLocation = 3 
	Set rsCustHasANYPriorityDelivery= cnnCustHasANYPriorityDelivery.Execute(SQLCustHasANYPriorityDelivery)
	
	'Response.write(SQLCustHasANYPriorityDelivery)
	
	If not rsCustHasANYPriorityDelivery.eof then resultCustHasANYPriorityDelivery = True
		
	rsCustHasANYPriorityDelivery.Close
	set rsCustHasANYPriorityDelivery= Nothing
	cnnCustHasANYPriorityDelivery.Close	
	set cnnCustHasANYPriorityDelivery = Nothing
	
	CustHasANYPriorityDelivery = resultCustHasANYPriorityDelivery 
	
End Function


Function GetRemainingAMStopsByUserNo(passedUserNo)

	Set cnnGetRemainingAMStopsByUserNo = Server.CreateObject("ADODB.Connection")
	cnnGetRemainingAMStopsByUserNo.open Session("ClientCnnString")

	resultGetRemainingAMStopsByUserNo = 0 
		
	SQLGetRemainingAMStopsByUserNo = "SELECT Distinct CustNum FROM RT_DeliveryBoard WHERE TruckNumber = '" & GetTruckNumberByUser(passedUserNo) & "' AND AMorPM = 'AM'"
 
	Set rsGetRemainingAMStopsByUserNo = Server.CreateObject("ADODB.Recordset")
	rsGetRemainingAMStopsByUserNo.CursorLocation = 3 
	Set rsGetRemainingAMStopsByUserNo= cnnGetRemainingAMStopsByUserNo.Execute(SQLGetRemainingAMStopsByUserNo)
	
	If not rsGetRemainingAMStopsByUserNo.eof then
		Do While Not rsGetRemainingAMStopsByUserNo.eof
			If CustHasANYAMDelivery(rsGetRemainingAMStopsByUserNo("CustNum"),passedUserNo) = False Then resultGetRemainingAMStopsByUserNo = resultGetRemainingAMStopsByUserNo + 1
	 		rsGetRemainingAMStopsByUserNo.Movenext
 		Loop
 	End If
		
	rsGetRemainingAMStopsByUserNo.Close
	set rsGetRemainingAMStopsByUserNo= Nothing
	cnnGetRemainingAMStopsByUserNo.Close	
	set cnnGetRemainingAMStopsByUserNo = Nothing
	
	GetRemainingAMStopsByUserNo = resultGetRemainingAMStopsByUserNo 
	
End Function


Function CustHasANYAMDelivery(passedCustomer,passedUserNo)

	Set cnnCustHasANYAMDelivery = Server.CreateObject("ADODB.Connection")
	cnnCustHasANYAMDelivery.open Session("ClientCnnString")

	resultCustHasANYAMDelivery = False
		
	SQLCustHasANYAMDelivery = "SELECT * FROM RT_DeliveryBoard WHERE TruckNumber = '" & GetTruckNumberByUser(passedUserNo) & "'"
	'SQLCustHasANYAMDelivery = SQLCustHasANYAMDelivery & " AND CustNum = '" & passedCustomer & "' AND AMorPM = 'AM' AND (DeliveryStatus IS NOT NULL OR DeliveryInProgress = 1)"
	SQLCustHasANYAMDelivery = SQLCustHasANYAMDelivery & " AND CustNum = '" & passedCustomer & "' AND AMorPM = 'AM'"
 
	Set rsCustHasANYAMDelivery = Server.CreateObject("ADODB.Recordset")
	rsCustHasANYAMDelivery.CursorLocation = 3 
	Set rsCustHasANYAMDelivery= cnnCustHasANYAMDelivery.Execute(SQLCustHasANYAMDelivery)
	
	If not rsCustHasANYAMDelivery.eof then resultCustHasANYAMDelivery = True
		
	rsCustHasANYAMDelivery.Close
	set rsCustHasANYAMDelivery= Nothing
	cnnCustHasANYAMDelivery.Close	
	set cnnCustHasANYAMDelivery = Nothing
	
	CustHasANYAMDelivery = resultCustHasANYAMDelivery 
	
End Function


Function CustHasANYDelivery(passedCustomer,passedUserNo)

	Set cnnCustHasANYDelivery = Server.CreateObject("ADODB.Connection")
	cnnCustHasANYDelivery.open Session("ClientCnnString")

	resultCustHasANYDelivery = False
		
	SQLCustHasANYDelivery = "SELECT * FROM RT_DeliveryBoard WHERE TruckNumber = '" & GetTruckNumberByUser(passedUserNo) & "'"
	SQLCustHasANYDelivery = SQLCustHasANYDelivery & " AND CustNum = '" & passedCustomer & "' AND (DeliveryStatus IS NOT NULL or DeliveryInProgress = 1)"
 
	Set rsCustHasANYDelivery = Server.CreateObject("ADODB.Recordset")
	rsCustHasANYDelivery.CursorLocation = 3 
	Set rsCustHasANYDelivery= cnnCustHasANYDelivery.Execute(SQLCustHasANYDelivery)
	
	If not rsCustHasANYDelivery.eof then resultCustHasANYDelivery = True
		
	rsCustHasANYDelivery.Close
	set rsCustHasANYDelivery= Nothing
	cnnCustHasANYDelivery.Close	
	set cnnCustHasANYDelivery = Nothing
	
	CustHasANYDelivery = resultCustHasANYDelivery 
	
End Function


Function CustHasANYDeliveryByCustAndTruck(passedCustomer,passedTruck)

	Set cnnCustHasANYDeliveryByCustAndTruck = Server.CreateObject("ADODB.Connection")
	cnnCustHasANYDeliveryByCustAndTruck.open Session("ClientCnnString")

	resultCustHasANYDeliveryByCustAndTruck = False
		
	SQLCustHasANYDeliveryByCustAndTruck = "SELECT * FROM RT_DeliveryBoard WHERE TruckNumber = '" & passedTruck & "'"
	SQLCustHasANYDeliveryByCustAndTruck = SQLCustHasANYDeliveryByCustAndTruck & " AND CustNum = '" & passedCustomer & "' AND (DeliveryStatus IS NOT NULL OR DeliveryInProgress = 1)"
 
	Set rsCustHasANYDeliveryByCustAndTruck = Server.CreateObject("ADODB.Recordset")
	rsCustHasANYDeliveryByCustAndTruck.CursorLocation = 3 
	Set rsCustHasANYDeliveryByCustAndTruck= cnnCustHasANYDeliveryByCustAndTruck.Execute(SQLCustHasANYDeliveryByCustAndTruck)
	
	If not rsCustHasANYDeliveryByCustAndTruck.eof then resultCustHasANYDeliveryByCustAndTruck = True
		
	rsCustHasANYDeliveryByCustAndTruck.Close
	set rsCustHasANYDeliveryByCustAndTruck= Nothing
	cnnCustHasANYDeliveryByCustAndTruck.Close	
	set cnnCustHasANYDeliveryByCustAndTruck = Nothing
	
	CustHasANYDeliveryByCustAndTruck = resultCustHasANYDeliveryByCustAndTruck 
	
End Function

Function GetNextCustomerStopByTruck(passedTruck)


	Set cnnGetNextCustomerStopByTruck = Server.CreateObject("ADODB.Connection")
	cnnGetNextCustomerStopByTruck.open Session("ClientCnnString")
	Set rsGetNextCustomerStopByTruck = Server.CreateObject("ADODB.Recordset")
	rsGetNextCustomerStopByTruck.CursorLocation = 3 

	resultGetNextCustomerStopByTruck = 0 

	'First See if they have specified a Next Stop Manually		
	SQLGetNextCustomerStopByTruck = "SELECT CustNum FROM RT_DeliveryBoard WHERE TruckNumber = '" &  passedTruck & "' AND ManualNextStop = 1"
 
	Set rsGetNextCustomerStopByTruck= cnnGetNextCustomerStopByTruck.Execute(SQLGetNextCustomerStopByTruck)
	
	If not rsGetNextCustomerStopByTruck.eof then
		'They manually specified a next stop
		resultGetNextCustomerStopByTruck = rsGetNextCustomerStopByTruck("CustNum")
	Else
		If DelBoardDontUseStopSequencing() = False Then ' If they have sequencing off, then only the manually set ones count
			'See if no stops at all have been made yet, then it's just the first seq
			SQLGetNextCustomerStopByTruck = "SELECT Count(*) AS Expr1 FROM RT_DeliveryBoard WHERE TruckNumber = '" & passedTruck & "' AND DeliveryStatus Is Not NULL"
			Set rsGetNextCustomerStopByTruck= cnnGetNextCustomerStopByTruck.Execute(SQLGetNextCustomerStopByTruck)
			
			If rsGetNextCustomerStopByTruck("Expr1") = 0 then
				'No stops at all have been made, get the first in the sequence
				SQLGetNextCustomerStopByTruck = "SELECT Top 1 CustNum FROM RT_DeliveryBoard WHERE TruckNumber = '" & passedTruck & "' Order By SequenceNumber, CustNum"
				Set rsGetNextCustomerStopByTruck= cnnGetNextCustomerStopByTruck.Execute(SQLGetNextCustomerStopByTruck)
				If Not rsGetNextCustomerStopByTruck.EOF Then resultGetNextCustomerStopByTruck = rsGetNextCustomerStopByTruck("CustNum")
			Else
				'Some stops have been made so we need to figure it out
				SQLGetNextCustomerStopByTruck = "SELECT CustNum FROM RT_DeliveryBoard WHERE TruckNumber = '" & passedTruck & "' Order By SequenceNumber, CustNum"
				Set rsGetNextCustomerStopByTruck= cnnGetNextCustomerStopByTruck.Execute(SQLGetNextCustomerStopByTruck)
				Do While Not rsGetNextCustomerStopByTruck.eof
					If CustHasANYDeliveryByCustAndTruck(rsGetNextCustomerStopByTruck("CustNum"),passedTruck) = False Then
						resultGetNextCustomerStopByTruck = rsGetNextCustomerStopByTruck("CustNum")
						Exit Do
					End IF
		 			rsGetNextCustomerStopByTruck.Movenext
		 		Loop
			End If
		End If
	End If
		
	rsGetNextCustomerStopByTruck.Close
	set rsGetNextCustomerStopByTruck= Nothing
	cnnGetNextCustomerStopByTruck.Close	
	set cnnGetNextCustomerStopByTruck = Nothing
	
	GetNextCustomerStopByTruck = resultGetNextCustomerStopByTruck 
	
End Function

Function DeliveryAlertSet(passedInvoice,passedUserNo)

	Set cnnDeliveryAlertSet = Server.CreateObject("ADODB.Connection")
	cnnDeliveryAlertSet.open Session("ClientCnnString")

	resultDeliveryAlertSet = False
		
	SQLDeliveryAlertSet = "SELECT * FROM SC_Alerts WHERE AlertType = 'DeliveryBoard' AND "
	SQLDeliveryAlertSet = SQLDeliveryAlertSet & " CreatedByUserNo = " & passedUserNo
	SQLDeliveryAlertSet = SQLDeliveryAlertSet & " AND  ReferenceValue = '" & passedInvoice & "'"

	Set rsDeliveryAlertSet = Server.CreateObject("ADODB.Recordset")
	rsDeliveryAlertSet.CursorLocation = 3 
	Set rsDeliveryAlertSet = cnnDeliveryAlertSet.Execute(SQLDeliveryAlertSet)
	
	If not rsDeliveryAlertSet.eof then resultDeliveryAlertSet = True
		
	rsDeliveryAlertSet.Close
	set rsDeliveryAlertSet= Nothing
	cnnDeliveryAlertSet.Close	
	set cnnDeliveryAlertSet = Nothing
	
	DeliveryAlertSet = resultDeliveryAlertSet 
	
End Function

Function DeliveryAlertCondition(passedInvoice,passedUserNo)

	Set cnnDeliveryAlertCondition = Server.CreateObject("ADODB.Connection")
	cnnDeliveryAlertCondition.open Session("ClientCnnString")

	resultDeliveryAlertCondition = ""
		
	SQLDeliveryAlertCondition = "SELECT * FROM SC_Alerts WHERE AlertType = 'DeliveryBoard' AND "
	SQLDeliveryAlertCondition = SQLDeliveryAlertCondition & " CreatedByUserNo = " & passedUserNo
	SQLDeliveryAlertCondition = SQLDeliveryAlertCondition & " AND  ReferenceValue = '" & passedInvoice & "'"

	Set rsDeliveryAlertCondition = Server.CreateObject("ADODB.Recordset")
	rsDeliveryAlertCondition.CursorLocation = 3 
	Set rsDeliveryAlertCondition = cnnDeliveryAlertCondition.Execute(SQLDeliveryAlertCondition)
	
	If not rsDeliveryAlertCondition.eof then resultDeliveryAlertCondition = rsDeliveryAlertCondition("Condition")
		
	rsDeliveryAlertCondition.Close
	set rsDeliveryAlertCondition= Nothing
	cnnDeliveryAlertCondition.Close	
	set cnnDeliveryAlertCondition = Nothing
	
	DeliveryAlertCondition = resultDeliveryAlertCondition 
	
End Function

Function GetCustNumberByInvoiceNumDelBoard(passedInvoiceNumber)

	Set cnnGetCustNumberByInvoiceNumDelBoard = Server.CreateObject("ADODB.Connection")
	cnnGetCustNumberByInvoiceNumDelBoard.open Session("ClientCnnString")

	resultGetCustNumberByInvoiceNumDelBoard = ""
		
	SQLGetCustNumberByInvoiceNumDelBoard = "Select CustNum from " & Session("SQL_Owner") & ".RT_DeliveryBoard where IvsNum = " & passedInvoiceNumber
	 
	Set rsGetCustNumberByInvoiceNumDelBoard = Server.CreateObject("ADODB.Recordset")
	rsGetCustNumberByInvoiceNumDelBoard.CursorLocation = 3 
	Set rsGetCustNumberByInvoiceNumDelBoard= cnnGetCustNumberByInvoiceNumDelBoard.Execute(SQLGetCustNumberByInvoiceNumDelBoard)
	
	
	If not rsGetCustNumberByInvoiceNumDelBoard.eof then resultGetCustNumberByInvoiceNumDelBoard = rsGetCustNumberByInvoiceNumDelBoard("CustNum")
	
	set rsGetCustNumberByInvoiceNumDelBoard= Nothing
	set cnnGetCustNumberByInvoiceNumDelBoard= Nothing
	
	GetCustNumberByInvoiceNumDelBoard = resultGetCustNumberByInvoiceNumDelBoard
	
End Function


Function GetCustNumberByInvoiceNumDelBoardHistory(passedInvoiceNumber)

	Set cnnGetCustNumberByInvoiceNumDelBoardHistory = Server.CreateObject("ADODB.Connection")
	cnnGetCustNumberByInvoiceNumDelBoardHistory.open Session("ClientCnnString")

	resultGetCustNumberByInvoiceNumDelBoardHistory = ""
		
	SQLGetCustNumberByInvoiceNumDelBoardHistory = "Select CustNum from " & Session("SQL_Owner") & ".RT_DeliveryBoardHistory where IvsNum = " & passedInvoiceNumber
	 
	Set rsGetCustNumberByInvoiceNumDelBoardHistory = Server.CreateObject("ADODB.Recordset")
	rsGetCustNumberByInvoiceNumDelBoardHistory.CursorLocation = 3 
	Set rsGetCustNumberByInvoiceNumDelBoardHistory= cnnGetCustNumberByInvoiceNumDelBoardHistory.Execute(SQLGetCustNumberByInvoiceNumDelBoardHistory)
	
	
	If not rsGetCustNumberByInvoiceNumDelBoardHistory.eof then resultGetCustNumberByInvoiceNumDelBoardHistory = rsGetCustNumberByInvoiceNumDelBoardHistory("CustNum")
	
	set rsGetCustNumberByInvoiceNumDelBoardHistory= Nothing
	set cnnGetCustNumberByInvoiceNumDelBoardHistory= Nothing
	
	GetCustNumberByInvoiceNumDelBoardHistory = resultGetCustNumberByInvoiceNumDelBoardHistory
	
End Function



Function GetTruckByInvoiceNumDelBoard(passedInvoiceNumber)

	Set cnnGetTruckByInvoiceNumDelBoard = Server.CreateObject("ADODB.Connection")
	cnnGetTruckByInvoiceNumDelBoard.open Session("ClientCnnString")

	resultGetTruckByInvoiceNumDelBoard = ""
		
	SQLGetTruckByInvoiceNumDelBoard = "Select TruckNumber from " & Session("SQL_Owner") & ".RT_DeliveryBoard where IvsNum = " & passedInvoiceNumber
	 
	Set rsGetTruckByInvoiceNumDelBoard = Server.CreateObject("ADODB.Recordset")
	rsGetTruckByInvoiceNumDelBoard.CursorLocation = 3 
	Set rsGetTruckByInvoiceNumDelBoard= cnnGetTruckByInvoiceNumDelBoard.Execute(SQLGetTruckByInvoiceNumDelBoard)
	
	
	If not rsGetTruckByInvoiceNumDelBoard.eof then resultGetTruckByInvoiceNumDelBoard = rsGetTruckByInvoiceNumDelBoard("TruckNumber")
	
	set rsGetTruckByInvoiceNumDelBoard= Nothing
	set cnnGetTruckByInvoiceNumDelBoard= Nothing
	
	GetTruckByInvoiceNumDelBoard = resultGetTruckByInvoiceNumDelBoard
	
End Function

Function GetLastDeliveryStatusChangeBYInvoiceNumDelBoard(passedInvoiceNumber)

	Set cnnGetLastDeliveryStatusChangeBYInvoiceNumDelBoard = Server.CreateObject("ADODB.Connection")
	cnnGetLastDeliveryStatusChangeBYInvoiceNumDelBoard.open Session("ClientCnnString")

	resultGetLastDeliveryStatusChangeBYInvoiceNumDelBoard = ""
		
	SQLGetLastDeliveryStatusChangeBYInvoiceNumDelBoard = "Select TOP 1 LastDeliveryStatusChange from " & Session("SQL_Owner") & ".RT_DeliveryBoard where IvsNum = " & passedInvoiceNumber & "ORDER BY LastDeliveryStatusChange DESC"
	 
	Set rsGetLastDeliveryStatusChangeBYInvoiceNumDelBoard = Server.CreateObject("ADODB.Recordset")
	rsGetLastDeliveryStatusChangeBYInvoiceNumDelBoard.CursorLocation = 3 
	Set rsGetLastDeliveryStatusChangeBYInvoiceNumDelBoard= cnnGetLastDeliveryStatusChangeBYInvoiceNumDelBoard.Execute(SQLGetLastDeliveryStatusChangeBYInvoiceNumDelBoard)
	
	
	If not rsGetLastDeliveryStatusChangeBYInvoiceNumDelBoard.eof then resultGetLastDeliveryStatusChangeBYInvoiceNumDelBoard = rsGetLastDeliveryStatusChangeBYInvoiceNumDelBoard("LastDeliveryStatusChange")
	
	set rsGetLastDeliveryStatusChangeBYInvoiceNumDelBoard= Nothing
	set cnnGetLastDeliveryStatusChangeBYInvoiceNumDelBoard= Nothing
	
	GetLastDeliveryStatusChangeBYInvoiceNumDelBoard = resultGetLastDeliveryStatusChangeBYInvoiceNumDelBoard
	
End Function

Function GetNumberOfInvoicesByTruckNumber(passedTruckNumber)

	Set cnnGetNumberOfInvoicesByTruckNumber = Server.CreateObject("ADODB.Connection")
	cnnGetNumberOfInvoicesByTruckNumber.open Session("ClientCnnString")

	resultGetNumberOfInvoicesByTruckNumber = 0
		
	SQLGetNumberOfInvoicesByTruckNumber = "Select Count(IvsNum) as TotalCount from RT_DeliveryBoard where TruckNumber = '" & passedTruckNumber & "'"
	 
	Set rsGetNumberOfInvoicesByTruckNumber = Server.CreateObject("ADODB.Recordset")
	rsGetNumberOfInvoicesByTruckNumber.CursorLocation = 3 
	Set rsGetNumberOfInvoicesByTruckNumber= cnnGetNumberOfInvoicesByTruckNumber.Execute(SQLGetNumberOfInvoicesByTruckNumber)
	
	
	If not rsGetNumberOfInvoicesByTruckNumber.eof then resultGetNumberOfInvoicesByTruckNumber = rsGetNumberOfInvoicesByTruckNumber("TotalCount")
	
	set rsGetNumberOfInvoicesByTruckNumber= Nothing
	set cnnGetNumberOfInvoicesByTruckNumber= Nothing
	
	GetNumberOfInvoicesByTruckNumber = resultGetNumberOfInvoicesByTruckNumber
	
End Function

Function GetNumberOfCustomersByTruckNumber(passedTruckNumber)

	Set cnnGetNumberOfCustomersByTruckNumber = Server.CreateObject("ADODB.Connection")
	cnnGetNumberOfCustomersByTruckNumber.open Session("ClientCnnString")

	resultGetNumberOfCustomersByTruckNumber = 0
		
	SQLGetNumberOfCustomersByTruckNumber = "Select Count(Distinct CustNum) as TotalCount from RT_DeliveryBoard where TruckNumber = '" & passedTruckNumber & "'"
	 
	Set rsGetNumberOfCustomersByTruckNumber = Server.CreateObject("ADODB.Recordset")
	rsGetNumberOfCustomersByTruckNumber.CursorLocation = 3 
	Set rsGetNumberOfCustomersByTruckNumber= cnnGetNumberOfCustomersByTruckNumber.Execute(SQLGetNumberOfCustomersByTruckNumber)
	
	
	If not rsGetNumberOfCustomersByTruckNumber.eof then resultGetNumberOfCustomersByTruckNumber = rsGetNumberOfCustomersByTruckNumber("TotalCount")
	
	set rsGetNumberOfCustomersByTruckNumber= Nothing
	set cnnGetNumberOfCustomersByTruckNumber= Nothing
	
	GetNumberOfCustomersByTruckNumber = resultGetNumberOfCustomersByTruckNumber
	
End Function

Function GetValueOfDeliveriesByTruckNumber(passedTruckNumber)

	Set cnnGetValueOfDeliveriesByTruckNumber = Server.CreateObject("ADODB.Connection")
	cnnGetValueOfDeliveriesByTruckNumber.open Session("ClientCnnString")

	resultGetValueOfDeliveriesByTruckNumber = 0
		
	SQLGetValueOfDeliveriesByTruckNumber = "Select Sum(Value) as TotalVal from RT_DeliveryBoard where TruckNumber = '" & passedTruckNumber & "'"
	 
	Set rsGetValueOfDeliveriesByTruckNumber = Server.CreateObject("ADODB.Recordset")
	rsGetValueOfDeliveriesByTruckNumber.CursorLocation = 3 
	Set rsGetValueOfDeliveriesByTruckNumber= cnnGetValueOfDeliveriesByTruckNumber.Execute(SQLGetValueOfDeliveriesByTruckNumber)
	
	
	If not rsGetValueOfDeliveriesByTruckNumber.eof then resultGetValueOfDeliveriesByTruckNumber = rsGetValueOfDeliveriesByTruckNumber("TotalVal")
	
	set rsGetValueOfDeliveriesByTruckNumber= Nothing
	set cnnGetValueOfDeliveriesByTruckNumber= Nothing
	
	GetValueOfDeliveriesByTruckNumber = resultGetValueOfDeliveriesByTruckNumber
	
End Function


Function GetNumberOfInvoicesByTruckNumberHistorical(passedTruckNumber, passedDeliveryDate)

	Set cnnGetNumberOfInvoicesByTruckNumberHistorical = Server.CreateObject("ADODB.Connection")
	cnnGetNumberOfInvoicesByTruckNumberHistorical.open Session("ClientCnnString")

	resultGetNumberOfInvoicesByTruckNumberHistorical = 0
		
	SQLGetNumberOfInvoicesByTruckNumberHistorical = "Select Count(IvsNum) as TotalCount from RT_DeliveryBoardHistory where TruckNumber = '" & passedTruckNumber & "' AND "
	SQLGetNumberOfInvoicesByTruckNumberHistorical = SQLGetNumberOfInvoicesByTruckNumberHistorical & " Year(LastDeliveryStatusChange) = " & Year(passedDeliveryDate) & " AND "
	SQLGetNumberOfInvoicesByTruckNumberHistorical = SQLGetNumberOfInvoicesByTruckNumberHistorical & " Month(LastDeliveryStatusChange) = " & Month(passedDeliveryDate) & " AND "
	SQLGetNumberOfInvoicesByTruckNumberHistorical = SQLGetNumberOfInvoicesByTruckNumberHistorical & " Day(LastDeliveryStatusChange) = " & Day(passedDeliveryDate)

	'Response.write("<br><br>" & SQLGetNumberOfInvoicesByTruckNumberHistorical & "<br><br>")

	Set rsGetNumberOfInvoicesByTruckNumberHistorical = Server.CreateObject("ADODB.Recordset")
	rsGetNumberOfInvoicesByTruckNumberHistorical.CursorLocation = 3 
	Set rsGetNumberOfInvoicesByTruckNumberHistorical= cnnGetNumberOfInvoicesByTruckNumberHistorical.Execute(SQLGetNumberOfInvoicesByTruckNumberHistorical)
	
	
	If not rsGetNumberOfInvoicesByTruckNumberHistorical.eof then 
		resultGetNumberOfInvoicesByTruckNumberHistorical = rsGetNumberOfInvoicesByTruckNumberHistorical("TotalCount")
	Else
		resultGetNumberOfInvoicesByTruckNumberHistorical = 0
	End If
	
	set rsGetNumberOfInvoicesByTruckNumberHistorical= Nothing
	set cnnGetNumberOfInvoicesByTruckNumberHistorical= Nothing
	
	GetNumberOfInvoicesByTruckNumberHistorical = resultGetNumberOfInvoicesByTruckNumberHistorical
	
End Function



Function GetNumberOfCustomersByTruckNumberHistorical(passedTruckNumber, passedDeliveryDate)

	Set cnnGetNumberOfCustomersByTruckNumberHistorical = Server.CreateObject("ADODB.Connection")
	cnnGetNumberOfCustomersByTruckNumberHistorical.open Session("ClientCnnString")

	resultGetNumberOfCustomersByTruckNumberHistoricalHistorical = 0
	
		
	SQLGetNumberOfCustomersByTruckNumberHistorical = "Select Count(Distinct CustNum) as TotalCount from RT_DeliveryBoardHistory where TruckNumber = '" & passedTruckNumber & "' AND "
	SQLGetNumberOfCustomersByTruckNumberHistorical = SQLGetNumberOfCustomersByTruckNumberHistorical & " Year(LastDeliveryStatusChange) = " & Year(passedDeliveryDate) & " AND "
	SQLGetNumberOfCustomersByTruckNumberHistorical = SQLGetNumberOfCustomersByTruckNumberHistorical & " Month(LastDeliveryStatusChange) = " & Month(passedDeliveryDate) & " AND "
	SQLGetNumberOfCustomersByTruckNumberHistorical = SQLGetNumberOfCustomersByTruckNumberHistorical & " Day(LastDeliveryStatusChange) = " & Day(passedDeliveryDate)
	
	Set rsGetNumberOfCustomersByTruckNumberHistorical = Server.CreateObject("ADODB.Recordset")
	rsGetNumberOfCustomersByTruckNumberHistorical.CursorLocation = 3 
	Set rsGetNumberOfCustomersByTruckNumberHistorical= cnnGetNumberOfCustomersByTruckNumberHistorical.Execute(SQLGetNumberOfCustomersByTruckNumberHistorical)
	
	
	If not rsGetNumberOfCustomersByTruckNumberHistorical.eof then 
		resultGetNumberOfCustomersByTruckNumberHistorical = rsGetNumberOfCustomersByTruckNumberHistorical("TotalCount")
	Else
		rsGetNumberOfCustomersByTruckNumberHistorical = 0
	End If
	
	set rsGetNumberOfCustomersByTruckNumberHistorical= Nothing
	set cnnGetNumberOfCustomersByTruckNumberHistorical= Nothing
	
	GetNumberOfCustomersByTruckNumberHistorical = resultGetNumberOfCustomersByTruckNumberHistorical
	
End Function



Function GetValueOfDeliveriesByTruckNumberHistorical(passedTruckNumber, passedDeliveryDate)

	Set cnnGetValueOfDeliveriesByTruckNumberHistorical = Server.CreateObject("ADODB.Connection")
	cnnGetValueOfDeliveriesByTruckNumberHistorical.open Session("ClientCnnString")

	resultGetValueOfDeliveriesByTruckNumberHistorical = 0
		
	SQLGetValueOfDeliveriesByTruckNumberHistorical = "Select Sum(Value) as TotalVal from RT_DeliveryBoardHistory where TruckNumber = '" & passedTruckNumber & "' AND "
	SQLGetValueOfDeliveriesByTruckNumberHistorical = SQLGetValueOfDeliveriesByTruckNumberHistorical & " Year(LastDeliveryStatusChange) = " & Year(passedDeliveryDate) & " AND "
	SQLGetValueOfDeliveriesByTruckNumberHistorical = SQLGetValueOfDeliveriesByTruckNumberHistorical & " Month(LastDeliveryStatusChange) = " & Month(passedDeliveryDate) & " AND "
	SQLGetValueOfDeliveriesByTruckNumberHistorical = SQLGetValueOfDeliveriesByTruckNumberHistorical & " Day(LastDeliveryStatusChange) = " & Day(passedDeliveryDate)
	

	Set rsGetValueOfDeliveriesByTruckNumberHistorical = Server.CreateObject("ADODB.Recordset")
	rsGetValueOfDeliveriesByTruckNumberHistorical.CursorLocation = 3 
	Set rsGetValueOfDeliveriesByTruckNumberHistorical= cnnGetValueOfDeliveriesByTruckNumberHistorical.Execute(SQLGetValueOfDeliveriesByTruckNumberHistorical)
	
	
	If not rsGetValueOfDeliveriesByTruckNumberHistorical.eof then 
		resultGetValueOfDeliveriesByTruckNumberHistorical = rsGetValueOfDeliveriesByTruckNumberHistorical("TotalVal")
	Else
		resultGetValueOfDeliveriesByTruckNumberHistorical = 0
	End If
	
	set rsGetValueOfDeliveriesByTruckNumberHistorical= Nothing
	set cnnGetValueOfDeliveriesByTruckNumberHistorical= Nothing
	
	GetValueOfDeliveriesByTruckNumberHistorical = resultGetValueOfDeliveriesByTruckNumberHistorical
	
End Function


Function DelBoardHistMostRecentDate()

	Set cnnDelBoardHistMostRecentDate = Server.CreateObject("ADODB.Connection")
	cnnDelBoardHistMostRecentDate.open Session("ClientCnnString")

	resultDelBoardHistMostRecentDate=""
		
	SQLDelBoardHistMostRecentDate = "Select Max(DeliveryDate) AS Expr1 from RT_DeliveryBoardHistory"
	 
	Set rsDelBoardHistMostRecentDate = Server.CreateObject("ADODB.Recordset")
	rsDelBoardHistMostRecentDate.CursorLocation = 3 
	Set rsDelBoardHistMostRecentDate= cnnDelBoardHistMostRecentDate.Execute(SQLDelBoardHistMostRecentDate)
	
	
	If not rsDelBoardHistMostRecentDate.eof then resultDelBoardHistMostRecentDate = rsDelBoardHistMostRecentDate("Expr1")
	
	DelBoardHistMostRecentDate = resultDelBoardHistMostRecentDate 
	
End Function

Function GetNumberOutOfSequenceByTruckNumber(passedTruckNumber)

	Set cnnGetNumberOutOfSequenceByTruckNumber = Server.CreateObject("ADODB.Connection")
	cnnGetNumberOutOfSequenceByTruckNumber.open Session("ClientCnnString")

	resultGetNumberOutOfSequenceByTruckNumber = 0
		
	SQLGetNumberOutOfSequenceByTruckNumber = "Select Count(*) as TotalCount from zOutOfSequenceReport_" & Trim(Session("userNo")) & " where TruckNumber = '" & passedTruckNumber & "' AND "
	SQLGetNumberOutOfSequenceByTruckNumber  = SQLGetNumberOutOfSequenceByTruckNumber  & " SequenceNumber <> ActualDeliverySequence"
	 
	Set rsGetNumberOutOfSequenceByTruckNumber = Server.CreateObject("ADODB.Recordset")
	rsGetNumberOutOfSequenceByTruckNumber.CursorLocation = 3 
	Set rsGetNumberOutOfSequenceByTruckNumber= cnnGetNumberOutOfSequenceByTruckNumber.Execute(SQLGetNumberOutOfSequenceByTruckNumber)
	
	
	If not rsGetNumberOutOfSequenceByTruckNumber.eof then resultGetNumberOutOfSequenceByTruckNumber = rsGetNumberOutOfSequenceByTruckNumber("TotalCount")
	
	set rsGetNumberOutOfSequenceByTruckNumber= Nothing
	set cnnGetNumberOutOfSequenceByTruckNumber= Nothing
	
	GetNumberOutOfSequenceByTruckNumber = resultGetNumberOutOfSequenceByTruckNumber
	
End Function


Function AutoPromptNextStopON()

	Set cnnAutoPromptNextStopON = Server.CreateObject("ADODB.Connection")
	cnnAutoPromptNextStopON.open Session("ClientCnnString")

	resultAutoPromptNextStopON = False
		
	SQLAutoPromptNextStopON = "Select * from Settings_Global"
	 
	Set rsAutoPromptNextStopON = Server.CreateObject("ADODB.Recordset")
	rsAutoPromptNextStopON.CursorLocation = 3 
	Set rsAutoPromptNextStopON= cnnAutoPromptNextStopON.Execute(SQLAutoPromptNextStopON)
	
	
	If not rsAutoPromptNextStopON.eof then
		If rsAutoPromptNextStopON("AutoPromptNextStop") = 0 Then resultAutoPromptNextStopON = False Else resultAutoPromptNextStopON = True
	End If
	AutoPromptNextStopON= resultAutoPromptNextStopON 
	
End Function


Function AutoForceSelectNextStopON(passedUserNo)

	'First will lookup the setting for the user, if set to
	'Use Global, will go get the global setting

	Set cnnAutoForceSelectNextStop = Server.CreateObject("ADODB.Connection")
	cnnAutoForceSelectNextStop.open Session("ClientCnnString")

	resultAutoForceSelectNextStop = False
	
	SQLAutoForceSelectNextStop = "Select * from tblUsers WHERE UserNo = " & passedUserNo
	 
	Set rsAutoForceSelectNextStop = Server.CreateObject("ADODB.Recordset")
	rsAutoForceSelectNextStop.CursorLocation = 3 
	Set rsAutoForceSelectNextStop= cnnAutoForceSelectNextStop.Execute(SQLAutoForceSelectNextStop)

	If not rsAutoForceSelectNextStop.eof then
		If rsAutoForceSelectNextStop("userForceNextStopSelectionOverride") = "Yes" Then resultAutoForceSelectNextStop = True
		If rsAutoForceSelectNextStop("userForceNextStopSelectionOverride") = "No" Then resultAutoForceSelectNextStop = False		
	End If

	
	If rsAutoForceSelectNextStop("userForceNextStopSelectionOverride") = "Use Global" Then	
		
		SQLAutoForceSelectNextStop = "Select * from Settings_Global"
		 
		Set rsAutoForceSelectNextStop = Server.CreateObject("ADODB.Recordset")
		rsAutoForceSelectNextStop.CursorLocation = 3 
		Set rsAutoForceSelectNextStop= cnnAutoForceSelectNextStop.Execute(SQLAutoForceSelectNextStop)
		
		
		If not rsAutoForceSelectNextStop.eof then
			If rsAutoForceSelectNextStop("AutoForceSelectNextStop") = 0 Then resultAutoForceSelectNextStop = False Else resultAutoForceSelectNextStop = True
		End If
	
	End If
	
	Set rsAutoForceSelectNextStop = Nothing
	cnnAutoForceSelectNextStop.Close
	Set cnnAutoForceSelectNextStop = Nothing
	
	AutoForceSelectNextStopON = resultAutoForceSelectNextStop 
	
End Function


Function GetLastInvoiceMarkedByTruckNumber(passedTruckNumber)

	Set cnnGetLastInvoiceMarkedByTruckNumber = Server.CreateObject("ADODB.Connection")
	cnnGetLastInvoiceMarkedByTruckNumber.open Session("ClientCnnString")

	resultGetLastInvoiceMarkedByTruckNumber = ""
		
	SQLGetLastInvoiceMarkedByTruckNumber = "Select Top 1 * FROM " & Session("SQL_Owner") & ".RT_DeliveryBoard where TruckNumber = '" & passedTruckNumber & "' AND (DeliveryStatus IS NOT NULL or DeliveryInProgress = 1) ORDER BY LastDeliveryStatusChange DESC"
	 
	Set rsGetLastInvoiceMarkedByTruckNumber = Server.CreateObject("ADODB.Recordset")
	rsGetLastInvoiceMarkedByTruckNumber.CursorLocation = 3 
	Set rsGetLastInvoiceMarkedByTruckNumber= cnnGetLastInvoiceMarkedByTruckNumber.Execute(SQLGetLastInvoiceMarkedByTruckNumber)
	
	
	If not rsGetLastInvoiceMarkedByTruckNumber.eof then
		resultGetLastInvoiceMarkedByTruckNumber = rsGetLastInvoiceMarkedByTruckNumber("IvsNum")
	End If
	
	set rsGetLastInvoiceMarkedByTruckNumber= Nothing
	set cnnGetLastInvoiceMarkedByTruckNumber= Nothing
	
	GetLastInvoiceMarkedByTruckNumber = resultGetLastInvoiceMarkedByTruckNumber
	
End Function

Function DelBoardDontUseStopSequencing()

	resultDelBoardDontUseStopSequencing = False

	Set cnnDelBoardDontUseStopSequencing = Server.CreateObject("ADODB.Connection")
	cnnDelBoardDontUseStopSequencing.open (Session("ClientCnnString"))
	Set rsDelBoardDontUseStopSequencing = Server.CreateObject("ADODB.Recordset")
	rsDelBoardDontUseStopSequencing.CursorLocation = 3 

	SQLDelBoardDontUseStopSequencing = "SELECT DelBoardDontUseStopSequencing FROM Settings_Global"

	Set rsDelBoardDontUseStopSequencing = cnnDelBoardDontUseStopSequencing.Execute(SQLDelBoardDontUseStopSequencing)

	If not rsDelBoardDontUseStopSequencing.EOF Then
		If rsDelBoardDontUseStopSequencing("DelBoardDontUseStopSequencing") = 1 Then resultDelBoardDontUseStopSequencing = True
	End If
	set rsDelBoardDontUseStopSequencing = Nothing
	cnnDelBoardDontUseStopSequencing.close
	set cnnDelBoardDontUseStopSequencing = Nothing
	
	DelBoardDontUseStopSequencing = resultDelBoardDontUseStopSequencing

End Function


Function DelBoardDontShowDeliveryLineItems()

	resultDelBoardDontShowDeliveryLineItems = False

	Set cnnDelBoardDontShowDeliveryLineItems = Server.CreateObject("ADODB.Connection")
	cnnDelBoardDontShowDeliveryLineItems.open (Session("ClientCnnString"))
	Set rsDelBoardDontShowDeliveryLineItems = Server.CreateObject("ADODB.Recordset")
	rsDelBoardDontShowDeliveryLineItems.CursorLocation = 3 

	SQLDelBoardDontShowDeliveryLineItems = "SELECT DoNotShowDeliveryLineItems FROM Settings_Global"

	Set rsDelBoardDontShowDeliveryLineItems = cnnDelBoardDontShowDeliveryLineItems.Execute(SQLDelBoardDontShowDeliveryLineItems)

	If not rsDelBoardDontShowDeliveryLineItems.EOF Then
		If rsDelBoardDontShowDeliveryLineItems("DoNotShowDeliveryLineItems") = 1 Then resultDelBoardDontShowDeliveryLineItems = True
	End If
	set rsDelBoardDontShowDeliveryLineItems = Nothing
	cnnDelBoardDontShowDeliveryLineItems.close
	set cnnDelBoardDontShowDeliveryLineItems = Nothing
	
	DelBoardDontShowDeliveryLineItems = resultDelBoardDontShowDeliveryLineItems

End Function


Function DelBoardIgnoreThisRoute(passedTruckNumber)

	resultDelBoardIgnoreThisRoute = False

	Set cnnDelBoardIgnoreThisRoute = Server.CreateObject("ADODB.Connection")
	cnnDelBoardIgnoreThisRoute.open Session("ClientCnnString")

	resultAutoPromptNextStop = False
		
	SQLDelBoardIgnoreThisRoute = "Select * from Settings_Global"
	 
	Set rsDelBoardIgnoreThisRoute = Server.CreateObject("ADODB.Recordset")
	rsDelBoardIgnoreThisRoute.CursorLocation = 3 
	Set rsDelBoardIgnoreThisRoute = cnnDelBoardIgnoreThisRoute.Execute(SQLDelBoardIgnoreThisRoute)
	
	If Not rsDelBoardIgnoreThisRoute.Eof Then
		RoutesToIgnore = rsDelBoardIgnoreThisRoute("DelBoardRoutesToIgnore")
		
		If RoutesToIgnore <> "" Then
			RoutesToIgnoreArray = split(RoutesToIgnore,",")
			For i = 0 to ubound(RoutesToIgnoreArray)
				If trim(passedTruckNumber) = trim(RoutesToIgnoreArray(i)) Then resultDelBoardIgnoreThisRoute = True
			Next 
		End If
	End If
	
	Set rsDelBoardIgnoreThisRoute = Nothing
	cnnDelBoardIgnoreThisRoute.Close
	Set cnnDelBoardIgnoreThisRoute = Nothing
	
	DelBoardIgnoreThisRoute = resultDelBoardIgnoreThisRoute 
	
End Function



Function DriverNumberHasNagAlerts(passedDriverUserNo)

	resultDriverNumberHasNagAlerts = False

	Set cnnDriverNumberHasNagAlerts = Server.CreateObject("ADODB.Connection")
	cnnDriverNumberHasNagAlerts.open Session("ClientCnnString")

	resultAutoPromptNextStop = False
		
	SQLDriverNumberHasNagAlerts = "SELECT * FROM SC_NagSkipUsers WHERE UserNo = " & passedDriverUserNo & " AND (NagType = 'routingNoActivity' OR NagType = 'routingNoActivity')"
	 
	Set rsDriverNumberHasNagAlerts = Server.CreateObject("ADODB.Recordset")
	rsDriverNumberHasNagAlerts.CursorLocation = 3 
	Set rsDriverNumberHasNagAlerts = cnnDriverNumberHasNagAlerts.Execute(SQLDriverNumberHasNagAlerts)
	
	If NOT rsDriverNumberHasNagAlerts.EOF Then
		resultDriverNumberHasNagAlerts = True
	End If
	
	Set rsDriverNumberHasNagAlerts = Nothing
	cnnDriverNumberHasNagAlerts.Close
	Set cnnDriverNumberHasNagAlerts = Nothing
	
	DriverNumberHasNagAlerts = resultDriverNumberHasNagAlerts 
	
End Function

Function GetDriverCommentsByInvoiceNumber(passedInvoiceNumber)

	Set cnnGetDriverCommentsByInvoiceNumber = Server.CreateObject("ADODB.Connection")
	cnnGetDriverCommentsByInvoiceNumber.open Session("ClientCnnString")

	resultGetDriverCommentsByInvoiceNumber = ""
		
	SQLGetDriverCommentsByInvoiceNumber = "Select DriverComments from " & Session("SQL_Owner") & ".RT_DeliveryBoard where IvsNum = " & passedInvoiceNumber
	 
	Set rsGetDriverCommentsByInvoiceNumber = Server.CreateObject("ADODB.Recordset")
	rsGetDriverCommentsByInvoiceNumber.CursorLocation = 3 
	Set rsGetDriverCommentsByInvoiceNumber= cnnGetDriverCommentsByInvoiceNumber.Execute(SQLGetDriverCommentsByInvoiceNumber)
	
	
	If not rsGetDriverCommentsByInvoiceNumber.eof then resultGetDriverCommentsByInvoiceNumber = rsGetDriverCommentsByInvoiceNumber("DriverComments")
	
	set rsGetDriverCommentsByInvoiceNumber= Nothing
	set cnnGetDriverCommentsByInvoiceNumber= Nothing
	
	GetDriverCommentsByInvoiceNumber = resultGetDriverCommentsByInvoiceNumber
	
End Function

Function InvoiceIsNextStop(passedIvsNum)

	Set cnnInvoiceIsNextStop = Server.CreateObject("ADODB.Connection")
	cnnInvoiceIsNextStop.open Session("ClientCnnString")

	resultInvoiceIsNextStop= "False"
		
	SQLInvoiceIsNextStop = "Select * from RT_DeliveryBoard WHERE IvsNum = " & passedIvsNum
	 
	Set rsInvoiceIsNextStop = Server.CreateObject("ADODB.Recordset")
	rsInvoiceIsNextStop.CursorLocation = 3 
	Set rsInvoiceIsNextStop= cnnInvoiceIsNextStop.Execute(SQLInvoiceIsNextStop)
	
	If not rsInvoiceIsNextStop.eof then 
		If rsInvoiceIsNextStop("ManualNextStop") = 1 Then resultInvoiceIsNextStop= "True"
	End If
		
	rsInvoiceIsNextStop.Close
	set rsInvoiceIsNextStop= Nothing
	cnnInvoiceIsNextStop.Close	
	set cnnInvoiceIsNextStop = Nothing
	
	InvoiceIsNextStop = resultInvoiceIsNextStop 
	
End Function

Function GetLastInvoiceMarkedDATETIMEByTruckNumber(passedTruckNumber)

	Set cnnGetLastInvoiceMarkedDATETIMEByTruckNumber = Server.CreateObject("ADODB.Connection")
	cnnGetLastInvoiceMarkedDATETIMEByTruckNumber.open Session("ClientCnnString")

	resultGetLastInvoiceMarkedDATETIMEByTruckNumber = ""
		
	SQLGetLastInvoiceMarkedDATETIMEByTruckNumber = "Select Top 1 * FROM " & Session("SQL_Owner") & ".RT_DeliveryBoard where TruckNumber = '" & passedTruckNumber & "' AND (DeliveryStatus IS NOT NULL or DeliveryInProgress = 1) ORDER BY LastDeliveryStatusChange DESC"
	 
	Set rsGetLastInvoiceMarkedDATETIMEByTruckNumber = Server.CreateObject("ADODB.Recordset")
	rsGetLastInvoiceMarkedDATETIMEByTruckNumber.CursorLocation = 3 
	Set rsGetLastInvoiceMarkedDATETIMEByTruckNumber= cnnGetLastInvoiceMarkedDATETIMEByTruckNumber.Execute(SQLGetLastInvoiceMarkedDATETIMEByTruckNumber)
	
	
	If not rsGetLastInvoiceMarkedDATETIMEByTruckNumber.eof then
		resultGetLastInvoiceMarkedDATETIMEByTruckNumber = rsGetLastInvoiceMarkedDATETIMEByTruckNumber("LastDeliveryStatusChange")
	End If
	
	set rsGetLastInvoiceMarkedDATETIMEByTruckNumber= Nothing
	set cnnGetLastInvoiceMarkedDATETIMEByTruckNumber= Nothing
	
	GetLastInvoiceMarkedDATETIMEByTruckNumber = resultGetLastInvoiceMarkedDATETIMEByTruckNumber
	
End Function

Function GetNumberOfNagMessagesSentOnDate(passedUserNo, passedNagType, passedDate)

	Set cnnGetNumberOfNagMessagesSentOnDate = Server.CreateObject("ADODB.Connection")
	cnnGetNumberOfNagMessagesSentOnDate.open Session("ClientCnnString")

	resultGetNumberOfNagMessagesSentOnDateHistorical = 0
	
		
	SQLGetNumberOfNagMessagesSentOnDate = "SELECT COUNT(*) As Expr1 FROM SC_NagsSent WHERE "
	SQLGetNumberOfNagMessagesSentOnDate = SQLGetNumberOfNagMessagesSentOnDate & "UserNoSentToIfApplicable = " & passedUserNo & " AND "
	SQLGetNumberOfNagMessagesSentOnDate = SQLGetNumberOfNagMessagesSentOnDate & "NagType = '" & passedNagType & "' AND "
	SQLGetNumberOfNagMessagesSentOnDate = SQLGetNumberOfNagMessagesSentOnDate & "Year(RecordCreationDateTime) = " & Year(passedDate) & " AND "
	SQLGetNumberOfNagMessagesSentOnDate = SQLGetNumberOfNagMessagesSentOnDate & "Month(RecordCreationDateTime) = " & Month(passedDate) & " AND "
	SQLGetNumberOfNagMessagesSentOnDate = SQLGetNumberOfNagMessagesSentOnDate & "Day(RecordCreationDateTime) = " & Day(passedDate)

	Set rsGetNumberOfNagMessagesSentOnDate = Server.CreateObject("ADODB.Recordset")
	rsGetNumberOfNagMessagesSentOnDate.CursorLocation = 3 
	Set rsGetNumberOfNagMessagesSentOnDate= cnnGetNumberOfNagMessagesSentOnDate.Execute(SQLGetNumberOfNagMessagesSentOnDate)
	
	If not rsGetNumberOfNagMessagesSentOnDate.eof then resultGetNumberOfNagMessagesSentOnDate = rsGetNumberOfNagMessagesSentOnDate("Expr1")

	set rsGetNumberOfNagMessagesSentOnDate= Nothing
	set cnnGetNumberOfNagMessagesSentOnDate= Nothing
	
	GetNumberOfNagMessagesSentOnDate = resultGetNumberOfNagMessagesSentOnDate
	
End Function

Function GetLastNagSentTime(passedUserNo, passedNagType)

	Set cnnGetLastNagSentTime = Server.CreateObject("ADODB.Connection")
	cnnGetLastNagSentTime.open Session("ClientCnnString")

	resultGetLastNagSentTimeHistorical = ""
	
	SQLGetLastNagSentTime = "SELECT Top 1 RecordCreationDateTime FROM SC_NagsSent WHERE "
	SQLGetLastNagSentTime = SQLGetLastNagSentTime & "UserNoSentToIfApplicable = " & passedUserNo & " AND "
	SQLGetLastNagSentTime = SQLGetLastNagSentTime & "NagType = '" & passedNagType & "' Order By RecordCreationDateTime DESC"

	Set rsGetLastNagSentTime = Server.CreateObject("ADODB.Recordset")
	rsGetLastNagSentTime.CursorLocation = 3 
	Set rsGetLastNagSentTime= cnnGetLastNagSentTime.Execute(SQLGetLastNagSentTime)
	
	If not rsGetLastNagSentTime.eof then resultGetLastNagSentTime = rsGetLastNagSentTime("RecordCreationDateTime")

	set rsGetLastNagSentTime= Nothing
	set cnnGetLastNagSentTime= Nothing
	
	GetLastNagSentTime = resultGetLastNagSentTime
	
End Function

Function GetLastDeliveryStatusChangeBYTruck(passedTruckNumber)

	Set cnnGetLastDeliveryStatusChangeBYTruck = Server.CreateObject("ADODB.Connection")
	cnnGetLastDeliveryStatusChangeBYTruck.open Session("ClientCnnString")

	resultGetLastDeliveryStatusChangeBYTruck = ""
		
	SQLGetLastDeliveryStatusChangeBYTruck = "Select Top 1 LastDeliveryStatusChange from " & Session("SQL_Owner") & ".RT_DeliveryBoard where TruckNumber = '" & passedTruckNumber & "' ORDER BY LastDeliveryStatusChange DESC"

	Set rsGetLastDeliveryStatusChangeBYTruck = Server.CreateObject("ADODB.Recordset")
	rsGetLastDeliveryStatusChangeBYTruck.CursorLocation = 3 
	Set rsGetLastDeliveryStatusChangeBYTruck= cnnGetLastDeliveryStatusChangeBYTruck.Execute(SQLGetLastDeliveryStatusChangeBYTruck)
	
	
	If not rsGetLastDeliveryStatusChangeBYTruck.eof then resultGetLastDeliveryStatusChangeBYTruck = rsGetLastDeliveryStatusChangeBYTruck("LastDeliveryStatusChange")
	
	set rsGetLastDeliveryStatusChangeBYTruck= Nothing
	set cnnGetLastDeliveryStatusChangeBYTruck= Nothing
	
	GetLastDeliveryStatusChangeBYTruck = resultGetLastDeliveryStatusChangeBYTruck
	
End Function

Function GetNumberOfNagMessagesSentSinceDateTime(passedUserNo, passedNagType, passedDate)

	Set cnnGetNumberOfNagMessagesSentSinceDateTime = Server.CreateObject("ADODB.Connection")
	cnnGetNumberOfNagMessagesSentSinceDateTime.open Session("ClientCnnString")

	resultGetNumberOfNagMessagesSentSinceDateTimeHistorical = 0
	
		
	SQLGetNumberOfNagMessagesSentSinceDateTime = "SELECT COUNT(*) As Expr1 FROM SC_NagsSent WHERE "
	SQLGetNumberOfNagMessagesSentSinceDateTime = SQLGetNumberOfNagMessagesSentSinceDateTime & "UserNoSentToIfApplicable = " & passedUserNo & " AND "
	SQLGetNumberOfNagMessagesSentSinceDateTime = SQLGetNumberOfNagMessagesSentSinceDateTime & "NagType = '" & passedNagType & "' AND "
	SQLGetNumberOfNagMessagesSentSinceDateTime = SQLGetNumberOfNagMessagesSentSinceDateTime & "RecordCreationDateTime > '" & passedDate & "'"

	Set rsGetNumberOfNagMessagesSentSinceDateTime = Server.CreateObject("ADODB.Recordset")
	rsGetNumberOfNagMessagesSentSinceDateTime.CursorLocation = 3 
	Set rsGetNumberOfNagMessagesSentSinceDateTime= cnnGetNumberOfNagMessagesSentSinceDateTime.Execute(SQLGetNumberOfNagMessagesSentSinceDateTime)
	
	If not rsGetNumberOfNagMessagesSentSinceDateTime.eof then resultGetNumberOfNagMessagesSentSinceDateTime = rsGetNumberOfNagMessagesSentSinceDateTime("Expr1")

	set rsGetNumberOfNagMessagesSentSinceDateTime= Nothing
	set cnnGetNumberOfNagMessagesSentSinceDateTime= Nothing
	
	GetNumberOfNagMessagesSentSinceDateTime = resultGetNumberOfNagMessagesSentSinceDateTime
	
End Function

Function DriverInNagSkipTable(passedDriverUserNo,passedNagType)

	resultDriverInNagSkipTable = False

	Set cnnDriverInNagSkipTable = Server.CreateObject("ADODB.Connection")
	cnnDriverInNagSkipTable.open Session("ClientCnnString")

	SQLDriverInNagSkipTable = "SELECT * FROM SC_NagSkipUsers WHERE UserNo = " & passedDriverUserNo & " AND NagType = '" & passedNagType & "'"
	 
	Set rsDriverInNagSkipTable = Server.CreateObject("ADODB.Recordset")
	rsDriverInNagSkipTable.CursorLocation = 3 
	Set rsDriverInNagSkipTable = cnnDriverInNagSkipTable.Execute(SQLDriverInNagSkipTable)
	
	If NOT rsDriverInNagSkipTable.EOF Then resultDriverInNagSkipTable = True
	
	Set rsDriverInNagSkipTable = Nothing
	
	cnnDriverInNagSkipTable.Close
	Set cnnDriverInNagSkipTable = Nothing
	
	DriverInNagSkipTable = resultDriverInNagSkipTable 
	
End Function

Function DeliveryInProgress(passedIvsNum)

	Set cnnDeliveryInProgress = Server.CreateObject("ADODB.Connection")
	cnnDeliveryInProgress.open Session("ClientCnnString")

	resultDeliveryInProgress= False
		
	SQLDeliveryInProgress = "Select * from RT_DeliveryBoard WHERE IvsNum = " & passedIvsNum
	 
	Set rsDeliveryInProgress = Server.CreateObject("ADODB.Recordset")
	rsDeliveryInProgress.CursorLocation = 3 
	Set rsDeliveryInProgress= cnnDeliveryInProgress.Execute(SQLDeliveryInProgress)
	
	If not rsDeliveryInProgress.eof then 
		If rsDeliveryInProgress("DeliveryInProgress") = 1 Then resultDeliveryInProgress =  True
	End If
		
	rsDeliveryInProgress.Close
	set rsDeliveryInProgress= Nothing
	cnnDeliveryInProgress.Close	
	set cnnDeliveryInProgress = Nothing
	
	DeliveryInProgress = resultDeliveryInProgress 
	
End Function

Function DeliveryInProgressByCust(passedCustID)

	Set cnnDeliveryInProgressByCust = Server.CreateObject("ADODB.Connection")
	cnnDeliveryInProgressByCust.open Session("ClientCnnString")

	resultDeliveryInProgressByCust= False
		
	SQLDeliveryInProgressByCust = "Select * from RT_DeliveryBoard WHERE CustNum = '" & passedCustID & "' AND DeliveryInProgress = 1"
	 
	Set rsDeliveryInProgressByCust = Server.CreateObject("ADODB.Recordset")
	rsDeliveryInProgressByCust.CursorLocation = 3 
	Set rsDeliveryInProgressByCust= cnnDeliveryInProgressByCust.Execute(SQLDeliveryInProgressByCust)
	
	If not rsDeliveryInProgressByCust.eof then resultDeliveryInProgressByCust =  True
		
	rsDeliveryInProgressByCust.Close
	set rsDeliveryInProgressByCust= Nothing
	cnnDeliveryInProgressByCust.Close	
	set cnnDeliveryInProgressByCust = Nothing
	
	DeliveryInProgressByCust = resultDeliveryInProgressByCust 
	
End Function


Function DeliveryIsPriority(passedIvsNum)

	Set cnnDeliveryIsPriority = Server.CreateObject("ADODB.Connection")
	cnnDeliveryIsPriority.open Session("ClientCnnString")

	resultDeliveryIsPriority = False
		
	SQLDeliveryIsPriority = "Select * from RT_DeliveryBoard WHERE IvsNum = " & passedIvsNum
	 
	Set rsDeliveryIsPriority = Server.CreateObject("ADODB.Recordset")
	rsDeliveryIsPriority.CursorLocation = 3 
	Set rsDeliveryIsPriority= cnnDeliveryIsPriority.Execute(SQLDeliveryIsPriority)
	
	If not rsDeliveryIsPriority.eof then 
		If rsDeliveryIsPriority("Priority") = 1 Then resultDeliveryIsPriority =  True
	End If
		
	rsDeliveryIsPriority.Close
	set rsDeliveryIsPriority= Nothing
	cnnDeliveryIsPriority.Close	
	set cnnDeliveryIsPriority = Nothing
	
	DeliveryIsPriority = resultDeliveryIsPriority 
	
End Function



Function DeliveryIsAM(passedIvsNum)

	Set cnnDeliveryIsAM = Server.CreateObject("ADODB.Connection")
	cnnDeliveryIsAM.open Session("ClientCnnString")

	resultDeliveryIsAM = False
		
	SQLDeliveryIsAM = "Select * from RT_DeliveryBoard WHERE IvsNum = " & passedIvsNum
	 
	Set rsDeliveryIsAM = Server.CreateObject("ADODB.Recordset")
	rsDeliveryIsAM.CursorLocation = 3 
	Set rsDeliveryIsAM= cnnDeliveryIsAM.Execute(SQLDeliveryIsAM)
	
	If not rsDeliveryIsAM.eof then 
		If rsDeliveryIsAM("AMorPM") = "AM" Then resultDeliveryIsAM =  True
	End If
		
	rsDeliveryIsAM.Close
	set rsDeliveryIsAM= Nothing
	cnnDeliveryIsAM.Close	
	set cnnDeliveryIsAM = Nothing
	
	DeliveryIsAM = resultDeliveryIsAM 
	
End Function



%>

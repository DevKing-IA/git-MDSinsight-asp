<%	

	
	Set cnnCheckRTDeliveryBoard = Server.CreateObject("ADODB.Connection")
	cnnCheckRTDeliveryBoard.open (Session("ClientCnnString"))
	Set rsCheckRTDeliveryBoard = Server.CreateObject("ADODB.Recordset")
	rsCheckRTDeliveryBoard.CursorLocation = 3 
	
	SQL_CheckRTDeliveryBoard = "SELECT COL_LENGTH('RT_DeliveryBoard', 'Priority') AS IsItThere"
	Set rsCheckRTDeliveryBoard = cnnCheckRTDeliveryBoard.Execute(SQL_CheckRTDeliveryBoard)
	If IsNull(rsCheckRTDeliveryBoard("IsItThere")) Then
		SQL_CheckRTDeliveryBoard = "ALTER TABLE RT_DeliveryBoard ADD Priority INT NOT NULL DEFAULT 0"
		Set rsCheckRTDeliveryBoard = cnnCheckRTDeliveryBoard.Execute(SQL_CheckRTDeliveryBoard)
	End If
	
	SQL_CheckRTDeliveryBoard = "SELECT COL_LENGTH('RT_DeliveryBoardPending', 'Priority') AS IsItThere"
	Set rsCheckRTDeliveryBoard = cnnCheckRTDeliveryBoard.Execute(SQL_CheckRTDeliveryBoard)
	If IsNull(rsCheckRTDeliveryBoard("IsItThere")) Then
		SQL_CheckRTDeliveryBoard = "ALTER TABLE RT_DeliveryBoardPending ADD Priority INT NOT NULL DEFAULT 0"
		Set rsCheckRTDeliveryBoard = cnnCheckRTDeliveryBoard.Execute(SQL_CheckRTDeliveryBoard)
	End If

	SQL_CheckRTDeliveryBoard = "SELECT COL_LENGTH('RT_DeliveryBoardHistory', 'Priority') AS IsItThere"
	Set rsCheckRTDeliveryBoard = cnnCheckRTDeliveryBoard.Execute(SQL_CheckRTDeliveryBoard)
	If IsNull(rsCheckRTDeliveryBoard("IsItThere")) Then
		SQL_CheckRTDeliveryBoard = "ALTER TABLE RT_DeliveryBoardHistory ADD Priority INT NOT NULL DEFAULT 0"
		Set rsCheckRTDeliveryBoard = cnnCheckRTDeliveryBoard.Execute(SQL_CheckRTDeliveryBoard)
	End If
	
	Set rsCheckRTDeliveryBoard = Nothing
	cnnCheckRTDeliveryBoard.Close
	Set cnnCheckRTDeliveryBoard = Nothing

				
%>
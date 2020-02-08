<%	

	on error goto 0
	
	Set cnnCheckRTDeliveryBoardHistory = Server.CreateObject("ADODB.Connection")
	cnnCheckRTDeliveryBoardHistory.open (Session("ClientCnnString"))
	Set rsCheckRTDeliveryBoardHistory = Server.CreateObject("ADODB.Recordset")
	rsCheckRTDeliveryBoardHistory.CursorLocation = 3 
	
	SQL_CheckRTDeliveryBoardHistory = "SELECT COL_LENGTH('RT_DeliveryBoardHistory', 'FieldServiceNotesReportSendTime') AS IsItThere"
	Set rsCheckRTDeliveryBoardHistory = cnnCheckRTDeliveryBoardHistory.Execute(SQL_CheckRTDeliveryBoardHistory)
	If NOT IsNull(rsCheckRTDeliveryBoardHistory("IsItThere")) Then
		SQL_CheckRTDeliveryBoardHistory = "ALTER TABLE RT_DeliveryBoardHistory DROP Column FieldServiceNotesReportSendTime"
'		Response.Write(CheckRTDeliveryBoardHistory & "<br>")
		Set rsCheckRTDeliveryBoardHistory = cnnCheckRTDeliveryBoardHistory.Execute(SQL_CheckRTDeliveryBoardHistory)
	End If
	
	SQL_CheckRTDeliveryBoardHistory = "SELECT COL_LENGTH('RT_DeliveryBoardHistory', 'FSBoardKioskGlobalColorDispatchAcknowledged') AS IsItThere"
	Set rsCheckRTDeliveryBoardHistory = cnnCheckRTDeliveryBoardHistory.Execute(SQL_CheckRTDeliveryBoardHistory)
	If NOT IsNull(rsCheckRTDeliveryBoardHistory("IsItThere")) Then
		SQL_CheckRTDeliveryBoardHistory = "ALTER TABLE RT_DeliveryBoardHistory DROP Column FSBoardKioskGlobalColorDispatchAcknowledged"
'		Response.Write(CheckRTDeliveryBoardHistory & "<br>")		
		Set rsCheckRTDeliveryBoardHistory = cnnCheckRTDeliveryBoardHistory.Execute(SQL_CheckRTDeliveryBoardHistory)
	End If
	
	SQL_CheckRTDeliveryBoardHistory = "SELECT COL_LENGTH('RT_DeliveryBoardHistory', 'FSBoardDispatchedColor') AS IsItThere"
	Set rsCheckRTDeliveryBoardHistory = cnnCheckRTDeliveryBoardHistory.Execute(SQL_CheckRTDeliveryBoardHistory)
	If NOT IsNull(rsCheckRTDeliveryBoardHistory("IsItThere")) Then
		SQL_CheckRTDeliveryBoardHistory = "ALTER TABLE RT_DeliveryBoardHistory DROP Column FSBoardDispatchedColor"
'		Response.Write(CheckRTDeliveryBoardHistory & "<br>")		
		Set rsCheckRTDeliveryBoardHistory = cnnCheckRTDeliveryBoardHistory.Execute(SQL_CheckRTDeliveryBoardHistory)
	End If
	
	
	Set rsCheckRTDeliveryBoardHistory = Nothing
	cnnCheckRTDeliveryBoardHistory.Close
	Set cnnCheckRTDeliveryBoardHistory = Nothing

				
%>
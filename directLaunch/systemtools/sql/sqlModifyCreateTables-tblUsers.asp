<%	
		
	Set cnnCheckUsers = Server.CreateObject("ADODB.Connection")
	cnnCheckUsers.open (Session("ClientCnnString"))
	Set rsCheckUsers = Server.CreateObject("ADODB.Recordset")
	rsCheckUsers.CursorLocation = 3 
			
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userLoginDisableAccessHolidays') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userLoginDisableAccessHolidays INT NOT NULL DEFAULT 0"
		Set rsCheckUsers= cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userMobileInventoryControlAccess') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userMobileInventoryControlAccess BIT NOT NULL DEFAULT 0"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
		SQLCheckUsers = "UPDATE tblUsers SET userMobileInventoryControlAccess = 0"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userForceNextStopSelectionOverride') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userForceNextStopSelectionOverride varchar(50) NOT NULL DEFAULT 'Use Global'"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
		SQLCheckUsers = "UPDATE tblUsers SET userForceNextStopSelectionOverride = 'Use Global'"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNextStopNagMessageOverride') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNextStopNagMessageOverride Varchar(255) NOT NULL DEFAULT 'Use Global'"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
		SQLCheckUsers = "UPDATE tblUsers SET userNextStopNagMessageOverride = 'Use Global'"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNextStopNagMinutes') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNextStopNagMinutes INT NOT NULL DEFAULT 30"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNextStopNagIntervalMinutes') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNextStopNagIntervalMinutes INT NOT NULL DEFAULT 30"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNextStopNagMessageMaxToSendPerStop') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNextStopNagMessageMaxToSendPerStop INT NOT NULL DEFAULT 10"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNextStopNagMessageMaxToSendThisDriverPerDay') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNextStopNagMessageMaxToSendThisDriverPerDay INT NOT NULL DEFAULT 10"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNextStopNagMessageSendMethod') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNextStopNagMessageSendMethod Varchar(255) NOT NULL DEFAULT 'Text'"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNoActivityNagMessageOverride') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNoActivityNagMessageOverride Varchar(255) NOT NULL DEFAULT 'Use Global'"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNoActivityNagMinutes') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNoActivityNagMinutes INT NOT NULL DEFAULT 30"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNoActivityNagIntervalMinutes') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNoActivityNagIntervalMinutes INT NOT NULL DEFAULT 30"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNoActivityNagMessageMaxToSendPerStop') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNoActivityNagMessageMaxToSendPerStop INT NOT NULL DEFAULT 10"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNoActivityNagMessageMaxToSendPerDriverPerDay') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNoActivityNagMessageMaxToSendPerDriverPerDay INT NOT NULL DEFAULT 10"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNoActivityNagMessageSendMethod') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNoActivityNagMessageSendMethod Varchar(255) NOT NULL DEFAULT 'Text'"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNoActivityNagTimeOfDay') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNoActivityNagTimeOfDay Varchar(255) NULL"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userVMS_ID') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userVMS_ID Varchar(255) NULL"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userLoginDisableAccessHolidays') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userLoginDisableAccessHolidays INT NOT NULL DEFAULT 0"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userInventoryControlAccessType') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userInventoryControlAccessType Varchar(255) NULL"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
		SQLCheckUsers = "UPDATE tblUsers SET userInventoryControlAccessType = 'NONE'"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userForceNextStopSelectionOverride_FS') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userForceNextStopSelectionOverride_FS varchar(50) NOT NULL DEFAULT 'Use Global'"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
		SQLCheckUsers = "UPDATE tblUsers SET userForceNextStopSelectionOverride_FS = 'Use Global'"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNextStopNagMessageOverride_FS') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNextStopNagMessageOverride_FS Varchar(255) NOT NULL DEFAULT 'Use Global'"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
		SQLCheckUsers = "UPDATE tblUsers SET userNextStopNagMessageOverride_FS = 'Use Global'"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNextStopNagMinutes_FS') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNextStopNagMinutes_FS INT NOT NULL DEFAULT 30"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNextStopNagIntervalMinutes_FS') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNextStopNagIntervalMinutes_FS INT NOT NULL DEFAULT 30"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNextStopNagMessageMaxToSendPerStop_FS') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNextStopNagMessageMaxToSendPerStop_FS INT NOT NULL DEFAULT 10"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNextStopNagMessageMaxToSendThisDriverPerDay_FS') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNextStopNagMessageMaxToSendThisDriverPerDay_FS INT NOT NULL DEFAULT 10"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNextStopNagMessageSendMethod_FS') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNextStopNagMessageSendMethod_FS Varchar(255) NOT NULL DEFAULT 'Text'"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNoActivityNagMessageOverride_FS') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNoActivityNagMessageOverride_FS Varchar(255) NOT NULL DEFAULT 'Use Global'"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNoActivityNagMinutes_FS') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNoActivityNagMinutes_FS INT NOT NULL DEFAULT 30"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNoActivityNagIntervalMinutes_FS') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNoActivityNagIntervalMinutes_FS INT NOT NULL DEFAULT 30"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNoActivityNagMessageMaxToSendPerStop_FS') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNoActivityNagMessageMaxToSendPerStop_FS INT NOT NULL DEFAULT 10"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNoActivityNagMessageMaxToSendPerDriverPerDay_FS') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNoActivityNagMessageMaxToSendPerDriverPerDay_FS INT NOT NULL DEFAULT 10"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNoActivityNagMessageSendMethod_FS') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNoActivityNagMessageSendMethod_FS Varchar(255) NOT NULL DEFAULT 'Text'"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNoActivityNagTimeOfDay_FS') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNoActivityNagTimeOfDay_FS Varchar(255) NULL"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userMobileInventoryControlAccess') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userMobileInventoryControlAccess BIT NOT NULL DEFAULT 0"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
		SQLCheckUsers = "UPDATE tblUsers SET userMobileInventoryControlAccess = 0"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userForceNextStopSelectionOverride') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userForceNextStopSelectionOverride varchar(50) NOT NULL DEFAULT 'Use Global'"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
		SQLCheckUsers = "UPDATE tblUsers SET userForceNextStopSelectionOverride = 'Use Global'"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNextStopNagMessageOverride') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNextStopNagMessageOverride Varchar(255) NOT NULL DEFAULT 'Use Global'"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
		SQLCheckUsers = "UPDATE tblUsers SET userNextStopNagMessageOverride = 'Use Global'"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNextStopNagMinutes') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNextStopNagMinutes INT NOT NULL DEFAULT 30"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNextStopNagIntervalMinutes') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNextStopNagIntervalMinutes INT NOT NULL DEFAULT 30"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNextStopNagMessageMaxToSendPerStop') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNextStopNagMessageMaxToSendPerStop INT NOT NULL DEFAULT 10"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNextStopNagMessageMaxToSendThisDriverPerDay') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNextStopNagMessageMaxToSendThisDriverPerDay INT NOT NULL DEFAULT 10"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNextStopNagMessageSendMethod') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNextStopNagMessageSendMethod Varchar(255) NOT NULL DEFAULT 'Text'"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNoActivityNagMessageOverride') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNoActivityNagMessageOverride  Varchar(255) NOT NULL DEFAULT 'Use Global'"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNoActivityNagMinutes') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNoActivityNagMinutes INT NOT NULL DEFAULT 30"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNoActivityNagIntervalMinutes') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNoActivityNagIntervalMinutes INT NOT NULL DEFAULT 30"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNoActivityNagMessageMaxToSendPerStop') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNoActivityNagMessageMaxToSendPerStop INT NOT NULL DEFAULT 10"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNoActivityNagMessageMaxToSendPerDriverPerDay') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNoActivityNagMessageMaxToSendPerDriverPerDay INT NOT NULL DEFAULT 10"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNoActivityNagMessageSendMethod') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNoActivityNagMessageSendMethod Varchar(255) NOT NULL DEFAULT 'Text'"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNoActivityNagTimeOfDay') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNoActivityNagTimeOfDay Varchar(255) NULL"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userVMS_ID') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userVMS_ID Varchar(255) NULL"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userInventoryControlAccessType') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userInventoryControlAccessType Varchar(255) NULL"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
		SQLCheckUsers = "UPDATE tblUsers SET userInventoryControlAccessType = 'NONE'"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userForceNextStopSelectionOverride_FS') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userForceNextStopSelectionOverride_FS varchar(50) NOT NULL DEFAULT 'Use Global'"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
		SQLCheckUsers = "UPDATE tblUsers SET userForceNextStopSelectionOverride_FS = 'Use Global'"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNextStopNagMessageOverride_FS') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNextStopNagMessageOverride_FS Varchar(255) NOT NULL DEFAULT 'Use Global'"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
		SQLCheckUsers = "UPDATE tblUsers SET userNextStopNagMessageOverride_FS = 'Use Global'"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNextStopNagMinutes_FS') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNextStopNagMinutes_FS INT NOT NULL DEFAULT 30"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNextStopNagIntervalMinutes_FS') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNextStopNagIntervalMinutes_FS INT NOT NULL DEFAULT 30"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNextStopNagMessageMaxToSendPerStop_FS') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNextStopNagMessageMaxToSendPerStop_FS INT NOT NULL DEFAULT 10"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNextStopNagMessageMaxToSendThisDriverPerDay_FS') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNextStopNagMessageMaxToSendThisDriverPerDay_FS INT NOT NULL DEFAULT 10"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNextStopNagMessageSendMethod_FS') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNextStopNagMessageSendMethod_FS Varchar(255) NOT NULL DEFAULT 'Text'"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNoActivityNagMessageOverride_FS') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNoActivityNagMessageOverride_FS Varchar(255) NOT NULL DEFAULT 'Use Global'"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNoActivityNagMinutes_FS') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNoActivityNagMinutes_FS INT NOT NULL DEFAULT 30"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNoActivityNagIntervalMinutes_FS') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNoActivityNagIntervalMinutes_FS INT NOT NULL DEFAULT 30"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNoActivityNagMessageMaxToSendPerStop_FS') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNoActivityNagMessageMaxToSendPerStop_FS INT NOT NULL DEFAULT 10"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNoActivityNagMessageMaxToSendPerDriverPerDay_FS') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNoActivityNagMessageMaxToSendPerDriverPerDay_FS INT NOT NULL DEFAULT 10"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNoActivityNagMessageSendMethod_FS') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNoActivityNagMessageSendMethod_FS Varchar(255) NOT NULL DEFAULT 'Text'"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userNoActivityNagTimeOfDay_FS') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userNoActivityNagTimeOfDay_FS Varchar(255) NULL"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userSalesPersonNumber') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userSalesPersonNumber int NULL"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	'***********************************************
	'If there are no user records in the users table
	'create the rsmith user
	'***********************************************
	On Error Resume Next
	rsCheckUsers.Close 
	On Error Goto 0
	rsCheckUsers.CursorLocation = 3
	SQLCheckUsers = "SELECT * FROM tblUsers WHERE userEmail = 'rsmith@ocsaccess.com'"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If rsCheckUsers.EOF Then
		SQLCheckUsers = "INSERT INTO tblUsers ("
		SQLCheckUsers = SQLCheckUsers & "userFirstName "
		SQLCheckUsers = SQLCheckUsers & ", userLastName "
		SQLCheckUsers = SQLCheckUsers & ", userEmail "				
		SQLCheckUsers = SQLCheckUsers & ", userPassword "				
		SQLCheckUsers = SQLCheckUsers & ", userEnabled "				
		SQLCheckUsers = SQLCheckUsers & ", userAdmin "				
		SQLCheckUsers = SQLCheckUsers & ", userDisplayname "
		SQLCheckUsers = SQLCheckUsers & ", userCellNumber "				
		SQLCheckUsers = SQLCheckUsers & ", userType "				
		SQLCheckUsers = SQLCheckUsers & ", userArchived "								
		SQLCheckUsers = SQLCheckUsers & ", userLicense "				
		SQLCheckUsers = SQLCheckUsers & ", userLicenseExpiration "				
		SQLCheckUsers = SQLCheckUsers & ") VALUES ("
		SQLCheckUsers = SQLCheckUsers & "'Rich'"			
		SQLCheckUsers = SQLCheckUsers & ",'Smith'"					
		SQLCheckUsers = SQLCheckUsers & ",'rsmith@ocsaccess.com'"					
		SQLCheckUsers = SQLCheckUsers & ",'changeme'"			
		SQLCheckUsers = SQLCheckUsers & ",1"			
		SQLCheckUsers = SQLCheckUsers & ",1"			
		SQLCheckUsers = SQLCheckUsers & ",'Rich'"			
		SQLCheckUsers = SQLCheckUsers & ",'6099294430'"			
		SQLCheckUsers = SQLCheckUsers & ",'Admin'"			
		SQLCheckUsers = SQLCheckUsers & ",0"			
		SQLCheckUsers = SQLCheckUsers & ",'Programming'"			
		SQLCheckUsers = SQLCheckUsers & ",'" & DateAdd("y",1,Now())& "'"			
		SQLCheckUsers = SQLCheckUsers & ")"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
		
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userSalesPersonNumber2') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userSalesPersonNumber2 int NULL"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userEditEqpOnTheFly') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userEditEqpOnTheFly INT NOT NULL DEFAULT 0"
		Set rsCheckUsers= cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'UseNewMobileLogic') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD UseNewMobileLogic INT NOT NULL DEFAULT 0"
		Set rsCheckUsers= cnnCheckUsers.Execute(SQLCheckUsers)
		SQLCheckUsers = "UPDATE tblUsers SET UseNewMobileLogic = 0"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If

	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userReceivePartsRequestEmails') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userReceivePartsRequestEmails INT NOT NULL DEFAULT 0"
		Set rsCheckUsers= cnnCheckUsers.Execute(SQLCheckUsers)
		SQLCheckUsers = "UPDATE tblUsers SET userReceivePartsRequestEmails = 0"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
		SQLCheckUsers = "UPDATE tblUsers SET userReceivePartsRequestEmails = 1 WHERE userType = 'Service Manager'"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userEditCRMOnTheFly') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userEditCRMOnTheFly BIT NOT NULL DEFAULT 0"
		Set rsCheckUsers= cnnCheckUsers.Execute(SQLCheckUsers)
	End If

	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'loginLandingPageURL') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD loginLandingPageURL varchar(1000) NULL"
		Set rsCheckUsers= cnnCheckUsers.Execute(SQLCheckUsers)
	End If

	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userEmailServer') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userEmailServer varchar(1000) NULL"
		Set rsCheckUsers= cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
on error goto 0	
	'This code will rename a column
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userProspectingAddEditAccess') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If NOT IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "sp_rename 'tblUsers.userProspectingAddEditAccess','userProspectingAddEditAccess','COLUMN'"
		Set rsCheckUsers= cnnCheckUsers.Execute(SQLCheckUsers)
	End If

	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userRegionsToViewService') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userRegionsToViewService varchar(1000) NULL"
		Set rsCheckUsers= cnnCheckUsers.Execute(SQLCheckUsers)
	End If

	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userCreateEquipmentSymptomCodesOnTheFly') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userCreateEquipmentSymptomCodesOnTheFly INT NOT NULL DEFAULT 0"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If

	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userCreateEquipmentResolutionCodesOnTheFly') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userCreateEquipmentResolutionCodesOnTheFly INT NOT NULL DEFAULT 0"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If
	
	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userCreateEquipmentProblemCodesOnTheFly') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userCreateEquipmentProblemCodesOnTheFly INT NOT NULL DEFAULT 0"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If


	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userCreateNewServiceTicket') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userCreateNewServiceTicket INT NOT NULL DEFAULT 0"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If

	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userAccessServiceDispatchCenter') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userAccessServiceDispatchCenter INT NOT NULL DEFAULT 0"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If

	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userAccessServiceActionsModalButton') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userAccessServiceActionsModalButton INT NOT NULL DEFAULT 0"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If

	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userAccessServiceDispatchButton') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userAccessServiceDispatchButton INT NOT NULL DEFAULT 0"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If

	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userAccessServiceCloseCancelButton') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userAccessServiceCloseCancelButton INT NOT NULL DEFAULT 0"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If

	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userLeftNavAPIModule') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userLeftNavAPIModule INT NOT NULL DEFAULT 0"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If

	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userLeftNavBIModule') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userLeftNavBIModule INT NOT NULL DEFAULT 0"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If

	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userLeftNavProspectingModule') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userLeftNavProspectingModule INT NOT NULL DEFAULT 0"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If

	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userLeftNavCustomerServiceModule') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userLeftNavCustomerServiceModule INT NOT NULL DEFAULT 0"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If

	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userLeftNavEquipmentModule') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userLeftNavEquipmentModule INT NOT NULL DEFAULT 0"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If

	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userLeftNavInventoryControlModule') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userLeftNavInventoryControlModule INT NOT NULL DEFAULT 0"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If

	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userLeftNavAccountsReceivableModule') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userLeftNavAccountsReceivableModule INT NOT NULL DEFAULT 0"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If

	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userLeftNavAccountsPayableModule') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userLeftNavAccountsPayableModule INT NOT NULL DEFAULT 0"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If

	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userLeftNavServiceModule') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userLeftNavServiceModule INT NOT NULL DEFAULT 0"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If

	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userLeftNavRoutingModule') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userLeftNavRoutingModule INT NOT NULL DEFAULT 0"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If

	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userLeftNavQuickbooksModule') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userLeftNavQuickbooksModule INT NOT NULL DEFAULT 0"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If

	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userLeftNavFiltertraxModule') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userLeftNavFiltertraxModule INT NOT NULL DEFAULT 0"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If

	SQLCheckUsers = "SELECT COL_LENGTH('tblUsers', 'userLeftNavSystem') AS IsItThere"
	Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	If IsNull(rsCheckUsers("IsItThere")) Then
		SQLCheckUsers = "ALTER TABLE tblUsers ADD userLeftNavSystem INT NOT NULL DEFAULT 0"
		Set rsCheckUsers = cnnCheckUsers.Execute(SQLCheckUsers)
	End If

	Set rsCheckUsers = Nothing
	cnnCheckUsers.Close
	Set cnnCheckUsers = Nothing
		
				
%>
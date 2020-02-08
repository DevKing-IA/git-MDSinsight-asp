<%	

	On error goto 0
	
	Set cnnCheckSettingsGlobal = Server.CreateObject("ADODB.Connection")
	cnnCheckSettingsGlobal.CommandTimeout = 120
	cnnCheckSettingsGlobal.open (Session("ClientCnnString"))
	Set rsCheckSettingsGlobal = Server.CreateObject("ADODB.Recordset")
	rsCheckSettingsGlobal.CursorLocation = 3 
			
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'FSWApp_SignatOpt') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD FSWApp_SignatOpt Int NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET FSWApp_SignatOpt = 0"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'FSWApp_HideAsset') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD FSWApp_HideAsset Int NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET FSWApp_HideAsset = 0"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'FSWApp_NoUnableDdwn') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD FSWApp_NoUnableDdwn Int NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET FSWApp_NoUnableDdwn = 0"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'FS_TechCanDecline') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD FS_TechCanDecline int NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET FS_TechCanDecline= 0"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
		
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'FS_TechCanDecline') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD FS_TechCanDecline int NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET FS_TechCanDecline= 0"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
			
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'ServiceColorsOn') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD ServiceColorsOn Int NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET ServiceColorsOn = 0"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'ServiceNormalAlertColor') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD ServiceNormalAlertColor varchar(50)"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'ServicePriorityColor') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD ServicePriorityColor varchar(50)"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'ServicePriorityAlertColor') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD ServicePriorityAlertColor varchar(50)"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'DoNotShowDeliveryLineItems') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD DoNotShowDeliveryLineItems INT NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'AutoForceSelectNextStop') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD AutoForceSelectNextStop INT NOT NULL DEFAULT 0"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'ProspSnapshotOnOff') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD ProspSnapshotOnOff INT NOT NULL DEFAULT 0"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'ProspSnapshotInsideSales') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD ProspSnapshotInsideSales INT NOT NULL DEFAULT 0"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'ProspSnapshotOutsideSales') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD ProspSnapshotOutsideSales INT NOT NULL DEFAULT 0"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'ProspSnapshotUserNos') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD ProspSnapshotUserNos varchar(1000)"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'ProspSnapshotSalesRepDisplayUserNos') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD ProspSnapshotSalesRepDisplayUserNos varchar(1000)"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'ProspSnapshotAdditionalEmails') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD ProspSnapshotAdditionalEmails varchar(1000)"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'ProspSnapshotEmailSubject') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD ProspSnapshotEmailSubject varchar(1000)"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'NextStopNagMessageONOFF') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD NextStopNagMessageONOFF INT NOT NULL DEFAULT 0"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'NextStopNagMinutes') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD NextStopNagMinutes INT NOT NULL DEFAULT 30"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'NextStopNagIntervalMinutes') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD NextStopNagIntervalMinutes INT NOT NULL DEFAULT 30"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'NextStopNagMessageMaxToSendPerStop') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD NextStopNagMessageMaxToSendPerStop INT NOT NULL DEFAULT 10"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'NextStopNagMessageMaxToSendPerDriverPerDay') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD NextStopNagMessageMaxToSendPerDriverPerDay INT NOT NULL DEFAULT 10"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'NextStopNagMessageSendMethod') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD NextStopNagMessageSendMethod Varchar(255) NOT NULL DEFAULT 'Text'"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'NoActivityNagMessageONOFF') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD NoActivityNagMessageONOFF INT NOT NULL DEFAULT 0"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'NoActivityNagMinutes') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD NoActivityNagMinutes INT NOT NULL DEFAULT 30"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'NoActivityNagIntervalMinutes') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD NoActivityNagIntervalMinutes INT NOT NULL DEFAULT 30"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'NoActivityNagMessageMaxToSendPerStop') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD NoActivityNagMessageMaxToSendPerStop INT NOT NULL DEFAULT 10"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'NoActivityNagMessageMaxToSendPerDriverPerDay') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD NoActivityNagMessageMaxToSendPerDriverPerDay INT NOT NULL DEFAULT 10"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'NoActivityNagMessageSendMethod') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD NoActivityNagMessageSendMethod Varchar(255) NOT NULL DEFAULT 'Text'"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'NoActivityNagTimeOfDay') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD NoActivityNagTimeOfDay Varchar(255) NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'NoActivityNagMessageONOFF_FS') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD NoActivityNagMessageONOFF_FS INT NOT NULL DEFAULT 0"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET NoActivityNagMessageONOFF_FS = 0"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'NoActivityNagMinutes_FS') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD NoActivityNagMinutes_FS INT NOT NULL DEFAULT 30"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET NoActivityNagMinutes_FS = 0"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'NoActivityNagIntervalMinutes_FS') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD NoActivityNagIntervalMinutes_FS INT NOT NULL DEFAULT 30"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET NoActivityNagIntervalMinutes_FS = 30"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'NoActivityNagMessageMaxToSendPerStop_FS') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD NoActivityNagMessageMaxToSendPerStop_FS INT NOT NULL DEFAULT 10"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET NoActivityNagMessageMaxToSendPerStop_FS = 10"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'NoActivityNagMessageMaxToSendPerDriverPerDay_FS') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD NoActivityNagMessageMaxToSendPerDriverPerDay_FS INT NOT NULL DEFAULT 10"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET NoActivityNagMessageMaxToSendPerDriverPerDay_FS = 10"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'NoActivityNagMessageSendMethod_FS') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD NoActivityNagMessageSendMethod_FS Varchar(255) NOT NULL DEFAULT 'Text'"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET NoActivityNagMessageSendMethod_FS = 'Text'"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'NoActivityNagTimeOfDay_FS') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD NoActivityNagTimeOfDay_FS Varchar(255) NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'NextStopNagMessageONOFF_FS') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD NextStopNagMessageONOFF_FS INT NOT NULL DEFAULT 0"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET NextStopNagMessageONOFF_FS = 0"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'NextStopNagMinutes_FS') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD NextStopNagMinutes_FS INT NOT NULL DEFAULT 30"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET NextStopNagMinutes_FS = 30"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'NextStopNagIntervalMinutes_FS') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD NextStopNagIntervalMinutes_FS INT NOT NULL DEFAULT 30"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET NextStopNagIntervalMinutes_FS = 30"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'NextStopNagMessageMaxToSendPerStop_FS') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD NextStopNagMessageMaxToSendPerStop_FS INT NOT NULL DEFAULT 10"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET NextStopNagMessageMaxToSendPerStop_FS = 10"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'NextStopNagMessageMaxToSendPerDriverPerDay_FS') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD NextStopNagMessageMaxToSendPerDriverPerDay_FS INT NOT NULL DEFAULT 10"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET NextStopNagMessageMaxToSendPerDriverPerDay_FS = 10"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'NextStopNagMessageSendMethod_FS') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD NextStopNagMessageSendMethod_FS Varchar(255) NOT NULL DEFAULT 'Text'"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET NextStopNagMessageSendMethod_FS = 'Text'"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'OrderAPIRepostONOFF') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD OrderAPIRepostONOFF INT NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET OrderAPIRepostONOFF = 0"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'OrderAPIRepostURL') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD OrderAPIRepostURL Varchar(255) NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'APIDailyActivityReportOnOff') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD APIDailyActivityReportOnOff INT NOT NULL DEFAULT 0"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'APIDailyActivityReportUserNos') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD APIDailyActivityReportUserNos varchar(1000)"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'InvoiceAPIRepostONOFF') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD InvoiceAPIRepostONOFF INT NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET InvoiceAPIRepostONOFF = 0"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'InvoiceAPIRepostURL') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD InvoiceAPIRepostURL Varchar(255) NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'CMAPIRepostONOFF') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD CMAPIRepostONOFF Int NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET CMAPIRepostONOFF = 0"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	
	End If
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'CMAPIRepostURL') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD CMAPIRepostURL Varchar(255) NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'RAAPIRepostONOFF') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD RAAPIRepostONOFF Int NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET RAAPIRepostONOFF = 0"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'APIDailyActivityReportAdditionalEmails') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD APIDailyActivityReportAdditionalEmails varchar(1000)"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'SumInvAPIRepostONOFF') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD SumInvAPIRepostONOFF Int NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET SumInvAPIRepostONOFF = 0"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'RAAPIRepostURL') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD RAAPIRepostURL Varchar(255) NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'SumInvAPIRepostURL') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD SumInvAPIRepostURL Varchar(255) NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'APIDailyActivityReportEmailSubject') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD APIDailyActivityReportEmailSubject varchar(1000)"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'OrderAPIRepostMode') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD OrderAPIRepostMode varchar(10)"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET OrderAPIRepostMode = 'TEST'"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'InvoiceAPIRepostMode') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD InvoiceAPIRepostMode varchar(10)"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET InvoiceAPIRepostMode = 'TEST'"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'RAAPIRepostMode') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD RAAPIRepostMode varchar(10)"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET RAAPIRepostMode = 'TEST'"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'CMAPIRepostMode') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD CMAPIRepostMode varchar(10)"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET CMAPIRepostMode= 'TEST'"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'SumInvAPIRepostMode') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD SumInvAPIRepostMode varchar(10)"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET SumInvAPIRepostMode= 'TEST'"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'SendInvoiceType') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD SendInvoiceType varchar(10)"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET SendInvoiceType = 'UNPOSTED'"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'OrderAPIOffsetDays') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD OrderAPIOffsetDays Int NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET OrderAPIOffsetDays = 1"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'InvoiceAPIOffsetDays') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD InvoiceAPIOffsetDays Int NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET InvoiceAPIOffsetDays = 1"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'RAAPIOffsetDays') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD RAAPIOffsetDays Int NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET RAAPIOffsetDays = 1"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'CMAPIOffsetDays') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD CMAPIOffsetDays Int NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET CMAPIOffsetDays = 1"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'SumInvAPIOffsetDays') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD SumInvAPIOffsetDays Int NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET SumInvAPIOffsetDays = 1"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'OrderCutoffTime') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD OrderCutoffTime varchar(255)"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET OrderCutoffTime = '1800'"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'InvoiceCutoffTime') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD InvoiceCutoffTime varchar(255)"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET InvoiceCutoffTime = '1800'"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'RACutoffTime') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD RACutoffTime varchar(255)"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET RACutoffTime = '1800'"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'CMCutoffTime') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD CMCutoffTime varchar(255)"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET CMCutoffTime = '1800'"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'FSWApp_SignatOpt') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD FSWApp_SignatOpt Int NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET FSWApp_SignatOpt = 0"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'FSWApp_HideAsset') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD FSWApp_HideAsset Int NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET FSWApp_HideAsset = 0"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'FSWApp_NoUnableDdwn') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD FSWApp_NoUnableDdwn Int NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET FSWApp_NoUnableDdwn = 0"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'InventoryAPIRepostONOFF') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD InventoryAPIRepostONOFF INT NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET InventoryAPIRepostONOFF = 0"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'InventoryAPIRepostMode') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD InventoryAPIRepostMode varchar(10)"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET InventoryAPIRepostMode = 'TEST'"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'InventoryAPIRepostURL') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD InventoryAPIRepostURL Varchar(255) NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If

	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'InventoryAPIRepostOnHandONOFF') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD InventoryAPIRepostOnHandONOFF INT NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET InventoryAPIRepostOnHandONOFF = 0"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'InventoryAPIRepostOnHandMode') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD InventoryAPIRepostOnHandMode varchar(10)"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET InventoryAPIRepostOnHandMode = 'TEST'"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'InventoryAPIRepostOnHandURL') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD InventoryAPIRepostOnHandURL Varchar(255) NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'InventoryAPIDailyActivityReportOnOff') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD InventoryAPIDailyActivityReportOnOff INT NOT NULL DEFAULT 0"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'InventoryAPIDailyActivityReportUserNos') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD InventoryAPIDailyActivityReportUserNos varchar(1000)"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'InventoryAPIDailyActivityReportAdditionalEmails') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD InventoryAPIDailyActivityReportAdditionalEmails varchar(1000)"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'InventoryAPIDailyActivityReportEmailSubject') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD InventoryAPIDailyActivityReportEmailSubject varchar(1000)"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If

'	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'BalancePeriods') AS IsItThere"
'	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
'	If NOT IsNull(rsCheckSettingsGlobal("IsItThere")) Then
'		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global DROP COLUMN BalancePeriods"
'		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
'	End If

'	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'MasterTPLYDollars') AS IsItThere"
'	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
'	If NOT IsNull(rsCheckSettingsGlobal("IsItThere")) Then
'		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global DROP COLUMN MasterTPLYDollars"
'		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
'	End If

'	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'MasterTPLYPercent') AS IsItThere"
'	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
'	If NOT IsNull(rsCheckSettingsGlobal("IsItThere")) Then
'		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global DROP COLUMN MasterTPLYPercent"
'		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
'	End If

'	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'MasterP3PDollars') AS IsItThere"
'	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
'	If NOT IsNull(rsCheckSettingsGlobal("IsItThere")) Then
'		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global DROP COLUMN MasterP3PDollars"
'		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
'	End If

'	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'MasterP3PPercent') AS IsItThere"
'	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
'	If NOT IsNull(rsCheckSettingsGlobal("IsItThere")) Then
'		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global DROP COLUMN MasterP3PPercent"
'		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
'	End If

'	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'MasterP12PDollars') AS IsItThere"
'	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
'	If NOT IsNull(rsCheckSettingsGlobal("IsItThere")) Then
'		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global DROP COLUMN MasterP12PDollars"
'		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
'	End If

'	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'MasterP12PPercent') AS IsItThere"
'	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
'	If NOT IsNull(rsCheckSettingsGlobal("IsItThere")) Then
'		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global DROP COLUMN MasterP12PPercent"
'		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
'	End If

'	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'DefaultSelectedCategoriesForVPandVPC') AS IsItThere"
'	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
'	If NOT IsNull(rsCheckSettingsGlobal("IsItThere")) Then
'		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global DROP COLUMN DefaultSelectedCategoriesForVPandVPC"
'		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
'	End If

	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'LeakageDataInvoiceTypes') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If NOT IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global DROP COLUMN LeakageDataInvoiceTypes"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If

	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'DelBoardPriorityColor') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD DelBoardPriorityColor varchar(255) NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET DelBoardPriorityColor = '#FF0000'"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If

	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'ServiceCarryOverReportEmailTo') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If NOT IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global DROP COLUMN ServiceCarryOverReportEmailTo"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If

	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'ServiceCarryOverReportEmailSubject') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If NOT IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global DROP COLUMN ServiceCarryOverReportEmailSubject"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'ServiceCarryOverReportONOFF') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If NOT IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global DROP COLUMN ServiceCarryOverReportONOFF"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'FSDefaultNotificationMethod') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD FSDefaultNotificationMethod varchar(255)"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET FSDefaultNotificationMethod = 'Text'"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'FS_SignatureOptional') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD FS_SignatureOptional int NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET FS_SignatureOptional = 0"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'InventoryWebAppPostOnHandMode') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD InventoryWebAppPostOnHandMode varchar(10)"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET InventoryWebAppPostOnHandMode= 'TEST'"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'InventoryWebAppPostOnHandURL') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD InventoryWebAppPostOnHandURL Varchar(255) NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If

	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'BackendInventoryPostsMode') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD BackendInventoryPostsMode varchar(10)"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET InventoryWebAppPostOnHandMode= 'TEST'"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'BackendInventoryPostsURL') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD BackendInventoryPostsURL Varchar(255) NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If


' This one is a DROP
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'ZeroRentalReportOnOff') AS IsItThere"
	Set rsCheckSettingsGlobal  = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If NOT IsNull(rsCheckSettingsGlobal("IsItThere")) Then
	
		On error resume next
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global DROP CONSTRAINT DF_Settings_Global_ZeroRentalReportOnOff"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		On error goto 0
		
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global DROP COLUMN ZeroRentalReportOnOff"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If


' This one is a DROP
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'ZeroRentalEmailSubject') AS IsItThere"
	Set rsCheckSettingsGlobal  = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If NOT IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global DROP COLUMN ZeroRentalEmailSubject"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If
	
' This one is a DROP
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'ZeroRentalsEmailTo') AS IsItThere"
	Set rsCheckSettingsGlobal  = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If NOT IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global DROP COLUMN ZeroRentalsEmailTo"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If

	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'OrderAPISwapAddressLines') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD OrderAPISwapAddressLines int NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET OrderAPISwapAddressLines = 0"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If

	' This one is a DROP
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'CMAPISwapAddressLines') AS IsItThere"
	Set rsCheckSettingsGlobal  = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If NOT IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global DROP COLUMN CMAPISwapAddressLines"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If

	' This one is a DROP
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'RAAPISwapAddressLines') AS IsItThere"
	Set rsCheckSettingsGlobal  = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If NOT IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global DROP COLUMN RAAPISwapAddressLines"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If

	' This one is a DROP
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'InvoiceAPISwapAddressLines') AS IsItThere"
	Set rsCheckSettingsGlobal  = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If NOT IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global DROP COLUMN InvoiceAPISwapAddressLines"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If


	' This one is a DROP
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'FSBoardRejectedColor') AS IsItThere"
	Set rsCheckSettingsGlobal  = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If NOT IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global DROP COLUMN FSBoardRejectedColor"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If

	' This one is a DROP
	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'FS_TechCanReject') AS IsItThere"
	Set rsCheckSettingsGlobal  = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If NOT IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global DROP COLUMN FS_TechCanReject"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If



		
	Set rsCheckSettingsGlobal = Nothing
	cnnCheckSettingsGlobal.Close
	Set cnnCheckSettingsGlobal = Nothing

	'**********************************************************************************
	'This code should be at the bottom after all the other structure change code is run
	'It will see if there is a record in the table & if not it will insert one with
	'a dummy value
	'Then it runs an update where all the proper default values are set
	'It was only done this way to make the code more readable
	'**********************************************************************************	
	
	Set cnnCheckSettingsGlobal = Server.CreateObject("ADODB.Connection")
	cnnCheckSettingsGlobal.open (Session("ClientCnnString"))
	Set rsCheckSettingsGlobal = Server.CreateObject("ADODB.Recordset")

	SQL_CheckSettingsGlobal = "SELECT COUNT (*) AS GlobalCount FROM Settings_Global"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)

	If rsCheckSettingsGlobal("GlobalCount") <> 1 Then

		SQL_CheckSettingsGlobal = "INSERT INTO Settings_Global (MasterTPLYDollars) VALUES (0)"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET "
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  MasterTPLYDollars = 500"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,MasterTPLYPercent = 6"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,MasterP3PDollars = 500"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,MasterP3PPercent = 5"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,MasterP12PDollars =  500"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,MasterP12PPercent = 7"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,BalancePeriods = 1"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,DefaultSelectedCategoriesForVPandVPC = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,POST_Serno = '" & ClientKey & "'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,POST_Mode = 'TEST'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,POST_ServiceMemoURL1 = 'http://apidev.mdsinsight.com/apiIn/receive_servicememo_xml.asp'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,POST_ServiceMemoURL1_MplexFormat = 1"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,POST_AssetLocationURL1 = 'http://apidev.mdsinsight.com/apiIn/receive_servicememo_xml.asp'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,POST_ServiceMemoURL2 = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,POST_AssetLocationURL2 = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,POST_ServiceMemoURL1ONOFF = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,POST_AssetLocationURL1ONOFF = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,POST_ServiceMemoURL2ONOFF = 0"		
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,POST_AssetLocationURL2ONOFF = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,EmailForNon200Responses = 'rich@ocsaccess.com'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,InternalEmail_MailDomain = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,STOPALLEMAIL = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,NotesScreenShowPopup = 1"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,FilterChangeDays = 10"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,FilterChangeDaysFieldService = 10"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,NeverPutOnHold = 1"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,DelBoardScheduledColor = '#ffffff'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,DelBoardCompletedColor = '#dbefd2'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,DelBoardSkippedColor = '#f6bfbf'" 
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,DelBoardProfitDollars = 500"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,DelBoardAtOrAboveProfitColor = '#6aa84f'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,DelBoardBelowProfitColor = '#ff4c4c'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,DelBoardUserAlertColor = '#ffff00'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,DelBoardNextStopColor = '#fce5cd'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,DelBoardAMColor = '#ff0000'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,DelBoardPriorityColor = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,DelBoardPieTimerColor = '#E73E97'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,DelBoardTitleGradientColor = '#cfe2f3'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,DelBoardTitleText = 'dev Deliveries for ~dow~, ~today~'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,DelBoardTitleTextFontColor = '#000000'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,DelBoardDontUseStopSequencing = 1"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,DelBoardRoutesToIgnore = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,DelBoardUPSRoutes = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,EZTextingID = 'ocsaccess'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,EZTextingPassword = 'tomato'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,NightBatchRunReportTime = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,NightBatchRunReportEmail = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,NightBatchRunReportOn = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,ServiceNormalAlertColor = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,ServicePriorityColor = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,ServicePriorityAlertColor = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,ServiceColorsOn = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,CRMTabLogColor = '#cc4125'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,CRMTabProductsColor = '#ff9900'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,CRMTabEquipmentColor = '#2ecc71'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,CRMTabDocumentsColor = '#3d85c6'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,CRMTabLocationColor = '#ff0000'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,CRMTabContactsColor = '#9900ff'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,CRMTabCompetitorsColor = '#666666'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,CRMTabOpportunityColor = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,CRMTabAuditTrailColor = '#f1c232'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,CRMMaxActivityDaysWarning = 5"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,CRMMaxActivityDaysPermitted = 12"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,CRMAutoCoordinateColors = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,CRMTileOfferingColor = '#4a86e8'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,CRMTileCompetitorColor = '#ad52ea'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,CRMTileDollarsColor = '#5cb85c'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,CRMTileActivityColor = '#ffd966'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,CRMTileOwnerColor = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,CRMTileCommentsColor = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,CRMTileStageColor = '#ff9900'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,CRMHideLocationTab = 1"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,AutoPromptNextStop = 1"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,CRMHideProductsTab = 1"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,CRMHideEquipmentTab = 1"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,EWSPostURL = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,EWSDefaultApptDuration = 45"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,EWSDefaultMeetingDuration = 120"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,OrderAPIRepostURL = 'http://98.6.75.158:3291/ocsmds/ocsapi'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,OrderAPIRepostONOFF = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,DoNotShowDeliveryLineItems = 1"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,AutoForceSelectNextStop = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,ProspSnapshotOnOff = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,ProspSnapshotInsideSales = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,ProspSnapshotOutsideSales = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,ProspSnapshotUserNos = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,ProspSnapshotSalesRepDisplayUserNos = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,ProspSnapshotAdditionalEmails = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,ProspSnapshotEmailSubject = 'Hooray! The report is here!'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,NextStopNagMessageONOFF = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,NextStopNagMinutes = 45"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,NextStopNagIntervalMinutes = 35"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,NextStopNagMessageMaxToSendPerStop = 3"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,NextStopNagMessageMaxToSendPerDriverPerDay = 5"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,NextStopNagMessageSendMethod = 'Text'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,NoActivityNagMessageONOFF = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,NoActivityNagMinutes = 95"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,NoActivityNagIntervalMinutes = 45"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,NoActivityNagMessageMaxToSendPerStop = 1"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,NoActivityNagMessageMaxToSendPerDriverPerDay = 20"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,NoActivityNagMessageSendMethod = 'Text'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,NoActivityNagTimeOfDay = '12:00'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,APIDailyActivityReportOnOff = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,APIDailyActivityReportUserNos = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,APIDailyActivityReportAdditionalEmails = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,APIDailyActivityReportEmailSubject = 'Daily API Activity Summary By Partner Report'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,InvoiceAPIRepostONOFF = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,InvoiceAPIRepostURL = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,CMAPIRepostONOFF = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,CMAPIRepostURL = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,RAAPIRepostONOFF = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,RAAPIRepostURL = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,SumInvAPIRepostONOFF = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,SumInvAPIRepostURL = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,OrderAPIRepostMode = 'Test'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,InvoiceAPIRepostMode = 'Test'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,RAAPIRepostMode = 'Test'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,CMAPIRepostMode = 'Test'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,SumInvAPIRepostMode = 'Test'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,SendInvoiceType = 'UNPOSTED'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,OrderAPIOffsetDays = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,InvoiceAPIOffsetDays = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,RAAPIOffsetDays = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,CMAPIOffsetDays = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,SumInvAPIOffsetDays = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,DelBoardInProgressColor = '#00ffff'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,OrderCutoffTime = '1800'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,InvoiceCutoffTime = '1800'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,CMCutoffTime = '1800'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,RACutoffTime = '1800'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,FieldServiceNotesReportOnOff = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,FieldServiceNotesReportUserNos = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,FieldServiceNotesReportAdditionalEmails = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,FieldServiceNotesReportEmailSubject = 'Service Ticket Tech Notes Report'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,FSWApp_SignatOpt = 1"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,FSWApp_HideAsset = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,FSWApp_NoUnableDdwn = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,NoActivityNagMessageONOFF_FS = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,NoActivityNagMinutes_FS = 30"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,NoActivityNagIntervalMinutes_FS = 30"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,NoActivityNagMessageMaxToSendPerStop_FS = 10"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,NoActivityNagMessageMaxToSendPerDriverPerDay_FS = 10"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,NoActivityNagMessageSendMethod_FS = 'Text'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,NoActivityNagTimeOfDay_FS = '12:00'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,NextStopNagMessageONOFF_FS = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,NextStopNagMinutes_FS = 30"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,NextStopNagIntervalMinutes_FS = 30"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,NextStopNagMessageMaxToSendPerStop_FS = 10"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,NextStopNagMessageMaxToSendPerDriverPerDay_FS = 10"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,NextStopNagMessageSendMethod_FS = 'Text'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,MasterNagMessageONOFF = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,FS_SignatureOptional = 1"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,FS_TechCanDecline = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,FSDefaultNotificationMethod = 'Text'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,InventoryAPIRepostONOFF = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,InventoryAPIRepostMode = 'TEST'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,InventoryAPIRepostURL = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,InventoryAPIDailyActivityReportOnOff = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,InventoryAPIDailyActivityReportUserNos = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,InventoryAPIDailyActivityReportAdditionalEmails = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,InventoryAPIDailyActivityReportEmailSubject = 'Inventory API DAily Activity'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,InventoryAPIRepostOnHandONOFF = 0"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,InventoryAPIRepostOnHandMode = 'TEST'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,InventoryAPIRepostOnHandURL = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,InventoryPostOnHandMode = 'TEST'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,InventoryPostOnHandURL = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,InventoryWebAppPostOnHandMode = 'TEST'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,InventoryWebAppPostOnHandURL = ''"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,BackendInventoryPostsMode = 'TEST'"
		SQL_CheckSettingsGlobal = SQL_CheckSettingsGlobal & "  ,BackendInventoryPostsURL = ''"

		'Response.Write("<br>" & SQL_CheckSettingsGlobal & "<br>")

		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	
	End If
	

	'This one is a DROP
	SQLCheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'POST_CustomerURL1') AS IsItThere"
	Set rsCheckSettingsGlobal  = cnnCheckSettingsGlobal.Execute(SQLCheckSettingsGlobal)
	If NOT IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQLCheckSettingsGlobal = "ALTER TABLE Settings_Global DROP COLUMN POST_CustomerURL1"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQLCheckSettingsGlobal)
	End If

	'This one is a DROP WITH A DROP CONSTRAINT
	SQLCheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'PMCallDays') AS IsItThere"
	Set rsCheckSettingsGlobal  = cnnCheckSettingsGlobal.Execute(SQLCheckSettingsGlobal)
	If NOT IsNull(rsCheckSettingsGlobal("IsItThere")) Then
	
		SQLCheckSettingsGlobal = "ALTER TABLE Settings_Global DROP CONSTRAINT [DF_tblSettings_Global_PMCallDays]"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQLCheckSettingsGlobal)
	
		SQLCheckSettingsGlobal = "ALTER TABLE Settings_Global DROP COLUMN PMCallDays"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQLCheckSettingsGlobal)
	End If

	'This one is a DROP WITH A DROP CONSTRAINT
	SQLCheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'PMCallDaysFieldService') AS IsItThere"
	Set rsCheckSettingsGlobal  = cnnCheckSettingsGlobal.Execute(SQLCheckSettingsGlobal)
	If NOT IsNull(rsCheckSettingsGlobal("IsItThere")) Then
	
		SQLCheckSettingsGlobal = "ALTER TABLE Settings_Global DROP CONSTRAINT [DF_tblSettings_Global_PMCallDaysFieldService]"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQLCheckSettingsGlobal)
	
		SQLCheckSettingsGlobal = "ALTER TABLE Settings_Global DROP COLUMN PMCallDaysFieldService"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQLCheckSettingsGlobal)
	End If


	'This one is a DROP
	SQLCheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'POST_CustomerURL2') AS IsItThere"
	Set rsCheckSettingsGlobal  = cnnCheckSettingsGlobal.Execute(SQLCheckSettingsGlobal)
	If NOT IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQLCheckSettingsGlobal = "ALTER TABLE Settings_Global DROP COLUMN POST_CustomerURL2"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQLCheckSettingsGlobal)
	End If
	
	
	'This one is a DROP WITH A DROP CONSTRAINT
	SQLCheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'POST_CustomerURL1ONOFF') AS IsItThere"
	Set rsCheckSettingsGlobal  = cnnCheckSettingsGlobal.Execute(SQLCheckSettingsGlobal)
	If NOT IsNull(rsCheckSettingsGlobal("IsItThere")) Then
	
		SQLCheckSettingsGlobal = "ALTER TABLE Settings_Global DROP CONSTRAINT [DF_tblSettings_Global_POST_CustomerURL1ONOFF]"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQLCheckSettingsGlobal)
	
		SQLCheckSettingsGlobal = "ALTER TABLE Settings_Global DROP COLUMN POST_CustomerURL1ONOFF"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQLCheckSettingsGlobal)
	End If

	'This one is a DROP WITH A DROP CONSTRAINT
	SQLCheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'POST_CustomerULR2ONOFF') AS IsItThere"
	Set rsCheckSettingsGlobal  = cnnCheckSettingsGlobal.Execute(SQLCheckSettingsGlobal)
	If NOT IsNull(rsCheckSettingsGlobal("IsItThere")) Then
	
		SQLCheckSettingsGlobal = "ALTER TABLE Settings_Global DROP CONSTRAINT [DF_tblSettings_Global_POST_CustomerULR2ONOFF]"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQLCheckSettingsGlobal)
	
		SQLCheckSettingsGlobal = "ALTER TABLE Settings_Global DROP COLUMN POST_CustomerULR2ONOFF"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQLCheckSettingsGlobal)
	End If


	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'PopulateLatAndLongONOFF') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD PopulateLatAndLongONOFF int NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET PopulateLatAndLongONOFF = 1"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If

	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'PopulateLatAndLongONOFF') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD PopulateLatAndLongONOFF int NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET PopulateLatAndLongONOFF = 1"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If

	SQL_CheckSettingsGlobal = "SELECT COL_LENGTH('Settings_Global', 'Schedule_PopulateLatAndLong') AS IsItThere"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If IsNull(rsCheckSettingsGlobal("IsItThere")) Then
		SQL_CheckSettingsGlobal = "ALTER TABLE Settings_Global ADD Schedule_PopulateLatAndLong varchar(1000) NULL"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET Schedule_PopulateLatAndLong = '1,1,1,1,1,1,1,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,0,0'"
		Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	End If

	' This code makes sure scheduled process information is not NULL
	SQL_CheckSettingsGlobal = "SELECT * FROM Settings_Global"
	Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
	If NOT rsCheckSettingsGlobal.EOF Then
		If IsNull(rsCheckSettingsGlobal ("Schedule_PopulateLatAndLong")) Then
			SQL_CheckSettingsGlobal = "UPDATE Settings_Global SET Schedule_PopulateLatAndLong = '1,1,1,1,1,1,1,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,0,0'"
			Set rsCheckSettingsGlobal = cnnCheckSettingsGlobal.Execute(SQL_CheckSettingsGlobal)
		End If
	End If


	Set rsCheckSettingsGlobal = Nothing
	cnnCheckSettingsGlobal.Close
	Set cnnCheckSettingsGlobal = Nothing
			
%>
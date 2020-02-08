<%	
	Set cnnSettings_NeedToKnow = Server.CreateObject("ADODB.Connection")
	cnnSettings_NeedToKnow.open (Session("ClientCnnString"))
	Set rsSettings_NeedToKnow = Server.CreateObject("ADODB.Recordset")
	rsSettings_NeedToKnow.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute("SELECT TOP 1 * FROM Settings_NeedToKnow")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLSettings_NeedToKnow = "CREATE TABLE [Settings_NeedToKnow]( "
			SQLSettings_NeedToKnow = SQLSettings_NeedToKnow & " [N2KAPIEmailToUserNos] [varchar](1000) NULL,  "
			SQLSettings_NeedToKnow = SQLSettings_NeedToKnow & " [N2KAPIUserNosToCC] [varchar](1000) NULL,  "
			SQLSettings_NeedToKnow = SQLSettings_NeedToKnow & " [N2KAPIEmailAddressesToCC] [varchar](8000) NULL,  "
			SQLSettings_NeedToKnow = SQLSettings_NeedToKnow & " [N2KAREmailToUserNos] [varchar](1000) NULL,  "
			SQLSettings_NeedToKnow = SQLSettings_NeedToKnow & " [N2KARUserNosToCC] [varchar](1000) NULL,  "
			SQLSettings_NeedToKnow = SQLSettings_NeedToKnow & " [N2KAREmailAddressesToCC] [varchar](8000) NULL,  "
			SQLSettings_NeedToKnow = SQLSettings_NeedToKnow & " [N2KEquipmentEmailToUserNos] [varchar](1000) NULL,  "
			SQLSettings_NeedToKnow = SQLSettings_NeedToKnow & " [N2KEquipmentUserNosToCC] [varchar](1000) NULL,  "
			SQLSettings_NeedToKnow = SQLSettings_NeedToKnow & " [N2KEquipmentEmailAddressesToCC] [varchar](8000) NULL,  "
			SQLSettings_NeedToKnow = SQLSettings_NeedToKnow & " [N2KGlobalSettingsEmailToUserNos] [varchar](1000) NULL,  "
			SQLSettings_NeedToKnow = SQLSettings_NeedToKnow & " [N2KGlobalSettingsUserNosToCC] [varchar](1000) NULL,  "
			SQLSettings_NeedToKnow = SQLSettings_NeedToKnow & " [N2KGlobalSettingsEmailAddressesToCC] [varchar](8000) NULL,  "

			SQLSettings_NeedToKnow = SQLSettings_NeedToKnow & ") ON [PRIMARY]"
			Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
			
			
		End If
	End If


	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KInventoryEmailToUserNos') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KInventoryEmailToUserNos varchar(1000) NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KInventoryUserNosToCC') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KInventoryUserNosToCC varchar(1000) NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KInventoryEmailAddressesToCC') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KInventoryEmailAddressesToCC varchar(8000) NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If
on error goto 0

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KInventoryReportONOFF') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KInventoryReportONOFF [int] NULL "
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		
		SQLSettings_NeedToKnow = "ALTER TABLE [Settings_NeedToKnow] ADD CONSTRAINT [DF_Settings_NeedToKnow_N2KInventoryReportONOFF]  DEFAULT ((0)) FOR [N2KInventoryReportONOFF]"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)

	End If
	
	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KARReportONOFF') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KARReportONOFF [int] NULL "
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		
		SQLSettings_NeedToKnow = "ALTER TABLE [Settings_NeedToKnow] ADD CONSTRAINT [DF_Settings_NeedToKnow_N2KARReportONOFF]  DEFAULT ((0)) FOR [N2KARReportONOFF]"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)

	End If
	
	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KAPIReportONOFF') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KAPIReportONOFF [int] NULL "
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		
		SQLSettings_NeedToKnow = "ALTER TABLE [Settings_NeedToKnow] ADD CONSTRAINT [DF_Settings_NeedToKnow_N2KAPIReportONOFF]  DEFAULT ((0)) FOR [N2KAPIReportONOFF]"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)

	End If
	
	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KEquipmentReportONOFF') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KEquipmentReportONOFF [int] NULL "
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		
		SQLSettings_NeedToKnow = "ALTER TABLE [Settings_NeedToKnow] ADD CONSTRAINT [DF_Settings_NeedToKnow_N2KEquipmentReportONOFF]  DEFAULT ((0)) FOR [N2KEquipmentReportONOFF]"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)

	End If
	
	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KGlobalSettingsReportONOFF') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KGlobalSettingsReportONOFF [int] NULL "
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		
		SQLSettings_NeedToKnow = "ALTER TABLE [Settings_NeedToKnow] ADD CONSTRAINT [DF_Settings_NeedToKnow_N2KGlobalSettingsReportONOFF]  DEFAULT ((0)) FOR [N2KGlobalSettingsReportONOFF]"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)

	End If
	
	
	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KInventoryAllowedDuplicateBins') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KInventoryAllowedDuplicateBins varchar(8000) NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If
	
	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'Schedule_APINeedToKnowReportGeneration') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD Schedule_APINeedToKnowReportGeneration varchar(1000) NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET Schedule_APINeedToKnowReportGeneration = '0,0,0,0,0,0,0,6:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,0,0'"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If
	
	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'Schedule_FinanceNeedToKnowReportGeneration') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD Schedule_FinanceNeedToKnowReportGeneration varchar(1000) NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET Schedule_FinanceNeedToKnowReportGeneration = '0,0,0,0,0,0,0,6:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,0,0'"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If
	
	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'Schedule_EquipmentNeedToKnowReportGeneration') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD Schedule_EquipmentNeedToKnowReportGeneration varchar(1000) NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET Schedule_EquipmentNeedToKnowReportGeneration = '0,0,0,0,0,0,0,6:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,0,0'"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'Schedule_GlobalSettingsNeedToKnowReportGeneration') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD Schedule_GlobalSettingsNeedToKnowReportGeneration varchar(1000) NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET Schedule_GlobalSettingsNeedToKnowReportGeneration = '0,0,0,0,0,0,0,6:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,0,0'"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'Schedule_InventoryNeedToKnowReportGeneration') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD Schedule_InventoryNeedToKnowReportGeneration varchar(1000) NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET Schedule_InventoryNeedToKnowReportGeneration = '0,0,0,0,0,0,0,6:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,0,0'"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KInventoryIncludeBlankCaseBin') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KInventoryIncludeBlankCaseBin int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KInventoryIncludeBlankCaseBin = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If
	
	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KInventoryIncludeBlankCaseUPCCode') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KInventoryIncludeBlankCaseUPCCode int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KInventoryIncludeBlankCaseUPCCode = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KInventoryIncludeBlankUnitandCaseUPCCode') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KInventoryIncludeBlankUnitandCaseUPCCode int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KInventoryIncludeBlankUnitandCaseUPCCode = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KInventoryIncludeBlankUnitBin') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KInventoryIncludeBlankUnitBin int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KInventoryIncludeBlankUnitBin = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KInventoryIncludeBlankUnitUPCCode') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KInventoryIncludeBlankUnitUPCCode int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KInventoryIncludeBlankUnitUPCCode = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KInventoryIncludeDuplicateUnitorCaseBin') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KInventoryIncludeDuplicateUnitorCaseBin int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KInventoryIncludeDuplicateUnitorCaseBin = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KInventoryIncludeDuplicateUPCCode') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KInventoryIncludeDuplicateUPCCode int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KInventoryIncludeDuplicateUPCCode = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KARIncludeEmptyAddress2') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KARIncludeEmptyAddress2 int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KARIncludeEmptyAddress2 = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KARIncludeEmptyCity') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KARIncludeEmptyCity int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KARIncludeEmptyCity = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KARIncludeNotAssignedToRegion') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KARIncludeNotAssignedToRegion int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KARIncludeNotAssignedToRegion = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If


	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KARIncludeEmptyCityStateZip') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KARIncludeEmptyCityStateZip int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KARIncludeEmptyCityStateZip = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KARIncludeEmptyCustomerName') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KARIncludeEmptyCustomerName int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KARIncludeEmptyCustomerName = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KARIncludeEmptyPhoneNumber') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KARIncludeEmptyPhoneNumber int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KARIncludeEmptyPhoneNumber = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KARIncludeEmptyState') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KARIncludeEmptyState int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KARIncludeEmptyState = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KARIncludeEmptyZip') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KARIncludeEmptyZip int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KARIncludeEmptyZip = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KARIncludeInvalidCityStateZip') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KARIncludeInvalidCityStateZip int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KARIncludeInvalidCityStateZip = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KARIncludeInvalidPhoneNumber') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KARIncludeInvalidPhoneNumber int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KARIncludeInvalidPhoneNumber = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KARIncludeInvalidState') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KARIncludeInvalidState int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KARIncludeInvalidState = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KARIncludeInvalidZipCode') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KARIncludeInvalidZipCode int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KARIncludeInvalidZipCode = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KARIncludeMissingcustomertype') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KARIncludeMissingcustomertype int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KARIncludeMissingcustomertype = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KARIncludeMissingprimarysalesman') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KARIncludeMissingprimarysalesman int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KARIncludeMissingprimarysalesman = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KARIncludeMissingsecondarysalesman') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KARIncludeMissingsecondarysalesman int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KARIncludeMissingsecondarysalesman = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KGlobalIncludeMissingClientLogoFile') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KGlobalIncludeMissingClientLogoFile int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KGlobalIncludeMissingClientLogoFile = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KGlobalIncludeMissingHolidayinCompanyCalendar') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KGlobalIncludeMissingHolidayinCompanyCalendar int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KGlobalIncludeMissingHolidayinCompanyCalendar = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KEqpIncludeBlankInsightAssetTagBrandPrefix') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KEqpIncludeBlankInsightAssetTagBrandPrefix int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KEqpIncludeBlankInsightAssetTagBrandPrefix = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KEqpIncludeBlankInsightAssetTagClassPrefix') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KEqpIncludeBlankInsightAssetTagClassPrefix int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KEqpIncludeBlankInsightAssetTagClassPrefix = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KEqpIncludeBlankInsightAssetTagManufacturerPrefix') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KEqpIncludeBlankInsightAssetTagManufacturerPrefix int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KEqpIncludeBlankInsightAssetTagManufacturerPrefix = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KEqpIncludeBlankInsightAssetTagModelPrefix') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KEqpIncludeBlankInsightAssetTagModelPrefix int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KEqpIncludeBlankInsightAssetTagModelPrefix = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KEqpIncludeUndefinedBrandExistsforEqp') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KEqpIncludeUndefinedBrandExistsforEqp int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KEqpIncludeUndefinedBrandExistsforEqp = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KEqpIncludeUndefinedClassExistsforEqp') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KEqpIncludeUndefinedClassExistsforEqp int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KEqpIncludeUndefinedClassExistsforEqp = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KEqpIncludeUndefinedConditionCodeExistsforEqp') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KEqpIncludeUndefinedConditionCodeExistsforEqp int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KEqpIncludeUndefinedConditionCodeExistsforEqp = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KEqpIncludeUndefinedGroupExistsforEqp') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KEqpIncludeUndefinedGroupExistsforEqp int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KEqpIncludeUndefinedGroupExistsforEqp = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KEqpIncludeUndefinedManufacturerExistsforEqp') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KEqpIncludeUndefinedManufacturerExistsforEqp int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KEqpIncludeUndefinedManufacturerExistsforEqp = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KEqpIncludeUndefinedModelExistsforEqp') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KEqpIncludeUndefinedModelExistsforEqp int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KEqpIncludeUndefinedModelExistsforEqp = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KEqpIncludeUndefinedStatusCodeExistsforEqp') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KEqpIncludeUndefinedStatusCodeExistsforEqp int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KEqpIncludeUndefinedStatusCodeExistsforEqp = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	SQLSettings_NeedToKnow = "SELECT COL_LENGTH('Settings_NeedToKnow', 'N2KEqpIncludeZeroDollarRentalsExistforEqp') AS IsItThere"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If IsNull(rsSettings_NeedToKnow("IsItThere")) Then
		SQLSettings_NeedToKnow = "ALTER TABLE Settings_NeedToKnow ADD N2KEqpIncludeZeroDollarRentalsExistforEqp int NULL"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET N2KEqpIncludeZeroDollarRentalsExistforEqp = 1"
		Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	End If

	' This code makes sure scheduled process information is not NULL
	SQLSettings_NeedToKnow = "SELECT * FROM Settings_NeedToKnow"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If NOT rsSettings_NeedToKnow.EOF Then
		If IsNull(rsSettings_NeedToKnow("Schedule_APINeedToKnowReportGeneration")) Then
			SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET Schedule_APINeedToKnowReportGeneration = '0,0,0,0,0,0,0,6:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,0,0'"
			Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		End If
	End If

	SQLSettings_NeedToKnow = "SELECT * FROM Settings_NeedToKnow"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If NOT rsSettings_NeedToKnow.EOF Then
		If IsNull(rsSettings_NeedToKnow("Schedule_FinanceNeedToKnowReportGeneration")) Then
			SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET Schedule_FinanceNeedToKnowReportGeneration = '0,0,0,0,0,0,0,6:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,0,0'"
			Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		End If
	End If

	SQLSettings_NeedToKnow = "SELECT * FROM Settings_NeedToKnow"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If NOT rsSettings_NeedToKnow.EOF Then
		If IsNull(rsSettings_NeedToKnow("Schedule_EquipmentNeedToKnowReportGeneration")) Then
			SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET Schedule_EquipmentNeedToKnowReportGeneration = '0,0,0,0,0,0,0,6:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,0,0'"
			Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		End If
	End If

	SQLSettings_NeedToKnow = "SELECT * FROM Settings_NeedToKnow"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If NOT rsSettings_NeedToKnow.EOF Then
		If IsNull(rsSettings_NeedToKnow("Schedule_GlobalSettingsNeedToKnowReportGeneration")) Then
			SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET Schedule_GlobalSettingsNeedToKnowReportGeneration = '0,0,0,0,0,0,0,6:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,0,0'"
			Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		End If
	End If

		SQLSettings_NeedToKnow = "SELECT * FROM Settings_NeedToKnow"
	Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
	If NOT rsSettings_NeedToKnow.EOF Then
		If IsNull(rsSettings_NeedToKnow("Schedule_InventoryNeedToKnowReportGeneration")) Then
			SQLSettings_NeedToKnow = "UPDATE Settings_NeedToKnow SET Schedule_InventoryNeedToKnowReportGeneration = '0,0,0,0,0,0,0,6:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,0,0'"
			Set rsSettings_NeedToKnow = cnnSettings_NeedToKnow.Execute(SQLSettings_NeedToKnow)
		End If
	End If

	set rsSettings_NeedToKnow = nothing
	cnnSettings_NeedToKnow.close
	set cnnSettings_NeedToKnow = nothing
				
%>
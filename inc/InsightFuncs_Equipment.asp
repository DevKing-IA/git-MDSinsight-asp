<%
'********************************
'List of all the functions & subs
'********************************
'Sub Record_EQ_Activity(passedEquipmentIntRecID,passedActivity,passedUserNo)

'Func GetTotalNumberOfEquipmentRecords()
'Func GetTotalNumberOfCustomerEquipmentRecords()
'Func GetTotalNumberOfManufacturers()
'Func GetTotalNumberOfBrands()
'Func GetTotalNumberOfModels()
'Func GetTotalNumberOfGroups()
'Func GetTotalNumberOfClasses()
'Func GetTotalNumberOfConditionCodes()
'Func GetTotalNumberOfStatusCodes()
'Func GetTotalNumberOfMovementCodes()
'Func GetTotalNumberOfAcquisitionCodes()
'Func GetTotalNumberOfModelsForCustomer(passedCustID,passedModelIntRecID)
'Func GetTotalValueOfEquipmentForCustomer(passedCustID)
'Func GetTotalValueOfModelsForCustomer(passedCustID,passedModelIntRecID)
'Func GetTotalValueOfRentalModelsForCustomer(passedCustID,passedModelIntRecID)
'Func CustHasEquipment(passedCustID)
'Func GetCustomerIDByEquipIntRecID(passedEquipIntRecID)
'Func GetAvailableForPlacementByEquipIntRecID(passedEquipIntRecID)


'Func NumberEquipmentRecsDefinedForManufacturer(passedManufIntRecID)
'Func NumberEquipmentRecsDefinedForBrand(passedBrandIntRecID)
'Func NumberEquipmentRecsDefinedForManufacturerAndBrand(passedManufIntRecID,passedBrandIntRecID)
'Func NumberEquipmentRecsDefinedForModel(passedModelIntRecID)
'Func NumberEquipmentRecsDefinedForCondition(passedConditionIntRecID)
'Func NumberEquipmentRecsDefinedForClass(passedClassIntRecID)
'Func NumberEquipmentRecsDefinedForGroup(passedGroupIntRecID)
'Func NumberEquipmentRecsDefinedForStatusCode(passedStatusCodeIntRecID)
'Func NumberEquipmentRecsDefinedForMovementCode(passedMovementCodeIntRecID)
'Func NumberEquipmentRecsDefinedForAcquisitionCode(passedAcquisitionCodeIntRecID)

'Func NumberCustomerEquipmentRecsDefinedForManufacturer(passedManufIntRecID)
'Func NumberCustomerEquipmentRecsDefinedForBrand(passedBrandIntRecID)
'Func NumberCustomerEquipmentRecsDefinedForModel(passedModelIntRecID)
'Func NumberCustomerEquipmentRecsDefinedForCondition(passedConditionIntRecID)
'Func NumberCustomerEquipmentRecsDefinedForClass(passedClassIntRecID)
'Func NumberCustomerEquipmentRecsDefinedForGroup(passedGroupIntRecID)
'Func NumberCustomerEquipmentRecsDefinedForStatusCode(passedStatusCodeIntRecID)
'Func NumberCustomerEquipmentRecsDefinedForMovementCode(passedMovementCodeIntRecID)
'Func NumberCustomerEquipmentRecsDefinedForAcquistionCode(passedAcquisitionCodeIntRecID)
'Func NumberModelsWithStatusCode(passedStatusCodeIntRecID,passedModelIntRecID)

'Func GetManufacturerNameByIntRecID(passedManufIntRecID)
'Func GetManufacturerNameByBrandIntRecID(passedBrandIntRecID)
'Func GetManufacturerNameByModelIntRecID(passedModelIntRecID)
'Func GetManufacturerIntRecIDByBrandIntRecID(passedBrandIntRecID)
'Func GetManufacturerIntRecIDByModelIntRecID(passedModelIntRecID)

'Func GetBrandNameByIntRecID(passedBrandIntRecID)
'Func GetBrandNameByModelIntRecID(passedModelIntRecID)
'Func GetBrandIntRecIDByModelIntRecID(passedModelIntRecID)

'Func GetModelNameByIntRecID(passedModelIntRecID)

'Func GetMovementCodeDescByIntRecID(passedMovementCodeIntRecID)
'Func GetMovementCodeByIntRecID(passedMovementCodeIntRecID)

'Func GetAcquisitionCodeDescByIntRecID(passedAcquisitionCodeIntRecID)
'Func GetAcquisitionCodeByIntRecID(passedAcquisitionCodeIntRecID)

'Func GetGroupNameByIntRecID(passedGroupIntRecID)

'Func GetClassNameByIntRecID(passedClassIntRecID)
'Func GetClassNameByModelIntRecID(passedModelIntRecID)
'Func GetClassIDByModelIntRecID(passedModelIntRecID)

'Func GetConditionNameByIntRecID(passedConditionIntRecID)
'Func GetStatusCodeNameByIntRecID(passedStatusCodeIntRecID)

'Func NumberOfDocumentsByModelIntRecID(passedModelIntRecID)
'Func NumberOfImagesByModelIntRecID(passedModelIntRecID)
'Func NumberOfLinksByModelIntRecID(passedModelIntRecID)

'Func GetEquipVendorNameByVendorID(passedVendorIntRecID)

'Func GetInsightAssetTagByEquipIntRecID(passedEquipIntRecID)

'************************************
'End List of all the functions & subs
'************************************

'**************************************************************************************************************************************
'**************************************************************************************************************************************
Sub Record_EQ_Activity(passedEquipmentIntRecID,passedActivity,passedUserNo)

	'Creates an entry in EQ_Audit
	
	SQLRecord_EQ_Activity = "INSERT INTO EQ_Audit (EQRecID,Activity,PerformedByUserNo) "
	SQLRecord_EQ_Activity = SQLRecord_EQ_Activity &  " VALUES (" & passedEquipmentIntRecID
	SQLRecord_EQ_Activity = SQLRecord_EQ_Activity & ",'"  & passedActivity & "'," & passedUserNo & ")"
	
	Set cnnRecord_EQ_Activity = Server.CreateObject("ADODB.Connection")
	cnnRecord_EQ_Activity.open (Session("ClientCnnString"))

	Set rsRecord_EQ_Activity = Server.CreateObject("ADODB.Recordset")
	rsRecord_EQ_Activity.CursorLocation = 3 
	Set rsRecord_EQ_Activity = cnnRecord_EQ_Activity.Execute(SQLRecord_EQ_Activity)
	set rsRecord_EQ_Activity = Nothing
	
End Sub
'**************************************************************************************************************************************
'**************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetTotalNumberOfEquipmentRecords()

	Set cnnTotalNumEquipRecords = Server.CreateObject("ADODB.Connection")
	cnnTotalNumEquipRecords.open Session("ClientCnnString")

	resultTotalNumEquipRecords = 0
		
	SQLTotalNumEquipRecords = "SELECT COUNT(*) AS EQUIPCOUNT FROM EQ_Equipment"
	 
	Set rsTotalNumEquipRecords = Server.CreateObject("ADODB.Recordset")
	rsTotalNumEquipRecords.CursorLocation = 3 
	
	rsTotalNumEquipRecords.Open SQLTotalNumEquipRecords,cnnTotalNumEquipRecords 
			
	resultTotalNumEquipRecords = rsTotalNumEquipRecords("EQUIPCOUNT")
	
	rsTotalNumEquipRecords.Close
	set rsTotalNumEquipRecords = Nothing
	cnnTotalNumEquipRecords.Close	
	set cnnTotalNumEquipRecords = Nothing
	
	GetTotalNumberOfEquipmentRecords = resultTotalNumEquipRecords
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetTotalNumberOfCustomerEquipmentRecords()

	Set cnnTotalNumEquipRecords = Server.CreateObject("ADODB.Connection")
	cnnTotalNumEquipRecords.open Session("ClientCnnString")

	resultTotalNumEquipRecords = 0
		
	SQLTotalNumEquipRecords = "SELECT COUNT(*) AS EQUIPCOUNT FROM EQ_CustomerEquipment"
	 
	Set rsTotalNumEquipRecords = Server.CreateObject("ADODB.Recordset")
	rsTotalNumEquipRecords.CursorLocation = 3 
	
	rsTotalNumEquipRecords.Open SQLTotalNumEquipRecords,cnnTotalNumEquipRecords 
			
	resultTotalNumEquipRecords = rsTotalNumEquipRecords("EQUIPCOUNT")
	
	rsTotalNumEquipRecords.Close
	set rsTotalNumEquipRecords = Nothing
	cnnTotalNumEquipRecords.Close	
	set cnnTotalNumEquipRecords = Nothing
	
	GetTotalNumberOfCustomerEquipmentRecords = resultTotalNumEquipRecords
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetTotalNumberOfManufacturers()

	Set cnnTotalNumManufacturers = Server.CreateObject("ADODB.Connection")
	cnnTotalNumManufacturers.open Session("ClientCnnString")

	resultTotalNumManufacturers = 0
		
	SQLTotalNumManufacturers = "SELECT COUNT(*) AS MANFCOUNT FROM EQ_Manufacturers"
	 
	Set rsTotalNumManufacturers = Server.CreateObject("ADODB.Recordset")
	rsTotalNumManufacturers.CursorLocation = 3 
	
	rsTotalNumManufacturers.Open SQLTotalNumManufacturers,cnnTotalNumManufacturers 
			
	resultTotalNumManufacturers = rsTotalNumManufacturers("MANFCOUNT")
	
	rsTotalNumManufacturers.Close
	set rsTotalNumManufacturers = Nothing
	cnnTotalNumManufacturers.Close	
	set cnnTotalNumManufacturers = Nothing
	
	GetTotalNumberOfManufacturers= resultTotalNumManufacturers
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetTotalNumberOfBrands()

	Set cnnTotalNumBrands = Server.CreateObject("ADODB.Connection")
	cnnTotalNumBrands.open Session("ClientCnnString")

	resultTotalNumBrands = 0
		
	SQLTotalNumBrands = "SELECT COUNT(*) AS BRANDCOUNT FROM EQ_Brands"
	 
	Set rsTotalNumBrands = Server.CreateObject("ADODB.Recordset")
	rsTotalNumBrands.CursorLocation = 3 
	
	rsTotalNumBrands.Open SQLTotalNumBrands,cnnTotalNumBrands 
			
	resultTotalNumBrands = rsTotalNumBrands("BRANDCOUNT")
	
	rsTotalNumBrands.Close
	set rsTotalNumBrands = Nothing
	cnnTotalNumBrands.Close	
	set cnnTotalNumBrands = Nothing
	
	GetTotalNumberOfBrands= resultTotalNumBrands
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************

'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetTotalNumberOfGroups()

	Set cnnTotalNumGroups = Server.CreateObject("ADODB.Connection")
	cnnTotalNumGroups.open Session("ClientCnnString")

	resultTotalNumGroups = 0
		
	SQLTotalNumGroups = "SELECT COUNT(*) AS GroupCOUNT FROM EQ_Groups"
	 
	Set rsTotalNumGroups = Server.CreateObject("ADODB.Recordset")
	rsTotalNumGroups.CursorLocation = 3 
	
	rsTotalNumGroups.Open SQLTotalNumGroups,cnnTotalNumGroups 
			
	resultTotalNumGroups = rsTotalNumGroups("GroupCOUNT")
	
	rsTotalNumGroups.Close
	set rsTotalNumGroups = Nothing
	cnnTotalNumGroups.Close	
	set cnnTotalNumGroups = Nothing
	
	GetTotalNumberOfGroups= resultTotalNumGroups
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetTotalNumberOfClasses()

	Set cnnTotalNumClasses = Server.CreateObject("ADODB.Connection")
	cnnTotalNumClasses.open Session("ClientCnnString")

	resultTotalNumClasses = 0
		
	SQLTotalNumClasses = "SELECT COUNT(*) AS ClassCOUNT FROM EQ_Classes"
	 
	Set rsTotalNumClasses = Server.CreateObject("ADODB.Recordset")
	rsTotalNumClasses.CursorLocation = 3 
	
	rsTotalNumClasses.Open SQLTotalNumClasses,cnnTotalNumClasses 
			
	resultTotalNumClasses = rsTotalNumClasses("ClassCOUNT")
	
	rsTotalNumClasses.Close
	set rsTotalNumClasses = Nothing
	cnnTotalNumClasses.Close	
	set cnnTotalNumClasses = Nothing
	
	GetTotalNumberOfClasses= resultTotalNumClasses
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetTotalNumberOfModels()

	Set cnnTotalNumModels = Server.CreateObject("ADODB.Connection")
	cnnTotalNumModels.open Session("ClientCnnString")

	resultTotalNumModels = 0
		
	SQLTotalNumModels = "SELECT COUNT(*) AS MODELCOUNT FROM EQ_Models"
	 
	Set rsTotalNumModels = Server.CreateObject("ADODB.Recordset")
	rsTotalNumModels.CursorLocation = 3 
	
	rsTotalNumModels.Open SQLTotalNumModels,cnnTotalNumModels 
			
	resultTotalNumModels = rsTotalNumModels("MODELCOUNT")
	
	rsTotalNumModels.Close
	set rsTotalNumModels = Nothing
	cnnTotalNumModels.Close	
	set cnnTotalNumModels = Nothing
	
	GetTotalNumberOfModels= resultTotalNumModels
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetTotalNumberOfConditionCodes()

	Set cnnTotalNumConditionCodes = Server.CreateObject("ADODB.Connection")
	cnnTotalNumConditionCodes.open Session("ClientCnnString")

	resultTotalNumConditionCodes = 0
		
	SQLTotalNumConditionCodes = "SELECT COUNT(*) AS ConditionCOUNT FROM EQ_ConditionCodes"
	 
	Set rsTotalNumConditionCodes = Server.CreateObject("ADODB.Recordset")
	rsTotalNumConditionCodes.CursorLocation = 3 
	
	rsTotalNumConditionCodes.Open SQLTotalNumConditionCodes,cnnTotalNumConditionCodes 
			
	resultTotalNumConditionCodes = rsTotalNumConditionCodes("ConditionCOUNT")
	
	rsTotalNumConditionCodes.Close
	set rsTotalNumConditionCodes = Nothing
	cnnTotalNumConditionCodes.Close	
	set cnnTotalNumConditionCodes = Nothing
	
	GetTotalNumberOfConditionCodes= resultTotalNumConditionCodes
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetTotalNumberOfStatusCodes()

	Set cnnTotalNumStatusCodes = Server.CreateObject("ADODB.Connection")
	cnnTotalNumStatusCodes.open Session("ClientCnnString")

	resultTotalNumStatusCodes = 0
		
	SQLTotalNumStatusCodes = "SELECT COUNT(*) AS StatusCOUNT FROM EQ_StatusCodes"
	 
	Set rsTotalNumStatusCodes = Server.CreateObject("ADODB.Recordset")
	rsTotalNumStatusCodes.CursorLocation = 3 
	
	rsTotalNumStatusCodes.Open SQLTotalNumStatusCodes,cnnTotalNumStatusCodes 
			
	resultTotalNumStatusCodes = rsTotalNumStatusCodes("StatusCOUNT")
	
	rsTotalNumStatusCodes.Close
	set rsTotalNumStatusCodes = Nothing
	cnnTotalNumStatusCodes.Close	
	set cnnTotalNumStatusCodes = Nothing
	
	GetTotalNumberOfStatusCodes= resultTotalNumStatusCodes
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetTotalNumberOfMovementCodes()

	Set cnnTotalNumMovementCodes = Server.CreateObject("ADODB.Connection")
	cnnTotalNumMovementCodes.open Session("ClientCnnString")

	resultTotalNumMovementCodes = 0
		
	SQLTotalNumMovementCodes = "SELECT COUNT(*) AS MovementCOUNT FROM EQ_MovementCodes"
	 
	Set rsTotalNumMovementCodes = Server.CreateObject("ADODB.Recordset")
	rsTotalNumMovementCodes.CursorLocation = 3 
	
	rsTotalNumMovementCodes.Open SQLTotalNumMovementCodes,cnnTotalNumMovementCodes 
			
	resultTotalNumMovementCodes = rsTotalNumMovementCodes("MovementCOUNT")
	
	rsTotalNumMovementCodes.Close
	set rsTotalNumMovementCodes = Nothing
	cnnTotalNumMovementCodes.Close	
	set cnnTotalNumMovementCodes = Nothing
	
	GetTotalNumberOfMovementCodes= resultTotalNumMovementCodes
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetTotalNumberOfAcquisitionCodes()

	Set cnnTotalNumAcquisitionCodes = Server.CreateObject("ADODB.Connection")
	cnnTotalNumAcquisitionCodes.open Session("ClientCnnString")

	resultTotalNumAcquisitionCodes = 0
		
	SQLTotalNumAcquisitionCodes = "SELECT COUNT(*) AS AcquisitionCOUNT FROM EQ_AcquisitionCodes"
	 
	Set rsTotalNumAcquisitionCodes = Server.CreateObject("ADODB.Recordset")
	rsTotalNumAcquisitionCodes.CursorLocation = 3 
	
	rsTotalNumAcquisitionCodes.Open SQLTotalNumAcquisitionCodes,cnnTotalNumAcquisitionCodes 
			
	resultTotalNumAcquisitionCodes = rsTotalNumAcquisitionCodes("AcquisitionCOUNT")
	
	rsTotalNumAcquisitionCodes.Close
	set rsTotalNumAcquisitionCodes = Nothing
	cnnTotalNumAcquisitionCodes.Close	
	set cnnTotalNumAcquisitionCodes = Nothing
	
	GetTotalNumberOfAcquisitionCodes= resultTotalNumAcquisitionCodes
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************

'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function NumberEquipmentRecsDefinedForManufacturer(passedManufIntRecID)

	Set cnnNumPcsEquipDefinedForManufacturer = Server.CreateObject("ADODB.Connection")
	cnnNumPcsEquipDefinedForManufacturer.open Session("ClientCnnString")

	resultNumPcsEquipDefinedForManufacturer = 0
		
	'SQLNumPcsEquipDefinedForManufacturer = "SELECT COUNT(*) AS EQUIPCOUNT FROM EQ_Equipment WHERE ManufacIntRecID = " & passedManufIntRecID
	
	
	SQLNumPcsEquipDefinedForManufacturer = "SELECT COUNT(*) AS EQUIPCOUNT FROM EQ_Equipment WHERE ModelIntRecID IN "
	SQLNumPcsEquipDefinedForManufacturer = 	SQLNumPcsEquipDefinedForManufacturer & "(SELECT InternalRecordIdentifier FROM EQ_Models WHERE ManufacIntRecID = " & passedManufIntRecID & ")"

	 
	Set rsNumPcsEquipDefinedForManufacturer = Server.CreateObject("ADODB.Recordset")
	rsNumPcsEquipDefinedForManufacturer.CursorLocation = 3 
	
	rsNumPcsEquipDefinedForManufacturer.Open SQLNumPcsEquipDefinedForManufacturer,cnnNumPcsEquipDefinedForManufacturer 
			
	resultNumPcsEquipDefinedForManufacturer = rsNumPcsEquipDefinedForManufacturer("EQUIPCOUNT")
	
	rsNumPcsEquipDefinedForManufacturer.Close
	set rsNumPcsEquipDefinedForManufacturer = Nothing
	cnnNumPcsEquipDefinedForManufacturer.Close	
	set cnnNumPcsEquipDefinedForManufacturer = Nothing
	
	NumberEquipmentRecsDefinedForManufacturer = resultNumPcsEquipDefinedForManufacturer
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function NumberEquipmentRecsDefinedForBrand(passedBrandIntRecID)

	Set cnnNumPcsEquipDefinedForBrand = Server.CreateObject("ADODB.Connection")
	cnnNumPcsEquipDefinedForBrand.open Session("ClientCnnString")

	resultNumPcsEquipDefinedForBrand = 0
		
	SQLNumPcsEquipDefinedForBrand = "SELECT COUNT(*) AS EQUIPCOUNT FROM EQ_Equipment WHERE ModelIntRecID IN "
	SQLNumPcsEquipDefinedForBrand =  SQLNumPcsEquipDefinedForBrand & "(SELECT InternalRecordIdentifier FROM EQ_Models WHERE BrandIntRecID = " & passedBrandIntRecID & ")"
	
	Set rsNumPcsEquipDefinedForBrand = Server.CreateObject("ADODB.Recordset")
	rsNumPcsEquipDefinedForBrand.CursorLocation = 3 
	
	rsNumPcsEquipDefinedForBrand.Open SQLNumPcsEquipDefinedForBrand,cnnNumPcsEquipDefinedForBrand 
			
	resultNumPcsEquipDefinedForBrand = rsNumPcsEquipDefinedForBrand("EQUIPCOUNT")
	
	rsNumPcsEquipDefinedForBrand.Close
	set rsNumPcsEquipDefinedForBrand = Nothing
	cnnNumPcsEquipDefinedForBrand.Close	
	set cnnNumPcsEquipDefinedForBrand = Nothing
	
	NumberEquipmentRecsDefinedForBrand = resultNumPcsEquipDefinedForBrand
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function NumberCustomerEquipmentRecsDefinedForBrand(passedBrandIntRecID)

	Set cnnNumCustPcsEquipDefinedForBrand = Server.CreateObject("ADODB.Connection")
	cnnNumCustPcsEquipDefinedForBrand.open Session("ClientCnnString")

	resultNumCustPcsEquipDefinedForBrand = 0
		
	SQLNumCustPcsEquipDefinedForBrand = "SELECT COUNT(*) AS EQUIPCOUNT FROM EQ_CustomerEquipment INNER JOIN EQ_Equipment ON "
	SQLNumCustPcsEquipDefinedForBrand = SQLNumCustPcsEquipDefinedForBrand & "EQ_Equipment.InternalRecordIdentifier = EQ_CustomerEquipment.EquipIntRecID "
	SQLNumCustPcsEquipDefinedForBrand = SQLNumCustPcsEquipDefinedForBrand & " WHERE EQ_Equipment.ModelIntRecID IN "
	SQLNumCustPcsEquipDefinedForBrand = SQLNumCustPcsEquipDefinedForBrand  & "(SELECT InternalRecordIdentifier FROM EQ_Models WHERE BrandIntRecID = " & passedBrandIntRecID & ")"
	 
	Set rsNumCustPcsEquipDefinedForBrand = Server.CreateObject("ADODB.Recordset")
	rsNumCustPcsEquipDefinedForBrand.CursorLocation = 3 
	
	rsNumCustPcsEquipDefinedForBrand.Open SQLNumCustPcsEquipDefinedForBrand,cnnNumCustPcsEquipDefinedForBrand 
			
	resultNumCustPcsEquipDefinedForBrand = rsNumCustPcsEquipDefinedForBrand("EQUIPCOUNT")
	
	rsNumCustPcsEquipDefinedForBrand.Close
	set rsNumCustPcsEquipDefinedForBrand = Nothing
	cnnNumCustPcsEquipDefinedForBrand.Close	
	set cnnNumCustPcsEquipDefinedForBrand = Nothing
	
	NumberCustomerEquipmentRecsDefinedForBrand = resultNumCustPcsEquipDefinedForBrand
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function NumberEquipmentRecsDefinedForManufacturerAndBrand(passedManufIntRecID, passedBrandIntRecID)

	Set cnnNumPcsEquipDefinedForManufacturer = Server.CreateObject("ADODB.Connection")
	cnnNumPcsEquipDefinedForManufacturer.open Session("ClientCnnString")

	resultNumPcsEquipDefinedForManufacturer = 0
		
	SQLNumPcsEquipDefinedForManufacturer = "SELECT COUNT(*) AS EQUIPCOUNT FROM EQ_Equipment WHERE ModelIntRecID IN "
	SQLNumPcsEquipDefinedForManufacturer =  SQLNumPcsEquipDefinedForManufacturer & "(SELECT InternalRecordIdentifier FROM EQ_Models "
	SQLNumPcsEquipDefinedForManufacturer =  SQLNumPcsEquipDefinedForManufacturer & "WHERE BrandIntRecID = " & passedBrandIntRecID & " AND ManufacIntRecID = " & passedManufIntRecID & ")"
	
	 
	Set rsNumPcsEquipDefinedForManufacturer = Server.CreateObject("ADODB.Recordset")
	rsNumPcsEquipDefinedForManufacturer.CursorLocation = 3 
	
	rsNumPcsEquipDefinedForManufacturer.Open SQLNumPcsEquipDefinedForManufacturer,cnnNumPcsEquipDefinedForManufacturer 
			
	resultNumPcsEquipDefinedForManufacturer = rsNumPcsEquipDefinedForManufacturer("EQUIPCOUNT")
	
	rsNumPcsEquipDefinedForManufacturer.Close
	set rsNumPcsEquipDefinedForManufacturer = Nothing
	cnnNumPcsEquipDefinedForManufacturer.Close	
	set cnnNumPcsEquipDefinedForManufacturer = Nothing
	
	NumberEquipmentRecsDefinedForManufacturerAndBrand = resultNumPcsEquipDefinedForManufacturer
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************




'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function NumberCustomerEquipmentRecsDefinedForManf(passedManufIntRecID)

	Set cnnNumCustPcsEquipDefinedForManf = Server.CreateObject("ADODB.Connection")
	cnnNumCustPcsEquipDefinedForManf.open Session("ClientCnnString")

	resultNumCustPcsEquipDefinedForManf = 0
		
	SQLNumCustPcsEquipDefinedForManf = "SELECT COUNT(*) AS EQUIPCOUNT FROM EQ_CustomerEquipment INNER JOIN EQ_Equipment ON "
	SQLNumCustPcsEquipDefinedForManf = SQLNumCustPcsEquipDefinedForManf & "EQ_Equipment.InternalRecordIdentifier = EQ_CustomerEquipment.EquipIntRecID "
	SQLNumCustPcsEquipDefinedForManf = SQLNumCustPcsEquipDefinedForManf & " WHERE EQ_Equipment.ModelIntRecID IN "
	SQLNumCustPcsEquipDefinedForManf = SQLNumCustPcsEquipDefinedForManf & "(SELECT InternalRecordIdentifier FROM EQ_Models WHERE ManufacIntRecID = " & passedManufIntRecID & ")"
	
	Set rsNumCustPcsEquipDefinedForManf = Server.CreateObject("ADODB.Recordset")
	rsNumCustPcsEquipDefinedForManf.CursorLocation = 3 
	
	rsNumCustPcsEquipDefinedForManf.Open SQLNumCustPcsEquipDefinedForManf,cnnNumCustPcsEquipDefinedForManf 
			
	resultNumCustPcsEquipDefinedForManf = rsNumCustPcsEquipDefinedForManf("EQUIPCOUNT")
	
	rsNumCustPcsEquipDefinedForManf.Close
	set rsNumCustPcsEquipDefinedForManf = Nothing
	cnnNumCustPcsEquipDefinedForManf.Close	
	set cnnNumCustPcsEquipDefinedForManf = Nothing
	
	NumberCustomerEquipmentRecsDefinedForManf = resultNumCustPcsEquipDefinedForManf
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************





'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function NumberEquipmentRecsDefinedForModel(passedModelIntRecID)

	Set cnnNumPcsEquipDefinedForModel = Server.CreateObject("ADODB.Connection")
	cnnNumPcsEquipDefinedForModel.open Session("ClientCnnString")

	resultNumPcsEquipDefinedForModel = 0
		
	SQLNumPcsEquipDefinedForModel = "SELECT COUNT(*) AS EQUIPCOUNT FROM EQ_Equipment WHERE ModelIntRecID = " & passedModelIntRecID
	 
	Set rsNumPcsEquipDefinedForModel = Server.CreateObject("ADODB.Recordset")
	rsNumPcsEquipDefinedForModel.CursorLocation = 3 
	
	rsNumPcsEquipDefinedForModel.Open SQLNumPcsEquipDefinedForModel,cnnNumPcsEquipDefinedForModel 
			
	resultNumPcsEquipDefinedForModel = rsNumPcsEquipDefinedForModel("EQUIPCOUNT")
	
	rsNumPcsEquipDefinedForModel.Close
	set rsNumPcsEquipDefinedForModel = Nothing
	cnnNumPcsEquipDefinedForModel.Close	
	set cnnNumPcsEquipDefinedForModel = Nothing
	
	NumberEquipmentRecsDefinedForModel = resultNumPcsEquipDefinedForModel
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function NumberCustomerEquipmentRecsDefinedForModel(passedModelIntRecID)

	Set cnnNumCustPcsEquipDefinedForModel = Server.CreateObject("ADODB.Connection")
	cnnNumCustPcsEquipDefinedForModel.open Session("ClientCnnString")

	resultNumCustPcsEquipDefinedForModel = 0
		
	SQLNumCustPcsEquipDefinedForModel = "SELECT COUNT(*) AS EQUIPCOUNT FROM EQ_CustomerEquipment INNER JOIN EQ_Equipment ON "
	SQLNumCustPcsEquipDefinedForModel = SQLNumCustPcsEquipDefinedForModel & "EQ_Equipment.InternalRecordIdentifier = EQ_CustomerEquipment.EquipIntRecID "
	SQLNumCustPcsEquipDefinedForModel = SQLNumCustPcsEquipDefinedForModel & " WHERE EQ_Equipment.ModelIntRecID = " & passedModelIntRecID
	 
	Set rsNumCustPcsEquipDefinedForModel = Server.CreateObject("ADODB.Recordset")
	rsNumCustPcsEquipDefinedForModel.CursorLocation = 3 
	
	rsNumCustPcsEquipDefinedForModel.Open SQLNumCustPcsEquipDefinedForModel,cnnNumCustPcsEquipDefinedForModel 
			
	resultNumCustPcsEquipDefinedForModel = rsNumCustPcsEquipDefinedForModel("EQUIPCOUNT")
	
	rsNumCustPcsEquipDefinedForModel.Close
	set rsNumCustPcsEquipDefinedForModel = Nothing
	cnnNumCustPcsEquipDefinedForModel.Close	
	set cnnNumCustPcsEquipDefinedForModel = Nothing
	
	NumberCustomerEquipmentRecsDefinedForModel = resultNumCustPcsEquipDefinedForModel
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************





'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function NumberEquipmentRecsDefinedForCondition(passedConditionIntRecID)

	Set cnnNumPcsEquipDefinedForCondition = Server.CreateObject("ADODB.Connection")
	cnnNumPcsEquipDefinedForCondition.open Session("ClientCnnString")

	resultNumPcsEquipDefinedForCondition = 0
		
	SQLNumPcsEquipDefinedForCondition = "SELECT COUNT(*) AS EQUIPCOUNT FROM EQ_Equipment WHERE CurrentConditionIntRecID = " & passedConditionIntRecID
	
	 
	Set rsNumPcsEquipDefinedForCondition = Server.CreateObject("ADODB.Recordset")
	rsNumPcsEquipDefinedForCondition.CursorLocation = 3 
	
	rsNumPcsEquipDefinedForCondition.Open SQLNumPcsEquipDefinedForCondition,cnnNumPcsEquipDefinedForCondition 
			
	resultNumPcsEquipDefinedForCondition = rsNumPcsEquipDefinedForCondition("EQUIPCOUNT")
	
	rsNumPcsEquipDefinedForCondition.Close
	set rsNumPcsEquipDefinedForCondition = Nothing
	cnnNumPcsEquipDefinedForCondition.Close	
	set cnnNumPcsEquipDefinedForCondition = Nothing
	
	NumberEquipmentRecsDefinedForCondition = resultNumPcsEquipDefinedForCondition
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function NumberCustomerEquipmentRecsDefinedForCondition(passedConditionIntRecID)

	Set cnnNumCustPcsEquipDefinedForCondition = Server.CreateObject("ADODB.Connection")
	cnnNumCustPcsEquipDefinedForCondition.open Session("ClientCnnString")

	resultNumCustPcsEquipDefinedForCondition = 0
	
	SQLNumCustPcsEquipDefinedForCondition = "SELECT COUNT(*) AS EQUIPCOUNT FROM EQ_CustomerEquipment INNER JOIN EQ_Equipment ON "
	SQLNumCustPcsEquipDefinedForCondition = SQLNumCustPcsEquipDefinedForCondition & "EQ_Equipment.InternalRecordIdentifier = EQ_CustomerEquipment.EquipIntRecID "
	SQLNumCustPcsEquipDefinedForCondition = SQLNumCustPcsEquipDefinedForCondition & " WHERE EQ_Equipment.CurrentConditionIntRecID = " & passedConditionIntRecID

	 
	Set rsNumCustPcsEquipDefinedForCondition = Server.CreateObject("ADODB.Recordset")
	rsNumCustPcsEquipDefinedForCondition.CursorLocation = 3 
	
	rsNumCustPcsEquipDefinedForCondition.Open SQLNumCustPcsEquipDefinedForCondition,cnnNumCustPcsEquipDefinedForCondition 
			
	resultNumCustPcsEquipDefinedForCondition = rsNumCustPcsEquipDefinedForCondition("EQUIPCOUNT")
	
	rsNumCustPcsEquipDefinedForCondition.Close
	set rsNumCustPcsEquipDefinedForCondition = Nothing
	cnnNumCustPcsEquipDefinedForCondition.Close	
	set cnnNumCustPcsEquipDefinedForCondition = Nothing
	
	NumberCustomerEquipmentRecsDefinedForCondition = resultNumCustPcsEquipDefinedForCondition
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************




'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function NumberEquipmentRecsDefinedForStatusCode(passedStatusCodeIntRecID)

	Set cnnNumPcsEquipDefinedForStatusCode = Server.CreateObject("ADODB.Connection")
	cnnNumPcsEquipDefinedForStatusCode.open Session("ClientCnnString")

	resultNumPcsEquipDefinedForStatusCode = 0
		
	SQLNumPcsEquipDefinedForStatusCode = "SELECT COUNT(*) AS EQUIPCOUNT FROM EQ_Equipment WHERE StatusCodeIntRecID = " & passedStatusCodeIntRecID
	
	 
	Set rsNumPcsEquipDefinedForStatusCode = Server.CreateObject("ADODB.Recordset")
	rsNumPcsEquipDefinedForStatusCode.CursorLocation = 3 
	
	rsNumPcsEquipDefinedForStatusCode.Open SQLNumPcsEquipDefinedForStatusCode,cnnNumPcsEquipDefinedForStatusCode 
			
	resultNumPcsEquipDefinedForStatusCode = rsNumPcsEquipDefinedForStatusCode("EQUIPCOUNT")
	
	rsNumPcsEquipDefinedForStatusCode.Close
	set rsNumPcsEquipDefinedForStatusCode = Nothing
	cnnNumPcsEquipDefinedForStatusCode.Close	
	set cnnNumPcsEquipDefinedForStatusCode = Nothing
	
	NumberEquipmentRecsDefinedForStatusCode = resultNumPcsEquipDefinedForStatusCode
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function NumberCustomerEquipmentRecsDefinedForStatusCode(passedStatusCodeIntRecID)

	Set cnnNumCustPcsEquipDefinedForStatusCode = Server.CreateObject("ADODB.Connection")
	cnnNumCustPcsEquipDefinedForStatusCode.open Session("ClientCnnString")

	resultNumCustPcsEquipDefinedForStatusCode = 0
	
	SQLNumCustPcsEquipDefinedForStatusCode = "SELECT COUNT(*) AS EQUIPCOUNT FROM EQ_CustomerEquipment INNER JOIN EQ_Equipment ON "
	SQLNumCustPcsEquipDefinedForStatusCode = SQLNumCustPcsEquipDefinedForStatusCode & "EQ_Equipment.InternalRecordIdentifier = EQ_CustomerEquipment.EquipIntRecID "
	SQLNumCustPcsEquipDefinedForStatusCode = SQLNumCustPcsEquipDefinedForStatusCode & " WHERE EQ_Equipment.StatusCodeIntRecID = " & passedStatusCodeIntRecID

	 
	Set rsNumCustPcsEquipDefinedForStatusCode = Server.CreateObject("ADODB.Recordset")
	rsNumCustPcsEquipDefinedForStatusCode.CursorLocation = 3 
	
	rsNumCustPcsEquipDefinedForStatusCode.Open SQLNumCustPcsEquipDefinedForStatusCode,cnnNumCustPcsEquipDefinedForStatusCode 
			
	resultNumCustPcsEquipDefinedForStatusCode = rsNumCustPcsEquipDefinedForStatusCode("EQUIPCOUNT")
	
	rsNumCustPcsEquipDefinedForStatusCode.Close
	set rsNumCustPcsEquipDefinedForStatusCode = Nothing
	cnnNumCustPcsEquipDefinedForStatusCode.Close	
	set cnnNumCustPcsEquipDefinedForStatusCode = Nothing
	
	NumberCustomerEquipmentRecsDefinedForStatusCode = resultNumCustPcsEquipDefinedForStatusCode
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************




'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function NumberModelsWithStatusCode(passedStatusCodeIntRecID,passedModelIntRecID)

	Set cnnNumberModelsWithStatusCode = Server.CreateObject("ADODB.Connection")
	cnnNumberModelsWithStatusCode.open Session("ClientCnnString")

	resultNumberModelsWithStatusCode = 0
	
	SQLNumberModelsWithStatusCode = "SELECT COUNT(*) AS EQUIPCOUNT FROM EQ_Equipment "
	SQLNumberModelsWithStatusCode = SQLNumberModelsWithStatusCode & " WHERE EQ_Equipment.StatusCodeIntRecID = " & passedStatusCodeIntRecID
	SQLNumberModelsWithStatusCode = SQLNumberModelsWithStatusCode & " AND EQ_Equipment.ModelIntRecID = " & passedModelIntRecID

	Set rsNumberModelsWithStatusCode = Server.CreateObject("ADODB.Recordset")
	rsNumberModelsWithStatusCode.CursorLocation = 3 
	
	rsNumberModelsWithStatusCode.Open SQLNumberModelsWithStatusCode,cnnNumberModelsWithStatusCode 
			
	resultNumberModelsWithStatusCode = rsNumberModelsWithStatusCode("EQUIPCOUNT")
	
	rsNumberModelsWithStatusCode.Close
	set rsNumberModelsWithStatusCode = Nothing
	cnnNumberModelsWithStatusCode.Close	
	set cnnNumberModelsWithStatusCode = Nothing
	
	NumberModelsWithStatusCode = resultNumberModelsWithStatusCode
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************





'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function NumberEquipmentRecsDefinedForMovementCode(passedMovementCodeIntRecID)

	Set cnnNumPcsEquipDefinedForMovementCode = Server.CreateObject("ADODB.Connection")
	cnnNumPcsEquipDefinedForMovementCode.open Session("ClientCnnString")

	resultNumPcsEquipDefinedForMovementCode = 0
		
	SQLNumPcsEquipDefinedForMovementCode = "SELECT COUNT(*) AS EQUIPCOUNT FROM EQ_Equipment WHERE MovementCodeIntRecID = " & passedMovementCodeIntRecID
	
	 
	Set rsNumPcsEquipDefinedForMovementCode = Server.CreateObject("ADODB.Recordset")
	rsNumPcsEquipDefinedForMovementCode.CursorLocation = 3 
	
	rsNumPcsEquipDefinedForMovementCode.Open SQLNumPcsEquipDefinedForMovementCode,cnnNumPcsEquipDefinedForMovementCode 
			
	resultNumPcsEquipDefinedForMovementCode = rsNumPcsEquipDefinedForMovementCode("EQUIPCOUNT")
	
	rsNumPcsEquipDefinedForMovementCode.Close
	set rsNumPcsEquipDefinedForMovementCode = Nothing
	cnnNumPcsEquipDefinedForMovementCode.Close	
	set cnnNumPcsEquipDefinedForMovementCode = Nothing
	
	NumberEquipmentRecsDefinedForMovementCode = resultNumPcsEquipDefinedForMovementCode
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function NumberCustomerEquipmentRecsDefinedForMovementCode(passedMovementCodeIntRecID)

	Set cnnNumCustPcsEquipDefinedForMovementCode = Server.CreateObject("ADODB.Connection")
	cnnNumCustPcsEquipDefinedForMovementCode.open Session("ClientCnnString")

	resultNumCustPcsEquipDefinedForMovementCode = 0
	
	SQLNumCustPcsEquipDefinedForMovementCode = "SELECT COUNT(*) AS EQUIPCOUNT FROM EQ_CustomerEquipment INNER JOIN EQ_Equipment ON "
	SQLNumCustPcsEquipDefinedForMovementCode = SQLNumCustPcsEquipDefinedForMovementCode & "EQ_Equipment.InternalRecordIdentifier = EQ_CustomerEquipment.EquipIntRecID "
	SQLNumCustPcsEquipDefinedForMovementCode = SQLNumCustPcsEquipDefinedForMovementCode & " WHERE EQ_Equipment.MovementCodeIntRecID = " & passedMovementCodeIntRecID

	 
	Set rsNumCustPcsEquipDefinedForMovementCode = Server.CreateObject("ADODB.Recordset")
	rsNumCustPcsEquipDefinedForMovementCode.CursorLocation = 3 
	
	rsNumCustPcsEquipDefinedForMovementCode.Open SQLNumCustPcsEquipDefinedForMovementCode,cnnNumCustPcsEquipDefinedForMovementCode 
			
	resultNumCustPcsEquipDefinedForMovementCode = rsNumCustPcsEquipDefinedForMovementCode("EQUIPCOUNT")
	
	rsNumCustPcsEquipDefinedForMovementCode.Close
	set rsNumCustPcsEquipDefinedForMovementCode = Nothing
	cnnNumCustPcsEquipDefinedForMovementCode.Close	
	set cnnNumCustPcsEquipDefinedForMovementCode = Nothing
	
	NumberCustomerEquipmentRecsDefinedForMovementCode = resultNumCustPcsEquipDefinedForMovementCode
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************







'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function NumberEquipmentRecsDefinedForAcquisitionCode(passedAcquisitionCodeIntRecID)

	Set cnnNumPcsEquipDefinedForAcquisitionCode = Server.CreateObject("ADODB.Connection")
	cnnNumPcsEquipDefinedForAcquisitionCode.open Session("ClientCnnString")

	resultNumPcsEquipDefinedForAcquisitionCode = 0
		
	SQLNumPcsEquipDefinedForAcquisitionCode = "SELECT COUNT(*) AS EQUIPCOUNT FROM EQ_Equipment WHERE AcquisitionCodeIntRecID = " & passedAcquisitionCodeIntRecID
	
	 
	Set rsNumPcsEquipDefinedForAcquisitionCode = Server.CreateObject("ADODB.Recordset")
	rsNumPcsEquipDefinedForAcquisitionCode.CursorLocation = 3 
	
	rsNumPcsEquipDefinedForAcquisitionCode.Open SQLNumPcsEquipDefinedForAcquisitionCode,cnnNumPcsEquipDefinedForAcquisitionCode 
			
	resultNumPcsEquipDefinedForAcquisitionCode = rsNumPcsEquipDefinedForAcquisitionCode("EQUIPCOUNT")
	
	rsNumPcsEquipDefinedForAcquisitionCode.Close
	set rsNumPcsEquipDefinedForAcquisitionCode = Nothing
	cnnNumPcsEquipDefinedForAcquisitionCode.Close	
	set cnnNumPcsEquipDefinedForAcquisitionCode = Nothing
	
	NumberEquipmentRecsDefinedForAcquisitionCode = resultNumPcsEquipDefinedForAcquisitionCode
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function NumberCustomerEquipmentRecsDefinedForAcquisitionCode(passedAcquisitionCodeIntRecID)

	Set cnnNumCustPcsEquipDefinedForAcquisitionCode = Server.CreateObject("ADODB.Connection")
	cnnNumCustPcsEquipDefinedForAcquisitionCode.open Session("ClientCnnString")

	resultNumCustPcsEquipDefinedForAcquisitionCode = 0
	
	SQLNumCustPcsEquipDefinedForAcquisitionCode = "SELECT COUNT(*) AS EQUIPCOUNT FROM EQ_CustomerEquipment INNER JOIN EQ_Equipment ON "
	SQLNumCustPcsEquipDefinedForAcquisitionCode = SQLNumCustPcsEquipDefinedForAcquisitionCode & "EQ_Equipment.InternalRecordIdentifier = EQ_CustomerEquipment.EquipIntRecID "
	SQLNumCustPcsEquipDefinedForAcquisitionCode = SQLNumCustPcsEquipDefinedForAcquisitionCode & " WHERE EQ_Equipment.AcquisitionCodeIntRecID = " & passedAcquisitionCodeIntRecID

	 
	Set rsNumCustPcsEquipDefinedForAcquisitionCode = Server.CreateObject("ADODB.Recordset")
	rsNumCustPcsEquipDefinedForAcquisitionCode.CursorLocation = 3 
	
	rsNumCustPcsEquipDefinedForAcquisitionCode.Open SQLNumCustPcsEquipDefinedForAcquisitionCode,cnnNumCustPcsEquipDefinedForAcquisitionCode 
			
	resultNumCustPcsEquipDefinedForAcquisitionCode = rsNumCustPcsEquipDefinedForAcquisitionCode("EQUIPCOUNT")
	
	rsNumCustPcsEquipDefinedForAcquisitionCode.Close
	set rsNumCustPcsEquipDefinedForAcquisitionCode = Nothing
	cnnNumCustPcsEquipDefinedForAcquisitionCode.Close	
	set cnnNumCustPcsEquipDefinedForAcquisitionCode = Nothing
	
	NumberCustomerEquipmentRecsDefinedForAcquisitionCode = resultNumCustPcsEquipDefinedForAcquisitionCode
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************










'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function NumberEquipmentRecsDefinedForGroup(passedGroupIntRecID)

	Set cnnNumPcsEquipDefinedForGroup = Server.CreateObject("ADODB.Connection")
	cnnNumPcsEquipDefinedForGroup.open Session("ClientCnnString")

	resultNumPcsEquipDefinedForGroup = 0
		
	SQLNumPcsEquipDefinedForGroup = "SELECT COUNT(*) AS EQUIPCOUNT FROM EQ_Equipment WHERE ModelIntRecID IN "
	SQLNumPcsEquipDefinedForGroup =  SQLNumPcsEquipDefinedForGroup & "(SELECT InternalRecordIdentifier FROM EQ_Models WHERE GroupIntRecID = " & passedGroupIntRecID & ")"
	
	Set rsNumPcsEquipDefinedForGroup = Server.CreateObject("ADODB.Recordset")
	rsNumPcsEquipDefinedForGroup.CursorLocation = 3 
	
	rsNumPcsEquipDefinedForGroup.Open SQLNumPcsEquipDefinedForGroup,cnnNumPcsEquipDefinedForGroup 
			
	resultNumPcsEquipDefinedForGroup = rsNumPcsEquipDefinedForGroup("EQUIPCOUNT")
	
	rsNumPcsEquipDefinedForGroup.Close
	set rsNumPcsEquipDefinedForGroup = Nothing
	cnnNumPcsEquipDefinedForGroup.Close	
	set cnnNumPcsEquipDefinedForGroup = Nothing
	
	NumberEquipmentRecsDefinedForGroup = resultNumPcsEquipDefinedForGroup
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function NumberCustomerEquipmentRecsDefinedForGroup(passedGroupIntRecID)

	Set cnnNumCustPcsEquipDefinedForGroup = Server.CreateObject("ADODB.Connection")
	cnnNumCustPcsEquipDefinedForGroup.open Session("ClientCnnString")

	resultNumCustPcsEquipDefinedForGroup = 0
	
	SQLNumCustPcsEquipDefinedForGroup = "SELECT COUNT(*) AS EQUIPCOUNT FROM EQ_CustomerEquipment INNER JOIN EQ_Equipment ON "
	SQLNumCustPcsEquipDefinedForGroup = SQLNumCustPcsEquipDefinedForGroup & "EQ_Equipment.InternalRecordIdentifier = EQ_CustomerEquipment.EquipIntRecID "
	SQLNumCustPcsEquipDefinedForGroup = SQLNumCustPcsEquipDefinedForGroup & " WHERE EQ_Equipment.ModelIntRecID IN "
	SQLNumCustPcsEquipDefinedForGroup = SQLNumCustPcsEquipDefinedForGroup & "(SELECT InternalRecordIdentifier FROM EQ_Models WHERE GroupIntRecID = " & passedGroupIntRecID & ")"
	 
	Set rsNumCustPcsEquipDefinedForGroup = Server.CreateObject("ADODB.Recordset")
	rsNumCustPcsEquipDefinedForGroup.CursorLocation = 3 
	
	rsNumCustPcsEquipDefinedForGroup.Open SQLNumCustPcsEquipDefinedForGroup,cnnNumCustPcsEquipDefinedForGroup 
			
	resultNumCustPcsEquipDefinedForGroup = rsNumCustPcsEquipDefinedForGroup("EQUIPCOUNT")
	
	rsNumCustPcsEquipDefinedForGroup.Close
	set rsNumCustPcsEquipDefinedForGroup = Nothing
	cnnNumCustPcsEquipDefinedForGroup.Close	
	set cnnNumCustPcsEquipDefinedForGroup = Nothing
	
	NumberCustomerEquipmentRecsDefinedForGroup = resultNumCustPcsEquipDefinedForGroup
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************





'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetManufacturerNameByIntRecID(passedManufIntRecID)

	Set cnnGetManufacturerNameByIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetManufacturerNameByIntRecID.open Session("ClientCnnString")

	resultGetManufacturerNameByIntRecID = ""
		
	SQLGetManufacturerNameByIntRecID = "SELECT * FROM EQ_Manufacturers WHERE InternalRecordIdentifier = " & passedManufIntRecID
	 
	Set rsGetManufacturerNameByIntRecID  = Server.CreateObject("ADODB.Recordset")
	rsGetManufacturerNameByIntRecID.CursorLocation = 3 
	
	rsGetManufacturerNameByIntRecID.Open SQLGetManufacturerNameByIntRecID,cnnGetManufacturerNameByIntRecID 
			
	resultGetManufacturerNameByIntRecID = rsGetManufacturerNameByIntRecID("ManufacturerName")
	
	rsGetManufacturerNameByIntRecID.Close
	set rsGetManufacturerNameByIntRecID = Nothing
	cnnGetManufacturerNameByIntRecID.Close	
	set cnnGetManufacturerNameByIntRecID = Nothing
	
	GetManufacturerNameByIntRecID  = resultGetManufacturerNameByIntRecID 
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetManufacturerNameByBrandIntRecID(passedBrandIntRecID)

	Set cnnGetManufacturerNameByBrandIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetManufacturerNameByBrandIntRecID.open Session("ClientCnnString")

	resultGetManufacturerNameByBrandIntRecID = ""
		
	SQLGetManufacturerNameByBrandIntRecID = "SELECT * FROM EQ_Brands WHERE InternalRecordIdentifier = " & passedBrandIntRecID
	 
	Set rsGetManufacturerNameByBrandIntRecID  = Server.CreateObject("ADODB.Recordset")
	rsGetManufacturerNameByBrandIntRecID.CursorLocation = 3 
	
	rsGetManufacturerNameByBrandIntRecID.Open SQLGetManufacturerNameByBrandIntRecID,cnnGetManufacturerNameByBrandIntRecID 
			
	manufIntRecID = rsGetManufacturerNameByBrandIntRecID("ManufacIntRecID")
	

			Set cnnGetManfName = Server.CreateObject("ADODB.Connection")
			cnnGetManfName.open Session("ClientCnnString")
		
			resultGetManfName = ""
				
			SQLGetManfName = "SELECT * FROM EQ_Manufacturers WHERE InternalRecordIdentifier = " & manufIntRecID
			 
			Set rsGetManfName  = Server.CreateObject("ADODB.Recordset")
			rsGetManfName.CursorLocation = 3 
			
			rsGetManfName.Open SQLGetManfName,cnnGetManfName 
					
			resultGetManufacturerNameByBrandIntRecID = rsGetManfName("ManufacturerName")
			
			rsGetManfName.Close
			set rsGetManfName = Nothing
			cnnGetManfName.Close	
			set cnnGetManfName = Nothing
	
	
	rsGetManufacturerNameByBrandIntRecID.Close
	set rsGetManufacturerNameByBrandIntRecID = Nothing
	cnnGetManufacturerNameByBrandIntRecID.Close	
	set cnnGetManufacturerNameByBrandIntRecID = Nothing
	
	GetManufacturerNameByBrandIntRecID  = resultGetManufacturerNameByBrandIntRecID 
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetManufacturerNameByModelIntRecID(passedModelIntRecID)

	Set cnnGetManufacturerNameByModelIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetManufacturerNameByModelIntRecID.open Session("ClientCnnString")

	resultGetManufacturerNameByModelIntRecID = ""
		
	SQLGetManufacturerNameByModelIntRecID = "SELECT * FROM EQ_MODELS WHERE InternalRecordIdentifier = " & passedModelIntRecID
	 
	Set rsGetManufacturerNameByModelIntRecID  = Server.CreateObject("ADODB.Recordset")
	rsGetManufacturerNameByModelIntRecID.CursorLocation = 3 
	
	rsGetManufacturerNameByModelIntRecID.Open SQLGetManufacturerNameByModelIntRecID,cnnGetManufacturerNameByModelIntRecID 
			
	manufIntRecID = rsGetManufacturerNameByModelIntRecID("ManufacIntRecID")
	

			Set cnnGetManfName = Server.CreateObject("ADODB.Connection")
			cnnGetManfName.open Session("ClientCnnString")
		
			resultGetManfName = ""
				
			SQLGetManfName = "SELECT * FROM EQ_Manufacturers WHERE InternalRecordIdentifier = " & manufIntRecID
			 
			Set rsGetManfName  = Server.CreateObject("ADODB.Recordset")
			rsGetManfName.CursorLocation = 3 
			
			rsGetManfName.Open SQLGetManfName,cnnGetManfName 
					
			resultGetManufacturerNameByModelIntRecID = rsGetManfName("ManufacturerName")
			
			rsGetManfName.Close
			set rsGetManfName = Nothing
			cnnGetManfName.Close	
			set cnnGetManfName = Nothing
	
	
	rsGetManufacturerNameByModelIntRecID.Close
	set rsGetManufacturerNameByModelIntRecID = Nothing
	cnnGetManufacturerNameByModelIntRecID.Close	
	set cnnGetManufacturerNameByModelIntRecID = Nothing
	
	GetManufacturerNameByModelIntRecID  = resultGetManufacturerNameByModelIntRecID 
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetManufacturerIntRecIDByBrandIntRecID(passedBrandIntRecID)

	Set cnnGetManufacturerIntRecIDByBrandIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetManufacturerIntRecIDByBrandIntRecID.open Session("ClientCnnString")

	resultGetManufacturerIntRecIDByBrandIntRecID = ""
		
	SQLGetManufacturerIntRecIDByBrandIntRecID = "SELECT * FROM EQ_Brands WHERE InternalRecordIdentifier = " & passedBrandIntRecID
	 
	Set rsGetManufacturerIntRecIDByBrandIntRecID  = Server.CreateObject("ADODB.Recordset")
	rsGetManufacturerIntRecIDByBrandIntRecID.CursorLocation = 3 
	
	rsGetManufacturerIntRecIDByBrandIntRecID.Open SQLGetManufacturerIntRecIDByBrandIntRecID,cnnGetManufacturerIntRecIDByBrandIntRecID 
			
	resultGetManufacturerIntRecIDByBrandIntRecID = rsGetManufacturerIntRecIDByBrandIntRecID("ManufacIntRecID")
	
	rsGetManufacturerIntRecIDByBrandIntRecID.Close
	set rsGetManufacturerIntRecIDByBrandIntRecID = Nothing
	cnnGetManufacturerIntRecIDByBrandIntRecID.Close	
	set cnnGetManufacturerIntRecIDByBrandIntRecID = Nothing
	
	GetManufacturerIntRecIDByBrandIntRecID  = resultGetManufacturerIntRecIDByBrandIntRecID 
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************




'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetManufacturerIntRecIDByModelIntRecID(passedModelIntRecID)

	Set cnnGetManufacturerIntRecIDByModelIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetManufacturerIntRecIDByModelIntRecID.open Session("ClientCnnString")

	resultGetManufacturerIntRecIDByModelIntRecID = ""
		
	SQLGetManufacturerIntRecIDByModelIntRecID = "SELECT * FROM EQ_Models WHERE InternalRecordIdentifier = " & passedModelIntRecID
	 
	Set rsGetManufacturerIntRecIDByModelIntRecID  = Server.CreateObject("ADODB.Recordset")
	rsGetManufacturerIntRecIDByModelIntRecID.CursorLocation = 3 
	
	rsGetManufacturerIntRecIDByModelIntRecID.Open SQLGetManufacturerIntRecIDByModelIntRecID,cnnGetManufacturerIntRecIDByModelIntRecID 
			
	resultGetManufacturerIntRecIDByModelIntRecID = rsGetManufacturerIntRecIDByModelIntRecID("ManufacIntRecID")
	
	rsGetManufacturerIntRecIDByModelIntRecID.Close
	set rsGetManufacturerIntRecIDByModelIntRecID = Nothing
	cnnGetManufacturerIntRecIDByModelIntRecID.Close	
	set cnnGetManufacturerIntRecIDByModelIntRecID = Nothing
	
	GetManufacturerIntRecIDByModelIntRecID  = resultGetManufacturerIntRecIDByModelIntRecID 
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************






'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetBrandIntRecIDByModelIntRecID(passedModelIntRecID)

	Set cnnGetBrandIntRecIDByModelIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetBrandIntRecIDByModelIntRecID.open Session("ClientCnnString")

	resultGetBrandIntRecIDByModelIntRecID = ""
		
	SQLGetBrandIntRecIDByModelIntRecID = "SELECT * FROM EQ_Models WHERE InternalRecordIdentifier = " & passedModelIntRecID
	 
	Set rsGetBrandIntRecIDByModelIntRecID  = Server.CreateObject("ADODB.Recordset")
	rsGetBrandIntRecIDByModelIntRecID.CursorLocation = 3 
	
	rsGetBrandIntRecIDByModelIntRecID.Open SQLGetBrandIntRecIDByModelIntRecID,cnnGetBrandIntRecIDByModelIntRecID 
			
	resultGetBrandIntRecIDByModelIntRecID = rsGetBrandIntRecIDByModelIntRecID("BrandIntRecID")
	
	rsGetBrandIntRecIDByModelIntRecID.Close
	set rsGetBrandIntRecIDByModelIntRecID = Nothing
	cnnGetBrandIntRecIDByModelIntRecID.Close	
	set cnnGetBrandIntRecIDByModelIntRecID = Nothing
	
	GetBrandIntRecIDByModelIntRecID  = resultGetBrandIntRecIDByModelIntRecID 
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************





'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetBrandNameByIntRecID(passedBrandIntRecID)

	Set cnnGetBrandNameByIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetBrandNameByIntRecID.open Session("ClientCnnString")

	resultGetBrandNameByIntRecID = ""
		
	SQLGetBrandNameByIntRecID = "SELECT * FROM EQ_Brands WHERE InternalRecordIdentifier = " & passedBrandIntRecID
	
	Set rsGetBrandNameByIntRecID  = Server.CreateObject("ADODB.Recordset")
	rsGetBrandNameByIntRecID.CursorLocation = 3 
	
	rsGetBrandNameByIntRecID.Open SQLGetBrandNameByIntRecID,cnnGetBrandNameByIntRecID 
			
	resultGetBrandNameByIntRecID = rsGetBrandNameByIntRecID("Brand")
	
	rsGetBrandNameByIntRecID.Close
	set rsGetBrandNameByIntRecID = Nothing
	cnnGetBrandNameByIntRecID.Close	
	set cnnGetBrandNameByIntRecID = Nothing
	
	GetBrandNameByIntRecID  = resultGetBrandNameByIntRecID 
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************




'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetBrandNameByModelIntRecID(passedModelIntRecID)

	Set cnnGetBrandNameByModelIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetBrandNameByModelIntRecID.open Session("ClientCnnString")

	resultGetBrandNameByModelIntRecID = ""
		
	SQLGetBrandNameByModelIntRecID = "SELECT * FROM EQ_MODELS WHERE InternalRecordIdentifier = " & passedModelIntRecID
	 
	Set rsGetBrandNameByModelIntRecID  = Server.CreateObject("ADODB.Recordset")
	rsGetBrandNameByModelIntRecID.CursorLocation = 3 
	
	rsGetBrandNameByModelIntRecID.Open SQLGetBrandNameByModelIntRecID,cnnGetBrandNameByModelIntRecID 
			
	brandIntRecID = rsGetBrandNameByModelIntRecID("BrandIntRecID")
	

			Set cnnGetBrandName = Server.CreateObject("ADODB.Connection")
			cnnGetBrandName.open Session("ClientCnnString")
		
			resultGetBrandName = ""
				
			SQLGetBrandName = "SELECT * FROM EQ_Brands WHERE InternalRecordIdentifier = " & brandIntRecID
			 
			Set rsGetBrandName  = Server.CreateObject("ADODB.Recordset")
			rsGetBrandName.CursorLocation = 3 
			
			rsGetBrandName.Open SQLGetBrandName,cnnGetBrandName 
					
			resultGetBrandNameByModelIntRecID = rsGetBrandName("Brand")
			
			rsGetBrandName.Close
			set rsGetBrandName = Nothing
			cnnGetBrandName.Close	
			set cnnGetBrandName = Nothing
	
	
	rsGetBrandNameByModelIntRecID.Close
	set rsGetBrandNameByModelIntRecID = Nothing
	cnnGetBrandNameByModelIntRecID.Close	
	set cnnGetBrandNameByModelIntRecID = Nothing
	
	GetBrandNameByModelIntRecID  = resultGetBrandNameByModelIntRecID 
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************






'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetModelNameByIntRecID(passedModelIntRecID)

	Set cnnGetModelNameByIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetModelNameByIntRecID.open Session("ClientCnnString")

	resultGetModelNameByIntRecID = ""
		
	SQLGetModelNameByIntRecID = "SELECT * FROM EQ_Models WHERE InternalRecordIdentifier = " & passedModelIntRecID
	 
	Set rsGetModelNameByIntRecID  = Server.CreateObject("ADODB.Recordset")
	rsGetModelNameByIntRecID.CursorLocation = 3 
	
	rsGetModelNameByIntRecID.Open SQLGetModelNameByIntRecID,cnnGetModelNameByIntRecID 
			
	resultGetModelNameByIntRecID = rsGetModelNameByIntRecID("Model")
	
	rsGetModelNameByIntRecID.Close
	set rsGetModelNameByIntRecID = Nothing
	cnnGetModelNameByIntRecID.Close	
	set cnnGetModelNameByIntRecID = Nothing
	
	GetModelNameByIntRecID  = resultGetModelNameByIntRecID 
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetMovementCodeDescByIntRecID(passedMovementCodeIntRecID)

	Set cnnGetMovementCodeDescByIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetMovementCodeDescByIntRecID.open Session("ClientCnnString")

	resultGetMovementCodeDescByIntRecID = ""
		
	SQLGetMovementCodeDescByIntRecID = "SELECT * FROM EQ_MovementCodes WHERE InternalRecordIdentifier = " & passedMovementCodeIntRecID
	 
	Set rsGetMovementCodeDescByIntRecID  = Server.CreateObject("ADODB.Recordset")
	rsGetMovementCodeDescByIntRecID.CursorLocation = 3 
	
	rsGetMovementCodeDescByIntRecID.Open SQLGetMovementCodeDescByIntRecID,cnnGetMovementCodeDescByIntRecID 
			
	resultGetMovementCodeDescByIntRecID = rsGetMovementCodeDescByIntRecID("movementDesc")
	
	rsGetMovementCodeDescByIntRecID.Close
	set rsGetMovementCodeDescByIntRecID = Nothing
	cnnGetMovementCodeDescByIntRecID.Close	
	set cnnGetMovementCodeDescByIntRecID = Nothing
	
	GetMovementCodeDescByIntRecID  = resultGetMovementCodeDescByIntRecID 
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************




'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetMovementCodeByIntRecID(passedMovementCodeIntRecID)

	Set cnnGetMovementCodeByIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetMovementCodeByIntRecID.open Session("ClientCnnString")

	resultGetMovementCodeByIntRecID = ""
		
	SQLGetMovementCodeByIntRecID = "SELECT * FROM EQ_MovementCodes WHERE InternalRecordIdentifier = " & passedMovementCodeIntRecID
	 
	Set rsGetMovementCodeByIntRecID  = Server.CreateObject("ADODB.Recordset")
	rsGetMovementCodeByIntRecID.CursorLocation = 3 
	
	rsGetMovementCodeByIntRecID.Open SQLGetMovementCodeByIntRecID,cnnGetMovementCodeByIntRecID 
			
	resultGetMovementCodeByIntRecID = rsGetMovementCodeByIntRecID("movementCode")
	
	rsGetMovementCodeByIntRecID.Close
	set rsGetMovementCodeByIntRecID = Nothing
	cnnGetMovementCodeByIntRecID.Close	
	set cnnGetMovementCodeByIntRecID = Nothing
	
	GetMovementCodeByIntRecID  = resultGetMovementCodeByIntRecID 
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************





'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetAcquisitionCodeDescByIntRecID(passedAcquisitionCodeIntRecID)

	Set cnnGetAcquisitionCodeDescByIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetAcquisitionCodeDescByIntRecID.open Session("ClientCnnString")

	resultGetAcquisitionCodeDescByIntRecID = ""
		
	SQLGetAcquisitionCodeDescByIntRecID = "SELECT * FROM EQ_AcquisitionCodes WHERE InternalRecordIdentifier = " & passedAcquisitionCodeIntRecID
	 
	Set rsGetAcquisitionCodeDescByIntRecID  = Server.CreateObject("ADODB.Recordset")
	rsGetAcquisitionCodeDescByIntRecID.CursorLocation = 3 
	
	rsGetAcquisitionCodeDescByIntRecID.Open SQLGetAcquisitionCodeDescByIntRecID,cnnGetAcquisitionCodeDescByIntRecID 
			
	resultGetAcquisitionCodeDescByIntRecID = rsGetAcquisitionCodeDescByIntRecID("AcquisitionDesc")
	
	rsGetAcquisitionCodeDescByIntRecID.Close
	set rsGetAcquisitionCodeDescByIntRecID = Nothing
	cnnGetAcquisitionCodeDescByIntRecID.Close	
	set cnnGetAcquisitionCodeDescByIntRecID = Nothing
	
	GetAcquisitionCodeDescByIntRecID  = resultGetAcquisitionCodeDescByIntRecID 
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************




'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetAcquisitionCodeByIntRecID(passedAcquisitionCodeIntRecID)

	Set cnnGetAcquisitionCodeByIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetAcquisitionCodeByIntRecID.open Session("ClientCnnString")

	resultGetAcquisitionCodeByIntRecID = ""
		
	SQLGetAcquisitionCodeByIntRecID = "SELECT * FROM EQ_AcquisitionCodes WHERE InternalRecordIdentifier = " & passedAcquisitionCodeIntRecID
	 
	Set rsGetAcquisitionCodeByIntRecID  = Server.CreateObject("ADODB.Recordset")
	rsGetAcquisitionCodeByIntRecID.CursorLocation = 3 
	
	rsGetAcquisitionCodeByIntRecID.Open SQLGetAcquisitionCodeByIntRecID,cnnGetAcquisitionCodeByIntRecID 
			
	resultGetAcquisitionCodeByIntRecID = rsGetAcquisitionCodeByIntRecID("AcquisitionCode")
	
	rsGetAcquisitionCodeByIntRecID.Close
	set rsGetAcquisitionCodeByIntRecID = Nothing
	cnnGetAcquisitionCodeByIntRecID.Close	
	set cnnGetAcquisitionCodeByIntRecID = Nothing
	
	GetAcquisitionCodeByIntRecID  = resultGetAcquisitionCodeByIntRecID 
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function NumberEquipmentRecsDefinedForModelByCondition(passedConditionIntRecID)

	Set cnnNumPcsEquipDefinedForCondition = Server.CreateObject("ADODB.Connection")
	cnnNumPcsEquipDefinedForCondition.open Session("ClientCnnString")

	resultNumPcsEquipDefinedForCondition = 0
		
	SQLNumPcsEquipDefinedForCondition = "SELECT COUNT(*) AS EQUIPCOUNT FROM EQ_Equipment WHERE CurrentConditionIntRecID = " & passedConditionIntRecID
	 
	Set rsNumPcsEquipDefinedForCondition = Server.CreateObject("ADODB.Recordset")
	rsNumPcsEquipDefinedForCondition.CursorLocation = 3 
	
	rsNumPcsEquipDefinedForCondition.Open SQLNumPcsEquipDefinedForCondition,cnnNumPcsEquipDefinedForCondition 
			
	resultNumPcsEquipDefinedForCondition = rsNumPcsEquipDefinedForCondition("EQUIPCOUNT")
	
	rsNumPcsEquipDefinedForCondition.Close
	set rsNumPcsEquipDefinedForCondition = Nothing
	cnnNumPcsEquipDefinedForCondition.Close	
	set cnnNumPcsEquipDefinedForCondition = Nothing
	
	NumberEquipmentRecsDefinedForModelByCondition = resultNumPcsEquipDefinedForCondition
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function NumberEquipmentRecsDefinedForClass(passedClassIntRecID)

	Set cnnNumPcsEquipDefinedForClass = Server.CreateObject("ADODB.Connection")
	cnnNumPcsEquipDefinedForClass.open Session("ClientCnnString")

	resultNumPcsEquipDefinedForClass = 0
		
	SQLNumPcsEquipDefinedForClass = "SELECT COUNT(*) AS EQUIPCOUNT FROM EQ_Equipment WHERE ModelIntRecID IN "
	SQLNumPcsEquipDefinedForClass =  SQLNumPcsEquipDefinedForClass & "(SELECT InternalRecordIdentifier FROM EQ_Models WHERE ClassIntRecID = " & passedClassIntRecID & ")"
	
	Set rsNumPcsEquipDefinedForClass = Server.CreateObject("ADODB.Recordset")
	rsNumPcsEquipDefinedForClass.CursorLocation = 3 
	
	rsNumPcsEquipDefinedForClass.Open SQLNumPcsEquipDefinedForClass,cnnNumPcsEquipDefinedForClass 
			
	resultNumPcsEquipDefinedForClass = rsNumPcsEquipDefinedForClass("EQUIPCOUNT")
	
	rsNumPcsEquipDefinedForClass.Close
	set rsNumPcsEquipDefinedForClass = Nothing
	cnnNumPcsEquipDefinedForClass.Close	
	set cnnNumPcsEquipDefinedForClass = Nothing
	
	NumberEquipmentRecsDefinedForClass = resultNumPcsEquipDefinedForClass
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************




'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function NumberCustomerEquipmentRecsDefinedForClass(passedClassIntRecID)

	Set cnnNumCustPcsEquipDefinedForClass = Server.CreateObject("ADODB.Connection")
	cnnNumCustPcsEquipDefinedForClass.open Session("ClientCnnString")

	resultNumCustPcsEquipDefinedForClass = 0
	
	SQLNumCustPcsEquipDefinedForClass = "SELECT COUNT(*) AS EQUIPCOUNT FROM EQ_CustomerEquipment INNER JOIN EQ_Equipment ON "
	SQLNumCustPcsEquipDefinedForClass = SQLNumCustPcsEquipDefinedForClass & "EQ_Equipment.InternalRecordIdentifier = EQ_CustomerEquipment.EquipIntRecID "
	SQLNumCustPcsEquipDefinedForClass = SQLNumCustPcsEquipDefinedForClass & " WHERE EQ_Equipment.ModelIntRecID IN "
	SQLNumCustPcsEquipDefinedForClass =  SQLNumCustPcsEquipDefinedForClass & "(SELECT InternalRecordIdentifier FROM EQ_Models WHERE ClassIntRecID = " & passedClassIntRecID & ")"

	
	Set rsNumCustPcsEquipDefinedForClass = Server.CreateObject("ADODB.Recordset")
	rsNumCustPcsEquipDefinedForClass.CursorLocation = 3 
	
	rsNumCustPcsEquipDefinedForClass.Open SQLNumCustPcsEquipDefinedForClass,cnnNumCustPcsEquipDefinedForClass 
			
	resultNumCustPcsEquipDefinedForClass = rsNumCustPcsEquipDefinedForClass("EQUIPCOUNT")
	
	rsNumCustPcsEquipDefinedForClass.Close
	set rsNumCustPcsEquipDefinedForClass = Nothing
	cnnNumCustPcsEquipDefinedForClass.Close	
	set cnnNumCustPcsEquipDefinedForClass = Nothing
	
	NumberCustomerEquipmentRecsDefinedForClass = resultNumCustPcsEquipDefinedForClass
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function NumberCustomerEquipmentRecsDefinedForManufacturer(passedManfIntRecID)

	Set cnnNumCustPcsEquipDefinedForManf = Server.CreateObject("ADODB.Connection")
	cnnNumCustPcsEquipDefinedForManf.open Session("ClientCnnString")

	resultNumCustPcsEquipDefinedForManf = 0
		
	SQLNumCustPcsEquipDefinedForManf = "SELECT COUNT(*) AS EQUIPCOUNT FROM EQ_CustomerEquipment INNER JOIN EQ_Equipment ON "
	SQLNumCustPcsEquipDefinedForManf = SQLNumCustPcsEquipDefinedForManf & "EQ_Equipment.InternalRecordIdentifier = EQ_CustomerEquipment.EquipIntRecID "
	SQLNumCustPcsEquipDefinedForManf = SQLNumCustPcsEquipDefinedForManf & " WHERE EQ_Equipment.ModelIntRecID IN "
	SQLNumCustPcsEquipDefinedForManf = SQLNumCustPcsEquipDefinedForManf & "(SELECT InternalRecordIdentifier FROM EQ_Models WHERE ManufacIntRecID = " & passedManfIntRecID & ")"
	 
	Set rsNumCustPcsEquipDefinedForManf = Server.CreateObject("ADODB.Recordset")
	rsNumCustPcsEquipDefinedForManf.CursorLocation = 3 
	
	rsNumCustPcsEquipDefinedForManf.Open SQLNumCustPcsEquipDefinedForManf,cnnNumCustPcsEquipDefinedForManf 
			
	resultNumCustPcsEquipDefinedForManf = rsNumCustPcsEquipDefinedForManf("EQUIPCOUNT")
	
	rsNumCustPcsEquipDefinedForManf.Close
	set rsNumCustPcsEquipDefinedForManf = Nothing
	cnnNumCustPcsEquipDefinedForManf.Close	
	set cnnNumCustPcsEquipDefinedForManf = Nothing
	
	NumberCustomerEquipmentRecsDefinedForManufacturer = resultNumCustPcsEquipDefinedForManf
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************




'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetClassNameByIntRecID(passedClassIntRecID)

	Set cnnGetClassNameByIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetClassNameByIntRecID.open Session("ClientCnnString")

	resultGetClassNameByIntRecID = ""
		
	SQLGetClassNameByIntRecID = "SELECT * FROM EQ_Classes WHERE InternalRecordIdentifier = " & passedClassIntRecID
	 
	Set rsGetClassNameByIntRecID  = Server.CreateObject("ADODB.Recordset")
	rsGetClassNameByIntRecID.CursorLocation = 3 
	
	rsGetClassNameByIntRecID.Open SQLGetClassNameByIntRecID,cnnGetClassNameByIntRecID 
			
	resultGetClassNameByIntRecID = rsGetClassNameByIntRecID("Class")
	
	rsGetClassNameByIntRecID.Close
	set rsGetClassNameByIntRecID = Nothing
	cnnGetClassNameByIntRecID.Close	
	set cnnGetClassNameByIntRecID = Nothing
	
	GetClassNameByIntRecID  = resultGetClassNameByIntRecID 
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetGroupNameByIntRecID(passedGroupIntRecID)

	Set cnnGetGroupNameByIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetGroupNameByIntRecID.open Session("ClientCnnString")

	resultGetGroupNameByIntRecID = ""
		
	SQLGetGroupNameByIntRecID = "SELECT * FROM EQ_Groups WHERE InternalRecordIdentifier = " & passedGroupIntRecID
	 
	Set rsGetGroupNameByIntRecID  = Server.CreateObject("ADODB.Recordset")
	rsGetGroupNameByIntRecID.CursorLocation = 3 
	
	rsGetGroupNameByIntRecID.Open SQLGetGroupNameByIntRecID,cnnGetGroupNameByIntRecID 
			
	resultGetGroupNameByIntRecID = rsGetGroupNameByIntRecID("GroupName")
	
	rsGetGroupNameByIntRecID.Close
	set rsGetGroupNameByIntRecID = Nothing
	cnnGetGroupNameByIntRecID.Close	
	set cnnGetGroupNameByIntRecID = Nothing
	
	GetGroupNameByIntRecID  = resultGetGroupNameByIntRecID 
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetConditionNameByIntRecID(passedConditionIntRecID)

	Set cnnGetConditionNameByIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetConditionNameByIntRecID.open Session("ClientCnnString")

	resultGetConditionNameByIntRecID = ""
		
	SQLGetConditionNameByIntRecID = "SELECT * FROM EQ_Condition WHERE InternalRecordIdentifier = " & passedConditionIntRecID
	 
	Set rsGetConditionNameByIntRecID  = Server.CreateObject("ADODB.Recordset")
	rsGetConditionNameByIntRecID.CursorLocation = 3 
	
	rsGetConditionNameByIntRecID.Open SQLGetConditionNameByIntRecID,cnnGetConditionNameByIntRecID 
			
	resultGetConditionNameByIntRecID = rsGetConditionNameByIntRecID("Condition")
	
	rsGetConditionNameByIntRecID.Close
	set rsGetConditionNameByIntRecID = Nothing
	cnnGetConditionNameByIntRecID.Close	
	set cnnGetConditionNameByIntRecID = Nothing
	
	GetConditionNameByIntRecID  = resultGetConditionNameByIntRecID 
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetStatusCodeNameByIntRecID(passedStatusCodeIntRecID)

	Set cnnGetStatusCodeNameByIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetStatusCodeNameByIntRecID.open Session("ClientCnnString")

	resultGetStatusCodeNameByIntRecID = ""
		
	SQLGetStatusCodeNameByIntRecID = "SELECT * FROM EQ_StatusCodes WHERE InternalRecordIdentifier = " & passedStatusCodeIntRecID
	 
	Set rsGetStatusCodeNameByIntRecID  = Server.CreateObject("ADODB.Recordset")
	rsGetStatusCodeNameByIntRecID.CursorLocation = 3 
	
	rsGetStatusCodeNameByIntRecID.Open SQLGetStatusCodeNameByIntRecID,cnnGetStatusCodeNameByIntRecID 
			
	resultGetStatusCodeNameByIntRecID = rsGetStatusCodeNameByIntRecID("statusDesc")
	
	rsGetStatusCodeNameByIntRecID.Close
	set rsGetStatusCodeNameByIntRecID = Nothing
	cnnGetStatusCodeNameByIntRecID.Close	
	set cnnGetStatusCodeNameByIntRecID = Nothing
	
	GetStatusCodeNameByIntRecID  = resultGetStatusCodeNameByIntRecID 
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetGroupNameByModelIntRecID(passedModelIntRecID)

	Set cnnGetGroupNameByModelIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetGroupNameByModelIntRecID.open Session("ClientCnnString")

	resultGetGroupNameByModelIntRecID = ""
		
	SQLGetGroupNameByModelIntRecID = "SELECT * FROM EQ_MODELS WHERE InternalRecordIdentifier = " & passedModelIntRecID
	 
	Set rsGetGroupNameByModelIntRecID  = Server.CreateObject("ADODB.Recordset")
	rsGetGroupNameByModelIntRecID.CursorLocation = 3 
	
	rsGetGroupNameByModelIntRecID.Open SQLGetGroupNameByModelIntRecID,cnnGetGroupNameByModelIntRecID 
			
	GroupIntRecID = rsGetGroupNameByModelIntRecID("GroupIntRecID")
	

			Set cnnGetGroupName = Server.CreateObject("ADODB.Connection")
			cnnGetGroupName.open Session("ClientCnnString")
		
			resultGetGroupName = ""
				
			SQLGetGroupName = "SELECT * FROM EQ_Groups WHERE InternalRecordIdentifier = " & GroupIntRecID
			 
			Set rsGetGroupName  = Server.CreateObject("ADODB.Recordset")
			rsGetGroupName.CursorLocation = 3 
			
			rsGetGroupName.Open SQLGetGroupName,cnnGetGroupName 
					
			resultGetGroupNameByModelIntRecID = rsGetGroupName("GroupName")
			
			rsGetGroupName.Close
			set rsGetGroupName = Nothing
			cnnGetGroupName.Close	
			set cnnGetGroupName = Nothing
	
	
	rsGetGroupNameByModelIntRecID.Close
	set rsGetGroupNameByModelIntRecID = Nothing
	cnnGetGroupNameByModelIntRecID.Close	
	set cnnGetGroupNameByModelIntRecID = Nothing
	
	GetGroupNameByModelIntRecID = resultGetGroupNameByModelIntRecID 
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetClassNameByModelIntRecID(passedModelIntRecID)

	Set cnnGetClassNameByModelIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetClassNameByModelIntRecID.open Session("ClientCnnString")

	resultGetClassNameByModelIntRecID = ""
		
	SQLGetClassNameByModelIntRecID = "SELECT * FROM EQ_MODELS WHERE InternalRecordIdentifier = " & passedModelIntRecID
	 
	Set rsGetClassNameByModelIntRecID  = Server.CreateObject("ADODB.Recordset")
	rsGetClassNameByModelIntRecID.CursorLocation = 3 
	
	rsGetClassNameByModelIntRecID.Open SQLGetClassNameByModelIntRecID,cnnGetClassNameByModelIntRecID 
			
	classIntRecID = rsGetClassNameByModelIntRecID("ClassIntRecID")
	

			Set cnnGetClassName = Server.CreateObject("ADODB.Connection")
			cnnGetClassName.open Session("ClientCnnString")
		
			resultGetClassName = ""
				
			SQLGetClassName = "SELECT * FROM EQ_Classes WHERE InternalRecordIdentifier = " & classIntRecID
			 
			Set rsGetClassName  = Server.CreateObject("ADODB.Recordset")
			rsGetClassName.CursorLocation = 3 
			
			rsGetClassName.Open SQLGetClassName,cnnGetClassName 
					
			resultGetClassNameByModelIntRecID = rsGetClassName("Class")
			
			rsGetClassName.Close
			set rsGetClassName = Nothing
			cnnGetClassName.Close	
			set cnnGetClassName = Nothing
	
	
	rsGetClassNameByModelIntRecID.Close
	set rsGetClassNameByModelIntRecID = Nothing
	cnnGetClassNameByModelIntRecID.Close	
	set cnnGetClassNameByModelIntRecID = Nothing
	
	GetClassNameByModelIntRecID = resultGetClassNameByModelIntRecID 
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************

'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetClassIDByModelIntRecID(passedModelIntRecID)

	Set cnnGetClassIDByModelIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetClassIDByModelIntRecID.open Session("ClientCnnString")

	resultGetClassIDByModelIntRecID = ""
		
	SQLGetClassIDByModelIntRecID = "SELECT * FROM EQ_MODELS WHERE InternalRecordIdentifier = " & passedModelIntRecID
	 
	Set rsGetClassIDByModelIntRecID  = Server.CreateObject("ADODB.Recordset")
	rsGetClassIDByModelIntRecID.CursorLocation = 3 
	
	rsGetClassIDByModelIntRecID.Open SQLGetClassIDByModelIntRecID,cnnGetClassIDByModelIntRecID 
			
	ClassIntRecID = rsGetClassIDByModelIntRecID("ClassIntRecID")
	

			Set cnnGetClassID = Server.CreateObject("ADODB.Connection")
			cnnGetClassID.open Session("ClientCnnString")
		
			resultGetClassID = ""
				
			SQLGetClassID = "SELECT * FROM EQ_Classes WHERE InternalRecordIdentifier = " & ClassIntRecID
			 
			Set rsGetClassID  = Server.CreateObject("ADODB.Recordset")
			rsGetClassID.CursorLocation = 3 
			
			rsGetClassID.Open SQLGetClassID,cnnGetClassID 
					
			resultGetClassIDByModelIntRecID = rsGetClassID("InternalRecordIdentifier")
			
			rsGetClassID.Close
			set rsGetClassID = Nothing
			cnnGetClassID.Close	
			set cnnGetClassID = Nothing
	
	
	rsGetClassIDByModelIntRecID.Close
	set rsGetClassIDByModelIntRecID = Nothing
	cnnGetClassIDByModelIntRecID.Close	
	set cnnGetClassIDByModelIntRecID = Nothing
	
	GetClassIDByModelIntRecID  = resultGetClassIDByModelIntRecID 
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetTotalNumberOfModelsForCustomer(passedCustID,passedModelIntRecID)

	Set cnnTotalNumModelsForCust = Server.CreateObject("ADODB.Connection")
	cnnTotalNumModelsForCust.open Session("ClientCnnString")

	resultTotalNumModelsForCust = 0
		
	SQLTotalNumModelsForCust = " SELECT COUNT(EQ_Equipment.ModelIntRecID) AS ModelCount "
	SQLTotalNumModelsForCust = SQLTotalNumModelsForCust & " FROM  EQ_CustomerEquipment INNER JOIN "
	SQLTotalNumModelsForCust = SQLTotalNumModelsForCust & " EQ_Equipment ON EQ_Equipment.InternalRecordIdentifier = EQ_CustomerEquipment.EquipIntRecID INNER JOIN "
	SQLTotalNumModelsForCust = SQLTotalNumModelsForCust & " EQ_Models ON EQ_Equipment.ModelIntRecID = EQ_Models.InternalRecordIdentifier "
	SQLTotalNumModelsForCust = SQLTotalNumModelsForCust & " WHERE (EQ_CustomerEquipment.CustID = '" & passedCustID & "') "
	SQLTotalNumModelsForCust = SQLTotalNumModelsForCust & " AND (EQ_Equipment.ModelIntRecID =  " & passedModelIntRecID & ") "
	'SQLTotalNumModelsForCust = SQLTotalNumModelsForCust & " GROUP BY EQ_Equipment.ModelIntRecID "
	
	'Response.Write("<br><br>" & SQLTotalNumModelsForCust  & "<br>")
	
		 
	Set rsTotalNumModelsForCust = Server.CreateObject("ADODB.Recordset")
	rsTotalNumModelsForCust.CursorLocation = 3 
	
	rsTotalNumModelsForCust.Open SQLTotalNumModelsForCust,cnnTotalNumModelsForCust 
			
	resultTotalNumModelsForCust = rsTotalNumModelsForCust("ModelCount")
	
	rsTotalNumModelsForCust.Close
	set rsTotalNumModelsForCust = Nothing
	cnnTotalNumModelsForCust.Close	
	set cnnTotalNumModelsForCust = Nothing
	
	GetTotalNumberOfModelsForCustomer = resultTotalNumModelsForCust
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************




'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetTotalValueOfModelsForCustomer(passedCustID,passedModelIntRecID)

	Set cnnTotalValueOfModelsForCust = Server.CreateObject("ADODB.Connection")
	cnnTotalValueOfModelsForCust.open Session("ClientCnnString")

	resultTotalValueOfModelsForCust = 0
		
	SQLTotalValueOfModelsForCust = " SELECT SUM(EQ_Equipment.PurchaseCost) AS Expr1 "
	SQLTotalValueOfModelsForCust = SQLTotalValueOfModelsForCust & " FROM  EQ_Equipment INNER JOIN "
	SQLTotalValueOfModelsForCust = SQLTotalValueOfModelsForCust & " EQ_CustomerEquipment ON EQ_CustomerEquipment.EquipIntRecID = EQ_Equipment.InternalRecordIdentifier INNER JOIN "
	SQLTotalValueOfModelsForCust = SQLTotalValueOfModelsForCust & " EQ_Models ON EQ_Equipment.ModelIntRecID = EQ_Models.InternalRecordIdentifier "
	SQLTotalValueOfModelsForCust = SQLTotalValueOfModelsForCust & " WHERE (EQ_CustomerEquipment.CustID = '" & passedCustID & "') "
	SQLTotalValueOfModelsForCust = SQLTotalValueOfModelsForCust & " AND (EQ_Equipment.ModelIntRecID =  " & passedModelIntRecID & ") "
		 
	Set rsTotalValueOfModelsForCust = Server.CreateObject("ADODB.Recordset")
	rsTotalValueOfModelsForCust.CursorLocation = 3 
	
	rsTotalValueOfModelsForCust.Open SQLTotalValueOfModelsForCust,cnnTotalValueOfModelsForCust 
			
	resultTotalValueOfModelsForCust = rsTotalValueOfModelsForCust("Expr1")
	
	rsTotalValueOfModelsForCust.Close
	set rsTotalValueOfModelsForCust = Nothing
	cnnTotalValueOfModelsForCust.Close	
	set cnnTotalValueOfModelsForCust = Nothing
	
	GetTotalValueOfModelsForCustomer = resultTotalValueOfModelsForCust
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************





'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetTotalValueOfRentalModelsForCustomer(passedCustID,passedModelIntRecID)

	Set cnnTotalValueOfRentalModelsForCust = Server.CreateObject("ADODB.Connection")
	cnnTotalValueOfRentalModelsForCust.open Session("ClientCnnString")

	resultTotalValueOfRentalModelsForCust = 0
		
	SQLTotalValueOfRentalModelsForCust = " SELECT SUM(EQ_CustomerEquipment.RentAmt) AS Expr1 "
	SQLTotalValueOfRentalModelsForCust = SQLTotalValueOfRentalModelsForCust & " FROM  EQ_Equipment INNER JOIN "
	SQLTotalValueOfRentalModelsForCust = SQLTotalValueOfRentalModelsForCust & " EQ_CustomerEquipment ON EQ_CustomerEquipment.EquipIntRecID = EQ_Equipment.InternalRecordIdentifier INNER JOIN "
	SQLTotalValueOfRentalModelsForCust = SQLTotalValueOfRentalModelsForCust & " EQ_Models ON EQ_Equipment.ModelIntRecID = EQ_Models.InternalRecordIdentifier "
	SQLTotalValueOfRentalModelsForCust = SQLTotalValueOfRentalModelsForCust & " WHERE (EQ_CustomerEquipment.CustID = '" & passedCustID & "') "
	SQLTotalValueOfRentalModelsForCust = SQLTotalValueOfRentalModelsForCust & " AND (EQ_Equipment.ModelIntRecID =  " & passedModelIntRecID & ") "
		 
	Set rsTotalValueOfRentalModelsForCust = Server.CreateObject("ADODB.Recordset")
	rsTotalValueOfRentalModelsForCust.CursorLocation = 3 
	
	rsTotalValueOfRentalModelsForCust.Open SQLTotalValueOfRentalModelsForCust,cnnTotalValueOfRentalModelsForCust 
			
	resultTotalValueOfRentalModelsForCust = rsTotalValueOfRentalModelsForCust("Expr1")
	
	rsTotalValueOfRentalModelsForCust.Close
	set rsTotalValueOfRentalModelsForCust = Nothing
	cnnTotalValueOfRentalModelsForCust.Close	
	set cnnTotalValueOfRentalModelsForCust = Nothing
	
	GetTotalValueOfRentalModelsForCustomer = resultTotalValueOfRentalModelsForCust
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************




'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function CustHasEquipment(passedCustID)

	Set cnnCustHasEquipment = Server.CreateObject("ADODB.Connection")
	cnnCustHasEquipment.open Session("ClientCnnString")

	resultCustHasEquipment = ""
		
	SQLCustHasEquipment = " SELECT COUNT(*) as Expr1 FROM  EQ_CustomerEquipment WHERE CustID = '" & passedCustID & "'"
		 
	Set rsCustHasEquipment = Server.CreateObject("ADODB.Recordset")
	rsCustHasEquipment.CursorLocation = 3 
	
	
	rsCustHasEquipment.Open SQLCustHasEquipment,cnnCustHasEquipment 

	If Not rsCustHasEquipment.EOF Then
		If Not ISNULL(rsCustHasEquipment("Expr1")) Then
			If rsCustHasEquipment("Expr1") < 1 Then resultCustHasEquipment = False Else resultCustHasEquipment = True
		Else
			resultCustHasEquipment = False
		End If
	Else
		resultCustHasEquipment = False
	End If

	rsCustHasEquipment.Close
	set rsCustHasEquipment = Nothing
	cnnCustHasEquipment.Close	
	set cnnCustHasEquipment = Nothing
	
	CustHasEquipment = resultCustHasEquipment
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetTotalValueOfEquipmentForCustomer(passedCustID)

	Set cnnTotalEqipValueForCust = Server.CreateObject("ADODB.Connection")
	cnnTotalEqipValueForCust.open Session("ClientCnnString")

	resultTotalEqipValueForCust = 0
		
	SQLTotalEqipValueForCust = "SELECT EQ_Classes.Class, EQ_Classes.InternalRecordIdentifier, SUM(EQ_Equipment.PurchaseCost) AS Expr1, Max(EQ_StatusCodes.StatusDesc) As Stat "
	SQLTotalEqipValueForCust = SQLTotalEqipValueForCust & " FROM EQ_CustomerEquipment INNER JOIN "
	SQLTotalEqipValueForCust = SQLTotalEqipValueForCust & " EQ_Equipment ON EQ_CustomerEquipment.EquipIntRecID = EQ_Equipment.InternalRecordIdentifier INNER JOIN "
	SQLTotalEqipValueForCust = SQLTotalEqipValueForCust & " EQ_Models ON EQ_Equipment.ModelIntRecID = EQ_Models.InternalRecordIdentifier INNER JOIN "
	SQLTotalEqipValueForCust = SQLTotalEqipValueForCust & " EQ_Classes ON EQ_Models.ClassIntRecID = EQ_Classes.InternalRecordIdentifier "
	
	SQLTotalEqipValueForCust = SQLTotalEqipValueForCust & " INNER JOIN EQ_StatusCodes ON EQ_StatusCodes.InternalRecordIdentifier = EQ_Equipment.StatusCodeIntRecID "
		
	SQLTotalEqipValueForCust = SQLTotalEqipValueForCust & " WHERE        (EQ_CustomerEquipment.CustID = '" & passedCustID & "') "
	SQLTotalEqipValueForCust = SQLTotalEqipValueForCust & " GROUP BY EQ_Classes.Class, EQ_Classes.InternalRecordIdentifier "
	SQLTotalEqipValueForCust = SQLTotalEqipValueForCust & " ORDER BY Expr1 DESC"

		 
	Set rsTotalEqipValueForCust = Server.CreateObject("ADODB.Recordset")
	rsTotalEqipValueForCust.CursorLocation = 3 
	
	rsTotalEqipValueForCust.Open SQLTotalEqipValueForCust,cnnTotalEqipValueForCust 
		
	resultTotalEqipValueForCust = 0	
			
	If NOT rsTotalEqipValueForCust.EOF Then
		Do While NOT rsTotalEqipValueForCust.EOF
			'If Ucase( rsTotalEqipValueForCust("Stat")) <> "PURCHASED" Then
				resultTotalEqipValueForCust = resultTotalEqipValueForCust + rsTotalEqipValueForCust("Expr1")
			'End If
			rsTotalEqipValueForCust.MoveNext
		Loop
	End If

	
	rsTotalEqipValueForCust.Close
	set rsTotalEqipValueForCust = Nothing
	cnnTotalEqipValueForCust.Close	
	set cnnTotalEqipValueForCust = Nothing
	
	GetTotalValueOfEquipmentForCustomer = resultTotalEqipValueForCust
	
End Function

'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function AllCustomerEquipmentIsPurchased(passedCustID)

	Set cnnAllCustomerEquipmentIsPurchased = Server.CreateObject("ADODB.Connection")
	cnnAllCustomerEquipmentIsPurchased.open Session("ClientCnnString")

	resultAllCustomerEquipmentIsPurchased = ""
		
	SQLAllCustomerEquipmentIsPurchased = "SELECT EQ_StatusCodes.StatusDesc "
	SQLAllCustomerEquipmentIsPurchased = SQLAllCustomerEquipmentIsPurchased & " FROM EQ_CustomerEquipment INNER JOIN "
	SQLAllCustomerEquipmentIsPurchased = SQLAllCustomerEquipmentIsPurchased & " EQ_Equipment ON EQ_CustomerEquipment.EquipIntRecID = EQ_Equipment.InternalRecordIdentifier "
	SQLAllCustomerEquipmentIsPurchased = SQLAllCustomerEquipmentIsPurchased & " INNER JOIN EQ_StatusCodes ON EQ_StatusCodes.InternalRecordIdentifier = EQ_Equipment.StatusCodeIntRecID "
	SQLAllCustomerEquipmentIsPurchased = SQLAllCustomerEquipmentIsPurchased & " WHERE (EQ_CustomerEquipment.CustID = '" & passedCustID & "') "
	
	Set rsAllCustomerEquipmentIsPurchased = Server.CreateObject("ADODB.Recordset")
	rsAllCustomerEquipmentIsPurchased.CursorLocation = 3 
	
	rsAllCustomerEquipmentIsPurchased.Open SQLAllCustomerEquipmentIsPurchased,cnnAllCustomerEquipmentIsPurchased 

	If NOT rsAllCustomerEquipmentIsPurchased.EOF Then

		resultAllCustomerEquipmentIsPurchased = True
		
		Do While NOT rsAllCustomerEquipmentIsPurchased.EOF
			If Ucase( rsAllCustomerEquipmentIsPurchased("StatusDesc")) <> "PURCHASED" Then
				resultAllCustomerEquipmentIsPurchased = False
				Exit Do ' If one is not purchased, we're done
			End If
			rsAllCustomerEquipmentIsPurchased.MoveNext
		Loop
	End If

	
	rsAllCustomerEquipmentIsPurchased.Close
	set rsAllCustomerEquipmentIsPurchased = Nothing
	cnnAllCustomerEquipmentIsPurchased.Close	
	set cnnAllCustomerEquipmentIsPurchased = Nothing
	
	AllCustomerEquipmentIsPurchased = resultAllCustomerEquipmentIsPurchased
	
End Function

'**************************************************************************************************************************************
'**************************************************************************************************************************************


Function NumberOfDocumentsByModelIntRecID(passedModelIntRecID)

	resultNumberOfDocumentsByModel = 0

	Set cnnNumberOfDocumentsByModel = Server.CreateObject("ADODB.Connection")
	cnnNumberOfDocumentsByModel.open Session("ClientCnnString")
		
	SQLNumberOfDocumentsByModel = "SELECT COUNT(*) AS DocumentCount FROM EQ_ModelDocuments WHERE ModelIntRecID = " & passedModelIntRecID
 
	Set rsNumberOfDocumentsByModel = Server.CreateObject("ADODB.Recordset")
	rsNumberOfDocumentsByModel.CursorLocation = 3 
	Set rsNumberOfDocumentsByModel = cnnNumberOfDocumentsByModel.Execute(SQLNumberOfDocumentsByModel)
			 
	If not rsNumberOfDocumentsByModel.EOF Then resultNumberOfDocumentsByModel = rsNumberOfDocumentsByModel("DocumentCount")
	
	rsNumberOfDocumentsByModel.Close
	set rsNumberOfDocumentsByModel= Nothing
	cnnNumberOfDocumentsByModel.Close	
	set cnnNumberOfDocumentsByModel= Nothing
	
	NumberOfDocumentsByModelIntRecID = resultNumberOfDocumentsByModel
	
End Function

'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************


Function NumberOfImagesByModelIntRecID(passedModelIntRecID)

	resultNumberOfImagesByModel = 0

	Set cnnNumberOfImagesByModel = Server.CreateObject("ADODB.Connection")
	cnnNumberOfImagesByModel.open Session("ClientCnnString")
		
	SQLNumberOfImagesByModel = "SELECT COUNT(*) AS ImageCount FROM EQ_ModelImages WHERE ModelIntRecID = " & passedModelIntRecID
 
	Set rsNumberOfImagesByModel = Server.CreateObject("ADODB.Recordset")
	rsNumberOfImagesByModel.CursorLocation = 3 
	Set rsNumberOfImagesByModel = cnnNumberOfImagesByModel.Execute(SQLNumberOfImagesByModel)
			 
	If not rsNumberOfImagesByModel.EOF Then resultNumberOfImagesByModel = rsNumberOfImagesByModel("ImageCount")
	
	rsNumberOfImagesByModel.Close
	set rsNumberOfImagesByModel= Nothing
	cnnNumberOfImagesByModel.Close	
	set cnnNumberOfImagesByModel= Nothing
	
	NumberOfImagesByModelIntRecID = resultNumberOfImagesByModel
	
End Function

'**************************************************************************************************************************************
'**************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************


Function NumberOfLinksByModelIntRecID(passedModelIntRecID)

	resultNumberOfLinksByModel = 0

	Set cnnNumberOfLinksByModel = Server.CreateObject("ADODB.Connection")
	cnnNumberOfLinksByModel.open Session("ClientCnnString")
		
	SQLNumberOfLinksByModel = "SELECT COUNT(*) AS LinkCount FROM EQ_ModelLinks WHERE ModelIntRecID = " & passedModelIntRecID
	
	
 
	Set rsNumberOfLinksByModel = Server.CreateObject("ADODB.Recordset")
	rsNumberOfLinksByModel.CursorLocation = 3 
	Set rsNumberOfLinksByModel = cnnNumberOfLinksByModel.Execute(SQLNumberOfLinksByModel)
			 
	If not rsNumberOfLinksByModel.EOF Then resultNumberOfLinksByModel = rsNumberOfLinksByModel("LinkCount")
		
	rsNumberOfLinksByModel.Close
	set rsNumberOfLinksByModel= Nothing
	cnnNumberOfLinksByModel.Close	
	set cnnNumberOfLinksByModel= Nothing
	
	NumberOfLinksByModelIntRecID = resultNumberOfLinksByModel
	
End Function

'**************************************************************************************************************************************
'**************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetEquipVendorNameByVendorID(passedVendorIntRecID)

	Set cnnGetEquipVendorNameByVendorID = Server.CreateObject("ADODB.Connection")
	cnnGetEquipVendorNameByVendorID.open Session("ClientCnnString")

	resultGetEquipVendorNameByVendorID = ""
		
	SQLGetEquipVendorNameByVendorID = "SELECT * FROM AP_Vendor WHERE InternalRecordIdentifier = " & passedVendorIntRecID
	 
	Set rsGetEquipVendorNameByVendorID  = Server.CreateObject("ADODB.Recordset")
	rsGetEquipVendorNameByVendorID.CursorLocation = 3 
	
	rsGetEquipVendorNameByVendorID.Open SQLGetEquipVendorNameByVendorID,cnnGetEquipVendorNameByVendorID 
			
	resultGetEquipVendorNameByVendorID = rsGetEquipVendorNameByVendorID("vendorCompanyName")
		
	rsGetEquipVendorNameByVendorID.Close
	set rsGetEquipVendorNameByVendorID = Nothing
	cnnGetEquipVendorNameByVendorID.Close	
	set cnnGetEquipVendorNameByVendorID = Nothing
	
	GetEquipVendorNameByVendorID = resultGetEquipVendorNameByVendorID 
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetCustomerIDByEquipIntRecID(passedEquipIntRecID)

	Set cnnGetCustomerIDByEquipIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetCustomerIDByEquipIntRecID.open Session("ClientCnnString")

	resultGetCustomerIDByEquipIntRecID = ""
		
	SQLGetCustomerIDByEquipIntRecID = "SELECT CustID FROM EQ_CustomerEquipment WHERE EquipIntRecID = " & passedEquipIntRecID
	 
	Set rsGetCustomerIDByEquipIntRecID  = Server.CreateObject("ADODB.Recordset")
	rsGetCustomerIDByEquipIntRecID.CursorLocation = 3 
	
	rsGetCustomerIDByEquipIntRecID.Open SQLGetCustomerIDByEquipIntRecID,cnnGetCustomerIDByEquipIntRecID 
			
	resultGetCustomerIDByEquipIntRecID = rsGetCustomerIDByEquipIntRecID("CustID")
	
	rsGetCustomerIDByEquipIntRecID.Close
	set rsGetCustomerIDByEquipIntRecID = Nothing
	cnnGetCustomerIDByEquipIntRecID.Close	
	set cnnGetCustomerIDByEquipIntRecID = Nothing
	
	If resultGetCustomerIDByEquipIntRecID = "" Then resultGetCustomerIDByEquipIntRecID = "NOT PLACED AT AN ACCOUNT"	
	GetCustomerIDByEquipIntRecID  = resultGetCustomerIDByEquipIntRecID 
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetAvailableForPlacementByEquipIntRecID(passedEquipIntRecID)

	Set cnnGetAvailableForPlacementByEquipIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetAvailableForPlacementByEquipIntRecID.open Session("ClientCnnString")

	resultGetAvailableForPlacementByEquipIntRecID = 0
		
	SQLGetAvailableForPlacementByEquipIntRecID = "SELECT * FROM EQ_StatusCodes INNER JOIN EQ_Equipment on EQ_StatusCodes.InternalRecordIdentifier = EQ_Equipment.StatusCodeIntRecID "
	SQLGetAvailableForPlacementByEquipIntRecID = SQLGetAvailableForPlacementByEquipIntRecID  & " WHERE EQ_Equipment.InternalRecordIdentifier = " & passedEquipIntRecID
	 
	Set rsGetAvailableForPlacementByEquipIntRecID  = Server.CreateObject("ADODB.Recordset")
	rsGetAvailableForPlacementByEquipIntRecID.CursorLocation = 3 
	
	rsGetAvailableForPlacementByEquipIntRecID.Open SQLGetAvailableForPlacementByEquipIntRecID,cnnGetAvailableForPlacementByEquipIntRecID 
			
	resultGetAvailableForPlacementByEquipIntRecID = rsGetAvailableForPlacementByEquipIntRecID("statusAvailableForPlacement")
	
	rsGetAvailableForPlacementByEquipIntRecID.Close
	set rsGetAvailableForPlacementByEquipIntRecID = Nothing
	cnnGetAvailableForPlacementByEquipIntRecID.Close	
	set cnnGetAvailableForPlacementByEquipIntRecID = Nothing
	
	GetAvailableForPlacementByEquipIntRecID  = resultGetAvailableForPlacementByEquipIntRecID 
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'**************************************************************************************************************************************
'**************************************************************************************************************************************
Function GetInsightAssetTagByEquipIntRecID(passedEquipIntRecID)

	Set cnnGetInsightAssetTagByEquipIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetInsightAssetTagByEquipIntRecID.open Session("ClientCnnString")

	resultGetInsightAssetTagByEquipIntRecID = ""
	Set rsGetInsightAssetTagByEquipIntRecID  = Server.CreateObject("ADODB.Recordset")
	rsGetInsightAssetTagByEquipIntRecID.CursorLocation = 3 
	
		
	SQLGetInsightAssetTagByEquipIntRecID = "SELECT ModelIntRecID FROM EQ_Equipment WHERE InternalRecordIdentifier = " & passedEquipIntRecID
	rsGetInsightAssetTagByEquipIntRecID.Open SQLGetInsightAssetTagByEquipIntRecID,cnnGetInsightAssetTagByEquipIntRecID 
	ModelIntRecID = rsGetInsightAssetTagByEquipIntRecID("ModelIntRecID")
	rsGetInsightAssetTagByEquipIntRecID.Close
	
	
	SQLGetInsightAssetTagByEquipIntRecID = "SELECT ClassIntRecID, BrandIntRecID, ManufacIntRecID, InsightAssetTagPrefix FROM EQ_Models WHERE InternalRecordIdentifier = " & ModelIntRecID
	rsGetInsightAssetTagByEquipIntRecID.Open SQLGetInsightAssetTagByEquipIntRecID,cnnGetInsightAssetTagByEquipIntRecID 
	BrandIntRecID = rsGetInsightAssetTagByEquipIntRecID("BrandIntRecID")
	ManufacIntRecID = rsGetInsightAssetTagByEquipIntRecID("ManufacIntRecID")
	ClassIntRecID = rsGetInsightAssetTagByEquipIntRecID("ClassIntRecID")
	ModelInsightAssetTagPrefix = rsGetInsightAssetTagByEquipIntRecID("InsightAssetTagPrefix")
	rsGetInsightAssetTagByEquipIntRecID.Close
	
	
	SQLGetInsightAssetTagByEquipIntRecID = "SELECT InsightAssetTagPrefix FROM EQ_Classes WHERE InternalRecordIdentifier = " & ClassIntRecID
	rsGetInsightAssetTagByEquipIntRecID.Open SQLGetInsightAssetTagByEquipIntRecID,cnnGetInsightAssetTagByEquipIntRecID 
	ClassInsightAssetTagPrefix = rsGetInsightAssetTagByEquipIntRecID("InsightAssetTagPrefix")
	rsGetInsightAssetTagByEquipIntRecID.Close
	

	SQLGetInsightAssetTagByEquipIntRecID = "SELECT InsightAssetTagPrefix FROM EQ_Manufacturers WHERE InternalRecordIdentifier = " & ManufacIntRecID
	rsGetInsightAssetTagByEquipIntRecID.Open SQLGetInsightAssetTagByEquipIntRecID,cnnGetInsightAssetTagByEquipIntRecID 
	ManfInsightAssetTagPrefix = rsGetInsightAssetTagByEquipIntRecID("InsightAssetTagPrefix")
	rsGetInsightAssetTagByEquipIntRecID.Close
	

	SQLGetInsightAssetTagByEquipIntRecID = "SELECT InsightAssetTagPrefix FROM EQ_Brands WHERE InternalRecordIdentifier = " & BrandIntRecID
	rsGetInsightAssetTagByEquipIntRecID.Open SQLGetInsightAssetTagByEquipIntRecID,cnnGetInsightAssetTagByEquipIntRecID 
	BrandInsightAssetTagPrefix = rsGetInsightAssetTagByEquipIntRecID("InsightAssetTagPrefix")
	rsGetInsightAssetTagByEquipIntRecID.Close
	
	If ClassInsightAssetTagPrefix = "" OR IsNull(ClassInsightAssetTagPrefix) Then
		ClassInsightAssetTagPrefix = "***"
	End If
	
	If ManfInsightAssetTagPrefix = "" OR IsNull(ManfInsightAssetTagPrefix) Then 
		ManfInsightAssetTagPrefix = "***"
	End If
	
	If BrandInsightAssetTagPrefix = "" OR IsNull(BrandInsightAssetTagPrefix) Then
		BrandInsightAssetTagPrefix = "***"
	End If
	
	If ModelInsightAssetTagPrefix = "" OR IsNull(ModelInsightAssetTagPrefix) Then
		ModelInsightAssetTagPrefix = "***"
	End If
	
	InsightAssetTagPrefix = ClassInsightAssetTagPrefix & "-" & ManfInsightAssetTagPrefix & "-" & BrandInsightAssetTagPrefix & "-" & ModelInsightAssetTagPrefix & "-" & passedEquipIntRecID

	

	set rsGetInsightAssetTagByEquipIntRecID = Nothing
	cnnGetInsightAssetTagByEquipIntRecID.Close	
	set cnnGetInsightAssetTagByEquipIntRecID = Nothing
	
	If InsightAssetTagPrefix = "" OR Left(InsightAssetTagPrefix,12) ="***-***-***-***" Then InsightAssetTagPrefix = "NO INSIGHT ASSET TAG ABLE TO BE GENERATED"	
	GetInsightAssetTagByEquipIntRecID  = InsightAssetTagPrefix
	
End Function
'**************************************************************************************************************************************
'**************************************************************************************************************************************


%>
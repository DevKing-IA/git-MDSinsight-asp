<%
'********************************
'List of all the functions & subs
'********************************
'Func NumberICProductsInventoriedOrPickable()
'Func NumberOfSKUsDefinedForPartner(passedPartnerIntRecID)
'Func GetPartnerNameByIntRecID(passedPartnerIntRecID)
'Func ProductInventoryReportTableSet(passedReportNumber)
'Func GetProdSKUByUPC(passedScannedCode)
'Func GetProdDescByUPC(passedScannedCode)
'Func GetProdDescByprodSKU(passedprodSKU)
'Func GetProdUMByUPC(passedScannedCode)
'Func GetProdImage(passedprodSKU)
'Func GetProdCaseBinByprodSKU(passedprodSKU)
'Func GetProdUnitBinByprodSKU(passedprodSKU)
'Func GetFilterDescByFilterID(passedFilterID)
'************************************
'End List of all the functions & subs
'************************************

Function NumberICProductsInventoriedOrPickable()

	Set cnnNumberICProductsInventoriedOrPickable  = Server.CreateObject("ADODB.Connection")
	cnnNumberICProductsInventoriedOrPickable.open Session("ClientCnnString")

	resultNumberICProductsInventoriedOrPickable = 0
		
	SQLNumberICProductsInventoriedOrPickable  = "SELECT COUNT(*) AS SKUCOUNT FROM IC_Product WHERE prodInventoriedItem = 1 OR prodPickableItem = 1"
	 
	Set rsNumberICProductsInventoriedOrPickable  = Server.CreateObject("ADODB.Recordset")
	rsNumberICProductsInventoriedOrPickable.CursorLocation = 3 
	
	rsNumberICProductsInventoriedOrPickable.Open SQLNumberICProductsInventoriedOrPickable,cnnNumberICProductsInventoriedOrPickable 
			
	resultNumberICProductsInventoriedOrPickable = rsNumberICProductsInventoriedOrPickable("SKUCOUNT")
	
	rsNumberICProductsInventoriedOrPickable.Close
	set rsNumberICProductsInventoriedOrPickable = Nothing
	cnnNumberICProductsInventoriedOrPickable.Close	
	set cnnNumberICProductsInventoriedOrPickable = Nothing
	
	NumberICProductsInventoriedOrPickable = resultNumberICProductsInventoriedOrPickable
	
End Function

Function NumberOfSKUsDefinedForPartner(passedPartnerIntRecID)

	Set cnnNumberOfSKUsDefinedForPartnerNum  = Server.CreateObject("ADODB.Connection")
	cnnNumberOfSKUsDefinedForPartnerNum.open Session("ClientCnnString")

	resultNumberOfSKUsDefinedForPartnerNum = 0
		
	SQLNumberOfSKUsDefinedForPartnerNum  = "SELECT COUNT(*) AS SKUCOUNT FROM IC_ProductMapping WHERE PartnerIntRecID = " & passedPartnerIntRecID
	 
	Set rsNumberOfSKUsDefinedForPartnerNum  = Server.CreateObject("ADODB.Recordset")
	rsNumberOfSKUsDefinedForPartnerNum.CursorLocation = 3 
	
	rsNumberOfSKUsDefinedForPartnerNum.Open SQLNumberOfSKUsDefinedForPartnerNum,cnnNumberOfSKUsDefinedForPartnerNum 
			
	resultNumberOfSKUsDefinedForPartnerNum = rsNumberOfSKUsDefinedForPartnerNum("SKUCOUNT")
	
	rsNumberOfSKUsDefinedForPartnerNum.Close
	set rsNumberOfSKUsDefinedForPartnerNum = Nothing
	cnnNumberOfSKUsDefinedForPartnerNum.Close	
	set cnnNumberOfSKUsDefinedForPartnerNum = Nothing
	
	NumberOfSKUsDefinedForPartner = resultNumberOfSKUsDefinedForPartnerNum
	
End Function


Function GetPartnerNameByIntRecID(passedPartnerIntRecID)

	Set cnnGetPartnerNameByIntRecID  = Server.CreateObject("ADODB.Connection")
	cnnGetPartnerNameByIntRecID.open Session("ClientCnnString")

	resultGetPartnerNameByIntRecID = ""
		
	SQLGetPartnerNameByIntRecID  = "SELECT * FROM IC_Partners WHERE InternalRecordIdentifier = " & passedPartnerIntRecID
	 
	Set rsGetPartnerNameByIntRecID  = Server.CreateObject("ADODB.Recordset")
	rsGetPartnerNameByIntRecID.CursorLocation = 3 
	
	rsGetPartnerNameByIntRecID.Open SQLGetPartnerNameByIntRecID,cnnGetPartnerNameByIntRecID 
			
	resultGetPartnerNameByIntRecID = rsGetPartnerNameByIntRecID("partnerCompanyName")
	
	rsGetPartnerNameByIntRecID.Close
	set rsGetPartnerNameByIntRecID = Nothing
	cnnGetPartnerNameByIntRecID.Close	
	set cnnGetPartnerNameByIntRecID = Nothing
	
	GetPartnerNameByIntRecID  = resultGetPartnerNameByIntRecID 
	
End Function

Function ProductInventoryReportTableSet(passedReportNumber)

	resultProductInventoryReportTableSet = False

	SQLProductInventoryReportTableSet = "SELECT * from Settings_Reports where ReportNumber = " & passedReportNumber & " AND UserNo = " & Session("userNo")

	Set cnnProductInventoryReportTableSet = Server.CreateObject("ADODB.Connection")
	cnnProductInventoryReportTableSet.open (Session("ClientCnnString"))
	Set rsProductInventoryReportTableSet = Server.CreateObject("ADODB.Recordset")
	Set rsProductInventoryReportTableSet = cnnProductInventoryReportTableSet.Execute(SQLProductInventoryReportTableSet)

	If NOT rsProductInventoryReportTableSet.EOF Then resultProductInventoryReportTableSet = True


	Set rsProductInventoryReportTableSet = Nothing
	cnnProductInventoryReportTableSet.Close
	Set cnnProductInventoryReportTableSet = Nothing

	ProductInventoryReportTableSet = resultProductInventoryReportTableSet 

End Function

Function GetProdSKUByUPC(passedScannedCode)

	resultGetProdSKUByUPC = ""

	Set cnnGetProdSKUByUPC = Server.CreateObject("ADODB.Connection")
	cnnGetProdSKUByUPC.open (Session("ClientCnnString"))
	Set rsGetProdSKUByUPC = Server.CreateObject("ADODB.Recordset")
	rsGetProdSKUByUPC.CursorLocation = 3 
		
	SQL_GetProdSKUByUPC = "SELECT * FROM IC_Product WHERE prodUnitUPC = '" & passedScannedCode & "' OR prodCaseUPC = '" & passedScannedCode & "'"
	Set rsGetProdSKUByUPC = cnnGetProdSKUByUPC.Execute(SQL_GetProdSKUByUPC)
	
	If Not rsGetProdSKUByUPC.EOF Then resultGetProdSKUByUPC = rsGetProdSKUByUPC("prodSKU")
	
	Set rsGetProdSKUByUPC = Nothing
	cnnGetProdSKUByUPC.Close
	Set cnnGetProdSKUByUPC = Nothing
	
	GetProdSKUByUPC = resultGetProdSKUByUPC 

End Function


Function GetProdDescByUPC(passedScannedCode)

	resultGetProdDescByUPC = ""

	Set cnnGetProdDescByUPC = Server.CreateObject("ADODB.Connection")
	cnnGetProdDescByUPC.open (Session("ClientCnnString"))
	Set rsGetProdDescByUPC = Server.CreateObject("ADODB.Recordset")
	rsGetProdDescByUPC.CursorLocation = 3 
		
	SQL_GetProdDescByUPC = "SELECT * FROM IC_Product WHERE prodUnitUPC = '" & passedScannedCode & "' OR prodCaseUPC = '" & passedScannedCode & "'"
	Set rsGetProdDescByUPC = cnnGetProdDescByUPC.Execute(SQL_GetProdDescByUPC)
	
	If Not rsGetProdDescByUPC.EOF Then resultGetProdDescByUPC = rsGetProdDescByUPC("prodDescription")
	
	Set rsGetProdDescByUPC = Nothing
	cnnGetProdDescByUPC.Close
	Set cnnGetProdDescByUPC = Nothing

	GetProdDescByUPC = resultGetProdDescByUPC 
	
End Function

Function GetProdUMByUPC(passedScannedCode)

	resultGetProdUMByUPC = ""

	Set cnnGetProdUMByUPC = Server.CreateObject("ADODB.Connection")
	cnnGetProdUMByUPC.open (Session("ClientCnnString"))
	Set rsGetProdUMByUPC = Server.CreateObject("ADODB.Recordset")
	rsGetProdUMByUPC.CursorLocation = 3 
		
	SQL_GetProdUMByUPC = "SELECT * FROM IC_Product WHERE prodUnitUPC = '" & passedScannedCode & "' OR prodCaseUPC = '" & passedScannedCode & "'"
	Set rsGetProdUMByUPC = cnnGetProdUMByUPC.Execute(SQL_GetProdUMByUPC)
	
	If Not rsGetProdUMByUPC.EOF Then resultGetProdUMByUPC = rsGetProdUMByUPC("prodCasePricing")
	
	Set rsGetProdUMByUPC = Nothing
	cnnGetProdUMByUPC.Close
	Set cnnGetProdUMByUPC = Nothing

	GetProdUMByUPC = resultGetProdUMByUPC 
	
End Function

Function GetProdImage(passedprodSKU)

	resultGetProdImage = ""

	Set cnnGetProdImage = Server.CreateObject("ADODB.Connection")
	cnnGetProdImage.open (Session("ClientCnnString"))
	Set rsGetProdImage = Server.CreateObject("ADODB.Recordset")
	rsGetProdImage.CursorLocation = 3 
		
	SQL_GetProdImage = "SELECT * FROM IC_ProductImages WHERE prodSKU = '" & passedprodSKU & "' AND imgType = 'I'"
	Set rsGetProdImage = cnnGetProdImage.Execute(SQL_GetProdImage)
	
	If Not rsGetProdImage.EOF Then
		If CheckRemoteURL("http://www.mdsinsight.com/clientfiles/" & Replace(UCASE(MUV_READ("SERNO")),"D","")  & "/prodImages/inventory/" & rsGetProdImage("imgFileName")) = True Then
			resultGetProdImage = "http://www.mdsinsight.com/clientfiles/" & Replace(UCASE(MUV_READ("SERNO")),"D","") & "/prodImages/inventory/" & rsGetProdImage("imgFileName")
		End If
	End If		
		
	Set rsGetProdImage = Nothing
	cnnGetProdImage.Close
	Set cnnGetProdImage = Nothing

	If resultGetProdImage = "" Then resultGetProdImage = baseURL & "/img/nopic.png"

	GetProdImage = resultGetProdImage 
	
End Function

Function CheckRemoteURL(fileURL)
    ON ERROR RESUME NEXT
    Dim xmlhttp

    Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")

    xmlhttp.open "GET", fileURL, False
    xmlhttp.send
    If(Err.Number<>0) then
        Response.Write "Could not connect to remote server"
    else
        Select Case Cint(xmlhttp.status)
            Case 200, 202, 302
                Set xmlhttp = Nothing
                CheckRemoteURL = True
            Case Else
                Set xmlhttp = Nothing
                CheckRemoteURL = False
        End Select
    end if
    ON ERROR GOTO 0
End Function

Function GetProdDescByprodSKU(passedprodSKU)

	resultGetProdDescByprodSKU = ""

	Set cnnGetProdDescByprodSKU = Server.CreateObject("ADODB.Connection")
	cnnGetProdDescByprodSKU.open (Session("ClientCnnString"))
	Set rsGetProdDescByprodSKU = Server.CreateObject("ADODB.Recordset")
	rsGetProdDescByprodSKU.CursorLocation = 3 
		
	SQL_GetProdDescByprodSKU = "SELECT prodDescription FROM IC_Product WHERE prodSKU = '" & passedprodSKU& "'"
	Set rsGetProdDescByprodSKU = cnnGetProdDescByprodSKU.Execute(SQL_GetProdDescByprodSKU)
	
	If Not rsGetProdDescByprodSKU.EOF Then resultGetProdDescByprodSKU = rsGetProdDescByprodSKU("prodDescription")
	
	Set rsGetProdDescByprodSKU = Nothing
	cnnGetProdDescByprodSKU.Close
	Set cnnGetProdDescByprodSKU = Nothing

	GetProdDescByprodSKU = resultGetProdDescByprodSKU 
	
End Function


Function GetProdUnitBinByprodSKU(passedprodSKU)

	resultGetProdUnitBinByprodSKU = ""

	Set cnnGetProdUnitBinByprodSKU = Server.CreateObject("ADODB.Connection")
	cnnGetProdUnitBinByprodSKU.open (Session("ClientCnnString"))
	Set rsGetProdUnitBinByprodSKU = Server.CreateObject("ADODB.Recordset")
	rsGetProdUnitBinByprodSKU.CursorLocation = 3 
		
	SQL_GetProdUnitBinByprodSKU = "SELECT prodUnitBin FROM IC_Product WHERE prodSKU = '" & passedprodSKU & "'" 
	Set rsGetProdUnitBinByprodSKU = cnnGetProdUnitBinByprodSKU.Execute(SQL_GetProdUnitBinByprodSKU)
	
	If Not rsGetProdUnitBinByprodSKU.EOF Then resultGetProdUnitBinByprodSKU = rsGetProdUnitBinByprodSKU("prodUnitBin")
	
	Set rsGetProdUnitBinByprodSKU = Nothing
	cnnGetProdUnitBinByprodSKU.Close
	Set cnnGetProdUnitBinByprodSKU = Nothing
	
	GetProdUnitBinByprodSKU = resultGetProdUnitBinByprodSKU 

End Function

Function GetProdCaseBinByprodSKU(passedprodSKU)

	resultGetProdCaseBinByprodSKU = ""

	Set cnnGetProdCaseBinByprodSKU = Server.CreateObject("ADODB.Connection")
	cnnGetProdCaseBinByprodSKU.open (Session("ClientCnnString"))
	Set rsGetProdCaseBinByprodSKU = Server.CreateObject("ADODB.Recordset")
	rsGetProdCaseBinByprodSKU.CursorLocation = 3 
		
	SQL_GetProdCaseBinByprodSKU = "SELECT prodCaseBin FROM IC_Product WHERE prodSKU = '" & passedprodSKU & "'" 
	Set rsGetProdCaseBinByprodSKU = cnnGetProdCaseBinByprodSKU.Execute(SQL_GetProdCaseBinByprodSKU)
	
	If Not rsGetProdCaseBinByprodSKU.EOF Then resultGetProdCaseBinByprodSKU = rsGetProdCaseBinByprodSKU("prodCaseBin")
	
	Set rsGetProdCaseBinByprodSKU = Nothing
	cnnGetProdCaseBinByprodSKU.Close
	Set cnnGetProdCaseBinByprodSKU = Nothing
	
	GetProdCaseBinByprodSKU = resultGetProdCaseBinByprodSKU 

End Function

Function GetFilterDescByFilterID(passedFilterID)

	resultGetFilterDescByFilterID = ""

	Set cnnGetFilterDescByFilterID = Server.CreateObject("ADODB.Connection")
	cnnGetFilterDescByFilterID.open (Session("ClientCnnString"))
	Set rsGetFilterDescByFilterID = Server.CreateObject("ADODB.Recordset")
	rsGetFilterDescByFilterID.CursorLocation = 3 
		
	SQL_GetFilterDescByFilterID = "SELECT Description FROM IC_Filters WHERE FilterID = '" & passedFilterID& "'"
	Set rsGetFilterDescByFilterID = cnnGetFilterDescByFilterID.Execute(SQL_GetFilterDescByFilterID)
	
	If Not rsGetFilterDescByFilterID.EOF Then resultGetFilterDescByFilterID = rsGetFilterDescByFilterID("Description")
	
	Set rsGetFilterDescByFilterID = Nothing
	cnnGetFilterDescByFilterID.Close
	Set cnnGetFilterDescByFilterID = Nothing

	GetFilterDescByFilterID = resultGetFilterDescByFilterID 
	
End Function
%>
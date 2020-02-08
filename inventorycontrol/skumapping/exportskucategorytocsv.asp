
<!--#include file="../../inc/InsightFuncs_InventoryControl.asp"-->
<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/Insightfuncs.asp"-->

<%
						
	PartnerInternalRecordIdentifier = Request.QueryString("i")
	CategoryID = Request.QueryString("c")
	
    

    

	If CategoryID = "all" Then
		CategoryName = "ALL CATEGORIES"
	Else
		CategoryName = GetCategoryByID(CategoryID)
	End If
    DIM fs,tfile
    set fs=Server.CreateObject("Scripting.FileSystemObject")
    set tfile=fs.CreateTextFile(Server.MapPath("/clientfiles/" & trim(MUV_Read("SERNO"))& "/csv/prodMap_"+REPLACE(CategoryName," ","_")+".csv"),2)
	Set cnnProductsTable = Server.CreateObject("ADODB.Connection")
	cnnProductsTable.open (Session("ClientCnnString"))
	Set rsProductsTable = Server.CreateObject("ADODB.Recordset")
	rsProductsTable.CursorLocation = 3 
	
	Set cnnEquivalentSKUs = Server.CreateObject("ADODB.Connection")
	cnnEquivalentSKUs.open (Session("ClientCnnString"))
	Set rsEquivalentSKUs = Server.CreateObject("ADODB.Recordset")
	rsEquivalentSKUs.CursorLocation = 3 
	
    DIM stringToCSV
    stringToCSV=""

	If CategoryID = "all" Then				
		SQLProductsTable = "SELECT * FROM Product ORDER BY PartNo ASC"
	Else
		SQLProductsTable = "SELECT * FROM Product WHERE Category = " & CategoryID & " ORDER BY PartNo ASC"
	End If
	
	Set rsProductsTable = cnnProductsTable.Execute(SQLProductsTable)
	stringToCSV="partnerIntRecID,SKU,UM,CategoryID,partnerEquivalentSKU1,partnerEquivalentSKU2,partnerEquivalentSKU3,partnerEquivalentSKU4,partnerEquivalentSKU5,partnerEquivalentSKU6"
    tfile.WriteLine(stringToCSV)
   
	If NOT rsProductsTable.EOF Then
								
   
		Do While Not rsProductsTable.EOF
	
			CategoryIDToPass = rsProductsTable("Category")
			SKUFromProductsTable = rsProductsTable("PartNo")
			UMFromProductsTable = rsProductsTable("CasePricing")
			DESCFromProductsTable = rsProductsTable("Description") 
			
			SQLEquivalentSKUs = "SELECT * FROM IC_ProductMapping WHERE SKU = '" & SKUFromProductsTable & "' AND partnerIntRecID = " & PartnerInternalRecordIdentifier
			
			'Response.Write(SQLEquivalentSKUs & "<br>")
			
			Set rsEquivalentSKUs = cnnEquivalentSKUs.Execute(SQLEquivalentSKUs)
			
			If NOT rsEquivalentSKUs.EOF Then
			
			    
				Do While Not rsEquivalentSKUs.EOF
		            stringToCSV=""
                    stringToCSV=""""&(rsEquivalentSKUs("partnerIntRecID"))&"""" 
                    stringToCSV=stringToCSV+","
                    stringToCSV=stringToCSV+""""&(rsEquivalentSKUs("SKU"))&""""
                    stringToCSV=stringToCSV+","
                    stringToCSV=stringToCSV+""""&(rsEquivalentSKUs("UM"))&"""" 
                    stringToCSV=stringToCSV+","
                    stringToCSV=stringToCSV+""""&(rsEquivalentSKUs("CategoryID"))&"""" 
                    stringToCSV=stringToCSV+","
                    stringToCSV=stringToCSV+""""&(rsEquivalentSKUs("partnerEquivalentSKU1"))&"""" 
                    stringToCSV=stringToCSV+","
                    stringToCSV=stringToCSV+""""&(rsEquivalentSKUs("partnerEquivalentSKU2"))&"""" 
                    stringToCSV=stringToCSV+","
                    stringToCSV=stringToCSV+""""&(rsEquivalentSKUs("partnerEquivalentSKU3"))&"""" 
                    stringToCSV=stringToCSV+","
                    stringToCSV=stringToCSV+""""&(rsEquivalentSKUs("partnerEquivalentSKU4"))&"""" 
                    stringToCSV=stringToCSV+","
                    stringToCSV=stringToCSV+""""&(rsEquivalentSKUs("partnerEquivalentSKU5"))&"""" 
                    stringToCSV=stringToCSV+","
                    stringToCSV=stringToCSV+""""&(rsEquivalentSKUs("partnerEquivalentSKU6"))&"""" 
                    tfile.WriteLine(stringToCSV)
					
					rsEquivalentSKUs.movenext
				Loop
				
			End If
            rsProductsTable.MoveNext
        LOOP
    END IF
    tfile.close
    SET fs=nothing
    cnnProductsTable.Close
    cnnEquivalentSKUs.Close
   
    'Response.ContentType = "text/csv"
    'Response.AddHeader "Content-Disposition", "filename=myfile.csv;"
    'Response.Write(output)
    response.Redirect("/clientfiles/" & trim(MUV_Read("SERNO"))& "/csv/prodMap_"+REPLACE(CategoryName," ","_")+".csv")
    Response.End
	        %>
		                          
						
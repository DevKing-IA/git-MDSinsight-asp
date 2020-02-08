<!--#include file="../../inc/InsightFuncs_AR_AP.asp"-->
<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/Insightfuncs.asp"-->

<%
	IF Session("Userno") = "" Then
        Response.Write("Impossible to delete file. Please login before")
        ELSE
            InternalRecordIdentifier = Request.QueryString("i")
            PartnerInternalRecordIdentifier = Request.QueryString("i")
            FirstLetter = Request.QueryString("letter")

            If  FirstLetter = "all" Then
		        CategoryName = "ALL List"
	            Else
		            CategoryName = "Beginning with "+FirstLetter
	        End If
            DIM fs,tfile
            set fs=Server.CreateObject("Scripting.FileSystemObject")
            set tfile=fs.CreateTextFile(Server.MapPath("/clientfiles/" & trim(MUV_Read("SERNO"))& "/csv/customerMap_"+REPLACE(CategoryName," ","_")+".csv"),2)
	        stringToCSV="partnerRecID,partnerCustID,ourCustID,DelivInstructions,ReferalCode,ArOldAcctNum"
            tfile.WriteLine(stringToCSV)

            
            
            Set cnnCustomerTable = Server.CreateObject("ADODB.Connection")
	        cnnCustomerTable.open (Session("ClientCnnString"))
	        Set rsCustomerTable = Server.CreateObject("ADODB.Recordset")
	        rsCustomerTable.CursorLocation = 3 
	
	        Set cnnEquivalentCustomers = Server.CreateObject("ADODB.Connection")
	        cnnEquivalentCustomers.open (Session("ClientCnnString"))
	        Set rsEquivalentCustomers = Server.CreateObject("ADODB.Recordset")
	        rsEquivalentCustomers.CursorLocation = 3 
	        If FirstLetter = "all" Then				
		        SQLCustomersTable = "SELECT * FROM AR_Customer WHERE AcctStatus = 'A' ORDER BY CONVERT(int, CustNum) ASC"
	            Else
		            SQLCustomersTable = "SELECT * FROM AR_Customer WHERE LEFT(Name,1) = '" & FirstLetter & "' AND AcctStatus = 'A' ORDER BY CONVERT(int, CustNum) ASC"
	        End If
            rsCustomerTable.Open SQLCustomersTable,cnnCustomerTable,3,3
	        If NOT rsCustomerTable.EOF Then
    			Do While Not rsCustomerTable.EOF
								
				    customerID = rsCustomerTable("CustNum")
					customerName = rsCustomerTable("Name") 
					customerAddr1 = rsCustomerTable("Addr1") 
					customerAddr2 = rsCustomerTable("Addr2") 
					customerCityStateZip = rsCustomerTable("CityStateZip") 
					customerPhone = rsCustomerTable("Phone")
								
					SQLEquivalentCustomers = "SELECT * FROM AR_CustomerMapping WHERE "
					SQLEquivalentCustomers = SQLEquivalentCustomers & "partnerRecID = " & PartnerInternalRecordIdentifier & " AND "
					SQLEquivalentCustomers = SQLEquivalentCustomers & "ourCustID = '" & customerID & "'"
					
								
					Set rsEquivalentCustomers = cnnEquivalentCustomers.Execute(SQLEquivalentCustomers)
					If NOT rsEquivalentCustomers.EOF Then
					    
                        stringToCSV=""
                        stringToCSV=""""&(rsEquivalentCustomers("partnerRecID"))&"""" 
                        stringToCSV=stringToCSV+","
                        stringToCSV=stringToCSV+""""&(rsEquivalentCustomers("partnerCustID"))&"""" 
                        stringToCSV=stringToCSV+","
                        
                        stringToCSV=stringToCSV+""""&(rsEquivalentCustomers("ourCustID"))&"""" 
                        stringToCSV=stringToCSV+","
                        stringToCSV=stringToCSV+""""&(rsEquivalentCustomers("DelivInstructions"))&"""" 
                        stringToCSV=stringToCSV+","
                        stringToCSV=stringToCSV+""""&(rsEquivalentCustomers("ReferalCode"))&"""" 
                        stringToCSV=stringToCSV+","
                        stringToCSV=stringToCSV+""""&(rsEquivalentCustomers("ArOldAcctNum"))&"""" 
                   
                        tfile.WriteLine(stringToCSV)
						
					End If
                    
					rsCustomerTable.movenext
	            Loop
							
        	End If
			
			set rsCustomerTable = Nothing
			cnnCustomerTable.close
			set cnnCustomerTable = Nothing
			
			set rsEquivalentCustomers = Nothing
			cnnEquivalentCustomers.close
			set cnnEquivalentCustomers = Nothing
	        tfile.close
            SET fs=nothing
   
   
            'Response.ContentType = "text/csv"
            'Response.AddHeader "Content-Disposition", "filename=myfile.csv;"
            'Response.Write(output)
            response.Redirect("/clientfiles/" & trim(MUV_Read("SERNO"))& "/csv/customerMap_"+REPLACE(CategoryName," ","_")+".csv")
            Response.End

    END IF
	
   
	        %>
		                          
						
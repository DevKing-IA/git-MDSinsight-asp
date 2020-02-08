<!--#include file="../inc/header-prospecting.asp"-->
<!-- css !-->
<style type="text/css">

.nav-tabs{
	font-size: 12px;
}

.the-tabs .nav>li>a{
	padding: 5px 10px;
	font-weight: bold;
}

.tab-content{
	margin-top:20px;
	font-size:12px;
}

.tab-content .table>tbody>tr>td, .table>tbody>tr>th, .table>tfoot>tr>td, .table>tfoot>tr>th, .table>thead>tr>td, .table>thead>tr>th{
	border:0px;
   }
   
   
.tab-content .split-arrows{
	 text-align:center;
	 margin-top:10px;
	 margin-bottom: 10px;
 }
 
 .tab-content .split-arrows a{
	 display:inline-block;
	 background:#f5f5f5;
	 padding:5px;
 }
 
  .tab-content .split-arrows a:hover{
	  background:#ccc;
	  text-decoration:none;
  }

 

.tab-content .red-line{
	border-left:3px solid red;
}   

.row-line{
	margin-bottom:15px;
}

.th-width{
	width: 30%;
	font-weight: normal;
}

.th-width2{
	width: 70%;
	font-weight: normal;
}

.first-name{
	width: 54%;
	display: inline-block;
}

.first-name-mr{
	width: 30%;
	display: inline-block;
	margin-right: 2px;
}

.table-responsive tr {
    font-size:16px;
    font-weight:bold;
}

.btn-xlarge {
    padding: 18px 28px;
    font-size: 22px;
    line-height: normal;
	-webkit-border-radius: 8px;
	-moz-border-radius: 8px;
	border-radius: 8px;
}
</style>
<!-- eof css !-->

 	
<h1 class="page-header"><i class="fa fa-fw fa-asterisk"></i> <%= GetTerm("Check For Duplicates") %>
	<!-- customize !-->
	<div class="col pull-right">
	</div>
	<!-- eof customize !-->
</h1>

<SCRIPT LANGUAGE="JavaScript">
<!--

	$(window).load(function()
	{
	   var phones = [{ "mask": "(###) ###-####"}];
	    $('#txtPhoneNumber').inputmask({ 
	        mask: phones, 
	        placeholder: "(___) ___-____",
	        greedy: false, 
	        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
	});

	function isValid(p) {
	  //var phoneRe = /^[2-9]\d{2}[2-9]\d{2}\d{4}$/;
	  //var phoneRe = /^[\+]?[(]?[0-9]{3}[)]?[-\s\.]?[0-9]{3}[-\s\.]?[0-9]{4,6}$/;
	  var phoneRe = /^(1\s|1|)?((\(\d{3}\))|\d{3})(\-|\s)?(\d{3})(\-|\s)?(\d{4})$/;
	  var digits = p.replace(/\D/g, "");
	  return phoneRe.test(digits);
	}

   function validatePrequalifyProspectForm()
    {
        //if (document.frmAddProspectPrequalify.txtFirstName.value == "") {
            //swal("First name cannot be blank.");
            //return false;
        //}
        //if (document.frmAddProspectPrequalify.txtLastName.value == "") {
            //swal("Last name cannot be blank.");
            //return false;
        //}

        if (document.frmAddProspectPrequalify.txtCompanyName.value == "") {
            swal("Company name cannot be blank.");
            return false;
        }

        //if (document.frmAddProspectPrequalify.txtPhoneNumber.value == "") {
            //swal("Phone number cannot be blank.");
            //return false;
        //}

        //if (document.frmAddProspectPrequalify.txtAddress.value == "") {
            //swal("Address cannot be blank.");
            //return false;
        //}

        //if (isValid(document.frmAddProspectPrequalify.txtPhoneNumber.value) == false) {
           //swal("The phone number is invalid. Please enter any format like the following: 555-555-5555, (555)555-5555, (555) 555-5555, 555 555 5555, 5555555555, 1 555 555 5555.");
           //return false;
        //}

        return true;

    }
// -->
</SCRIPT>   

<%
CreatedByUserNo = Session("UserNo")
If CreatedByUserNo <> "" Then 
	CreatedByUserName = GetUserDisplayNameByUserNo(CreatedByUserNo)
Else
	CreatedByUserName = ""
End If

Function stripNonNumeric(inputString)
	Set regEx = New RegExp
	regEx.Global = True
	regEx.Pattern = "\D"
	stripNonNumeric = regEx.Replace(inputString,"")
End Function
  
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then 

	notACurrentProspect = true
	
	txtFirstName = Request.form("txtFirstName")
	txtLastName = Request.form("txtLastName")
	txtCompanyName = Request.form("txtCompanyName")
	txtPhoneNumber = Request.form("txtPhoneNumber")
	txtAddress = Request.form("txtAddress")
	
	txtLastName = Replace(txtLastName,"'","''")
	txtCompanyName = Replace(txtCompanyName,"'","''")
	
	SQL8 = "SELECT * FROM Settings_Global"
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs8 = Server.CreateObject("ADODB.Recordset")
	rs8.CursorLocation = 3 

	on error resume next
	SQL8 = "DROP TABLE zProspectPrequal_" & Trim(Session("userNo"))
	Set rs8 = cnn8.Execute(SQL8)
	on error goto 0
	
	on error resume next
	SQL8 = "DROP TABLE zProspectPrequalWeighted_" & Trim(Session("userNo"))
	Set rs8 = cnn8.Execute(SQL8)
	on error goto 0

	SQL8 = "CREATE TABLE zProspectPrequalWeighted_" & Trim(Session("userNo"))
	SQL8 = SQL8 & "("
	SQL8 = SQL8 & "                [Source] [varchar](255) NULL, "
	SQL8 = SQL8 & "                [Pool] [varchar](255) NULL, "
	SQL8 = SQL8 & "                [ProspectRecordID] [int] NULL, "
	SQL8 = SQL8 & "                [Custid] [varchar](255) NULL, "
	SQL8 = SQL8 & "                [CompanyName] [varchar](255) NULL, "
	SQL8 = SQL8 & "                [Address] [varchar](255) NULL, "		
	SQL8 = SQL8 & "                [PhoneNumber] [varchar](255) NULL, "	
	SQL8 = SQL8 & "                [ContactFirstName] [varchar](255) NULL, "	
	SQL8 = SQL8 & "                [ContactLastName] [varchar](255) NULL, "
	SQL8 = SQL8 & "                [ContactFullName] [varchar](255) NULL, "
	SQL8 = SQL8 & "                [Match_Company] [float] NULL, "
	SQL8 = SQL8 & "                [Match_Name] [float] NULL, "
	SQL8 = SQL8 & "                [Match_Address] [float] NULL, "
	SQL8 = SQL8 & "                [Match_Phone] [float] NULL)"

	Set rs8 = cnn8.Execute(SQL8)

	'****************************
	'Customer file - company name
	'****************************
	If txtCompanyName <> "" Then
		SQL8  = "SELECT 'Customer' as Source, '' AS Pool, 0 AS ProspectRecordID, Custnum AS CustID, Name AS CompanyName, Addr2 AS Address, Phone AS PhoneNumber, ContactFirstName AS ContactFirstName, ContactLastName AS ContactLastName, Contact AS ContactFullName "
		SQL8 = SQL8 & "FROM AR_Customer WHERE Name Like '%" & txtCompanyName & "%'"
		Set rs8 = cnn8.Execute(SQL8)
		If not rs8.Eof Then
			Do While Not rs8.Eof

					If rs8("CompanyName") <> "" Then CompanyName = Replace(rs8("CompanyName"),"'","''") Else CompanyName = ""
					If rs8("Address") <> "" Then Address = Replace(rs8("Address"),"'","''") Else Address = ""
					If rs8("ContactFirstName") <> "" Then ContactFirstName = Replace(rs8("ContactFirstName"),"'","''") Else ContactFirstName = ""
					If rs8("ContactLastName") <> "" Then ContactLastName = Replace(rs8("ContactLastName"),"'","''") Else ContactLastName = ""
					If rs8("ContactFullName") <> "" Then ContactFullName = Replace(rs8("ContactFullName"),"'","''") Else ContactFullName = ""
			
					SQLWeighted = "SELECT * FROM zProspectPrequalWeighted_" & Trim(Session("userNo")) & " WHERE CustID='" & rs8("CustID") & "'"
					Set rsWeighted = cnn8.Execute(SQLWeighted) 
					
					If rsWeighted.EOF Then 
						SQLWeighted = "INSERT INTO zProspectPrequalWeighted_" & Trim(Session("userNo")) & " " 
						SQLWeighted = SQLWeighted & "(Source, Pool, ProspectRecordID, CustID, CompanyName, Address, PhoneNumber, "
						SQLWeighted = SQLWeighted & " ContactFirstName, ContactLastName, ContactFullName, "
						SQLWeighted = SQLWeighted & " Match_Company, Match_Name, Match_Address, Match_Phone) "
						SQLWeighted = SQLWeighted & "VALUES ( "
						SQLWeighted = SQLWeighted & "'" & rs8("Source") & "','"  & rs8("Pool") & "'," & rs8("ProspectRecordID") & ",'" & rs8("CustID") & "','"  & CompanyName & "','"  & Address & "','" & rs8("phoneNumber")& "', "
						SQLWeighted = SQLWeighted & "'" & ContactFirstName & "','"  & ContactLastName & "','"  & ContactFullName & "', "
						If Ucase(CompanyName) = Ucase(txtCompanyName) Then
							SQLWeighted = SQLWeighted & "2,0,0,0)" ' Exact match gets an extra point
						Else
							SQLWeighted = SQLWeighted & "1,0,0,0)"
						End IF
					Else
						If Ucase(CompanyName) = Ucase(txtCompanyName) Then
							SQLWeighted = "UPDATE zProspectPrequalWeighted_" & Trim(Session("userNo")) & " SET Match_Company = Match_Company + 2 WHERE CustID='" & rs8("CustID") & "'" ' Exact match gets an extra point
						Else
							SQLWeighted = "UPDATE zProspectPrequalWeighted_" & Trim(Session("userNo")) & " SET Match_Company = Match_Company + 1 WHERE CustID='" & rs8("CustID") & "'"
						End IF
					End If
						
					Set rsWeighted = cnn8.Execute(SQLWeighted)  
				rs8.MoveNext
			Loop
		End IF
	End If

	'****************************
	'Customer file - address
	'****************************
	If txtAddress <> "" Then
		SQL8  = "SELECT 'Customer' as Source, '' AS Pool, 0 AS ProspectRecordID, Custnum AS CustID, Name AS CompanyName, Addr2 AS Address, Phone AS PhoneNumber, ContactFirstName AS ContactFirstName, ContactLastName AS ContactLastName, Contact AS ContactFullName "
		SQL8 = SQL8 & "FROM AR_Customer WHERE Addr2 Like '%" & txtAddress & "%'"
		Set rs8 = cnn8.Execute(SQL8)
		
		If not rs8.Eof Then
			Do While Not rs8.Eof
			
					If rs8("CompanyName") <> "" Then CompanyName = Replace(rs8("CompanyName"),"'","''") Else CompanyName = ""
					If rs8("Address") <> "" Then Address = Replace(rs8("Address"),"'","''") Else Address = ""
					If rs8("ContactFirstName") <> "" Then ContactFirstName = Replace(rs8("ContactFirstName"),"'","''") Else ContactFirstName = ""
					If rs8("ContactLastName") <> "" Then ContactLastName = Replace(rs8("ContactLastName"),"'","''") Else ContactLastName = ""
					If rs8("ContactFullName") <> "" Then ContactFullName = Replace(rs8("ContactFullName"),"'","''") Else ContactFullName = ""
			
					SQLWeighted = "SELECT * FROM zProspectPrequalWeighted_" & Trim(Session("userNo")) & " WHERE CustID='" & rs8("CustID") & "'"
					Set rsWeighted = cnn8.Execute(SQLWeighted) 
					
					If rsWeighted.EOF Then 
						SQLWeighted = "INSERT INTO zProspectPrequalWeighted_" & Trim(Session("userNo")) & " " 
						SQLWeighted = SQLWeighted & "(Source, Pool, ProspectRecordID, CustID, CompanyName, Address, PhoneNumber, "
						SQLWeighted = SQLWeighted & " ContactFirstName, ContactLastName, ContactFullName, "
						SQLWeighted = SQLWeighted & " Match_Company, Match_Name, Match_Address, Match_Phone) "
						SQLWeighted = SQLWeighted & "VALUES ( "
						SQLWeighted = SQLWeighted & "'" & rs8("Source") & "','"  & rs8("Pool") & "'," & rs8("ProspectRecordID") & ",'" & rs8("CustID") & "','"  & CompanyName & "','"  & Address & "','" & rs8("phoneNumber")& "', "
						SQLWeighted = SQLWeighted & "'" & ContactFirstName & "','"  & ContactLastName & "','"  & ContactFullName & "', "
						If Ucase(Address) = Ucase(txtAddress) Then
							SQLWeighted = SQLWeighted & "0,0,2,0)" ' Exact match gets an extra point
						Else
							SQLWeighted = SQLWeighted & "0,0,1,0)"
						End IF
					Else
						If Ucase(Address) = Ucase(txtAddress) Then
							SQLWeighted = "UPDATE zProspectPrequalWeighted_" & Trim(Session("userNo")) & " SET Match_Address = Match_Address + 2 WHERE CustID='" & rs8("CustID") & "'" ' Exact match gets an extra point
						Else
							SQLWeighted = "UPDATE zProspectPrequalWeighted_" & Trim(Session("userNo")) & " SET Match_Address = Match_Address + 1 WHERE CustID='" & rs8("CustID") & "'"						
						End IF
					End If
					Set rsWeighted = cnn8.Execute(SQLWeighted)  
				rs8.MoveNext
			Loop
		End IF
	End If

	'****************************
	'Customer file - phone
	'****************************
	If txtPhoneNumber <> "" Then
		SQL8  = "SELECT 'Customer' as Source, '' AS Pool, 0 AS ProspectRecordID, Custnum AS CustID, Name AS CompanyName, Addr2 AS Address, Phone AS PhoneNumber, ContactFirstName AS ContactFirstName, ContactLastName AS ContactLastName, Contact AS ContactFullName "
		SQL8 = SQL8 & "FROM AR_Customer WHERE Phone Like '%" & txtPhoneNumber & "%'"
		Set rs8 = cnn8.Execute(SQL8)
		
		If not rs8.Eof Then
			Do While Not rs8.Eof
			
					If rs8("CompanyName") <> "" Then CompanyName = Replace(rs8("CompanyName"),"'","''") Else CompanyName = ""
					If rs8("Address") <> "" Then Address = Replace(rs8("Address"),"'","''") Else Address = ""
					If rs8("ContactFirstName") <> "" Then ContactFirstName = Replace(rs8("ContactFirstName"),"'","''") Else ContactFirstName = ""
					If rs8("ContactLastName") <> "" Then ContactLastName = Replace(rs8("ContactLastName"),"'","''") Else ContactLastName = ""
					If rs8("ContactFullName") <> "" Then ContactFullName = Replace(rs8("ContactFullName"),"'","''") Else ContactFullName = ""
			
					SQLWeighted = "SELECT * FROM zProspectPrequalWeighted_" & Trim(Session("userNo")) & " WHERE CustID='" & rs8("CustID") & "'"
					Set rsWeighted = cnn8.Execute(SQLWeighted) 
					
					If rsWeighted.EOF Then 
						SQLWeighted = "INSERT INTO zProspectPrequalWeighted_" & Trim(Session("userNo")) & " " 
						SQLWeighted = SQLWeighted & "(Source, Pool, ProspectRecordID, CustID, CompanyName, Address, PhoneNumber, "
						SQLWeighted = SQLWeighted & " ContactFirstName, ContactLastName, ContactFullName, "
						SQLWeighted = SQLWeighted & " Match_Company, Match_Name, Match_Address, Match_Phone) "
						SQLWeighted = SQLWeighted & "VALUES ( "
						SQLWeighted = SQLWeighted & "'" & rs8("Source") & "','"  & rs8("Pool") & "'," & rs8("ProspectRecordID") & ",'" & rs8("CustID") & "','"  & CompanyName & "','"  & Address & "','" & rs8("phoneNumber")& "', "
						SQLWeighted = SQLWeighted & "'" & ContactFirstName & "','"  & ContactLastName & "','"  & ContactFullName & "', "
						If Ucase(rs8("phoneNumber")) = Ucase(txtPhoneNumber) Then
							SQLWeighted = SQLWeighted & "0,0,0,2)" ' Exact match gets an extra point
						Else
							SQLWeighted = SQLWeighted & "0,0,0,1)"
						End IF
					Else
						If Ucase(rs8("phoneNumber")) = Ucase(txtPhoneNumber) Then
							SQLWeighted = "UPDATE zProspectPrequalWeighted_" & Trim(Session("userNo")) & " SET Match_Phone = Match_Phone + 2 WHERE CustID='" & rs8("CustID") & "'" ' Exact match gets an extra point
						Else
							SQLWeighted = "UPDATE zProspectPrequalWeighted_" & Trim(Session("userNo")) & " SET Match_Phone = Match_Phone + 1 WHERE CustID='" & rs8("CustID") & "'"
						End If
					End If
					Set rsWeighted = cnn8.Execute(SQLWeighted)  
				rs8.MoveNext
			Loop
		End IF
	End If

	'****************************
	'Customer file - last name
	'****************************
	If txtLastName <> "" Then
		SQL8  = "SELECT 'Customer' as Source, '' AS Pool, 0 AS ProspectRecordID, Custnum AS CustID, Name AS CompanyName, Addr2 AS Address, Phone AS PhoneNumber, ContactFirstName AS ContactFirstName, ContactLastName AS ContactLastName, Contact AS ContactFullName "
		SQL8 = SQL8 & "FROM AR_Customer WHERE ContactLastName Like '%" & txtLastName & "%'"
		Set rs8 = cnn8.Execute(SQL8)
		
		If not rs8.Eof Then
			Do While Not rs8.Eof
			
					If rs8("CompanyName") <> "" Then CompanyName = Replace(rs8("CompanyName"),"'","''") Else CompanyName = ""
					If rs8("Address") <> "" Then Address = Replace(rs8("Address"),"'","''") Else Address = ""
					If rs8("ContactFirstName") <> "" Then ContactFirstName = Replace(rs8("ContactFirstName"),"'","''") Else ContactFirstName = ""
					If rs8("ContactLastName") <> "" Then ContactLastName = Replace(rs8("ContactLastName"),"'","''") Else ContactLastName = ""
					If rs8("ContactFullName") <> "" Then ContactFullName = Replace(rs8("ContactFullName"),"'","''") Else ContactFullName = ""
			
					SQLWeighted = "SELECT * FROM zProspectPrequalWeighted_" & Trim(Session("userNo")) & " WHERE CustID='" & rs8("CustID") & "'"
					Set rsWeighted = cnn8.Execute(SQLWeighted) 
					
					If rsWeighted.EOF Then 
						SQLWeighted = "INSERT INTO zProspectPrequalWeighted_" & Trim(Session("userNo")) & " " 
						SQLWeighted = SQLWeighted & "(Source, Pool, ProspectRecordID, CustID, CompanyName, Address, PhoneNumber, "
						SQLWeighted = SQLWeighted & " ContactFirstName, ContactLastName, ContactFullName, "
						SQLWeighted = SQLWeighted & " Match_Company, Match_Name, Match_Address, Match_Phone) "
						SQLWeighted = SQLWeighted & "VALUES ( "
						SQLWeighted = SQLWeighted & "'" & rs8("Source") & "','"  & rs8("Pool") & "'," & rs8("ProspectRecordID") & ",'" & rs8("CustID") & "','"  & CompanyName & "','"  & Address & "','" & rs8("phoneNumber")& "', "
						SQLWeighted = SQLWeighted & "'" & ContactFirstName & "','"  & ContactLastName & "','"  & ContactFullName & "', "
						If Ucase(ContactLastName) = Ucase(txtLastName) Then
							SQLWeighted = SQLWeighted & "0,2,0,0)" ' Exact match gets an extra point
						Else
							SQLWeighted = SQLWeighted & "0,1,0,0)"
						End IF
					Else
						If Ucase(ContactLastName) = Ucase(txtLastName) Then
							SQLWeighted = "UPDATE zProspectPrequalWeighted_" & Trim(Session("userNo")) & " SET Match_Name = Match_Name + 2 WHERE CustID='" & rs8("CustID") & "'" ' Exact match gets an extra point
						Else
							SQLWeighted = "UPDATE zProspectPrequalWeighted_" & Trim(Session("userNo")) & " SET Match_Name = Match_Name + 1 WHERE CustID='" & rs8("CustID") & "'"
						End If
					End If
					Set rsWeighted = cnn8.Execute(SQLWeighted)  
				rs8.MoveNext
			Loop
		End IF
	End If

	'**************************************
	'Prospects with Contacts - company name
	'**************************************
	If txtCompanyName <> "" Then
		SQL8 = "SELECT 'Prospect Contact' as Source, PR_Prospects.Pool AS Pool, PR_Prospects.InternalRecordIdentifier AS ProspectRecordID, '' AS CustID, PR_Prospects.Company AS CompanyName, '' AS Address, PR_ProspectContacts.Phone AS PhoneNumber, "
		SQL8 = SQL8 & "PR_ProspectContacts.FirstName AS ContactFirstName, PR_ProspectContacts.LastName AS ContactLastName, PR_ProspectContacts.FirstName + ' ' + PR_ProspectContacts.LastName AS ContactFullName FROM  PR_ProspectContacts "
		SQL8 = SQL8 & "INNER JOIN PR_Prospects ON PR_ProspectContacts.ProspectIntRecID = PR_Prospects.InternalRecordIdentifier WHERE PR_Prospects.Company Like '%" & txtCompanyName & "%'"
		'Response.write(SQL8)
		Set rs8 = cnn8.Execute(SQL8)

		If not rs8.Eof Then
			Do While Not rs8.Eof
					If rs8("CompanyName") <> "" Then CompanyName = Replace(rs8("CompanyName"),"'","''") Else CompanyName = ""
					If rs8("Address") <> "" Then Address = Replace(rs8("Address"),"'","''") Else Address = ""
					If rs8("ContactFirstName") <> "" Then ContactFirstName = Replace(rs8("ContactFirstName"),"'","''") Else ContactFirstName = ""
					If rs8("ContactLastName") <> "" Then ContactLastName = Replace(rs8("ContactLastName"),"'","''") Else ContactLastName = ""
					If rs8("ContactFullName") <> "" Then ContactFullName = Replace(rs8("ContactFullName"),"'","''") Else ContactFullName = ""
			
					SQLWeighted = "SELECT * FROM zProspectPrequalWeighted_" & Trim(Session("userNo")) & " WHERE ProspectRecordID =" & rs8("ProspectRecordID") 
					Set rsWeighted = cnn8.Execute(SQLWeighted) 

					If rsWeighted.EOF Then 
						SQLWeighted = "INSERT INTO zProspectPrequalWeighted_" & Trim(Session("userNo")) & " " 
						SQLWeighted = SQLWeighted & "(Source, Pool, ProspectRecordID, CustID, CompanyName, Address, PhoneNumber, "
						SQLWeighted = SQLWeighted & " ContactFirstName, ContactLastName, ContactFullName, "
						SQLWeighted = SQLWeighted & " Match_Company, Match_Name, Match_Address, Match_Phone) "
						SQLWeighted = SQLWeighted & "VALUES ( "
						SQLWeighted = SQLWeighted & "'" & rs8("Source") & "','"  & rs8("Pool") & "'," & rs8("ProspectRecordID") & ",'" & rs8("CustID") & "','"  & CompanyName & "','"  & Address & "','" & rs8("phoneNumber")& "', "
						SQLWeighted = SQLWeighted & "'" & ContactFirstName & "','"  & ContactLastName & "','"  & ContactFullName & "', "
						If Ucase(CompanyName) = Ucase(txtCompanyName) Then
							SQLWeighted = SQLWeighted & "2,0,0,0)" ' Exact match gets an extra point
						Else
							SQLWeighted = SQLWeighted & "1,0,0,0)"
						End IF
					Else
						If Ucase(CompanyName) = Ucase(txtCompanyName) Then
							SQLWeighted = "UPDATE zProspectPrequalWeighted_" & Trim(Session("userNo")) & " SET Match_Company = Match_Company + 2 WHERE ProspectRecordID =" & rs8("ProspectRecordID") ' Exact match gets an extra point
						Else
							SQLWeighted = "UPDATE zProspectPrequalWeighted_" & Trim(Session("userNo")) & " SET Match_Company = Match_Company + 1 WHERE ProspectRecordID =" & rs8("ProspectRecordID") 
						End If
					End If
					Set rsWeighted = cnn8.Execute(SQLWeighted)  
  
				rs8.MoveNext
			Loop
		End IF
	End IF
	'*********************************
	'Prospects with Contacts - address
	'*********************************
	If txtAddress <> "" Then
		SQL8 = "SELECT 'Prospect Contact' as Source, PR_Prospects.Pool AS Pool, PR_Prospects.InternalRecordIdentifier AS ProspectRecordID, '' AS CustID, PR_Prospects.Company AS CompanyName, '' AS Address, PR_ProspectContacts.Phone AS PhoneNumber, "
		SQL8 = SQL8 & "PR_ProspectContacts.FirstName AS ContactFirstName, PR_ProspectContacts.LastName AS ContactLastName, PR_ProspectContacts.FirstName + ' ' + PR_ProspectContacts.LastName AS ContactFullName FROM  PR_ProspectContacts "
		SQL8 = SQL8 & "INNER JOIN PR_Prospects ON PR_ProspectContacts.ProspectIntRecID = PR_Prospects.InternalRecordIdentifier WHERE PR_Prospects.Street Like '%" & txtAddress & "%'"
		Set rs8 = cnn8.Execute(SQL8)
		If not rs8.Eof Then
			Do While Not rs8.Eof
					If rs8("CompanyName") <> "" Then CompanyName = Replace(rs8("CompanyName"),"'","''") Else CompanyName = ""
					If rs8("Address") <> "" Then Address = Replace(rs8("Address"),"'","''") Else Address = ""
					If rs8("ContactFirstName") <> "" Then ContactFirstName = Replace(rs8("ContactFirstName"),"'","''") Else ContactFirstName = ""
					If rs8("ContactLastName") <> "" Then ContactLastName = Replace(rs8("ContactLastName"),"'","''") Else ContactLastName = ""
					If rs8("ContactFullName") <> "" Then ContactFullName = Replace(rs8("ContactFullName"),"'","''") Else ContactFullName = ""
			
					SQLWeighted = "SELECT * FROM zProspectPrequalWeighted_" & Trim(Session("userNo")) & " WHERE ProspectRecordID =" & rs8("ProspectRecordID") 
					Set rsWeighted = cnn8.Execute(SQLWeighted) 

					If rsWeighted.EOF Then 
						SQLWeighted = "INSERT INTO zProspectPrequalWeighted_" & Trim(Session("userNo")) & " " 
						SQLWeighted = SQLWeighted & "(Source, Pool, ProspectRecordID, CustID, CompanyName, Address, PhoneNumber, "
						SQLWeighted = SQLWeighted & " ContactFirstName, ContactLastName, ContactFullName, "
						SQLWeighted = SQLWeighted & " Match_Company, Match_Name, Match_Address, Match_Phone) "
						SQLWeighted = SQLWeighted & "VALUES ( "
						SQLWeighted = SQLWeighted & "'" & rs8("Source") & "','"  & rs8("Pool") & "'," & rs8("ProspectRecordID") & ",'" & rs8("CustID") & "','"  & CompanyName & "','"  & Address & "','" & rs8("phoneNumber")& "', "
						SQLWeighted = SQLWeighted & "'" & ContactFirstName & "','"  & ContactLastName & "','"  & ContactFullName & "', "
						If Ucase(Address) = Ucase(txtAddress) Then
							SQLWeighted = SQLWeighted & "0,0,2,0)" ' Exact match gets an extra point
						Else
							SQLWeighted = SQLWeighted & "0,0,1,0)"
						End IF
					Else
						If Ucase(Address) = Ucase(txtAddress) Then
							SQLWeighted = "UPDATE zProspectPrequalWeighted_" & Trim(Session("userNo")) & " SET Match_Address = Match_Address + 2 WHERE ProspectRecordID =" & rs8("ProspectRecordID") ' Exact match gets an extra point
						Else
							SQLWeighted = "UPDATE zProspectPrequalWeighted_" & Trim(Session("userNo")) & " SET Match_Address = Match_Address + 1 WHERE ProspectRecordID =" & rs8("ProspectRecordID") 						
						End If
					End If
					Set rsWeighted = cnn8.Execute(SQLWeighted)  
 
				rs8.MoveNext
			Loop
		End IF
	End IF
	
	'*******************************
	'Prospects with Contacts - phone
	'*******************************
	If txtPhoneNumber <> "" Then
		SQL8 = "SELECT 'Prospect Contact' as Source, PR_Prospects.Pool AS Pool, PR_Prospects.InternalRecordIdentifier AS ProspectRecordID, '' AS CustID, PR_Prospects.Company AS CompanyName, '' AS Address, PR_ProspectContacts.Phone AS PhoneNumber, "
		SQL8 = SQL8 & "PR_ProspectContacts.FirstName AS ContactFirstName, PR_ProspectContacts.LastName AS ContactLastName, PR_ProspectContacts.FirstName + ' ' + PR_ProspectContacts.LastName AS ContactFullName FROM  PR_ProspectContacts "
		SQL8 = SQL8 & "INNER JOIN PR_Prospects ON PR_ProspectContacts.ProspectIntRecID = PR_Prospects.InternalRecordIdentifier WHERE PR_ProspectContacts.Phone Like '%" & txtPhoneNumber & "%'"
		Set rs8 = cnn8.Execute(SQL8)
		
		If not rs8.Eof Then
			Do While Not rs8.Eof
					If rs8("CompanyName") <> "" Then CompanyName = Replace(rs8("CompanyName"),"'","''") Else CompanyName = ""
					If rs8("Address") <> "" Then Address = Replace(rs8("Address"),"'","''") Else Address = ""
					If rs8("ContactFirstName") <> "" Then ContactFirstName = Replace(rs8("ContactFirstName"),"'","''") Else ContactFirstName = ""
					If rs8("ContactLastName") <> "" Then ContactLastName = Replace(rs8("ContactLastName"),"'","''") Else ContactLastName = ""
					If rs8("ContactFullName") <> "" Then ContactFullName = Replace(rs8("ContactFullName"),"'","''") Else ContactFullName = ""
			
					SQLWeighted = "SELECT * FROM zProspectPrequalWeighted_" & Trim(Session("userNo")) & " WHERE ProspectRecordID =" & rs8("ProspectRecordID") 
					Set rsWeighted = cnn8.Execute(SQLWeighted) 

					If rsWeighted.EOF Then 
						SQLWeighted = "INSERT INTO zProspectPrequalWeighted_" & Trim(Session("userNo")) & " " 
						SQLWeighted = SQLWeighted & "(Source, Pool, ProspectRecordID, CustID, CompanyName, Address, PhoneNumber, "
						SQLWeighted = SQLWeighted & " ContactFirstName, ContactLastName, ContactFullName, "
						SQLWeighted = SQLWeighted & " Match_Company, Match_Name, Match_Address, Match_Phone) "
						SQLWeighted = SQLWeighted & "VALUES ( "
						SQLWeighted = SQLWeighted & "'" & rs8("Source") & "','"  & rs8("Pool") & "'," & rs8("ProspectRecordID") & ",'" & rs8("CustID") & "','"  & CompanyName & "','"  & Address & "','" & rs8("phoneNumber")& "', "
						SQLWeighted = SQLWeighted & "'" & ContactFirstName & "','"  & ContactLastName & "','"  & ContactFullName & "', "
						If Ucase(rs8("phoneNumber")) = Ucase(txtPhoneNumber) Then
							SQLWeighted = SQLWeighted & "0,0,0,2)" ' Exact match gets an extra point
						Else
							SQLWeighted = SQLWeighted & "0,0,0,1)"
						End IF
					Else
						If Ucase(rs8("phoneNumber")) = Ucase(txtPhoneNumber) Then
							SQLWeighted = "UPDATE zProspectPrequalWeighted_" & Trim(Session("userNo")) & " SET Match_Phone = Match_Phone + 2 WHERE ProspectRecordID =" & rs8("ProspectRecordID") ' Exact match gets an extra point
						Else
							SQLWeighted = "UPDATE zProspectPrequalWeighted_" & Trim(Session("userNo")) & " SET Match_Phone = Match_Phone + 1 WHERE ProspectRecordID =" & rs8("ProspectRecordID") 						
						End If
					End If
					Set rsWeighted = cnn8.Execute(SQLWeighted)  
 
				rs8.MoveNext
			Loop
		End IF
	End If

	'***********************************
	'Prospects with Contacts - last name
	'***********************************
	If txtLastName <> "" Then
		SQL8 = "SELECT 'Prospect Contact' as Source, PR_Prospects.Pool AS Pool, PR_Prospects.InternalRecordIdentifier AS ProspectRecordID, '' AS CustID, PR_Prospects.Company AS CompanyName, '' AS Address, PR_ProspectContacts.Phone AS PhoneNumber, "
		SQL8 = SQL8 & "PR_ProspectContacts.FirstName AS ContactFirstName, PR_ProspectContacts.LastName AS ContactLastName, PR_ProspectContacts.FirstName + ' ' + PR_ProspectContacts.LastName AS ContactFullName FROM  PR_ProspectContacts "
		SQL8 = SQL8 & "INNER JOIN PR_Prospects ON PR_ProspectContacts.ProspectIntRecID = PR_Prospects.InternalRecordIdentifier WHERE PR_ProspectContacts.LastName Like '%" & txtLastName & "%'"
		Set rs8 = cnn8.Execute(SQL8)
		
		If not rs8.Eof Then
			Do While Not rs8.Eof
					If rs8("CompanyName") <> "" Then CompanyName = Replace(rs8("CompanyName"),"'","''") Else CompanyName = ""
					If rs8("Address") <> "" Then Address = Replace(rs8("Address"),"'","''") Else Address = ""
					If rs8("ContactFirstName") <> "" Then ContactFirstName = Replace(rs8("ContactFirstName"),"'","''") Else ContactFirstName = ""
					If rs8("ContactLastName") <> "" Then ContactLastName = Replace(rs8("ContactLastName"),"'","''") Else ContactLastName = ""
					If rs8("ContactFullName") <> "" Then ContactFullName = Replace(rs8("ContactFullName"),"'","''") Else ContactFullName = ""
			
					SQLWeighted = "SELECT * FROM zProspectPrequalWeighted_" & Trim(Session("userNo")) & " WHERE ProspectRecordID =" & rs8("ProspectRecordID") 
					Set rsWeighted = cnn8.Execute(SQLWeighted) 

					If rsWeighted.EOF Then 
						SQLWeighted = "INSERT INTO zProspectPrequalWeighted_" & Trim(Session("userNo")) & " " 
						SQLWeighted = SQLWeighted & "(Source, Pool, ProspectRecordID, CustID, CompanyName, Address, PhoneNumber, "
						SQLWeighted = SQLWeighted & " ContactFirstName, ContactLastName, ContactFullName, "
						SQLWeighted = SQLWeighted & " Match_Company, Match_Name, Match_Address, Match_Phone) "
						SQLWeighted = SQLWeighted & "VALUES ( "
						SQLWeighted = SQLWeighted & "'" & rs8("Source") & "','"  & rs8("Pool") & "'," & rs8("ProspectRecordID") & ",'" & rs8("CustID") & "','"  & CompanyName & "','"  & Address & "','" & rs8("phoneNumber")& "', "
						SQLWeighted = SQLWeighted & "'" & ContactFirstName & "','"  & ContactLastName & "','"  & ContactFullName & "', "
						If Ucase(ContactLastName) = Ucase(txtContactLastName) Then
							SQLWeighted = SQLWeighted & "0,2,0,0)" ' Exact match gets an extra point
						Else
							SQLWeighted = SQLWeighted & "0,1,0,0)"
						End IF
					Else
						If Ucase(ContactLastName) = Ucase(txtContactLastName) Then
							SQLWeighted = "UPDATE zProspectPrequalWeighted_" & Trim(Session("userNo")) & " SET Match_Name = Match_Name + 2 WHERE ProspectRecordID =" & rs8("ProspectRecordID") ' Exact match gets an extra point
						Else
							SQLWeighted = "UPDATE zProspectPrequalWeighted_" & Trim(Session("userNo")) & " SET Match_Name = Match_Name + 1 WHERE ProspectRecordID =" & rs8("ProspectRecordID") 						
						End If
					End If
					Set rsWeighted = cnn8.Execute(SQLWeighted)  
 
				rs8.MoveNext
			Loop
		End IF
	End If

	'*****************************************
	'Prospects WITHOUT Contacts - company name
	'*****************************************
	If txtCompanyName <> "" Then
		SQL8 = "SELECT 'Prospects' as Source, PR_Prospects.Pool AS Pool, PR_Prospects.InternalRecordIdentifier AS ProspectRecordID, '' AS CustID, PR_Prospects.Company AS CompanyName, PR_Prospects.Street AS Address, '' AS PhoneNumber, "
		SQL8 = SQL8 & "'' AS ContactFirstName, '' AS ContactLastName, '' AS ContactFullName FROM "
		SQL8 = SQL8 & "PR_Prospects WHERE PR_Prospects.Company Like '%" & txtCompanyName & "%'"
		Set rs8 = cnn8.Execute(SQL8)

		If not rs8.Eof Then
			Do While Not rs8.Eof
					If rs8("CompanyName") <> "" Then CompanyName = Replace(rs8("CompanyName"),"'","''") Else CompanyName = ""
					If rs8("Address") <> "" Then Address = Replace(rs8("Address"),"'","''") Else Address = ""
					If rs8("ContactFirstName") <> "" Then ContactFirstName = Replace(rs8("ContactFirstName"),"'","''") Else ContactFirstName = ""
					If rs8("ContactLastName") <> "" Then ContactLastName = Replace(rs8("ContactLastName"),"'","''") Else ContactLastName = ""
					If rs8("ContactFullName") <> "" Then ContactFullName = Replace(rs8("ContactFullName"),"'","''") Else ContactFullName = ""
			
					SQLWeighted = "SELECT * FROM zProspectPrequalWeighted_" & Trim(Session("userNo")) & " WHERE ProspectRecordID =" & rs8("ProspectRecordID") 
					Set rsWeighted = cnn8.Execute(SQLWeighted) 

					If rsWeighted.EOF Then 
						SQLWeighted = "INSERT INTO zProspectPrequalWeighted_" & Trim(Session("userNo")) & " " 
						SQLWeighted = SQLWeighted & "(Source, Pool, ProspectRecordID, CustID, CompanyName, Address, PhoneNumber, "
						SQLWeighted = SQLWeighted & " ContactFirstName, ContactLastName, ContactFullName, "
						SQLWeighted = SQLWeighted & " Match_Company, Match_Name, Match_Address, Match_Phone) "
						SQLWeighted = SQLWeighted & "VALUES ( "
						SQLWeighted = SQLWeighted & "'" & rs8("Source") & "','"  & rs8("Pool") & "'," & rs8("ProspectRecordID") & ",'" & rs8("CustID") & "','"  & CompanyName & "','"  & Address & "','" & rs8("phoneNumber")& "', "
						SQLWeighted = SQLWeighted & "'" & ContactFirstName & "','"  & ContactLastName & "','"  & ContactFullName & "', "
						If Ucase(CompanyName) = Ucase(txtCompanyName) Then
							SQLWeighted = SQLWeighted & "2,0,0,0)" ' Exact match gets an extra point
						Else
							SQLWeighted = SQLWeighted & "1,0,0,0)"
						End IF
					Else
						If Ucase(CompanyName) = Ucase(txtCompanyName) Then
							SQLWeighted = "UPDATE zProspectPrequalWeighted_" & Trim(Session("userNo")) & " SET Match_Company = Match_Company + 2 WHERE ProspectRecordID =" & rs8("ProspectRecordID") ' Exact match gets an extra point
						Else
							SQLWeighted = "UPDATE zProspectPrequalWeighted_" & Trim(Session("userNo")) & " SET Match_Company = Match_Company + 1 WHERE ProspectRecordID =" & rs8("ProspectRecordID") 
						End If
					End If
					Set rsWeighted = cnn8.Execute(SQLWeighted)  
  
				rs8.MoveNext
			Loop
		End IF
	End IF
	

	SQL8 = "SELECT * FROM zProspectPrequalWeighted_" & Trim(Session("userNo")) & " ORDER BY Match_Phone + Match_Company + Match_Name + Match_Address DESC"
    Set rs8 = cnn8.Execute(SQL8)         

	If NOT rs8.EOF Then 
	
		rowCount = 0
		prospectMatchFound = True 
		%>
		
		<div class="col-lg-8" id="prospectMatches">
		<div class="row">
			<h2 style="margin-bottom:15px;"><%= GetTerm("Potential Duplicates Found") %>:</h2>
			<table class="table table-sm">
			  <thead>
			    <tr>
			      <th>Company</th>
			      <th>Street Address</th>
			      <th>Contact First Name</th>
			      <th>Contact Last Name</th>			      
			      <th>Phone</th>
			      <th>Source</th>
			    </tr>
			  </thead>
			  <tbody>
			    <tr>

		<%
		FirstRec = True
		
		Do While Not rs8.EOF
		
			ProspectRecordID = rs8("ProspectRecordID")
			CustID = rs8("CustID")
			enteredPhoneNumber = stripNonNumeric(txtPhoneNumber)
			FullName = UCASE(txtFirstName) & " " & UCASE(txtLastName)

				%>
			    <tr>
			      <% 
			      If rs8("Match_Company") <> 0 Then %>
					<td><strong><font color="red"><%= rs8("CompanyName") %></font></strong></td>
			      <% Else %>
			      	<td><%= rs8("CompanyName") %></td>			      
			      <% End If 
			      If rs8("Match_Address") <> 0 Then %>
					<td><strong><font color="red"><%= rs8("Address") %></font></strong></td>
			      <% Else %>
			      	<td><%= rs8("Address") %></td>			      
			      <% End If
			      If rs8("Match_Name") <> 0 Then %>
					<td><%= rs8("ContactFirstName") %></td>
					<td><strong><font color="red"><%= rs8("ContactLastName") %></font></strong></td>
			      <% Else %>
			      	<td><%= rs8("ContactFirstName") %></td>			      
			      	<td><%= rs8("ContactLastName") %></td>
			      <% End If
			      
			      If rs8("Match_Phone") <> 0 Then %>
					<td><strong><font color="red"><%= rs8("PhoneNumber") %></font></strong></td>
			      <% Else %>
			      	<td><%= rs8("PhoneNumber") %></td>			      
			      <% End If 
			      
			      	'***************************************************************************************************************************************************************
					'They need further clarification because a match in the customer file can be for an active or inactive customer. So, can you change it so that 
					'if the match is from the customer file (AR_Customer) we lookup the field AR_Customer.AcctStatus. There are 3 possible values: A,I,X.
					
					'A = Customer (Active)
					'I  = Customer (Inactive)
					'X = Customer (Closed status X)
					'***************************************************************************************************************************************************************
	      			
	      			AcctStatus = ""
	      			
		      		If CustID <> "" Then
						
						SQLCustAcctStatus = "SELECT AcctStatus FROM AR_Customer WHERE CustNum = '" & CustID & "'"
						
						Set cnnCustAcctStatus = Server.CreateObject("ADODB.Connection")
						cnnCustAcctStatus.open (Session("ClientCnnString"))
						
						Set rsCustAcctStatus = Server.CreateObject("ADODB.Recordset")
						rsCustAcctStatus.CursorLocation = 3 
						Set rsCustAcctStatus = cnnCustAcctStatus.Execute(SQLCustAcctStatus)
						
						If NOT rsCustAcctStatus.EOF Then
							AcctStatus = rsCustAcctStatus("AcctStatus")
						End If
						
						set rsCustAcctStatus = Nothing
						cnnCustAcctStatus.Close
						set cnnCustAcctStatus = Nothing

		      		End If
		      		
		      		If AcctStatus = "X" Then
		      			AcctStatus = "- Closed Acct"
		      		ElseIf AcctStatus = "I" Then
		      			AcctStatus = "- Inactive Acct"
		      		ElseIf AcctStatus = "A" Then
		      			AcctStatus = "- Active Acct"
		      		Else
		      			AcctStatus = ""
		      		End If		      			
			      
			      If rs8("Source") = "Prospects" OR rs8("Source") = "Prospect Contact" Then
			      
			      		
   			      		If  rs8("Pool") = "Won" Then 
			      			SourceDesc = "Won"
			      			%><td><a href="<%= BaseURL %>prospecting/viewProspectDetailWonPool.asp?i=<%= ProspectRecordID %>" target="_blank"><%= SourceDesc %> <%= AcctStatus %></a></td><%
			      		End If
			      		If  rs8("Pool") = "Live" Then 
			      			SourceDesc = "Prospect"
			      			%><td><a href="<%= BaseURL %>prospecting/viewProspectDetail.asp?i=<%= ProspectRecordID %>" target="_blank"><%= SourceDesc %> <%= AcctStatus %></a></td><%
			      		End If
			      		If  rs8("Pool") = "Dead" Then
			      			SourceDesc = GetTerm("Recycle Pool") %>
			      			<td><a href="<%= BaseURL %>prospecting/viewProspectDetailRecyclePool.asp?i=<%= ProspectRecordID %>" target="_blank"><%= SourceDesc %> <%= AcctStatus %></a></td>
			      		<% End If %>
			      <% Else %>
			      		<td><%= rs8("Source")%> <%= AcctStatus %></td>
			      <% End If %>
			      
			    </tr>
				<%

			rs8.MoveNext
		Loop
		
		End If

		txtFirstName = Request.form("txtFirstName")
		txtLastName = Request.form("txtLastName")
		txtCompanyName = Request.form("txtCompanyName")
		txtPhoneNumber = Request.form("txtPhoneNumber")
		txtAddress = Request.form("txtAddress")

			
		
		If prospectMatchFound = True Then
			%>
				</tbody>
				</table>
				<form action="<%= BaseURL %>prospecting/addProspect.asp" method="POST" name="frmAddProspect" id="frmAddProspect">	
					<input type="hidden" name="txtFirstName" id="txtFirstName" value="<%= txtFirstName %>">
					<input type="hidden" name="txtLastName" id="txtLastName" value="<%= txtLastName %>">
					<input type="hidden" name="txtCompanyName" id="txtCompanyName" value="<%= txtCompanyName %>">
					<input type="hidden" name="txtPhoneNumber" id="txtPhoneNumber" value="<%= txtPhoneNumber %>">
					<input type="hidden" name="txtAddress" id="txtAddress" value="<%= txtAddress%>">
					<button type="submit" class="btn btn-success pull-right btn-xlarge" id="btnPreqialifyProspectContinue">
				        <i class="fa fa-arrow-right"></i>&nbsp;Continue To Add Prospect Anyway
				    </button>
				</form>
				<!-- tabs start here !-->
			</div>
			</div>
 
			
		<% End If %>
	<%	
	If prospectMatchFound = False Then
		%>
		<form action="<%= BaseURL %>prospecting/addProspect.asp" method="POST" name="frmAddProspectAuto" id="frmAddProspectAuto">	
			<input type="hidden" name="txtFirstName" id="txtFirstName" value="<%= txtFirstName %>">
			<input type="hidden" name="txtLastName" id="txtLastName" value="<%= txtLastName %>">
			<input type="hidden" name="txtCompanyName" id="txtCompanyName" value="<%= txtCompanyName %>">
			<input type="hidden" name="txtPhoneNumber" id="txtPhoneNumber" value="<%= txtPhoneNumber %>">
			<input type="hidden" name="txtAddress" id="txtAddress" value="<%= txtAddress%>">
		</form>
		<script language="JavaScript">
			document.forms['frmAddProspectAuto'].submit();
		</script>
		<%
	End If
End If
%>

<div class="row the-tabs">
 
	<!-- tab content !-->
	<div class="tab-content">
		<form action="<%= BaseURL %>prospecting/addProspectPrequalify.asp" method="POST" name="frmAddProspectPrequalify" id="frmAddProspectPrequalify" onsubmit="return validatePrequalifyProspectForm();">	
	

		<!-- lead information  !-->
		<div role="tabpanel" class="tab-pane active" id="leadinformation">
			
			<div class="col-lg-8">
			<div class="row" style="border: 1px solid #999;margin-top:30px; padding:20px">
			
				<% If txtFirstName <> "" Then %>
					<h2 style="margin-bottom:15px;"><%= GetTerm("Check For Duplicates") %> Again:</h2>
					<hr class="style7">
				<% End If %>
				

				<!-- left column !-->
				<div class="col-lg-6">
					<div class="table-responsive">
					  <table class="table">
					  	<tbody>
						  	
						  	<!-- line !-->
						  	<tr style="margin-bottom:30px;">
							  	<th class="th-width"><strong>Created By</strong></th>
							  	<th class="th-width2"><%= CreatedByUserName %></th>
						  	</tr>
						  	<!-- eof line !-->
						  							  	
						  	<!-- line !-->
						  	<tr>
							  	<th class="th-width"><strong>Company</strong></th>
							  	<th class="th-width2"><input type="text" class="form-control red-line" name="txtCompanyName" id="txtCompanyName" value="<%= txtCompanyName %>"></th>
						  	</tr>
						  	<!-- eof line !-->
						  	
						  	<!-- line !-->
						  	<tr>
							  	<th class="th-width"><strong>Street Address</strong></th>
							  	<th class="th-width2"><input type="text" class="form-control" name="txtAddress" id="txtAddress" value="<%= txtAddress %>"></th>
						  	</tr>
						  	<!-- eof line !-->
						  	
						  	<!-- line !-->
						  	<tr>
							  	<th class="th-width"><strong>Contact First Name</strong></th>
							  	<th class="th-width2"><input type="text" class="form-control" name="txtFirstName" id="txtFirstName" value="<%= txtFirstName %>"></th>
						  	</tr>
						  	<!-- eof line !-->
						  	
						  	<!-- line !-->
						  	<tr>
							  	<th class="th-width"><strong>Contact Last Name</strong></th>
							  	<th class="th-width2"><input type="text" class="form-control" name="txtLastName" id="txtLastName" value="<%= txtLastName %>"></th>
						  	</tr>
						  	<!-- eof line !-->
						  							  	
						  	<!-- line !-->
						  	<tr>
							  	<th class="th-width"><strong>Phone Number</strong></th>
							  	<th class="th-width2"><input type="text" class="form-control" name="txtPhoneNumber" id="txtPhoneNumber" value="<%= txtPhoneNumber %>"></th>
						  	</tr>
						  	<!-- eof line !-->
						   						  	
					  	</tbody>
					  </table>
			  
					</div>
				</div>
				<div class="row">
			
				
				<!-- left column !-->
				<div class="col-lg-8">
					<div class="table-responsive">
					  <table class="table">
					  	<tbody>
					  	
						  	<tr>	
						  		<th class="th-width">
								    <a class="btn btn-primary pull-left btn-xlarge" href="main.asp" role="button"><i class="fa fa-arrow-left"></i>&nbsp; Back To <%= GetTerm("Prospect") %> List</a>
						  		</th>
						  		<th class="th-width">		    
									<button type="submit" class="btn btn-success pull-right btn-xlarge" id="btnPreqialifyProspect">
								        <i class="fa fa-user"></i>&nbsp;Click To <%= GetTerm("Check For Duplicates") %>
								    </button>
							    </th>
							</tr>
						<!-- eof line !-->
						   						  	
					  	</tbody>
					  </table>
					</div>
				</div>
				<!-- eof right column !-->		  	
						  				
			</div>
			</div>
		</div>
		<!-- eof lead information !-->
		
		 </form>
		
		 
		
	</div>
	<!-- eof tab content !-->
	
</div>
<!-- tabs end here !-->

 
 <!-- tabs js !-->
 <script type="text/javascript">
 $('a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
  e.target // newly activated tab
  e.relatedTarget // previous active tab
})
 </script>
 <!-- eof tabs js !-->

<!--#include file="../inc/footer-main.asp"-->
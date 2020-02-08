<!--#include file="inc/header-tech-and-driver.asp"-->



<style type="text/css">
	body{
		overflow-x: hidden;
	}
	
	.fieldservice-heading{
		line-height: 1;
		padding-top: 5px;
		padding-bottom: 5px;
	}
	
	.btn-home{
		font-size: 11px;
	}
	
 ul{
	 color: #000;
	 font-size: 11px;
	 text-transform: uppercase;
	 list-style-type: none;
	     -webkit-margin-before: 0px;
    -webkit-margin-after: 0px;
    -webkit-margin-start: 0px;
    -webkit-margin-end: 0px;
    -webkit-padding-start: 0px;
 }
 
 .enroute{
	 color: green;
 }
 
 .btn-spacing{
	 margin-bottom: 40px;
 }
 
 .pull-left{
	 margin-left: 5px;
 }
 
.btn-block {
    width: auto;
    display: inline-block;
}

.container-options{
	margin-top: 20px;
	margin-bottom: 20px;
	font-size: 12px;
}
 
	</style>       


 
<h1 class="fieldservice-heading" ><a class="btn btn-default btn-home pull-left" href="main_menu.asp" role="button"><i class="fa fa-home"></i>
 Home</a>
<%If filterChangeModuleOn() = True and prevMaintModuleOn() <> True Then%>
	Filter Changes
<% ElseIf filterChangeModuleOn() <> True and prevMaintModuleOn() = True Then %>
	PM Calls
<% ElseIf filterChangeModuleOn() = True and prevMaintModuleOn() = True Then %>
	Filter Changes / PM Calls
<%End If%>
</h1>

<% If Request.Form("optWhatToShow") <> "" Then
	Session("MulitUseVar") = Request.Form("optWhatToShow")
Else
	Session("MulitUseVar")="All"
End If %>

<form method="post" action="filterchanges.asp" name="frmOptionClick" id="frmOptionClick">
	<div class="row">
		<div class="container-fluid container-options">
	
		<!-- col !-->
		<div class="col-lg-3 col-md-3 col-sm-3 col-xs-3">
		    <% If Session("MulitUseVar")="All" Then %>
				<input type="radio" name="optWhatToShow" id="optWhatToShow" value="All" checked>
			<% Else %>
				<input type="radio" name="optWhatToShow" id="optWhatToShow" value="All">
			<%End If %>
			Show All 
		</div>
		<!-- eof col !-->
			
		<!-- col !-->
		<div class="col-lg-3 col-md-3 col-sm-3 col-xs-3">
			<% If Session("MulitUseVar")="PMOnly" Then %>
				<input type="radio" name="optWhatToShow" id="optWhatToShow" value="PMOnly" checked>
			<% Else %>
				<input type="radio" name="optWhatToShow" id="optWhatToShow" value="PMOnly">
			<% End If %>
			PMs Only 
		</div>
		<!-- eof col !-->
		
		<!-- col !-->
		<div class="col-lg-3 col-md-3 col-sm-3 col-xs-3">
		<% If Session("MulitUseVar")="FiltersOnly" Then %>
			<input type="radio" name="optWhatToShow" id="optWhatToShow" value="FiltersOnly" checked>
		<% Else %>
			<input type="radio" name="optWhatToShow" id="optWhatToShow" value="FiltersOnly">
		<% End If %>
			Filter Changes Only 
		</div>
		<!-- eof col !-->
		
		<!-- col !-->
		<div class="col-lg-3 col-md-3 col-sm-3 col-xs-3">
			<% If Session("MulitUseVar")="Both" Then %>
				<input type="radio" name="optWhatToShow" id="optWhatToShow" value="Both" checked>
			<% Else %>
				<input type="radio" name="optWhatToShow" id="optWhatToShow" value="Both">
			<%End If%>			
			Clients With PMs AND Filter Changes (Both)
		</div>
		<!-- eof col !-->
		
		<!-- col !-->
		<div class="col-lg-12">
		 <button class="btn btn-info btn-sm" type="submit">APPLY</button> 
		</div>
		<!-- eof col !-->
		
		</div>
	</div>
</form>


<div class="container-fluid fieldservice-container">

<%
'Remember, it reads the setting FieldServiceDays from tblSetting_Global
'to determine how many days to use in the evaluation
SQLtmp = "SELECT * FROM Settings_Global"
Set cnntmp = Server.CreateObject("ADODB.Connection")
cnntmp.open (Session("ClientCnnString"))
Set rstmp = Server.CreateObject("ADODB.Recordset")
rstmp.CursorLocation = 3 
Set rstmp = cnntmp.Execute(SQLtmp)
If not rstmp.EOF Then 
	If filterChangeModuleOn() = True Then FilterChangeDaysFieldService = rstmp("FilterChangeDaysFieldService")
	If prevMaintModuleOn() = True Then PMCallDaysFieldService = rstmp("PMCallDaysFieldService")
Else
	FilterChangeDaysFieldService = 15
	PMCallDaysFieldService = 15
End If
set rstmp = Nothing
cnntmp.close
set cnntmp = Nothing


SQL = "SELECT * FROM Assets "
SQL = SQL & "INNER JOIN EQ_ScheduledServiceDates ON Assets.assetNumber = EQ_ScheduledServiceDates.assetNumber "
SQL = SQL & "WHERE "
If filterChangeModuleOn() = True AND  prevMaintModuleOn() = True Then
	Select Case Session("MulitUseVar")
		Case "All"
			SQL = SQL & "(Assets.assetTypeNo = 335) AND "
			SQL = SQL & "(Assets.custAcctNum IN (SELECT CustNum FROM AR_Customer WHERE (AcctStatus = 'A')))  AND "
			SQL = SQL & "EQ_ScheduledServiceDates.nextDate1 IS NOT NULL "
			SQL = SQL & "AND EQ_ScheduledServiceDates.nextDate1 <= DateAdd(day," & FilterChangeDaysFieldService & ",getdate()) "
			SQL = SQL & " OR "
			SQL = SQL & "(Assets.assetTypeNo = 336) AND "
			SQL = SQL & "(Assets.custAcctNum IN (SELECT CustNum FROM AR_Customer WHERE (AcctStatus = 'A')))  AND "
			SQL = SQL & "EQ_ScheduledServiceDates.nextDate1 IS NOT NULL "
			SQL = SQL & "AND EQ_ScheduledServiceDates.nextDate1 <= DateAdd(day," & PMCallDaysFieldService & ",getdate()) "
			SQL = SQL & "ORDER BY EQ_ScheduledServiceDates.nextDate1, custAcctNum "
		Case "PMOnly"
			SQL = SQL & "(Assets.assetTypeNo = 336) AND "
			SQL = SQL & "(Assets.custAcctNum IN (SELECT CustNum FROM AR_Customer WHERE (AcctStatus = 'A')))  AND "
			SQL = SQL & "EQ_ScheduledServiceDates.nextDate1 IS NOT NULL "
			SQL = SQL & "AND EQ_ScheduledServiceDates.nextDate1 <= DateAdd(day," & PMCallDaysFieldService & ",getdate()) "
			SQL = SQL & "ORDER BY EQ_ScheduledServiceDates.nextDate1, custAcctNum "
		Case "FiltersOnly"
			SQL = SQL & "(Assets.assetTypeNo = 335) AND "
			SQL = SQL & "(Assets.custAcctNum IN (SELECT CustNum FROM AR_Customer WHERE (AcctStatus = 'A')))  AND "
			SQL = SQL & "EQ_ScheduledServiceDates.nextDate1 IS NOT NULL "
			SQL = SQL & "AND EQ_ScheduledServiceDates.nextDate1 <= DateAdd(day," & FilterChangeDaysFieldService & ",getdate()) "
			SQL = SQL & "ORDER BY EQ_ScheduledServiceDates.nextDate1, custAcctNum "
		Case "Both"
			'This one is a little more complex, we have to put stuff into a temp work file first
			Set cnn8 = Server.CreateObject("ADODB.Connection")
			cnn8.open (Session("ClientCnnString"))
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.CursorLocation = 3 
			' Drop & create temporary table
			on error resume next
			SQL = "DROP TABLE zAssetCustList_"  & Trim(Session("userNo"))
			Set rs = cnn8.Execute(SQL)
			on error goto 0
			SQL = "SELECT DISTINCT Assets.assetNumber, Assets.custAcctNum into zAssetCustList_"  & Trim(Session("userNo")) & " FROM Assets "
			SQL = SQL & "INNER JOIN EQ_ScheduledServiceDates ON Assets.assetNumber = EQ_ScheduledServiceDates.assetNumber "
			SQL = SQL & "WHERE "
			SQL = SQL & "Assets.assetTypeNo = 335 AND "
			SQL = SQL & "EQ_ScheduledServiceDates.nextDate1 IS NOT NULL "
			SQL = SQL & "AND EQ_ScheduledServiceDates.nextDate1 <= DateAdd(day," & FilterChangeDaysFieldService & ",getdate())  "
			SQL = SQL & "AND custAcctNum IN (SELECT CustNum FROM AR_Customer WHERE AcctStatus = 'A')  "
			Set rs = cnn8.Execute(SQL)
			'Step 2
			SQL = "INSERT INTO zAssetCustList_"  & Trim(Session("userNo")) & " (assetNumber, custAcctNum )  "
			SQL =  SQL & "SELECT DISTINCT Assets.assetNumber, Assets.custAcctNum FROM Assets "
			SQL = SQL & "INNER JOIN EQ_ScheduledServiceDates ON Assets.assetNumber = EQ_ScheduledServiceDates.assetNumber "
			SQL = SQL & "WHERE "
			SQL = SQL & "Assets.assetTypeNo = 336 AND "
			SQL = SQL & "EQ_ScheduledServiceDates.nextDate1 IS NOT NULL "
			SQL = SQL & "AND EQ_ScheduledServiceDates.nextDate1 <= DateAdd(day," & FilterChangeDaysFieldService & ",getdate()) "
			SQL = SQL & "AND custAcctNum IN (SELECT CustNum FROM AR_Customer WHERE AcctStatus = 'A')  "
			Set rs = cnn8.Execute(SQL)
			' Delete secondary work table
			on error resume next
			SQL = "DROP TABLE zAssetCustListTable2_"  & Trim(Session("userNo"))
			Set rs = cnn8.Execute(SQL)
			on error goto 0
			SQL = "SELECT custAcctNum, COUNT(custAcctNum) AS Expr1 into zAssetCustListTable2_"  & Trim(Session("userNo"))
			SQL = SQL & " FROM zAssetCustList_"  & Trim(Session("userNo")) & " GROUP BY custAcctNum"
			Set rs = cnn8.Execute(SQL)
			SQL = "DELETE FROM zAssetCustListTable2_"  & Trim(Session("userNo")) & " WHERE Expr1 < 2"
			Set rs = cnn8.Execute(SQL)
			' Winnow down the original work table
			SQL = "DELETE FROM zAssetCustList_"  & Trim(Session("userNo"))  & " WHERE custacctnum IN ( "
			SQL = SQL & "SELECT custAcctNum FROM zAssetCustList_"  & Trim(Session("userNo")) 
			SQL = SQL & " EXCEPT "
			SQL = SQL & "SELECT custAcctNum FROM zAssetCustListTable2_"  & Trim(Session("userNo")) & ")"
			Set rs = cnn8.Execute(SQL)
			Set rs = Nothing
			cnn8.close
			Set cnn8 = Nothing
			SQL = "SELECT * FROM Assets "
			SQL = SQL & "INNER JOIN EQ_ScheduledServiceDates ON Assets.assetNumber = EQ_ScheduledServiceDates.assetNumber "
			SQL = SQL & "WHERE "
			SQL = SQL & "Assets.assetNumber IN (Select assetNumber from zAssetCustList_"  & Trim(Session("userNo")) &") "
			SQL = SQL & "ORDER BY custAcctNum "
			'response.write(SQL)
		End Select
ElseIf filterChangeModuleOn() = True Then 
	SQL = SQL & "(Assets.assetTypeNo = 335) AND "
	SQL = SQL & "(Assets.custAcctNum IN (SELECT CustNum FROM AR_Customer WHERE (AcctStatus = 'A')))  AND "
	SQL = SQL & "EQ_ScheduledServiceDates.nextDate1 IS NOT NULL "
	SQL = SQL & "AND EQ_ScheduledServiceDates.nextDate1 <= DateAdd(day," & FilterChangeDaysFieldService & ",getdate()) "
	SQL = SQL & "ORDER BY EQ_ScheduledServiceDates.nextDate1, custAcctNum "
ElseIf prevMaintModuleOn() = True Then
	SQL = SQL & "(Assets.assetTypeNo = 336) AND "
	SQL = SQL & "(Assets.custAcctNum IN (SELECT CustNum FROM AR_Customer WHERE (AcctStatus = 'A')))  AND "
	SQL = SQL & "EQ_ScheduledServiceDates.nextDate1 IS NOT NULL "
	SQL = SQL & "AND EQ_ScheduledServiceDates.nextDate1 <= DateAdd(day," & PMCallDaysFieldService & ",getdate()) "
	SQL = SQL & "ORDER BY EQ_ScheduledServiceDates.nextDate1, custAcctNum "
End If



'Response.Write("<br>MultiUserVar:" & Session("MulitUseVar") &":DD<br>") 
'Response.Write(SQL)


Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rsCust = Server.CreateObject("ADODB.Recordset")
rsCust.CursorLocation = 3 


Set rs = cnn8.Execute(SQL)
			
	If not rs.EOF Then
		Do While Not rs.EOF
		
			If rs("assetTypeNo") = 335 Then 'This is a filter change, do PM calls below
			
					If FilterChangeSubmitted(rs("assetNumber"),rs("nextDate1")) <> True Then 
						
						If CustAROver(rs("custAcctNum"),60) <= 0 Then 
						
							If Instr(GetMyFilterRoutes(Session("Userno")),GetCustRouteNum(rs("custAcctNum"))) <> 0 Then ' Only customers matching my routes
								
								Response.Write("<div class='row alert alert-warning'>")%>
					
									<!-- client info !-->
									<div class="col-lg-8 col-md-8 col-sm-8 col-xs-8">
										
										
										<%'Lookup cust info
										SQL = "Select Name,Addr1,Addr2,CityStateZip,Phone,Contact from AR_Customer where CustNum='" & rs("custAcctNum") & "'"
										Set rsCust = cnn8.Execute(SQL)
										If not rsCust.Eof Then
										 	'Response.Write("User Ruotes:" & GetMyFilterRoutes(Session("Userno")) &"<br>")
 										 	'Response.Write("This Ruote:" & GetCustRouteNum(rs("custAcctNum")) &"<br>")
											Response.Write("<p style='font-size:11px'><strong>"& rs("custAcctNum") & " - " & rsCust("Name") & "</strong>")
											
								Response.Write("<ul>")
								Response.Write("<li>" & rsCust("Addr1") & "</li>")
								Response.Write("<li>" & rsCust("Addr2") & "</li>")
								Response.Write("<li>" & rsCust("CityStateZip") & "</li>")
								Response.Write("<li>" & rsCust("Phone") & "</li>")
								Response.Write("<li>" & rsCust("Contact") & "</li>")
								Response.Write("</ul>")

											
											Response.Write("<br>" & rs("Comment1") & "<br></p>")
											
											Response.Write("<p style='font-size:12px'>" & FormatDateTime(rs("nextDate1"))&"</p>")
											
											Response.Write("<p style='font-size:12px'>FILTER CHANGE</p>")
										End If%>
									</div>
								
									<!-- buttons !-->
									<div class="col-lg-4 col-md-4 col-sm-4 col-xs-4">
				
										<form method="post" action="viewfilterchanges.asp" name="frmFilterChange" id="frmFilterChange">
											<input type='hidden' id='txtAssetNumber' name='txtAssetNumber' value='<%=rs("AssetNumber")%>'>		 
											<button type="submit" class="btn btn-primary btn-block btn-spacing" >Details</button>
										</form>
										
									</div>
									<!-- eof buttons !-->
								</div>
								<!-- eof client box !-->	
								
								<hr />
								<%
								End If
							End If	
						End If
				End If
				
				If rs("assetTypeNo") = 336 Then 'This is a PM calls, filter changes are above
			
					If PMCallSubmitted(rs("assetNumber"),rs("nextDate1")) <> True Then 
						
						If CustAROver(rs("custAcctNum"),60) <= 0 Then 
						
							If Instr(GetMyFilterRoutes(Session("Userno")),GetCustRouteNum(rs("custAcctNum"))) <> 0 Then ' Only customers matching my routes
								
								Response.Write("<div class='row alert alert-warning'>")%>
					
									<!-- client info !-->
									<div class="col-lg-8 col-md-8 col-sm-8 col-xs-8">
										
										
										<%'Lookup cust info
										SQL = "Select Name,Addr1,Addr2,CityStateZip,Phone,Contact from AR_Customer where CustNum='" & rs("custAcctNum") & "'"
										Set rsCust = cnn8.Execute(SQL)
										If not rsCust.Eof Then
										 
											Response.Write("<p style='font-size:11px'><strong>"& rs("custAcctNum") & " - " & rsCust("Name") & "</strong>")
											
								Response.Write("<ul>")
								Response.Write("<li>" & rsCust("Addr1") & "</li>")
								Response.Write("<li>" & rsCust("Addr2") & "</li>")
								Response.Write("<li>" & rsCust("CityStateZip") & "</li>")
								Response.Write("<li>" & rsCust("Phone") & "</li>")
								Response.Write("<li>" & rsCust("Contact") & "</li>")
								Response.Write("</ul>")

											
											Response.Write("<br>" & rs("Comment1") & "<br></p>")

											
											Response.Write("<p style='font-size:12px'>" & FormatDateTime(rs("nextDate1"))&"</p>")
											
											Response.Write("<p style='font-size:12px'>PM CALL</p>")
										End If%>
									</div>
								
									<!-- buttons !-->
									<div class="col-lg-4 col-md-4 col-sm-4 col-xs-4">
				
										<form method="post" action="viewPMCall.asp" name="frmPMCall" id="frmPMCall">
											<input type='hidden' id='txtAssetNumber' name='txtAssetNumber' value='<%=rs("AssetNumber")%>'>		 
											<button type="submit" class="btn btn-primary btn-block btn-spacing" >Details</button>
										</form>
										
									</div>
									<!-- eof buttons !-->
								</div>
								<!-- eof client box !-->	
								
								<hr />
								<%
								End If
							End If	
						End If
				End If

				
			rs.movenext
		loop
	Else
		If filterChangeModuleOn() = True and prevMaintModuleOn() <> True Then%>
			No Filter Changes for you
		<% ElseIf filterChangeModuleOn() <> True and prevMaintModuleOn() = True Then %>
			No PM Calls for you
		<% ElseIf filterChangeModuleOn() = True and prevMaintModuleOn() = True Then %>
			No Filter Changes or PM Calls for you
		<%End If
	End IF

cnn8.close
Set rsCust = Nothing
Set rs = Nothing
Set cnn8 = Nothing				
%></div><!--#include file="inc/footer-tech-and-driver.asp"-->
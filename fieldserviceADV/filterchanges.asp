<!--#include file="../inc/header-field-service-mobile.asp"-->
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
Filter Changes
</h1>



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
Else
	FilterChangeDaysFieldService = 15

End If
set rstmp = Nothing
cnntmp.close
set cnntmp = Nothing


SQL = "SELECT * FROM Assets "
SQL = SQL & "INNER JOIN EQ_ScheduledServiceDates ON Assets.assetNumber = EQ_ScheduledServiceDates.assetNumber "
SQL = SQL & "WHERE "
SQL = SQL & "(Assets.assetTypeNo = 335) AND "
SQL = SQL & "(Assets.custAcctNum IN (SELECT CustNum FROM AR_Customer WHERE (AcctStatus = 'A')))  AND "
SQL = SQL & "EQ_ScheduledServiceDates.nextDate1 IS NOT NULL "
SQL = SQL & "AND EQ_ScheduledServiceDates.nextDate1 <= DateAdd(day," & FilterChangeDaysFieldService & ",getdate()) "
SQL = SQL & "ORDER BY EQ_ScheduledServiceDates.nextDate1, custAcctNum "



Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rsCust = Server.CreateObject("ADODB.Recordset")
rsCust.CursorLocation = 3 


Set rs = cnn8.Execute(SQL)
			
	If not rs.EOF Then
		Do While Not rs.EOF
		
			
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
				
			rs.movenext
		loop
	Else%>
		No Filter Changes for you

	<%End IF

cnn8.close
Set rsCust = Nothing
Set rs = Nothing
Set cnn8 = Nothing				
%></div><!--#include file="../inc/footer-field-service-noTimeout.asp"-->
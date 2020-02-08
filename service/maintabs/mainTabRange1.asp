<!--#include virtual="/inc/InsightFuncs.asp"-->
<!--#include virtual="/inc/InsightFuncs_Users.asp"-->
<!--#include virtual="/inc/InsightFuncs_Service.asp"-->
<!--#include virtual="/inc/InSightFuncs_BizIntel.asp"-->
<!--#include virtual="/inc/InsightFuncs_Equipment.asp"-->
<!--#include virtual="/inc/InsightFuncs_AR_AP.asp"-->
<%
dummy = MUV_Write("selectedServiceTab","#0to8")
Set cnn8 = Server.CreateObject("ADODB.Connection")
			cnn8.open (Session("ClientCnnString"))
Set cnnCustInfo = Server.CreateObject("ADODB.Connection")
			cnnCustInfo.open (Session("ClientCnnString"))			
			
			Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
	NumberOfMinutesInServiceDayVar = GetNumberOfMinutesInServiceDay()
%>
<div class="table-responsive">
    <table id="tableSuperSum7" class="food_planner table table-condensed sortable">
      <thead>
        <tr>
          <th class="sorttable_numeric">Date</th>
          <th width="5%">Ticket #</th>	   
          <th class="sorttable_nosort">&nbsp;</th>
          <th width="5%"><%=GetTerm("Customer")%><br>ID</th>
          <th width="15%">Company</th>
          <th class="sorttable_nosort" width="35%"><span id="td-padding">Description</span></th>
          <th>Stage</th>
          <th class="sorttable_numeric">Elapsed<br>Time</th>
          <th>Other Actions</th>
          <th>Submitted Via</th>
        </tr>
      </thead>
      
      <tbody class='searchable'>
      
	<%
	
	SQL = "SELECT * FROM FS_ServiceMemos WHERE CurrentStatus = 'OPEN' AND RecordSubType = 'OPEN' AND FilterChange <> 1 ORDER BY submissionDateTime DESC"

	Set rs = cnn8.Execute(SQL)

	DynamicFormCounter = 700
	
	If not rs.EOF Then
	
		LineX=1
		
		Do While Not rs.EOF

			ShowThisRec = True
			
			Set cnnUserRegionsForServiceBoard = Server.CreateObject("ADODB.Connection")
			cnnUserRegionsForServiceBoard.open (Session("ClientCnnString"))
			Set rsUserRegionsForServiceBoard = Server.CreateObject("ADODB.Recordset")
			rsUserRegionsForServiceBoard.CursorLocation = 3 
			
			SQLUserRegionsForServiceBoard = "SELECT UserRegionsToViewService FROM tblUsers WHERE UserNo = " & Session("UserNo")
			Set rsUserRegionsForServiceBoard = cnnUserRegionsForServiceBoard.Execute(SQLUserRegionsForServiceBoard)
		
			If IsNull(rsUserRegionsForServiceBoard("UserRegionsToViewService")) Then 
				UserRegionList  = ""
			Else
				UserRegionList = rsUserRegionsForServiceBoard("UserRegionsToViewService")
			End If
			
			set rsUserRegionsForServiceBoard = Nothing
			cnnUserRegionsForServiceBoard.close
			set cnnUserRegionsForServiceBoard = Nothing
			
			
			If UserRegionList <> "" Then
			
				CustRegion = GetCustRegionIntRecIDByCustID(rs.Fields("AccountNumber"))
				ShowThisRec = False
				
				RegionArray = Split(UserRegionList,",")
				
				For x = 0 to Ubound(RegionArray)
					If cint(RegionArray(x)) = cint(CustRegion) Then
						ShowThisRec = True
						Exit For
					End IF
				Next
			End If
		
			If ShowThisRec = True AND ServiceCallElapsedMinutesOpenTicket(rs.Fields("MemoNumber")) <= GetNumberOfMinutesInServiceDay() Then
			

				%>
					<!--#include file="mainTabTableDataAllTabsNotClosed.asp"-->
				<%	
				
				
				LineX=LineX+1
			
			End If 'End Elapsed Time Filter
				
				
			rs.movenext
		loop
		
	End If


	
	
    %>
      
      
      
      
      </tbody>
    </table>
  </div>



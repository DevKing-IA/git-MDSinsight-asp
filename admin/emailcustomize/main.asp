<!--#include file="../../inc/header.asp"-->




<style type="text/css">
	 
	body{
	 	overflow-x:hidden;
	}
	.page-header{
	 	margin-top: 0px;
	}
	  	  
	h3{
		 margin: 0px;
		 padding: 0px;
		 line-height: 1;
	}
	
	.ui-widget-header{
		background: #193048;
		border: 1px solid #193048;
	}
	
	.custom-row{
	  	margin-top: 10px;
	}
	
	.modal-link{
		cursor: pointer;
	}
	
	.table-history .table>thead>tr>th {
		vertical-align: bottom;
		border-bottom: 2px solid #ddd;
	}
	
	table.sortable th:not(.sorttable_sorted):not(.sorttable_sorted_reverse):not(.sorttable_nosort):after { 
		content: " \25B4\25BE" 
	}
	a.list-group-item {
		height:auto;
		min-height:100px;
	}
	a.list-group-item.active small {
		color:#fff;
	}
	a.list-group-item:nth-of-type(odd) {
	    background-color: #eee;
	}
	a.list-group-item:nth-of-type(odd):hover {
	    background-color: #ccc;
	}
	
</style>
<script language="javascript">
    function updateCustomOrDefault(id, type) {
        $.ajax({
            type: "POST",
            url: "../../inc/InSightFuncs_AjaxForAdminSettings.asp",
            data: "action=updateCustomOrDefault&id=" + encodeURIComponent(id) + "&type=" + encodeURIComponent(type),
            success: function (msg) {
               
            }
        })
    }
</script>

<br>
<h1 class="page-header"><i class="fa fa-envelope"></i> Customize System Emails</h1>
	
<!-- content starts here !-->
<div class="row">

	<!-- tabs start here !-->
	<div class="global-tabs">

		<ul class="nav nav-tabs responsive-tabs">
		<%
			SQL = "SELECT DISTINCT(emailModule) FROM SC_EmailCustomization GROUP BY emailModule"
			
			Set cnn8 = Server.CreateObject("ADODB.Connection")
			cnn8.open (Session("ClientCnnString"))
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.CursorLocation = 3 
			Set rs = cnn8.Execute(SQL)
			
			If NOT rs.EOF Then
			
				tabCounter = 1
			
				Do While NOT rs.EOF
					emailModule = rs("emailModule")
					
					If tabCounter = 1 Then
						tabCounter = 0
						%><li class="active"><a data-toggle="tab" href="#<%= LCASE(emailModule) %>"><%= UCASE(emailModule) %></a></li><%
					Else
						%><li><a data-toggle="tab" href="#<%= LCASE(emailModule) %>"><%= UCASE(emailModule) %></a></li><%
					End If 
					
					 
					
					rs.MoveNext
				Loop
			End If
		
			set rs = Nothing
			cnn8.close
			set cnn8 = Nothing
		%>		
		</ul>
		
		<!-- begin tab content !-->
		<div class="tab-content">
			
		<%
		SQL = "SELECT DISTINCT(emailModule) FROM SC_EmailCustomization GROUP BY emailModule"
		
		Set cnn8 = Server.CreateObject("ADODB.Connection")
		cnn8.open (Session("ClientCnnString"))
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = 3 
		Set rs = cnn8.Execute(SQL)
		
		If NOT rs.EOF Then
		
			tabCounter = 1
		
			Do While NOT rs.EOF
			
				emailModule = rs("emailModule")
				
				If tabCounter = 1 Then
					tabCounter = 0
					%><div class="tab-pane fade in active" id="<%= LCASE(emailModule) %>"><%
				Else
					%><div class="tab-pane fade" id="<%= LCASE(emailModule) %>"><%
				End If 
				%>
					<div class="container" style="width:100%; margin-top:20px;">
			        	<div class="list-group">
				<%
				
				SQLEmailModule = "SELECT * FROM SC_EmailCustomization WHERE emailModule = '" & emailModule & "' ORDER BY emailName ASC"
				
				Set cnnEmailModule = Server.CreateObject("ADODB.Connection")
				cnnEmailModule.open (Session("ClientCnnString"))
				Set rsEmailModule = Server.CreateObject("ADODB.Recordset")
				rsEmailModule.CursorLocation = 3 
				Set rsEmailModule = cnnEmailModule.Execute(SQLEmailModule)
				
				If NOT rsEmailModule.EOF Then
				
					Do While NOT rsEmailModule.EOF
					
						InternalRecordIdentifier = rsEmailModule("InternalRecordIdentifier")
						RecordCreationDateTime = rsEmailModule("RecordCreationDateTime")
						emailModule = rsEmailModule("emailModule")
						emailName = rsEmailModule("emailName")
						emailDescription = rsEmailModule("emailDescription")
						emailSubjectLine = rsEmailModule("emailSubjectLine")
						emailSubheaderText = rsEmailModule("emailSubheaderText")
						emailBodyCodePart1 = rsEmailModule("emailBodyCodePart1")
						emailBodyCodePart2 = rsEmailModule("emailBodyCodePart2")
						emailBodyCodePart3 = rsEmailModule("emailBodyCodePart3")
						emailAssociatedLink = rsEmailModule("emailAssociatedLink")
						emailAssociatedLinkButtonText = rsEmailModule("emailAssociatedLinkButtonText")
						emailType = rsEmailModule("emailType")
                        emailFileName = rsEmailModule("emailFileName")
                        customOrDefault = rsEmailModule("customOrDefault")
                        If isNull(customOrDefault) Then customOrDefault = "custom"
                        If NOT isNull(emailFileName) Then
                            emailFileNameDefault = Replace(emailFileName, ".txt","Default.txt")
                        Else
                            emailFileNameDefault = ""
                        End If
								
				
						%>
				          <a href="#" class="list-group-item">
				                <div class="col-md-7">
				                    <h4 class="list-group-item-heading"> <%= emailName %> 
				                    <% If emailType = "Internal" Then %>
				                    	<input class="btn btn-success btn-xs" type="button" value="<%= emailType %>">
				                    <% ElseIf emailType = "External" Then %>
				                    	<input class="btn btn-danger btn-xs" type="button" value="<%= emailType %>">
				                    <% ElseIf emailType = "Internal Only" Then %>
				                    	<input class="btn btn-warning btn-xs" type="button" value="<%= emailType %>">
				                    <% End If %>
				                    </h4>
				                    <p class="list-group-item-text"><%= emailDescription %></p>
                                    <% If emailAssociatedLink <> "" Then %>
				                    	<button type="button" class="btn btn-info btn-md" onclick="location.href='<%= emailAssociatedLink %>';"><i class="fa fa-chain"></i> <%= UCASE(emailAssociatedLinkButtonText) %></button>
				                    <% End If %>
				                </div>
                                <div class="col-md-2 text-center">
                                    <div class="row" style="margin-top: 7px;">
                                        <input type="radio" name="customOrDefault<%= InternalRecordIdentifier %>" <% If customOrDefault="custom" OR customOrDefault="" Then Response.Write("checked") End If %> onclick="updateCustomOrDefault(<%= InternalRecordIdentifier %>,'custom')" /> Use Custom Email
                                    </div>
                                    <div class="row" style="margin-top: 7px;">
                                        <input type="radio" name="customOrDefault<%= InternalRecordIdentifier %>" <% If customOrDefault="default" Then Response.Write("checked") End If %> onclick="updateCustomOrDefault(<%= InternalRecordIdentifier %>,'default')" /> Use Default Email
                                    </div>
                                </div>
				                <div class="col-md-3 text-center">
				                	<div class="row" style="margin-top: 1px;">
                                        <% 
                                        Dim isDisabled: isDisabled = ""
                                        If emailFileName<>"" Then 
                                           set fs=Server.CreateObject("Scripting.FileSystemObject")
                                           path = "C:\home\clientfilesV\" & MUV_READ("SERNO") & "\emails\"
                                           If NOT fs.FileExists(path & emailFileName) Then
                                            isDisabled = "disabled"
                                           End If
                                           set fs=Nothing
                                        Else
                                            isDisabled = "disabled"
                                        End If 
                                        %>				      
                                        <div class="col-md-6 pull-left"><button type="button" <%= isDisabled %> class="btn btn-primary btn-md btn-block" onclick="window.open('view_email.asp?email=<%= emailFileName %>');"><i class="fa fa-eye"></i> Preview Custom</button></div>
                                        <div class="col-md-6 pull-left"><button type="button" class="btn btn-primary btn-md btn-block" onclick="window.open('view_email.asp?type=default&email=<%= emailFileNameDefault %>');"><i class="fa fa-eye"></i> Preview Default</button></div>
				                    </div>
				                	<div class="row" style="margin-top: 7px;">
				                    	<div class="col-md-6 pull-left"><button type="button" class="btn btn-success btn-md btn-block" onclick="location.href='edit_email.asp?email=<%= emailFileName %>';"><i class="fa fa-pencil"></i> Edit Custom</button></div>
				                    	<!--div class="col-md-6 pull-left"><button type="button" class="btn btn-success btn-md btn-block" onclick="location.href='edit_email.asp?type=default&email=<%= emailFileNameDefault %>';"><i class="fa fa-pencil"></i> Load Default</button></div-->
				                    </div>
				                    
				                </div>
				          </a>      
						<%
					
						rsEmailModule.MoveNext
						Loop
						
					End If
					
						
					set rsEmailModule = Nothing
					cnnEmailModule.close
					set cnnEmailModule = Nothing
						
					%>					          
						
						</div><!-- eof container-->
		  			</div><!-- eof list-group-->	
				</div><!-- eof single tab !-->
				<%
						
					rs.MoveNext
				Loop
			End If
		
			set rs = Nothing
			cnn8.close
			set cnn8 = Nothing
			%>		
		
		</div><!-- eof tab-content (all tabs) -->
		
		
	</div><!-- eof global-tabs -->	
</div><!-- eof row -->
 
<!--#include file="../../inc/footer-main.asp"-->

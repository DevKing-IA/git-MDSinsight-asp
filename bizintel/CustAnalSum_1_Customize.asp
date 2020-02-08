<%
'Defaults
FilterSalesDollars = 100
FilterPercentage = 10

'************************
'Read Settings_Reports
'************************
SQL = "SELECT * from Settings_Reports where ReportNumber = 2100 AND UserNo = " & Session("userNo")
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs= cnn8.Execute(SQL)
UseSettings_Reports = False
If NOT rs.EOF Then
	UseSettings_Reports = True
	FilterSlsmn1 = rs("ReportSpecificData1")
	FilterSlsmn2 = rs("ReportSpecificData2")
	FilterReferral = rs("ReportSpecificData3")
	If FilterSlsmn1 <> "All" Then FilterSlsmn1 = CInt(FilterSlsmn1)
	If FilterSlsmn2 <> "All" Then FilterSlsmn2 = CInt(FilterSlsmn2)
	If FilterReferral <> "All" Then FilterReferral = CInt(FilterReferral)
	FilterSalesDollars = rs("ReportSpecificData5")
	FilterPercentage = rs("ReportSpecificData6")
	If FilterSalesDollars = "" Then FilterSalesDollars = 100
	If FilterPercentage = "" Then FilterPercentage = 10
End If


'****************************
'End Read Settings_Reports
'****************************
%>
<style type="text/css">
	 .ativa-scroll{
	 max-height: 360px
 }
</style>

<!-- modal scroll !-->
<script type="text/javascript">
	$(document).ready(ajustamodal);
	$(window).resize(ajustamodal);
	function ajustamodal() {
	var altura = $(window).height() - 200; //value corresponding to the modal heading + footer
	$(".ativa-scroll").css({"height":altura,"overflow-y":"auto"});
	}
</script>
<!-- eof modal scroll !-->
	
<div class="modal fade bs-example-modal-lg-customize" tabindex="-1" role="dialog" aria-labelledby="myLargeModalLabel" aria-hidden="true">
	<div class="modal-dialog modal-lg modal-height">
		<div class="modal-content">
			<div class="modal-header">
				<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
				<h4 class="modal-title" id="myModalLabel" align="center">Customize Customer Analysis Summary 1</h4>
			</div>

		<form method="post" action="CustAnalSum_1_Customize_SaveValues.asp" name="frmCustomerLeakageSummary_Customize">

			<div class="modal-body ativa-scroll">
 	      	
	 	      	<!-- filtering !-->
		      	<div class="container-fluid">
			      	<div class="row">
 		      	
				      	<!-- left column !-->
				      	<div class="col-lg-2 col-md-3 col-sm-12 col-xs-12 left-column">
			 		      	<h4><br>Filtering</h4>
				      	</div>
				      	<!-- eof left column !-->
 		      	
		 		      	<!-- right column !-->
		 		      	<div class="col-lg-10 col-md-9 col-sm-12 col-xs-12 right-column">
	 		      	
				      	<!-- row !-->
				      	<div class="row">
		     	
				      	<div class="col-lg-3 col-md-3 col-sm-12 col-xs-12">
					      	<strong>Slsmn 1</strong>
				      	</div>

		      		<div class="col-lg-3 col-md-3 col-sm-12 col-xs-12">
				      	<select class="form-control" name="selFilterSlsmn1">
						<% If UseSettings_Reports = False OR (UseSettings_Reports = True AND FilterSlsmn1="All") Then %>
					      	<option selected value="All">All</option>
					    <% Else %>
				      	  	<option value="All">All</option>
					    <% End IF %>  	
				      	<% 'Get all Slsmn 1 options
				      	  	SQL = "SELECT DISTINCT SalesmanSequence, Salesman.Name FROM Salesman "
				      	  	SQL = SQL & "Inner Join AR_Customer on Salesman = SalesmanSequence "
				      	  	SQL = SQL & "order by SalesmanSequence "
		
							Set cnn8 = Server.CreateObject("ADODB.Connection")
							cnn8.open (Session("ClientCnnString"))
							Set rs = Server.CreateObject("ADODB.Recordset")
							rs.CursorLocation = 3 
							Set rs = cnn8.Execute(SQL)
								
							If not rs.EOF Then
								Do
									Response.Write("<option ")
									If UseSettings_Reports = True Then
									 	If FilterSlsmn1 <> "All" Then
											If FilterSlsmn1 = rs("SalesmanSequence") Then Response.Write("selected ")
										End If
									End If
									Response.Write("value='" & rs("SalesmanSequence") & "'>" & rs("SalesmanSequence") & " - " & rs("Name") & "</option>")
									rs.movenext
								Loop until rs.eof
							End If
							set rs = Nothing
							cnn8.close
							set cnn8 = Nothing
				      	%>
						</select>
			      	</div>
		      	</div>
		      	<!-- eof row !-->
		      	
		      	<!-- row !-->
		      	<div class="row">
			      	<div class="col-lg-3 col-md-3 col-sm-12 col-xs-12">
			      		<strong>Slsmn 2</strong>
		     	 	</div>
		      	
		      		<div class="col-lg-3 col-md-3 col-sm-12 col-xs-12">
				      	<select class="form-control" name="selFilterSlsmn2">
						<% If UseSettings_Reports = False OR (UseSettings_Reports = True AND FilterSlsmn2="All") Then %>
					      	<option selected value="All">All</option>
					    <% Else %>
				      	  	<option value="All">All</option>
					    <% End IF %> 
				      	<% 'Get all Slsmn 2 options
				      	  	SQL = "SELECT DISTINCT SalesmanSequence, Salesman.Name FROM Salesman "
				      	  	SQL = SQL & "Inner Join AR_Customer on SecondarySalesman = SalesmanSequence "
				      	  	SQL = SQL & "order by SalesmanSequence "
	
		
							Set cnn8 = Server.CreateObject("ADODB.Connection")
							cnn8.open (Session("ClientCnnString"))
							Set rs = Server.CreateObject("ADODB.Recordset")
							rs.CursorLocation = 3 
							Set rs = cnn8.Execute(SQL)
								
							If not rs.EOF Then
								Do
									Response.Write("<option ")
									If UseSettings_Reports = True Then
									 	If FilterSlsmn2 <> "All" Then
											If FilterSlsmn2 = rs("SalesmanSequence") Then Response.Write("selected ")
										End If
									End If
									Response.Write("value='" & rs("SalesmanSequence") & "'>" & rs("SalesmanSequence") & " - " & rs("Name") & "</option>")
									rs.movenext
								Loop until rs.eof
							End If
							set rs = Nothing
							cnn8.close
							set cnn8 = Nothing
				      	%>
					    </select>
		      		</div>
		      	</div>
		      	<!-- eof row !-->
		      	
		      	<!-- row !-->
		      	<div class="row">
			      	<div class="col-lg-3 col-md-3 col-sm-12 col-xs-12">
				      	<strong>Referral</strong>
			      	</div>
		      	
		      		<div class="col-lg-3 col-md-3 col-sm-12 col-xs-12">
				      	<select class="form-control" name="selFilterReferral">
						<% If UseSettings_Reports = False OR (UseSettings_Reports = True AND FilterReferral="All") Then %>
					      	<option selected value="All">All</option>
					    <% Else %>
				      	  	<option value="All">All</option>
					    <% End IF %> 
				      	<% 'Get all Referral options
				      	  	SQL = "SELECT ReferalCode, Name FROM Referal"
		
							Set cnn8 = Server.CreateObject("ADODB.Connection")
							cnn8.open (Session("ClientCnnString"))
							Set rs = Server.CreateObject("ADODB.Recordset")
							rs.CursorLocation = 3 
							Set rs = cnn8.Execute(SQL)
								
							If not rs.EOF Then
								Do
									Response.Write("<option ")
									If UseSettings_Reports = True Then
									 	If FilterReferral <> "All" Then
											If FilterReferral = rs("ReferalCode") Then Response.Write("selected ")
										End If
									End If
									Response.Write("value='" & rs("ReferalCode") & "'>" & rs("ReferalCode") & " - " & rs("Name") & "</option>")
									rs.movenext
								Loop until rs.eof
							End If
							set rs = Nothing
							cnn8.close
							set cnn8 = Nothing
				      	%>
					    </select>
		      		</div>
		      	</div>
		      	<!-- eof row !-->
		      		      	
	      	</div>
	      	<!-- eof right column !-->
      	</div>
   	</div>

   	<div class="col-lg-12">
	   	<hr />
   	</div>

   	<div class="container-fluid">
	   	<div class="row">
 		      	
 		      	<!-- left column !-->
 		      	<div class="col-lg-2 col-md-3 col-sm-12 col-xs-12 left-column">
	 		      	<h4><br>Thresholds</h4>
 		      	</div>
 		      	<!-- eof left column !-->

				<!-- Threshold !-->
 		      	<div class="col-lg-10 col-md-9 col-sm-12 col-xs-12 right-column">
			      	<!-- row !-->
			      	<div class="row">
					    <div class="col-lg-8"><strong>Last closed period sales dollars is less than the prior three period average sales dollars by at least</strong></div>
			         	<div class="col-lg-3">
				         	<select class="form-control" name="selDollars">
					         	<% For x = 0 to 5000 Step 100
	   					         	Response.Write("<option value = " & x)
									If UseSettings_Reports = True Then 
										If cint(FilterSalesDollars) = cint(x) Then 
											Response.Write(" selected >$" & x & "&nbsp;</option>")
										Else
											Response.Write(">$" & x & "&nbsp;</option>")
										End If
									Else
										If UseSettings_Reports <> True AND x = 100 Then
			   					        	Response.Write(" selected >$" & x & "&nbsp;(default)</option>")
			   					        Else
			   					         	Response.Write(" >$" & x & "</option>")
			   					        End If
									End If			   					       
					         	Next %>
			         		</select></div>
			         	</div>
			         	<!-- eof row with data !-->
			         	
			         	<!-- row with data !-->
		         		<div class="row row-data">
							<div class="col-lg-8"><strong>The difference between the last closed period sales vs the prior three preriods average sales represents at least</strong></div>
				         	<div class="col-lg-3">
					         	<select class="form-control" name="selPercent">
					         	<% For x = 0 to 100
	   					         	Response.Write("<option value = " & x)
									If UseSettings_Reports = True Then
										If cint(FilterPercentage) = cint(x) Then
											Response.Write(" selected >" & x & "%&nbsp;</option>")
										Else
											Response.Write(">" & x & "%&nbsp;</option>")
										End If
									Else
										If UseSettings_Reports <> True AND  x = 10 Then
			   					        	Response.Write(" selected >" & x & "%&nbsp;(default)</option>")
			   					        Else
			   					         	Response.Write(" >" & x & "%</option>")
			   					        End If
		   					        End If
					         	Next %>
				         	</select>
			         	</div>
			        </div>
	         	</div>
	      	</div>
      	</div>
	</div>
      
	<div class="modal-footer">
		<button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
		<a href="#" onClick="document.frmCustomerLeakageSummary_Customize.submit()"><button type="button" class="btn btn-primary">Run Report</button></a>     
	</div>
</form>

</div>
</div>
</div>


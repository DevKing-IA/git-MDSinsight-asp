<%'******************
' **** Notes Tab ****
'********************
%>
<input type="hidden" name="txtProspectID" id="txtProspectID" value="<%= InternalRecordIdentifier %>">				
<script language="javascript">

	$(document).ready(function() {
		
		//we need to reload the page when the full email modal is closed because the hash tag
		//it leaves in the URL makes the page throw jQuery errors
	
	    $("[id^='myEmailModal']").on('hidden.bs.modal', function () {
	    	$(this).removeData('bs.modal');
	    	prospectID = $("#txtProspectID").val();
	        window.location.href = "viewProspectDetail.asp?id=" + prospectID;
	    });
	});
		
	
</script>
							
<div role="tabpanel" class="tab-pane fade in active" id="log">

		 <div class="row">
    	<div class="col-md-8 col-sm-12">
		<div class="input-group narrow-results"> <span class="input-group-addon">Narrow Results</span>
		    <input id="filter-notes" type="text" class="form-control filter-search-width" placeholder="Type here...">
		</div>
        </div>
        
        <div class="col-md-4 col-sm-12">
        <table align="right" style="width:100%;" border="0">
        	<tr>
            	<td><label><input type="radio" name="logtyperadiofilter" value="0" checked="checked">Show All</label></td>
                <td><label><input type="radio" name="logtyperadiofilter" value="1">Stage Changes Only</label></td>
                <td><label><input type="radio" name="logtyperadiofilter" value="2">Email Only</label></td>
            </tr>
            <tr>
            	<td><label><input type="radio" name="logtyperadiofilter" value="3">Notes Only</label></td>
                <td><label><input type="radio" name="logtyperadiofilter" value="4">Activities Only</label></td>
                <td>&nbsp;</td>
            </tr>
        </table>
        </div>
        
     </div> <!-- end of row-->
	  
		<div class="table-responsive">
            <table    class="table table-striped table-condensed table-hover table-bordered sortable standard-font">
              <thead  >
                <tr>
                  <th width="8%">Log Type</th>
                  <th width="5%">Date</th>
				  <th width="5%">Time</th>
				  <th width="8%">Entered By</th>
				  <th>Details</th>
                  <th class="sorttable_nosort text-center" style="width: 80px;">Actions</th>
                </tr>
              </thead>

			<tbody id="ajaxContainerLogNotes" class='searchable-notes ajax-loading'></tbody>
		</table>
	</div>
</div>
<%'**********************
' **** eof Notes Tab ****
'************************
%>

<script>
	$(document).ready(function () { ajaxLoadLogNotes(); });
	function ajaxRowNewLogNotes() {
		var value = {};
		value.id = 0;
		value.LogDetailType = "-";
		value.Date = "-";
		value.Time = "-";
		value.User = "-";
		value.LogNote = "";
		$('#ajaxRowLogNotes-' + 0 + '').remove();		
		$("#ajaxContainerLogNotes").prepend(ajaxRowHtmlNotes(value));
	}
	function ajaxRowHtmlNotes(value) {
		//If Log note type is note
		if (value.LogDetailType == '-') {
			var html = '\
				<tr id="ajaxRowLogNotes-' + value.id + '" class="' + (value.id==0?'ajaxRowEdit':'ajaxRowView') + ' note">\
					<td><i class="fa fa-sticky-note" aria-hidden="true"></i>&nbsp;Note</td>\
					<td>' + value.Date + '</td>\
					<td>' + value.Time + '</td>\
					<td>' + value.User + '</td>\
					<td>\
						<div class="visibleRowView">' + value.LogNote + '</div>\
						<div class="visibleRowEdit">' + value.LogNote + '</div>\
					</td>\
					<td class="text-center">&nbsp;</td>\
			   </tr>\
				';
		}
		
		else if (value.LogDetailType == 'Note') {
			var html = '\
				<tr id="ajaxRowLogNotes-' + value.id + '" class="' + (value.id==0?'ajaxRowEdit':'ajaxRowView') + ' note">\
					<td><i class="fa fa-sticky-note" aria-hidden="true"></i>&nbsp;' + value.LogDetailType + '</td>\
					<td>' + value.Date + '</td>\
					<td>' + value.Time + '</td>\
					<td>' + value.User + '</td>\
					<td>\
						<div class="visibleRowView">' + value.LogNote + '</div>\
						<div class="visibleRowEdit">' + value.LogNote + '</div>\
					</td>\
					<td class="text-center">&nbsp;</td>\
			   </tr>\
				';
		}

		//If Log note type is Email
		else if (value.LogDetailType == 'Email') {
			var html = '\
				<tr id="ajaxRowLogNotes-' + value.id + '" class="' + (value.id==0?'ajaxRowEdit':'ajaxRowView') + ' email">\
					<td><i class="fa fa-envelope" aria-hidden="true"></i>&nbsp;' + value.LogDetailType + '</td>\
					<td>' + value.Date + '</td>\
					<td>' + value.Time + '</td>\
					<td>' + value.User + '</td>\
					<td>\
						<div class="visibleRowView">' + value.LogNote + '</div>\
						<div class="visibleRowEdit"' + value.LogNote + '</div>\
					</td>\
					<td class="text-center"><a class="btn btn-primary btn-sm" data-toggle="modal" data-show="true" data-target="#myEmailModal' + value.id + '" href="viewProspect_displayFullEmailModal.asp?i=' + value.id + '"><i class="fa fa-envelope-open" aria-hidden="true"></i></a></td>\
			   </tr>\
				';
		}
		//If Log note type is Stage Change
		else if (value.LogDetailType == 'Stage Change') {
			var html = '\
				<tr id="ajaxRowLogNotes-' + value.id + '" class="' + (value.id==0?'ajaxRowEdit':'ajaxRowView') + ' stagechange">\
					<td><i class="fa fa-tasks" aria-hidden="true"></i>&nbsp;' + value.LogDetailType + '</td>\
					<td>' + value.Date + '</td>\
					<td>' + value.Time + '</td>\
					<td>' + value.User + '</td>\
					<td>\
						<div class="visibleRowView">' + value.LogNote + '</div>\
						<div class="visibleRowEdit"' + value.LogNote + '</div>\
					</td>\
					<td class="text-center">&nbsp;</td>\
			   </tr>\
				';
		}
			
		//else log note type is activity 
		else {
			var html = '\
				<tr id="ajaxRowLogNotes-' + value.id + '" class="' + (value.id==0?'ajaxRowEdit':'ajaxRowView') + ' activity">\
					<td><i class="fa fa-check-square" aria-hidden="true"></i>&nbsp;' + value.LogDetailType + '</td>\
					<td>' + value.Date + '</td>\
					<td>' + value.Time + '</td>\
					<td>' + value.User + '</td>\
					<td>\
						<div class="visibleRowView">' + value.LogNote + '</div>\
						<div class="visibleRowEdit"' + value.LogNote + '</div>\
					</td>\
					<td class="text-center">&nbsp;</td>\
			   </tr>\
				';
		}			
		
		return html;
	}
	function ajaxLoadLogNotes(updateAction, updateActionId, updateLogNoteType) {

		$("#ajaxContainerLogNotes").addClass("ajax-loading");
		var url = "ajax/pr_log.asp?i=<%= InternalRecordIdentifier %>";
		var jsondata = {};
		jsondata.updateAction = updateAction;
		jsondata.updateActionId = updateActionId;
		jsondata.updateLogNoteType = updateLogNoteType;
				
		$.ajax({
			type: "POST",
			url: url,
			dataType: "json",
			data: jsondata,
			success: function (data) {				
				var html = "";
				$.each(data, function (key, value) {
					html += ajaxRowHtmlNotes(value);
				});
				$("#ajaxContainerLogNotes").html(html);
				setTimeout(function(){
					$("#ajaxContainerLogNotes").removeClass("ajax-loading");
				}, 0);
				
			}
		});
	}
</script>


<%

SQLEmailLog = "SELECT * FROM PR_ProspectEmailLog WHERE ProspectRecID='"& InternalRecordIdentifier &"'"
Set cnnEmailLog = Server.CreateObject("ADODB.Connection")
cnnEmailLog.open (Session("ClientCnnString"))
Set rsEmailLog = Server.CreateObject("ADODB.Recordset")
rsEmailLog.CursorLocation = 3 
Set rsEmailLog = cnnEmailLog.Execute(SQLEmailLog)

If not rsEmailLog.EOF Then

	Do While Not rsEmailLog.EOF

%>

	<!-- modal  starts here !-->
	 <!-- Modal -->
	<div class="modal fade" id="myEmailModal<%= rsEmailLog("InternalRecordIdentifier") %>" tabindex="-1" role="dialog" aria-labelledby="myEmailModalLabel<%= rsEmailLog("InternalRecordIdentifier") %>" aria-hidden="true">
	    <div class="modal-dialog">
	        <div class="modal-content">
	            <div class="modal-body"></div>
	        </div>
	        <!-- /.modal-content -->
	    </div>
	    <!-- /.modal-dialog -->
	</div>
	<!-- /.modal -->
	<!-- modal  ends here !-->
<%							
		rsEmailLog.MoveNext						
	Loop
End If

Set rsEmailLog = Nothing
cnnEmailLog.Close
Set cnnEmailLog = Nothing


%> 


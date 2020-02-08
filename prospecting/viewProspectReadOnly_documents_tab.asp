<div role="tabpanel" class="tab-pane fade in" id="documents">


		<div class="input-group narrow-results"> <span class="input-group-addon">Narrow Results</span>
		    <input id="filter-documents" type="text" class="form-control filter-search-width" placeholder="Type here...">
		</div>
	  
		<div class="table-responsive">
            <table class="table table-striped table-condensed table-hover table-bordered sortable standard-font">
              <thead>
                <tr>
                  <th width="5%">&nbsp;</th>
                  <th width="5%">Date</th>
				  <th width="5%">Time</th>
				  <th width="8%">Attached By</th>
				  <th>Notes</th>
				  <th>File Attachment</th>
                </tr>
              </thead>
			<tbody id="ajaxContainerDocumentAttachment" class='searchable-documents ajax-loading'></tbody>
		</table>
	</div>
</div>
<%'**********************
' **** eof Notes Tab ****
'************************
%>

<script>
	$(document).ready(function () { ajaxLoadDocumentAttachment(); });
	function ajaxRowNewDocumentAttachment() {
		var value = {};
		value.id = 0;
		value.Date = "-";
		value.Time = "-";
		value.User = "-";
		value.DocumentNotes = "";
		value.DocumentAttachment = "";
		value.DocumentPath = "";
		value.DocumentExt = "";
		$('#ajaxRowDocumentAttachment-' + 0 + '').remove();		
		$("#ajaxContainerDocumentAttachment").prepend(ajaxRowHtmlDocuments(value));
	}
	function ajaxRowHtmlDocuments(value) {

		if (value.DocumentExt !== "") {				
			var html = '\
				<tr id="ajaxRowDocumentAttachment-' + value.id + '" class="' + (value.id==0?'ajaxRowEdit':'ajaxRowView') + '">\
					<td><img src="../../img/file-icons/' + value.DocumentExt + '.png" class="fileicon">&nbsp;' + value.DocumentExt + '</td>\
					<td>' + value.Date + '</td>\
					<td>' + value.Time + '</td>\
					<td>' + value.User + '</td>\
					<td>\
						<div class="visibleRowView">' + value.DocumentNotes + '</div>\
						<div class="visibleRowEdit"><input class="form-control" data-type="DocumentNotes" value="' + value.DocumentNotes + '" /></div>\
					</td>\
					<td>\
						<div class="visibleRowView"><a href="' + value.DocumentPath + '" target="_blank">' + value.DocumentAttachment + '</a></div>\
						<div class="visibleRowEdit"><input type="file" class="form-control" data-type="DocumentAttachment" value="' + value.DocumentAttachment + '" /></div>\
					</td>\
			   </tr>\
				';	
			return html;
		}
		else {
			var html = '\
			<tr id="ajaxRowDocumentAttachment-' + value.id + '" class="' + (value.id==0?'ajaxRowEdit':'ajaxRowView') + ' note">\
				<td><i class="fa fa-sticky-note" aria-hidden="true"></i>&nbsp;Note</td>\
				<td>' + value.Date + '</td>\
				<td>' + value.Time + '</td>\
				<td>' + value.User + '</td>\
				<td>\
					<div class="visibleRowView">' + value.DocumentNotes + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="DocumentNotes" value="' + value.DocumentNotes + '" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView"><a href="' + value.DocumentPath + '" target="_blank">' + value.DocumentAttachment + '</a></div>\
					<div class="visibleRowEdit"><input type="file" class="form-control" data-type="DocumentAttachment" value="' + value.DocumentAttachment + '" /></div>\
				</td>\
		   </tr>\
			';	
		return html;
		}
	}
	function ajaxLoadDocumentAttachment(updateAction, updateActionId) {

		
		$("#ajaxContainerDocumentAttachment").addClass("ajax-loading");
		var url = "ajax/pr_documents.asp?i=<%= InternalRecordIdentifier %>";
		var jsondata = {};
		jsondata.updateAction = updateAction;
		jsondata.updateActionId = updateActionId;
		$.ajax({
			type: "POST",
			url: url,
			dataType: "json",
			data: jsondata,
			success: function (data) {		
				var html = "";
				$.each(data, function (key, value) {
					html += ajaxRowHtmlDocuments(value);
				});
				$("#ajaxContainerDocumentAttachment").html(html);
				setTimeout(function(){
					$("#ajaxContainerDocumentAttachment").removeClass("ajax-loading");
				}, 0);
				
			}
		});
	}
</script>

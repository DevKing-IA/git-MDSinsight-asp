<div role="tabpanel" class="tab-pane fade in" id="links">

	<p><button type="button" class="btn btn-success" onclick="ajaxRowNewLink();">New Link</button></p>

		<div class="input-group narrow-results"> <span class="input-group-addon">Narrow Results</span>
		    <input id="filter-links" type="text" class="form-control filter-search-width" placeholder="Type here...">
		</div>
	  
		<div class="table-responsive">	        
            <table class="table table-striped table-condensed table-hover table-bordered sortable standard-font">
              <thead>
                <tr>
                  <th width="5%">&nbsp;</th>
                  <th width="5%">Date</th>
				  <th width="5%">Time</th>
				  <th width="8%">Added By</th>
				  <th>Link Notes</th>
				  <th>Link URL</th>
  				  <th width="3%">Sticky</th>
                  <th class="sorttable_nosort text-center" style="width: 80px;">Actions</th>
                </tr>
              </thead>
			<tbody id="ajaxContainerLinks" class='searchable-links ajax-loading'></tbody>
		</table>
	</div>
</div>
<%'**********************
' **** eof Notes Tab ****
'************************
%>

<script>
	

	$(document).ready(function () { 
	
		ajaxLoadModelLinks(); 
		
		$("#ajaxContainerLinks").delegate('#LinkURL', 'change', function(){
		
			if (!(is_valid_url($("#LinkURL").val()))) {
			    swal('Please enter a Valid URL');
			    $('#LinkURL').val('');
			}
		    
		});	
	
	});

	function ajaxRowNewLink() {
        var value = {};
        value.id = 0;
        value.Date = "-";
        value.Time = "-";
        value.User = "-";
        value.LinkNote = "";
        value.LinkURL = "";
		value.Sticky = "0";
		$('#ajaxRowModelLink-' + 0 + '').remove();		
		$("#ajaxContainerLinks").prepend(ajaxRowHtmlLinks(value));
	}
		
	function ajaxRowHtmlLinks(value) {
		var btns = '\
					<div class="visibleRowView btn-group btn-group-sm"><a class="btn btn-primary" onclick="ajaxRowMode(\'ModelLink\', ' + value.id + ', \'Edit\');"><i class="fa fa-edit"></i></a><a class="btn btn-danger" onclick="ajaxLoadModelLinks(\'delete\', ' + value.id + ');"><i class="fas fa-trash-alt"></i></a></div>\
					<div class="visibleRowEdit btn-group btn-group-sm"><a class="btn btn-success" onclick="ajaxLoadModelLinks(\'save\', ' + value.id + ');"><i class="fa fa-save"></i></a><a class="btn btn-default" onclick="ajaxRowMode(\'ModelLink\', ' + value.id + ', \'View\');"><i class="fa fa-times"></i></a></div>\
				';
				
		if(value.id==0)
			btns = '\
					<div class="visibleRowEdit btn-group btn-group-sm"><a class="btn btn-success" onclick="ajaxLoadModelLinks(\'insert\', ' + value.id + ');"><i class="fa fa-save"></i></a><a class="btn btn-default" onclick="ajaxRowMode(\'ModelLink\', ' + value.id + ', \'View\');"><i class="fa fa-times"></i></a></div>\
				';
			
        //alert("value.LinkURL : " + value.LinkURL);	
        
		if (value.LinkURL !== "") {				
			var htmlLinks = '\
				<tr id="ajaxRowModelLink-' + value.id + '" class="' + (value.id==0?'ajaxRowEdit':'ajaxRowView') + '">\
					<td><i class="fa fa-chain" aria-hidden="true"></i>&nbsp;Link</td>\
					<td>' + value.Date + '</td>\
					<td>' + value.Time + '</td>\
					<td>' + value.User + '</td>\
					<td>\
						<div class="visibleRowView">' + value.LinkNote + '</div>\
						<div class="visibleRowEdit"><input type="text" class="form-control" data-type="LinkNote" value="' + value.LinkNote + '" /></div>\
					</td>\
					<td>\
						<div class="visibleRowView"><a href="' + value.LinkURL + '" target="_blank">' + value.LinkURL + '</a></div>\
						<div class="visibleRowEdit"><input type="text" class="form-control" data-type="LinkURL" id="LinkURL" value="' + value.LinkURL + '" /></div>\
					</td>\
					<td>\
						<div class="visibleRowView"><a onclick="ajaxLoadModelLinks(\'Sticky-' + (value.Sticky == 1 ? '0' : '1') + '\', ' + value.id + ');" class="label label-' + (value.Sticky == 1 ? 'success' : 'danger') + '">' + (value.Sticky == 1 ? 'Yes' : 'No') + '</a></div>\
						<div class="visibleRowEdit"><input type="checkbox" name="Sticky" data-type="Sticky" ' + (value.Sticky==1?'checked':'') + '></div>\
					</td>\
					<td class="text-center">'+btns+'</td>\
			   </tr>\
				';	
			return htmlLinks;
		}
		else {
			var htmlLinks = '\
			<tr id="ajaxRowModelLink-' + value.id + '" class="' + (value.id==0?'ajaxRowEdit':'ajaxRowView') + ' note">\
				<td><i class="fa fa-chain" aria-hidden="true"></i>&nbsp;Link</td>\
				<td>' + value.Date + '</td>\
				<td>' + value.Time + '</td>\
				<td>' + value.User + '</td>\
				<td>\
					<div class="visibleRowView">' + value.LinkNote + '</div>\
					<div class="visibleRowEdit"><input type="text" class="form-control" data-type="LinkNote" value="' + value.LinkNote + '" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView"><a href="' + value.LinkURL + '" target="_blank">' + value.LinkURL + '</a></div>\
					<div class="visibleRowEdit"><input type="text" class="form-control" data-type="LinkURL" id="LinkURL" value="' + value.LinkURL + '" /></div>\
				</td>\
				<td>\
					<div class="visibleRowView"><a onclick="ajaxLoadModelLinks(\'Sticky-' + (value.Sticky == 1 ? '0' : '1') + '\', ' + value.id + ');" class="label label-' + (value.Sticky == 1 ? 'success' : 'danger') + '">' + (value.Sticky == 1 ? 'Yes' : 'No') + '</a></div>\
					<div class="visibleRowEdit"><input type="checkbox" data-type="Sticky" ' + (value.Sticky==1?'checked':'') + '></div>\
				</td>\
				<td class="text-center">'+btns+'</td>\
		   </tr>\
			';	
		return htmlLinks;
		}
	}
	
	function is_valid_url(url) 
	{
	   return /^(http(s)?:\/\/)?(www\.)?[a-z0-9]+([\-\.]{1}[a-z0-9]+)*\.[a-z]{2,5}(:[0-9]{1,5})?(\/.*)?$/.test(url);
	}	

	function ajaxLoadModelLinks(updateActionLinks, updateActionIdLinks) {
	
		if (updateActionLinks == "delete" && !confirm("Are you sure you want to delete this link?")) return;
		$("#ajaxContainerLinks").addClass("ajax-loading");
		var url = "ajax/eq_links.asp?i=<%= ModelIntRecID %>";
		var jsondata = {};
		jsondata.updateActionLinks = updateActionLinks;
		jsondata.updateActionIdLinks = updateActionIdLinks;
		
		if(updateActionLinks=="save" || updateActionLinks=="insert"){
			jsondata.LinkURL = $('#ajaxRowModelLink-' + updateActionIdLinks + ' [data-type="LinkURL"]').val();
			jsondata.LinkNote = $('#ajaxRowModelLink-' + updateActionIdLinks + ' [data-type="LinkNote"]').val();
			jsondata.Sticky = $('#ajaxRowModelLink-' + updateActionIdLinks + ' [data-type="Sticky"]').is(':checked')?1:0;
		}
		$.ajax({
			type: "POST",
			url: url,
			dataType: "json",
			data: jsondata,
			success: function (data) {
				//if (!data || data+""=="") { window.location.href = window.location.href + ""; return; }				
				var html = "";
				$.each(data, function (key, value) {
					html += ajaxRowHtmlLinks(value);
				});
				$("#ajaxContainerLinks").html(html);
				setTimeout(function(){
					$("#ajaxContainerLinks").removeClass("ajax-loading");
				}, 0);
				getNumberOfLinks();
				
			}
		});
	}

	function getNumberOfLinks() {
	    var url = "ajax/eq_get_links_number.asp?i=<%= ModelIntRecID %>";
	    $.ajax({
	        type: "POST",
	        url: url,
	        success: function (data) {
	            $("#linksNum").html(data);
	        }
	    });
	}


</script>

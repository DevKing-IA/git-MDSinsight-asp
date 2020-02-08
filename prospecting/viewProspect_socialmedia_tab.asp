<%'*************************
' **** Social Media Tab *****
'***************************
%>

<style>
 .decisionMaker {
	font-size:22px;
	font-weight:bold;
	font-color:#00FF00;
}

 .decisionMaker2 {
}

</style>
<div role="tabpanel" class="tab-pane fade in" id="socialmedia">
	  
	<p> <button type="button" class="btn btn-success" onclick="ajaxRowNewSocialMedia();">Add Social Media</button> </p>
	
	<div id="ajaxContainerSocialMediaMain"></div>


	  
	  <div class="table-responsive">
            <table  class="table table-striped table-condensed table-hover table-bordered sortable standard-font">
              <thead>
                <tr>
                  <th width="20%">Social Media Platform</th>
                  <th width="60%">Social Media Link</th>
                  <th class="sorttable_nosort" width="20%">Actions</th>
                </tr>
              </thead>
             
              <tbody id="ajaxContainerSocialMedia" class='searchable-socialmedia ajax-loading'></tbody>

		</table>
		

	</div>
</div>

<script>



	$(document).ready(function () { 
	
			ajaxLoadSocialMedia(); 
			
	});
	
	
	

	
 
	function ajaxRowNewSocialMedia() {
		var value = {};
		value.SocialMediaPlatform = "";
		value.SocialMediaLink = "";	
		value.SocialMediaID = 0;	
		$('#ajaxRowSocialMedia-' + 0 + '').remove();		
		$("#ajaxContainerSocialMedia").prepend(ajaxRowHtmlSocialMedia(value));
	}
	
	
	function ajaxRowHtmlSocialMedia(value) {
		var Platforms = { "Facebook": "Facebook","Twitter": "Twitter","Instagram": "Instagram","Linked In": "Linked In","Blog": "Blog","Youtube": "Youtube"};
		var SocialMediaPlatformSelect = '<select class="form-control" data-type="SocialMediaPlatform">';
		$.each(Platforms, function (key, platform) {
			SocialMediaPlatformSelect+='<option value="'+key+'" ' + (key +""==value.SocialMediaPlatform+""?'selected':'') + '>'+key+'</option>';
		});
		SocialMediaPlatformSelect+='</select>';
		var btns = '\
					<div class="visibleRowView btn-group btn-group-sm"><a class="btn btn-primary" onclick="ajaxRowMode(\'SocialMedia\', ' + value.SocialMediaID + ', \'Edit\');"><i class="fa fa-edit"></i></a><a class="btn btn-danger" onclick="ajaxLoadSocialMedia(\'delete\', ' + value.SocialMediaID + ');"><i class="fas fa-trash-alt"></i></a></div>\
					<div class="visibleRowEdit btn-group btn-group-sm"><a class="btn btn-success" onclick="ajaxLoadSocialMedia(\'save\', ' + value.SocialMediaID + ');"><i class="fa fa-save"></i></a><a class="btn btn-default" id="btnEdit" onclick="ajaxRowMode(\'SocialMedia\', ' + value.SocialMediaID + ', \'View\');"><i class="fa fa-times"></i></a></div>\
				';
		if(value.SocialMediaID==0)
			btns = '\
					<div class="visibleRowEdit btn-group btn-group-sm"><a class="btn btn-success" onclick="ajaxLoadSocialMedia(\'insert\', ' + value.SocialMediaID + ');"><i class="fa fa-save"></i></a><a class="btn btn-default" onclick="ajaxRowMode(\'SocialMedia\', ' + value.SocialMediaID + ', \'View\');"><i class="fa fa-times"></i></a></div>\
				';			
		var htmlSocialMedia = '\
			<tr id="ajaxRowSocialMedia-' + value.SocialMediaID + '" class="' + (value.SocialMediaID==0?'ajaxRowEdit':'ajaxRowView') + '">\
				<td>\
					<div class="visibleRowView"><img src="../img/socialmedia-icons/'+value.SocialMediaPlatform +'.png">&nbsp;' + value.SocialMediaPlatform + '</div>\
					<div class="visibleRowEdit">'+ SocialMediaPlatformSelect +'</div>\
				</td>\
				<td>\
					<div class="visibleRowView">' + value.SocialMediaLink + '</div>\
					<div class="visibleRowEdit"><input class="form-control" data-type="SocialMediaLink" value="' + value.SocialMediaLink.replace(/"/g, '&quot;') + '" /></div>\
				</td>\
				<td class="text-center">'+btns+'</td>\
		   </tr>\
			';
		return htmlSocialMedia;
	}
	function ajaxLoadSocialMedia(updateAction, updateActionId) {
		if (updateAction == "delete" && !confirm("Are your sure you want to delete this social media link?")) return;
		$("#ajaxContainerSocialMedia").addClass("ajax-loading");
		var url = "ajax/pr_socialmedia.asp?i=<%= InternalRecordIdentifier %>";
		var jsondata = {};
		jsondata.updateAction = updateAction;
		jsondata.updateActionId = updateActionId;

		if(updateAction=="save" || updateAction=="insert"){
			//jsondata.SocialMediaID			= $('#ajaxRowSocialMedia-' + updateActionId + ' [data-type="SocialMediaID"]').val();
			jsondata.SocialMediaPlatform	= $('#ajaxRowSocialMedia-' + updateActionId + ' [data-type="SocialMediaPlatform"]').val();
			jsondata.SocialMediaLink		= $('#ajaxRowSocialMedia-' + updateActionId + ' [data-type="SocialMediaLink"]').val();			
			if (jsondata.SocialMediaLink==''){
				alert("Social Media Link can not be empty!");
				$("#ajaxContainerSocialMedia").removeClass("ajax-loading");
				return false;	
			}
		}
		
		$.ajax({
			type: "POST",
			url: url,
			dataType: "json",
			data: jsondata,
			success: function (data) {
				//if (!data || data+""=="") { window.location.href = window.location.href + ""; return; }				
				var htmlSocialMedia = "";
				$.each(data, function (key, value) {
				
					//alert(key + " : " + value.PrimaryCompetitor);			
					htmlSocialMedia += ajaxRowHtmlSocialMedia(value);
				});
				$("#ajaxContainerSocialMedia").html(htmlSocialMedia);
				
				
				
				setTimeout(function(){
					$("#ajaxContainerSocialMedia").removeClass("ajax-loading");
				}, 0);
				
			},
			failure: function (data) {
				$("#ajaxContainerSocialMedia").html("Failed To Load Social Media");
				setTimeout(function(){
					$("#ajaxContainerSocialMedia").removeClass("ajax-loading");
				}, 0);
				
			}
			
		});
	}
</script>

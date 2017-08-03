
<%
Response.Charset="utf-8"
Dim htmlData
htmlData = Request.Form("content")
Function htmlspecialchars(str)
	str = Replace(str, "&", "&amp;")
	str = Replace(str, "<", "&lt;")
	str = Replace(str, ">", "&gt;")
	str = Replace(str, """", "&quot;")
	htmlspecialchars = str
End Function
%>
<link rel="stylesheet" href="../zmbjq/themes/default/default.css" />
	<link rel="stylesheet" href="../zmbjq/plugins/code/prettify.css" />
	<script charset="utf-8" src="../zmbjq/kindeditor-all.js"></script>
	<script charset="utf-8" src="../zmbjq/lang/zh_CN.js"></script>
	<script charset="utf-8" src="../zmbjq/plugins/code/prettify.js"></script>
	<script>
		KindEditor.ready(function(K) {
			var editor = K.create('textarea[name="content"]', {
				cssPath : '../zmbjq/plugins/code/prettify.css',
				uploadJson : '../zmbjq/asp/upload_json.asp',
				fileManagerJson : '../zmbjq/asp/file_manager_json.asp',
				allowFileManager : true,
	
			afterCreate : function() {
					var self = this;
					K.ctrl(document, 13, function() {
						self.sync();
						K('form[name=myform]')[0].submit();
					});
					K.ctrl(self.edit.doc, 13, function() {
						self.sync();
						K('form[name=myform]')[0].submit();
					});
					
				}
			});
			K('input[name=appendHtml]').click(function(e) {
					editor.insertHtml('[zmcmsfy]');
				});
			
			prettyPrint();
			
		});
		
	</script>



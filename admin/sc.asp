<%
Response.Charset="utf-8"%>

<script>
			KindEditor.ready(function(K) {
				var editor = K.editor({
				cssPath : '../zmbjq/plugins/code/prettify.css',
				uploadJson : '../zmbjq/asp/upload_json.asp',
				fileManagerJson : '../zmbjq/asp/file_manager_json.asp',
					allowFileManager : true
				});
				K('#image0').click(function() {
					editor.loadPlugin('image', function() {
						editor.plugin.imageDialog({
							showRemote : true,
							imageUrl : K('#url0').val(),
							clickFn : function(url, title, width, height, border, align) {
							url=url.replace(/.*(UpLoad)/ig, '$1');
								K('#url0').val(url);
								editor.hideDialog();
							}
						});
					});
				});
				
				
				
				
				K('#image1').click(function() {
					editor.loadPlugin('image', function() {
						editor.plugin.imageDialog({
							showRemote : true,
							imageUrl : K('#url1').val(),
							clickFn : function(url, title, width, height, border, align) {
							url=url.replace(/.*(UpLoad)/ig, '$1');
								K('#url1').val(url);
								editor.hideDialog();
							}
						});
					});
				});
				K('#image2').click(function() {
					editor.loadPlugin('image', function() {
						editor.plugin.imageDialog({
							showRemote : true,
							imageUrl : K('#url2').val(),
							clickFn : function(url, title, width, height, border, align) {
							url=url.replace(/.*(UpLoad)/ig, '$1');
								K('#url2').val(url);
								editor.hideDialog();
							}
						});
					});
				});
				K('#image3').click(function() {
					editor.loadPlugin('image', function() {
						editor.plugin.imageDialog({
							showRemote : true,
							imageUrl : K('#url3').val(),
							clickFn : function(url, title, width, height, border, align) {
							url=url.replace(/.*(UpLoad)/ig, '$1');
								K('#url3').val(url);
								editor.hideDialog();
							}
						});
					});
				});
				K('#image4').click(function() {
					editor.loadPlugin('image', function() {
						editor.plugin.imageDialog({
							showRemote : true,
							imageUrl : K('#url4').val(),
							clickFn : function(url, title, width, height, border, align) {
							url=url.replace(/.*(UpLoad)/ig, '$1');
								K('#url4').val(url);
								editor.hideDialog();
							}
						});
					});
				});
				K('#image5').click(function() {
					editor.loadPlugin('image', function() {
						editor.plugin.imageDialog({
							showRemote : true,
							imageUrl : K('#url5').val(),
							clickFn : function(url, title, width, height, border, align) {
							url=url.replace(/.*(UpLoad)/ig, '$1');
								K('#url5').val(url);
								editor.hideDialog();
							}
						});
					});
				});
					K('#image6').click(function() {
					editor.loadPlugin('image', function() {
						editor.plugin.imageDialog({
							showRemote : true,
							imageUrl : K('#url6').val(),
							clickFn : function(url, title, width, height, border, align) {
							url=url.replace(/.*(UpLoad)/ig, '$1');
								K('#url6').val(url);
								editor.hideDialog();
							}
						});
					});
				});
						K('#image7').click(function() {
					editor.loadPlugin('image', function() {
						editor.plugin.imageDialog({
							showRemote : true,
							imageUrl : K('#url7').val(),
							clickFn : function(url, title, width, height, border, align) {
						
								K('#url7').val(url);
								editor.hideDialog();
							}
						});
					});
				});
				
				K('#insertfile').click(function() {
					editor.loadPlugin('insertfile', function() {
						editor.plugin.fileDialog({
							fileUrl : K('#url').val(),
							clickFn : function(url, title) {
								K('#url').val(url);
								editor.hideDialog();
							}
						});
					});
				});
				
				
				

				
				
				
				
				
				
				
				
				
				
				
				
				
					K('#insertfile1').click(function() {
					editor.loadPlugin('insertfile', function() {
						editor.plugin.fileDialog({
							fileUrl : K('#url1').val(),
							clickFn : function(url, title) {
								K('#url1').val(url);
								editor.hideDialog();
							}
						});
					});
				});
				
				
			});
		</script>


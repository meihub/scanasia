<%
  Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload") 
  response.write Request(mySmartUpload.Form("fn"))
	'建立並準備使用上傳檔案元件 
	'允許上傳檔之副檔名及檔案格式限制，此例設定成只能傳gif等圖檔 
	mySmartUpload.AllowedFilesList = "gif,GIF,jpg,JPG,txt,doc" 
	'上傳檔之檔案大小上限，此例設定10k 
	mySmartUpload.MaxFileSize = 1000000000 
	'上傳檔之檔案大小總合之上限，此例設定100k 
	mySmartUpload.TotalMaxFileSize = 10000000000 
	'開始上傳 
	mySmartUpload.Upload 
	'對每個檔案分別動作 
	For each file In mySmartUpload.Files 	
	'存檔動作，目錄路徑要設定你自己的網站位址下的目錄，此例為/robert/ex09/upload/ 
	response.write Server.MapPath("\scanadmin\cb\images\"&file.FileName)&"<br>"
	file.SaveAs(Server.MapPath("\scanadmin\cb\images\"&file.FileName)) 
%>
	檔案已上傳<br>
	<table>
	  <tr><td>檔名：</td><td><%=file.FileName%></td></tr>
	  <tr><td>本機路徑：</td><td><%=file.FilePathName%></td></tr>
	  <tr><td>大小：</td><td><%=file.Size%></td></tr>			  
	</table>
<%	
    next 
%>

<%
  Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload") 
  response.write Request(mySmartUpload.Form("fn"))
	'�إߨ÷ǳƨϥΤW���ɮפ��� 
	'���\�W���ɤ����ɦW���ɮ׮榡����A���ҳ]�w���u���gif������ 
	mySmartUpload.AllowedFilesList = "gif,GIF,jpg,JPG,txt,doc" 
	'�W���ɤ��ɮפj�p�W���A���ҳ]�w10k 
	mySmartUpload.MaxFileSize = 1000000000 
	'�W���ɤ��ɮפj�p�`�X���W���A���ҳ]�w100k 
	mySmartUpload.TotalMaxFileSize = 10000000000 
	'�}�l�W�� 
	mySmartUpload.Upload 
	'��C���ɮפ��O�ʧ@ 
	For each file In mySmartUpload.Files 	
	'�s�ɰʧ@�A�ؿ����|�n�]�w�A�ۤv��������}�U���ؿ��A���Ҭ�/robert/ex09/upload/ 
	response.write Server.MapPath("\scanadmin\cb\images\"&file.FileName)&"<br>"
	file.SaveAs(Server.MapPath("\scanadmin\cb\images\"&file.FileName)) 
%>
	�ɮפw�W��<br>
	<table>
	  <tr><td>�ɦW�G</td><td><%=file.FileName%></td></tr>
	  <tr><td>�������|�G</td><td><%=file.FilePathName%></td></tr>
	  <tr><td>�j�p�G</td><td><%=file.Size%></td></tr>			  
	</table>
<%	
    next 
%>

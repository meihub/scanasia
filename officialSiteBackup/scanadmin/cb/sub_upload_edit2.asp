<!--#include FILE="upload.inc"-->
<%
dim upload,file,formName,formPath,iCount
set upload=new upload_F


'''''����ʱ��Ϊ�ļ��������ظ�������'''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	function MakedownName()
		dim fname
	  	fname = now()
		fname = replace(fname,"-","")
	 	fname = replace(fname," ","") 
		fname = replace(fname,":","")
	  	fname = replace(fname,"PM","")
	  	fname = replace(fname,"AM","")
		fname = replace(fname,"����","")
	  	fname = replace(fname,"����","")
	  	fname = int(fname) + int((10-1+1)*Rnd + 1)
		MakedownName=fname
	end function 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
formPath="../../ProImages/"

iCount=0
for each formName in upload.file ''�г������ϴ��˵��ļ�

 set file=upload.file(formName)  ''����һ���ļ�����
 
  if file.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
  newname=MakedownName()&"."&mid(file.FileName,InStrRev(file.FileName, ".")+1)

  file.SaveAs Server.mappath(formPath&newname)   ''�����ļ�
  
  response.Write"<script>parent.document.myform.works_image2.value='"&newname&"'</script>"
  
  iCount=iCount+1
  
 end if
 
next

  


%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css.css" type="text/css">
</head>

<body>

<br><table width="376" height="112" border="0" align="left" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td width="162" height="56" align="left" bgcolor="#FFFFFF">
	�ļ� <font color=red><%=newname%></font> 
	<br>
	�ϴ��� <font color=red>../../pro/ProImages</font> �ɹ�!	</td>
    <td width="184" rowspan="2" align="center" bgcolor="#FFFFFF">
	<img src="../../ProImages/<%=newname%>" width="205" height="180">	</td>
  </tr>
  <tr align="left">
    <td height="56" bgcolor="#FFFFFF">�ļ���С��
	<%response.write "<font color=red>"&cint(file.FileSize/1024)&"</font>"&" K" %></td>
  </tr>
</table>

<%
set file=nothing
set upload=nothing  ''ɾ���˶���
%>
</body>
</html>
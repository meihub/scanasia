<!--#include FILE="upload.inc"-->
<%
dim upload,file,formName,formPath,iCount
set upload=new upload_F


'''''根据时间为文件命不会重复的名字'''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	function MakedownName()
		dim fname
	  	fname = now()
		fname = replace(fname,"-","")
	 	fname = replace(fname," ","") 
		fname = replace(fname,":","")
	  	fname = replace(fname,"PM","")
	  	fname = replace(fname,"AM","")
		fname = replace(fname,"上午","")
	  	fname = replace(fname,"下午","")
	  	fname = int(fname) + int((10-1+1)*Rnd + 1)
		MakedownName=fname
	end function 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
formPath="../../ProImages/"

iCount=0
for each formName in upload.file ''列出所有上传了的文件

 set file=upload.file(formName)  ''生成一个文件对象
 
  if file.FileSize>0 then         ''如果 FileSize > 0 说明有文件数据
  newname=MakedownName()&"."&mid(file.FileName,InStrRev(file.FileName, ".")+1)

  file.SaveAs Server.mappath(formPath&newname)   ''保存文件
  
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
	文件 <font color=red><%=newname%></font> 
	<br>
	上传到 <font color=red>../../pro/ProImages</font> 成功!	</td>
    <td width="184" rowspan="2" align="center" bgcolor="#FFFFFF">
	<img src="../../ProImages/<%=newname%>" width="205" height="180">	</td>
  </tr>
  <tr align="left">
    <td height="56" bgcolor="#FFFFFF">文件大小：
	<%response.write "<font color=red>"&cint(file.FileSize/1024)&"</font>"&" K" %></td>
  </tr>
</table>

<%
set file=nothing
set upload=nothing  ''删除此对象
%>
</body>
</html>
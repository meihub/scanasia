<!--#include file="../ScriptLibrary/javascript.js"-->
<html>
<head>
<title>添加文件</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css.css" type="text/css">
</head>

<body>
<br>
<table width="376" height="116" border=0 align=left cellpadding=5 cellspacing=1 bgcolor="#CCCCCC">
 
<form action="sub_upload_edit.asp" method="post" enctype="multipart/form-data" name="form0" id="form0" >
<tr>
<td width=241 height=20 align="center" bgcolor="#FFFFFF"> 
<input name="src" type="file" class="wenbenkuang" id="src"  size="20" onChange="play()">

<input name="B1" type="submit" class="go-wenbenkuang" 
onClick="return Check_FileType()" value="上传" IsShowProcessBar="True"></TD>
<td width=110 align="center" bgcolor="#FFFFFF">
<img name="img" width="95" height="50" id="img" style="display:none"/>
<div id="zi">小图预览</div>
<div id="zi0" style="display:none">文件格式不正确</div>
</TD>
</tr>
</form>
</table>

</body>
</html>



<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<!--#include file="../../conn2.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="css.css" rel="stylesheet" type="text/css">
<title>產品二級菜單管理</title>
</head>
<%if session("admin")="" then
response.Redirect "../login.asp"
end if
%>
<% 
dim act,smallid

act=request("act")
smallid=request("smallid")

if act="修改" then
set rs=server.createobject("adodb.recordset")
sql="select * from scc where id="&smallid
rs.open sql,conn,1,3
rs("name")=request("name")
rs.update
rs.close
set rs=nothing
response.write"<script>window.location.href='fenlei2.asp?bigid="&request("bigid")&"&bigclassname="&request("bigclassname")&"'</script>"
response.end
end if


if act="删除" then
set rs=server.createobject("adodb.recordset")
sql="select * from scc where id="&smallid
rs.open sql,conn,1,3
rs.delete
rs.close
set rs=nothing

'-------------删除商品表中该小类对应的商品记录----------------
set rs=server.createobject("adodb.recordset")
sql="select * from c where scc="&smallid
rs.open sql,conn,1,3
if rs.recordcount>0 then
for i=1 to rs.recordcount
if rs.eof or rs.bof then exit for

'---------删除产品对应的图片\详细信息页面\显示图片页面--
productimage=rs("BigImg")
call deletefile(productimage)




'----------删除数据库中的记录------------------------
rs.delete
'----------------------------------------------------


rs.movenext
next
rs.close
set rs=nothing
response.write"<script>window.location.href='fenlei2.asp?bigid="&request("bigid")&"&bigclassname="&request("bigclassname")&"'</script>"
response.end
end if

end if


if act="addsmallclass1" then
set rs=server.createobject("adodb.recordset")
sql="select * from scc "
rs.open sql,conn,1,3
rs.addnew
rs("name")=request("name")
'rs("ename")=request("ename")
rs("ClassID")=request("bigid")

rs.update
rs.close
set rs=nothing
response.write"<script>window.location.href='fenlei2.asp?bigid="&request("bigid")&"&,bigclassname="&request("bigclassname")&"'</script>"
response.end
end if

 %>
<style type="text/css">
<!--
.style1 {color: #FFFFFF}
-->
</style>
<body>
<p>&nbsp;</p>
<p></p>
  <%
dim bigclassname,bigid
bigclassname=request("bigclassname")
bigid=request("bigid")



set rs=server.CreateObject("adodb.recordset")
sql="select * from scc where ClassID="&bigid
rs.open sql,conn,1,1
%>
<link href="../css.css" rel="stylesheet" type="text/css">
<table align="center"><tr>
  <td width="850" height="30" bgcolor="333333">&nbsp;&nbsp;<span class="style1">添加二級菜單</span></td>
</tr></table>
<table width="620" align="center" cellpadding="3" bordercolor="#CCCCCC">
  <tr bgcolor="#999999">
    <td class="xhx"><strong><span class="style1">產品類別<strong>&gt;&gt;</strong></span> <font color=red><%= bigclassname %></font> <span class="style1">包含的小類別</span></strong></td>
	<td align="right"> <strong> <a href="?act=addsmallclass&bigid=<%= bigid %>&bigclassname=<%= bigclassname %>" class="go-wenbenkuang">添加小類 + </a> &nbsp; &nbsp;&nbsp;<a href="fenlei1.asp" class="go-wenbenkuang">返回大類&hArr;</a></strong></td>
  </tr>
</table><br>

<%
if rs.recordcount=0 then
%>



<div align=center>該產品類別沒有對應的二級菜單！</div>
  
<% 
else
%>
<table width="420" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#999999">
  <tr>
    <td width="145" align="center" bgcolor="#e1e1e1"><strong>二級菜單名稱</strong></td>
    <!--td width="146" align="center" bgcolor="#e1e1e1">English Name </td-->
    <td width="105" align="center" bgcolor="#e1e1e1"><strong>操作</strong></td>

  </tr>
<%
while not rs.eof or rs.bof %> 
<form name="form1" action="?bigclassname=<%= bigclassname %>&bigid=<%= bigid %>&smallid=<%= rs("id") %>" method="post">
  <tr>
    <td align="center" bgcolor="#FFFFFF"><input name="name" type="text" class="wenbenkuang" id="name" value="<%=rs("name")%>" size="20" /></td>
    <!--td align="center" bgcolor="#FFFFFF"><input name="ename" type="text" class="wenbenkuang" id="ename" value="<%'=rs("ename")%>" size="20" /></td-->
    <td align="center" bgcolor="#FFFFFF"><input name="act" type="submit" class="go-wenbenkuang" id="act" value="修改" />
      <input name="act" type="submit" class="go-wenbenkuang" id="act" value="刪除" onClick="return confirm('你确定要删除此类别吗？')"/></td>
	
  </tr>
</form>

<%
rs.movenext
wend
end if

rs.close
set rs=nothing

%>
</table>
<br><br><p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p> 
<% if act="addsmallclass" then%>
<br>
<table width="420" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#999999">
  <form name="form2" action="?act=addsmallclass1&bigid=<%= bigid %>&bigclassname=<%= bigclassname %>" method="post">
  <tr>
    <td align="left" bgcolor="#e1e1e1"><span class="xhx">&nbsp;<strong>添加大類<font color=red><%= bigclassname %></font> 的二級菜單 </strong></span></td>
  </tr>
  <tr>
    <td height="40" align="left" bgcolor="#FFFFFF">&nbsp;<strong>二級菜單名稱：</strong>
      <input name="name" type="text" class="wenbenkuang" id="name" size="20" />
      <!--input name="ename" type="text" class="wenbenkuang" id="ename" size="20" /-->
      &nbsp;
      <input name="Submit" type="submit" class="go-wenbenkuang" value="提交" onClick="return check()"/>
	  </td>
  
  </tr>
  </form>
</table><br>
<table width="370" align="center" cellpadding="0">
  <tr>
    <td class="xhx">&nbsp;</td>
  </tr>

   </TABLE>
<p><br>
</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>  
  <%
rss.close
set rss=nothing
   End If %>
</p><A NAME="StranLink"></A>
</body><SCRIPT LANGUAGE=Javascript SRC="../../images/big5style.js"></SCRIPT>
</html>
<SCRIPT LANGUAGE="JavaScript">
<!--
function check()
{
   if(checkspace(document.form2.name.value)) {
	document.form2.name.focus();
    alert("请填写客户名称！");
	return false;
  }

}

function checkspace(checkstr) {
  var str = '';
  for(i = 0; i < checkstr.length; i++) {
    str = str + ' ';
  }
  return (str == checkstr);
}
//-->
</script>
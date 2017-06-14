<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/chao.asp"-->
<%if session("admin")="" then
response.Redirect "../login.asp"
end if
%>
<% 
dim action
action=request("action")

'=================删除大类部分(连同商品记录及图片一起删除)=====
if action="delete" then

'--------------删除大类中的记录------------
set rsbig=server.createobject("adodb.recordset")
sqlbig="select * from cc where ClassID="&request("bigclassid")
rsbig.open sqlbig,MM_chao_STRING,1,3
rsbig.delete
rsbig.close
set rsbig=nothing

'-------重新排列大类顺序-------
set rsor=server.createobject("adodb.recordset")
sqlor="select * from cc order by ClassID"
rsor.open sqlor,MM_chao_STRING,1,3
for i=1 to rsor.recordcount
rsor("ClassID")=i
rsor.update
rsor.movenext
next
rsor.close
set rsor=nothing


'-------------用循环删除大类对应的小类中的记录----
set rssmall=server.createobject("adodb.recordset")
sqlsmall="select * from scc where ClassID="&request("bigclassid")
rssmall.open sqlsmall,MM_chao_STRING,1,3
if rssmall.recordcount>0 then
for i=1 to rssmall.recordcount
if rssmall.eof or rssmall.bof then exit for
rssmall.delete
rssmall.movenext
next
rssmall.close
set rssmall=nothing
end if

'----------------用循环删除产品中属于此大类的产品,如果有图片则一起删除----
set rspro=server.createobject("adodb.recordset")
sqlpro="select * from c where ClassID="&request("bigclassid")
rspro.open sqlpro,MM_chao_STRING,1,3
if rspro.recordcount>0 then
for i=1 to rspro.recordcount
if rspro.eof or rspro.bof then exit for

'---------删除产品对应的图片\详细信息页面\显示图片页面--
productimage=rspro("BigImg")
call deletefile(productimage)


'----------删除数据库中的记录------------------------
rspro.delete
'----------------------------------------------------


rspro.movenext
next
rspro.close
set rspro=nothing
end if

response.Write"<script langue='javascript'>window.location.href='fenlei1.asp'</script>"
response.End

end if
'==============================================================


if action="addbigclass" then
set rs0=server.createobject("adodb.recordset")
sql0="select * from cc"
rs0.open sql0,MM_chao_STRING,1,3
rs0.addnew
rs0("ClassName")=request("name")
'rs0("ename")=request("ename")
rs0("ClassID")=request("idorder")
rs0.update
rs0.close
set rs0=nothing
response.Write"<script langue='javascript'>window.location.href='fenlei1.asp'</script>"
response.End
end if




if action="editbigclassid" then

idorder=cint(request("idorder"))


    
	'==========================如果是上升=========================================================
	if idorder<cint(request("oldorder")) then'注意要比较大小需变为同样数据类型，此处为数字型
    '=============================================================================================
	
	'================大于idorder部分需重新顺序排列========================
	set rs00=server.createobject("adodb.recordset")
	sql00="select * from cc where ClassID<>"&request("bigid")&" order by ClassID"
	rs00.open sql00,MM_chao_STRING,1,3
	
	for i=(idorder+1) to (idorder+rs00.recordcount)
	
	rs00("ClassID")=i
	rs00.update
	
	rs00.movenext
	next
	
	rs00.close
	set rs00=nothing
	'===========小于idorder部分即在idorder上不用修改，因为它的顺序不受idorder的影响===========
	
	set rs2=server.createobject("adodb.recordset")
	sql2="select * from cc where ClassID="&request("bigid")
	rs2.open sql2,MM_chao_STRING,1,3
	rs2("ClassName")=request("name")
	'rs2("ename")=request("ename")
	rs2("ClassID")=idorder
	rs2.update
	rs2.close
	set rs2=nothing
	response.Write"<script langue='javascript'>window.location.href='fenlei1.asp'</script>"
	response.End
	
	
	
	'======================如果不改变排序============================================================
	elseif idorder=cint(request("oldorder")) then
	
	set rs23=server.createobject("adodb.recordset")
	sql23="select * from cc where ClassID="&request("bigid")
	rs23.open sql23,MM_chao_STRING,1,3
	rs23("ClassName")=request("name")
	'rs23("ename")=request("ename")
	rs23.update
	rs23.close
	set rs23=nothing
	response.Write"<script langue='javascript'>window.location.href='fenlei1.asp'</script>"
	response.End
	'================================================================================================
	
	
	
	'==========================如果是下降位置===================================================
	elseif idorder>cint(request("oldorder")) then'注意要比较大小需变为同样数据类型，此处为数字型
	'===========================================================================================
	
	'===============小于idorder且不包括原来记录的部分需重新顺序排列=====
	set rsa=server.createobject("adodb.recordset")
	sqla="select * from cc where ClassID<>"&request("bigid")&" order by ClassID"
	rsa.open sqla,MM_chao_STRING,1,3
	for i=(idorder-rsa.recordcount) to (idorder-1)
	rsa("ClassID")=i
	rsa.update
	rsa.movenext
	next
	rsa.close
	set rsa=nothing
	'==========================大于idorder部分不用修改,因为它的顺序不受idorder的影响====
	set rsr=server.createobject("adodb.recordset")
	sqlr="select * from cc where ClassID="&request("bigid")
	rsr.open sqlr,MM_chao_STRING,1,3
	rsr("ClassName")=request("name")
	'rsr("ename")=request("ename")
	rsr("ClassID")=idorder
	rsr.update
	rsr.close
	set rs20=nothing
    response.Write"<script langue='javascript'>window.location.href='fenlei1.asp'</script>"
	response.End
	
	
	end if



end if

 %>
<style type="text/css">
<!--
.style1 {color: #FFFFFF}
-->
</style>
<link href="../css.css" rel="stylesheet" type="text/css">
<body>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script><link href="../css.css" rel="stylesheet" type="text/css">

<table align="center"><tr>
  <td width="850" height="30" bgcolor="333333">&nbsp;&nbsp;<span class="style1">&#39006;&#21029;&#31649;&#29702;</span></td>
</tr></table>
<table width=600 align="center" cellpadding="5">
<tr bgcolor="#999999">
  <td width="600" height="25" class="xhx"><span class="style1"><strong>&#22823;&#39006;&#21029;&#31649;&#29702;</strong></span> </td>
  <td width="600" align="right"><a href="?action=add" class="go-wenbenkuang"><strong>&#28155;&#21152;&#22823;&#39006;&#21029;</strong> + </a> </td>
</tr></table>

     <br>
<table width="520" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#e1e1e1">
    <td align="center">&#39006;&#21029;&#21517;&#23383; </td>
    <!--td width="105" align="center">English Name </td-->
    <td align="center"  >ClassID</td>
    <td align="center" >&#25805; &#20316;</td>
    <td align="center"  >&#28155;&#21152;&#23567;&#39006;&#21029;</td>
  </tr>
<%
set rs=server.createobject("adodb.recordset")
sql="select * from cc order by ClassID"
rs.open sql,MM_chao_STRING,1,1
if rs.recordcount>0 then
while not rs.eof or rs.bof 
%>


  <tr>
    <td align="center" bgcolor="#FFFFFF"><%= rs("ClassName") %></td>
    <!--td align="center" bgcolor="#FFFFFF"><%'= rs("ename") %></td-->
    <td align="center" bgcolor="#FFFFFF"><%= rs("ClassID") %></td>
    <td align="center" bgcolor="#FFFFFF"><a href="?action=delete&bigclassid=<%= rs("ClassID") %>" 
	onclick="return confirm('你确定要删除所选的大类吗？')">&#21034;&#38500;</a> | <a href="?action=edit&bigclassid=<%= rs("ClassID") %>">&#32232;&#36655;</a> | <a href="TVPphoto-4.asp?ClassID=<%= rs("ClassID") %>">&#35037;&#39166;&#22294;</a></td>
    <td align="center" bgcolor="#FFFFFF">
	
<a href="fenlei2.asp?bigclassname=<%= rs("ClassName") %>&bigid=<%= rs("ClassID") %>">&#28155;&#21152;&#23565;&#25033;&#23567;&#39006;</a>		</td>
  </tr>
  <%
  rs.movenext
  wend
  end if
  rs.close
  set rs=nothing
  %>
</table>
<br>


<% If action="add" Then 
set rss=server.createobject("adodb.recordset")
sqls="select Max(ClassID) as maxidorder from cc"
rss.open sqls,MM_chao_STRING,1,1
%>

<table width="420" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#999999">
  <form name="form11" action="?action=addbigclass" method="post">
  
  <tr>
    <td height="25" colspan="3" align="left" bgcolor="#FFFFFF"><strong>&nbsp;&#28155;&#21152;&#29986;&#21697;&#22823;&#39006;</strong></td>
    </tr>
  <tr bgcolor="#E1E1E1">
    <td width="163" height="22" align="center" bgcolor="#E1E1E1">&#22823;&#39006;&#21517;&#31281;</td>
    <!--td width="163" align="center" bgcolor="#E1E1E1">English Name</td-->
    <td width="78" height="22" align="center" bgcolor="#E1E1E1">ClassID</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td height="35" align="center"><input name="name" type="text" class="wenbenkuang" id="name" size="20" /></td>
    <!--td height="35" align="center"><input name="ename" type="text" class="wenbenkuang" id="ename" size="20" /></td -->
    <td height="35" align="center"><input name="idorder" type="text" class="wenbenkuang" id="idorder" value="<%=rss("maxidorder")+1%>" size="4" /></td>
  </tr>
  <tr>
    <td colspan="3" align="center" bgcolor="#FFFFFF"><input name="Submit" type="submit" class="go-wenbenkuang" value="&#28155;&#21152;" onClick="return check11()"/>
      &nbsp;
      <input name="Submit2" type="reset" class="go-wenbenkuang" value="&#37325;&#32622;" />
        
       </td>
    </tr>
  </form>
</table>
<% 
rss.close
set rss=nothing
End If %>


<% If action="edit" Then
set rs1=server.createobject("adodb.recordset")
sql1="select * from cc where ClassID="&request("bigclassid")
rs1.open sql1,MM_chao_STRING,1,1
 %>

<table width="420" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#999999">
  <form action="?action=editbigclassid&bigid=<%= rs1("ClassID") %>" method="post" name="form2" id="form2">
  
    <tr>
      <td height="25" colspan="3" align="left" bgcolor="#FFFFFF"><span class="xhx"><strong>&nbsp;&#20462;&#25913;&#22823;&#39006;</strong>>><font color=red><%= rs1("ClassName") %></font></span></td>
    </tr>
	<tr bgcolor="#E1E1E1">
      <td width="163" height="22" align="center" bgcolor="#E1E1E1">&#22823;&#39006;&#21517;&#31281;</td>
      <!--td width="164" height="22" align="center" bgcolor="#E1E1E1">English Name</td-->
      <td width="77" height="22" align="center">ClassID</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td height="35" align="center"><input name="name" type="text" class="wenbenkuang" id="name" value="<%=rs1("ClassName")%>" size="20" />
      <input name="oldorder" type="hidden" id="oldorder" value="<%=rs1("ClassID")%>" /></td>
      <!--td align="center"><input name="ename" type="text" class="wenbenkuang" id="ename" value="<%'=rs1("ename")%>" size="20" /></td-->
      <td height="35" align="center"><input name="idorder" type="text" class="wenbenkuang" id="idorder" value="<%=rs1("ClassID")%>" size="4"  readonly="true"/></td>
    </tr>
    <tr>
      <td colspan="3" align="center" bgcolor="#FFFFFF">
        
          <input name="Submit3" type="submit" class="go-wenbenkuang" value="&#20462;&#25913;" onClick="return check2()"/>
        &nbsp;
        <input name="Submit22" type="reset" class="go-wenbenkuang" value="&#37325;&#32622;" />
         
        </td>
    </tr>
  </form>
</table>
<%
rs1.close
set rs1=nothing
 End If %>
<p>&nbsp;</p>
<A NAME="StranLink"></A>
</body><SCRIPT LANGUAGE=Javascript SRC="../../images/big5style.js"></SCRIPT>

</html>
<SCRIPT LANGUAGE="JavaScript">

function check2()
{
   if(checkspace(document.form2.name.value)) {
	document.form2.name.focus();
    alert("请填写大类名称！");
	return false;
  }
  
     if(checkspace(document.form2.idorder.value)) {
	document.form2.idorder.focus();
    alert("请填写大类排序！");
	return false;
  }
  
     if(checknumber(document.form2.idorder.value))
	 {
	 document.form2.idorder.focus();
	 alert("排序只能输入数字！");
	 return false;
	 }

}

////验证输入是否为数字/////////////////////
function checknumber(String) 
{ 
var Letters = "1234567890"; 
var i; 
var c; 
for( i = 0; i < String.length; i ++ ) 
{ 
c = String.charAt( i ); 
if (Letters.indexOf( c ) ==-1) 
{ 
return true; 
} 
} 
return false; 
}
//////////////////////////////// //////




function checkspace(checkstr) {
  var str = '';
  for(i = 0; i < checkstr.length; i++) {
    str = str + ' ';
  }
  return (str == checkstr);
}


function check11()
{
   if(checkspace(document.form11.name.value)) {
	document.form11.name.focus();
    alert("请填写大类名称！");
	return false;
  }

}

</script>
<!-- #include file="conn.asp" -->
<!--#include file="check.asp"-->
<!--#include file="images/NK_SqlIn.Asp"-->
<%
reply_id=trim(request("reply_id"))
if reply_id="" then
response.Write("<script>alert('�Բ��𣡴��ݲ������󣡣�');history.back();</script>"):response.end
else
set rs=server.CreateObject("adodb.recordset")
sql="select * from reply where reply_id=" & reply_id
rs.open sql,conn,1,1
if rs.recordcount=0 then
response.Write("<script>alert('�Բ��𣡴��ݲ������󣡣�');history.back();</script>"):response.end
else
reply_name=rs("reply_name")
add_time=rs("reply_time")
content=rs("reply_content")
rs.close
end if
set rs=nothing
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style1 {color: #FF0000}
-->
</style>
<script language = "JavaScript">
function CheckForm()
{
    if (document.form1.add_time.value=="")
  {
    alert("���ڲ���Ϊ�գ�");
	document.form1.add_time.focus();
	return false;
  }
  if (document.form1.reply_name.value=="")
  {
    alert("���Ʋ���Ϊ�գ�");
	document.form1.reply_name.focus();
	return false;
  }
      if (document.form1.content.value=="")
  {
    alert("���ݲ���Ϊ�գ�");
	document.form1.content.focus();
	return false;
  }
  return true;  
}
</script>
</head>
<body >
<table width="600"  border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
                       <form action="reply_save.asp?action=edit&reply_id=<%=reply_id%>" method="post" onSubmit="return CheckForm();" name=form1 >
                        <tr bgcolor="#F7F7F7">
                          <td width="21%" height="34" align="right" valign="middle" bgcolor="#F7F7F7">nickname��<BR>                          </td>
                          <td width="79%"><input name="reply_name" type="text"  size="20" value="<%=reply_name%>"  ></td>
                        </tr>
                        <tr bgcolor="#F7F7F7">
                          <td width="21%" height="34" align="right" valign="middle" bgcolor="#F7F7F7">���ڣ�<BR>                          </td>
                          <td width="79%"><input name="add_time" type="text"  size="20" value="<%=add_time%>"  >
(eg:2005-12-12)*</td>
                        </tr>
                        <tr bgcolor="#F7F7F7">
                          <td width="21%" height="34" align="right" valign="middle" bgcolor="#F7F7F7">���ݣ�<br>                            <BR>                          </td>
                          <td width="79%"><textarea name="content" cols="70" rows="21" ><%=content%></textarea>
*</td>
                        </tr>
                        <tr align="center" valign="bottom" bgcolor="#F7F7F7">
                          <td height="34" colspan="2"><input name="�ύ" type="submit" class="button"  value="����">
                            ��
                              <input type="reset" class="button" value="������" onClick="javascript:history.back()"></td>
                        </tr></form>
<A NAME="StranLink"></A>
   </TABLE><SCRIPT LANGUAGE=Javascript SRC="../images/big5style.js"></SCRIPT>
</body>
</html>

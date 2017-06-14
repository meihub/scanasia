<!--#include file="../Connections/chao.asp"-->

<html>
<head>
<title>修改产品信息</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../css.css" rel="stylesheet" type="text/css">
</head>
<body>
<br><BR>

<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_chao_STRING
Recordset1.Source = "SELECT * FROM cc"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>
<%
Dim Recordset2
Dim Recordset2_numRows

Set Recordset2 = Server.CreateObject("ADODB.Recordset")
Recordset2.ActiveConnection = MM_chao_STRING
Recordset2.Source = "SELECT * FROM scc"
Recordset2.CursorType = 0
Recordset2.CursorLocation = 2
Recordset2.LockType = 1
Recordset2.Open()

Recordset2_numRows = 0
%>
<table width="650" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
<%
dim ID
ID=request("ID")

set rs= Server.CreateObject("adodb.recordset")
sql="select * from c where ID="&ID
rs.open sql,MM_chao_STRING,1,1
%>
  <tr> 
    <td height="27" align="center" bgcolor="#f1f1f1">[<font color="red">&nbsp;pro Iron Bathtub 产品信息修改  </font>]</td>
  </tr>
  <tr> 
    <td valign="middle" bgcolor="#FFFFFF"> <br>
	
      
        <table width="96%" border="0" align="center" cellpadding="6" cellspacing="1" bgcolor="#999999">
	<form action="save_edit.asp?ID=<%=rs("ID")%>" method="post" name="myform" id="myform">
          <tr>
            <td width="80" height="27" align="right" bgcolor="#FFFFFF">产品名称：</td>
            <td width="400" height="27" bgcolor="#FFFFFF" colspan="2"><input name="name" type="text" class="wenbenkuang" id="name" value="<%=rs("ProName")%>" size="21">
            </td>

          </tr>
		      <tr> 
                    <td class="t2" bgcolor="#FFFFFF"><div align="right">产品分类：</div></td>
                    <td  colspan="2" bgcolor="#FFFFFF"><select name="kind_id0" id="kind_id0">
                        <%
While (NOT Recordset1.EOF)
%>
                        <option value="<%=(Recordset1.Fields.Item("ClassID").Value)%>" <%If (Not isNull((rs.Fields.Item("ClassID").Value))) Then If (CStr(Recordset1.Fields.Item("ClassID").Value) = CStr((rs.Fields.Item("ClassID").Value))) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(Recordset1.Fields.Item("ClassName").Value)%></option>
                        <%
  Recordset1.MoveNext()
Wend
If (Recordset1.CursorType > 0) Then
  Recordset1.MoveFirst
Else
  Recordset1.Requery
End If
%>
                      </select>
					  <select name="kind_id1" id="kind_id1">
                        <%
While (NOT Recordset2.EOF)
%>
                        <option value="<%=(Recordset2.Fields.Item("id").Value)%>" <%If (Not isNull((rs.Fields.Item("scc").Value))) Then If (CStr(Recordset2.Fields.Item("id").Value) = CStr((rs.Fields.Item("scc").Value))) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(Recordset2.Fields.Item("name").Value)%></option>
                        <%
  Recordset2.MoveNext()
Wend
If (Recordset2.CursorType > 0) Then
  Recordset2.MoveFirst
Else
  Recordset2.Requery
End If
%>
                      </select>
					  
					  </td>
                  </tr>
		  
          <tr>
            <td width="80" height="27" align="right" bgcolor="#FFFFFF">原产品小图：</td>
            <td height="27" bgcolor="#FFFFFF">
			<input name="works_image" type="text" class="wenbenkuang" id="works_image" value="<%=rs("ProImg")%>" size="30" readonly="true">
            <input name="works_imageo" type="hidden" id="works_imageo" value="<%=rs("ProImg")%>"></td><td bgcolor="#FFFFFF"><img src="images\photo\<%=rs("ProImg")%>" alt="产品小图" width="95" height="50"></td>
          </tr>
		  
		  <tr>
            <td width="80" height="32" align="right" bgcolor="#FFFFFF">更新小图：</td>
            <td height="32" valign="bottom" bgcolor="#FFFFFF" colspan="2"><iframe src="upload_edit.asp" frameborder="0" scrolling="No" width="400" height="160"></iframe></td>
          </tr>
		       <tr>
            <td width="80" height="27" align="right" bgcolor="#FFFFFF">原产品大图：</td>
            <td height="27" bgcolor="#FFFFFF">
			<input name="works_image2" type="text" class="wenbenkuang" id="works_image2" value="<%=rs("BigImg")%>" size="30" readonly="true">
            <input name="works_imageo2" type="hidden" id="works_imageo2" value="<%=rs("BigImg")%>">
			</td><td bgcolor="#FFFFFF"><img src="images\photo\ProImages\<%=rs("BigImg")%>" alt="产品大图" width="165" height="120"></td>
          </tr>
		  <tr align="center">
            <td height="28" bgcolor="#FFFFFF" align="right">更新大图：</td>
            <td height="28"  align="left" bgcolor="#FFFFFF" colspan="2"><iframe src="upload_edit2.asp" frameborder="0" scrolling="No" width="400" height="160"></iframe></td>
          </tr>
		  <tr align="center">
      <td height="28" bgcolor="#FFFFFF" align="right">&nbsp;&nbsp;产品说明：</td>
      <td height="38" colspan="2" align="left" bgcolor="#FFFFFF">
	  <textarea name="works_info" cols="50" rows="6" id="works_info" style="display:yes"><%=rs("ProInfo")%></textarea>
                         
<!--iframe id="eWebEditor1" src="ewebeditor/ewebeditor.asp?id=works_info&amp;style=s_newssystem" frameborder="0"scrolling="No" width="550" height="200"></iframe-->	  </td>
    </tr>
		  <tr align="center">
            <td height="28" colspan="3" bgcolor="#FFFFFF"><input name="Submit3" type="button" class="go-wenbenkuang" value="返回管理" 
		onClick="window.location.href='TVPphoto-1.asp'">
&nbsp;
<input type="submit" class="go-wenbenkuang" name="Submit" value="修改产品信息" onClick="window.location.href='TVPphoto-1.asp'">
&nbsp;
<input name="Submit2" type="reset" class="go-wenbenkuang" value="重填信息"></td>
          </tr>
	      </form>	
      </table>
        <BR>
  </td>
  </tr>
  <% rs.close
  set rs=nothing
  %>
</table>

<table width="650" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="25" valign="bottom"><font color=red>-- 注意事项 --</font></td>
  </tr>
  <tr>
    <td height="25" valign="bottom">若需要更新产品图片，请在“更新图片”项中通过浏览后上传图片，旧图片将会被删除而被替换为新上传的产品图片！</td>
  </tr>
  <tr>
    <td height="325" valign="bottom">&nbsp;</td>
  </tr>
<A NAME="StranLink"></A>
   </TABLE><SCRIPT LANGUAGE=Javascript SRC="../../images/big5style.js"></SCRIPT>
</body>
</html>

<script language="javascript">
function checkdata()
{
if (document.myform.works_name.value=="")
	{
	  alert("对不起，请输入产品名称！")
	  document.myform.works_name.focus()
	  return false
	 }

}
</scri
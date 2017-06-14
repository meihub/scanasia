<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/chao.asp"-->
<%if session("admin")="" then
response.Redirect "../login.asp"
end if
%>
<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Request.QueryString("ID") <> "") Then 
  Recordset1__MMColParam = Request.QueryString("ID")
End If
%>
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_chao_STRING
Recordset1.Source = "SELECT * FROM c WHERE ID = " + Replace(Recordset1__MMColParam, "'", "''") + ""
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>图片展示</title>

<link href="../css.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" rightmargin="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <!--DWLayoutTable-->
  <tr> 
    <td width="768" height="369" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <!--DWLayoutTable-->
        <tr> 
          <td width="760" height="325" align="center" valign="top"><div align="center">
              <p>&nbsp;</p>
              <p><img src="images\photo\<%=(Recordset1.Fields.Item("BigImg").Value)%>"></p>
            </div></td>
        </tr>
      </table>
      <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
        <!--DWLayoutTable-->
        <tr> 
          <td width="760" height="369" valign="top"> <table width="100%" height="297" border="0">
              <tr> 
                <td height="127" valign="top"> <table width="99%" height="39" border="0" class="chao2">
                    <tr bgcolor="f7f7f7"> 
                      <td width="25%" height="35"> <div align="right"><font color="#FF0000">产品名称 ：</font></div></td>
                      <td width="75%"><%=(Recordset1.Fields.Item("ProName").Value)%></td>
                    </tr>
					  <tr bgcolor="f7f7f7"> 
                      <td width="25%" height="35"> <div align="right"><font color="#FF0000">产品说明 ：</font></div></td>
                      <td width="75%"><%=(Recordset1.Fields.Item("Proinfo").Value)%></td>
                    </tr>
					 <tr> 
    <td height="83" colspan="2"> <div align="center"><a href="#" class="whiteBox" onClick="window.history.back();">回上一页</a></div></td>
  </tr>
                  </table>
                </td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
 
</table>
<p>&nbsp;</p>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>

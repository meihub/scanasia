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
<%
' *** Edit Operations: declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Delete Record: declare variables

if (CStr(Request("MM_delete")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_chao_STRING
  MM_editTable = "c"
  MM_editColumn = "ID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "TVPphoto-1.asp"

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If
  
End If
%>
<%
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql delete statement
  MM_editQuery = "delete from " & MM_editTable & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the delete
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title> 图片管理</title>

<link href="../css.css" rel="stylesheet" type="text/css">
<style type=text/css>
a:link {text-decoration: none}
a:hover {text-decoration:none}
a:visited {text-decoration: none}
</style>
</head>

<body rightmargin="0" leftmargin="0" topmargin="0">
<table width="800" border="0" cellpadding="0" cellspacing="0">
  <!--DWLayoutTable-->
  <tr bgcolor="f7f7f7"> 
    
    <td height="33" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" background="../images2/c_04.jpg">
        <!--DWLayoutTable-->
        <tr> 
          <td width="233" height="33" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <!--DWLayoutTable-->
              <tr> 
                <td width="233" height="33"><table width="88%" border="0" align="center">
                    <tr> 
                      <td>您的位置是： 图片展示 
                     </td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
          <td width="567"><table width="60%" border="0" align="right">
              <tr> 
                <td width="23%"><div align="right"><A HREF="TVPphoto-1.asp">图片管理</A></div></td>
                <td width="24%"><div align="center"><A HREF="TVPphoto-2.asp">添加图片</A></div></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="800" height="410"><table width="98%" height="574" border="0" align="left">
        <tr> 
          <td valign="top"> <div align="center"> 
              <form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1">
                <table width="89%" border="0">
                  <tr> 
                    <td height="26" colspan="3" class="a1">&nbsp;</td>
                  </tr>
                  <tr valign="middle" class="t3"> 
                    <td width="24%" height="17"> <div align="center"></div>
                      <div align="center"></div></td>
                    <td colspan="2"> <div align="center"></div></td>
                  </tr>
                  <tr class="t2"> 
                    <td height="21"> <div align="right">图片名称： </div></td>
                    <td colspan="2" class="t2"> <div align="center"></div>
                      <div align="left"><%=(Recordset1.Fields.Item("ProName").Value)%> </div></td>
                  </tr>
                  <tr> 
                    <td class="t2"><div align="right">录入日期：</div></td>
                    <td colspan="2" class="t2"><%=(Recordset1.Fields.Item("DateTime").Value)%> </td>
                  </tr>
                  <tr> 
                    <td class="t2"><div align="right">上传缩图：</div></td>
                    <td width="20%">&nbsp;<img src="images\photo\<%=(Recordset1.Fields.Item("ProImg").Value)%>" alt="这是显示上传预览图片的位置" name="showImg1" width="127" height="100" id="showImg1" onClick='javascript:alert("这是显示上传预览图片的位置");'> 
                    </td>
                    <td width="56%">&nbsp;</td>
                  </tr>
                  <tr> 
                    <td class="t2"><div align="right"></div></td>
                    <td colspan="2" class="t2">&nbsp;</td>
                  </tr>
                  <tr> 
                    <td class="t2"><div align="right"></div></td>
                    <td colspan="2"><table width="20%" border="0">
                        <tr> 
                          <td> <input type="submit" name="Submit3" value="删除图片"></td>
                          <td><input name="Submit32" type="button" onClick="location='TVPphoto-1.asp'" value="返回上页"></td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr> 
                    <td><div align="center"><span class="t3"> </span></div></td>
                    <td colspan="2">&nbsp;</td>
                  </tr>
                </table>
                <input type="hidden" name="MM_delete" value="form1">
                <input type="hidden" name="MM_recordId" value="<%= Recordset1.Fields.Item("ID").Value %>">
              </form>
              <p class="t2">&nbsp;</p>
            </div></td>
        </tr>
      </table></td>
  </tr>
  <tr> <td height="79" colspan="2" valign="middle" align="center" width="1000" bgcolor="#f7f7f7" > 
</td>
  </tr>
<A NAME="StranLink"></A>
   </TABLE><SCRIPT LANGUAGE=Javascript SRC="../../images/big5style.js"></SCRIPT>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>

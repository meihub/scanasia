<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/chao.asp"-->
<%if session("admin")="" then
response.Redirect "../login.asp"
end if
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
  MM_editTable = "cc"
  MM_editColumn = "ClassID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "TVPfen-1.asp"

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

<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Request.QueryString("ClassID") <> "") Then 
  Recordset1__MMColParam = Request.QueryString("ClassID")
End If
%>
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_chao_STRING
Recordset1.Source = "SELECT * FROM cc WHERE ClassID = " + Replace(Recordset1__MMColParam, "'", "''") + ""
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>删除产品分类</title>
<link href="css/chao.css" rel="stylesheet" type="text/css">
<link href="chao-css/chao.css" rel="stylesheet" type="text/css">
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
  <tr> 

    <td height="33" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" background="../images2/products_04.jpg">
        <!--DWLayoutTable-->
        <tr> 
          <td width="233" height="33" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <!--DWLayoutTable-->
              <tr> 
                <td width="233" height="33"><table width="88%" border="0" align="center">
                    <tr> 
                      <td><font color="#666666" size="2"><strong>您的位置是：( 产品展示 
                        )</strong></font></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
          <td width="567"><table width="60%" border="0" align="right">
              <tr> 
                <td width="27%"><div align="center"><a href="TVPfen-1.asp"><font color="#FF6600" size="2">分类管理</font></a></div></td>
                <td width="26%"><div align="center"><font color="#666666" size="2"><a href="TVPfen-2.asp"><font color="#FF6600">添加分类</font></a></font> 
                  </div></td>
                <td width="23%"><div align="center"><font color="#666666" size="2"><A HREF="TVPphoto-1.asp"><font color="#FF6600">产品管理</font></A></font></div></td>
                <td width="24%"><div align="center"><font color="#666666" size="2"><A HREF="TVPphoto-2.asp"><font color="#FF6600">添加产品</font></A></font></div></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="800" height="410"><table width="98%" height="574" border="0" align="left">
        <tr> 
          <td valign="top"><div align="center"> 
              <form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1">
                <p>&nbsp;</p>
                <table width="90%" border="0">
                  <tr> 
                    <td><table width="90%" height="193" border="0" align="left" bgcolor="#F2F2F2">
                        <tr> 
                          <td width="27%" height="27" class="t2"> <div align="right">分类排序：</div></td>
                          <td width="73%" class="t3"><%=(Recordset1.Fields.Item("ClassID").Value)%></td>
                        </tr>
                        <tr> 
                          <td class="t2"><div align="right">分类名称：</div></td>
                          <td class="t3"><%=(Recordset1.Fields.Item("ClassName").Value)%> </td>
                        </tr>
                
                        <tr> 
                          <td height="67" class="t2"> <div align="right"></div></td>
                          <td> <table width="60%" border="0" align="center">
                              <tr> 
                                <td width="27%"><div align="center"> 
                                    <input type="submit" name="Submit3" value="删除分类">
                                  </div></td>
                                <td width="73%"><input name="Submit32" type="button" onClick="location='TVPfen-1.asp'" value="返回上页"></td>
                              </tr>
                            </table></td>
                        </tr>
                      </table>
                      <p class="t2">&nbsp;</p></td>
                  </tr>
                </table>
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                <input type="hidden" name="MM_recordId" value="<%= Recordset1.Fields.Item("ClassID").Value %>">
                <input type="hidden" name="MM_delete" value="form1">
              </form>
            </div></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
        <td height="12" colspan="5" valign="top" width="1000"><img src="../images/back.jpg"></td>
  </tr>
</table>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>


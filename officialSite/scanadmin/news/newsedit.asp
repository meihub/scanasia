<%@LANGUAGE="VBSCRIPT"%> 

<!--#include file="../../conn2.asp" -->
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
' *** Update Record: set variables

If (CStr(Request("MM_update")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = conn
  MM_editTable = "news"
  MM_editColumn = "ID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "newsadmin.asp"
  MM_fieldsStr  = "new_A1|value|new_A2|value|new_A3|value|content|value"
  MM_columnsStr = "new_A1|',none,''|new_A2|',none,''|new_A3|',none,''|new_A4|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next

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
' *** Update Record: construct a sql update statement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql update statement
  MM_editQuery = "update " & MM_editTable & " set "
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_formVal = MM_fields(MM_i+1)
    MM_typeArray = Split(MM_columns(MM_i+1),",")
    MM_delim = MM_typeArray(0)
    If (MM_delim = "none") Then MM_delim = ""
    MM_altVal = MM_typeArray(1)
    If (MM_altVal = "none") Then MM_altVal = ""
    MM_emptyVal = MM_typeArray(2)
    If (MM_emptyVal = "none") Then MM_emptyVal = ""
    If (MM_formVal = "") Then
      MM_formVal = MM_emptyVal
    Else
      If (MM_altVal <> "") Then
        MM_formVal = MM_altVal
      ElseIf (MM_delim = "'") Then  ' escape quotes
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(MM_i) & " = " & MM_formVal
  Next
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the update
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
If (Request.QueryString("ID") <> "") Then 
  Recordset1__MMColParam = Request.QueryString("ID")
End If
%>
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = conn
Recordset1.Source = "SELECT * FROM news WHERE ID = " + Replace(Recordset1__MMColParam, "'", "''") + ""
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>新聞編輯</title>

<style type=text/css>
a:link {text-decoration: none}
a:hover {text-decoration:none}
a:visited {text-decoration: none}
.style1 {color: #FFFFFF}
</style>
</head>
<link rel="stylesheet" href="../css.css" type="text/css">
<script language="JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
<body background="images/bg_3.gif"><table><tr>
  <td width="850" height="30" bgcolor="333333">&nbsp;&nbsp;<span class="style1">新聞編輯</span></td>
</tr></table>
<table width="98%" height="574" border="0" align="left">
          <tr> 
            <td valign="top"><div align="center"> 
                <form name="form1" method="POST" action="<%=MM_editAction%>">
                  <table width="100%" border="0">
                    <tr> 
                      <td><table width="99%" height="341" border="0" align="left" bgcolor="#F2F2F2">
                          <tr> 
                            <td width="20%" height="47">&nbsp;</td>
                            <td width="80%">&nbsp;</td>
                          </tr>
                          <tr> 
                            <td class="t2"><div align="right">新聞標題：</div></td>
                            <td><input name="new_A1" type="text" id="new_A12" value="<%=(Recordset1.Fields.Item("new_A1").Value)%>" size="30"></td>
                          </tr>
                          <tr> 
                            <td class="t2"><div align="right">日期：</div></td>
                            <td><input name="new_A2" type="text" id="new_A22" value="<%=(Recordset1.Fields.Item("new_A2").Value)%>" size="30"></td>
                          </tr>
                          <tr> 
                            <td class="t2"><div align="right">編輯人：</div></td>
                            <td><input name="new_A3" type="text" id="new_A32" value="<%=(Recordset1.Fields.Item("new_A3").Value)%>" size="30"></td>
                          </tr>
                          <tr> 
                            <td class="t2"><div align="right">新聞内容：</div></td>
                            <td><table width="100%" height="42" border="0" align="center">
                                <tr> 
                                  <td> 
                                    <!-- 编辑器调用开始 -->
                                    <!-- 注意 TEXTAREA 的 NAME 应与 ID=??? 相对应-->
                                    <textarea name="content" style="display:none"><%=(Recordset1.Fields.Item("new_A4").Value)%></textarea> 
                                    <iframe id="Editor" name="Editor" src='../ewebeditor/ewebeditor.asp?id=content&style=standard' frameborder=0 scrolling=no width='550' HEIGHT='400'></iframe> 
								    <!-- 编辑器调用结束 -->
                                    <div align="center"></div></td>
                                </tr>
                              </table></td>
                          </tr>
                          <tr> 
                            <td height="67" class="t2"> <div align="right"></div></td>
                            <td> <table width="60%" border="0" align="center">
                                <tr> 
                                  <td width="27%"><div align="center"> 
                                      <input type="submit" name="Submit3" value="修改新聞">
                                    </div></td>
                                  <td width="73%"><input name="Submit32" type="button" onClick="window.history.back()" value="返回上頁"></td>
                                </tr>
                              </table></td>
                          </tr>
                        </table>
                        <p class="t2">&nbsp;</p></td>
                    </tr>
                  </table>
                  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                  <input type="hidden" name="MM_update" value="form1">
                  <input type="hidden" name="MM_recordId" value="<%= Recordset1.Fields.Item("ID").Value %>">
                </form>
                <p>&nbsp;</p>
                <p>&nbsp;</p>
                <p>&nbsp;</p>
            </div></td>
          </tr>
<A NAME="StranLink"></A>
   </TABLE><SCRIPT LANGUAGE=Javascript SRC="../../images/big5style.js"></SCRIPT>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>


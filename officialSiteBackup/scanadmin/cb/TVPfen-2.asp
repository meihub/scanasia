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
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) = "form1") Then

  MM_editConnection = MM_chao_STRING
  MM_editTable = "cc"
  MM_editRedirectUrl = "TVPfen-1.asp"
  MM_fieldsStr  = "fen_4|value|fen_1|value"
  MM_columnsStr = "ClassID|',none,''|ClassName|',none,''"

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
' *** Insert Record: construct a sql insert statement and execute it

Dim MM_tableValues
Dim MM_dbValues

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
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
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End If
    MM_tableValues = MM_tableValues & MM_columns(MM_i)
    MM_dbValues = MM_dbValues & MM_formVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then
    ' execute the insert
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
<title>添加宣传册</title>
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
                      <td>您的位置是： 宣传册添加 
                        </td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
          <td width="567"><table width="60%" border="0" align="right">
              <tr> 
                <td width="27%"><div align="center"><a href="TVPfen-1.asp">宣传册管理</a></div></td>
                <td width="26%"><div align="center"><a href="TVPfen-2.asp">添加宣传册 </a>
                  </div></td>
                <td width="23%"><div align="center"><A HREF="TVPphoto-1.asp">内页管理</A></div></td>
                <td width="24%"><div align="center"><A HREF="TVPphoto-2.asp">添加内页</A></div></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="800" height="410" valign="top">
<table width="98%" height="367" border="0" align="left">
        <tr> 
          <td height="363" valign="top">
<div align="center"> 
              <form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1">
                <p>&nbsp;</p>
                <table width="90%" border="0">
                  <tr> 
                    <td><table width="90%" height="189" border="0" align="left" bgcolor="#F2F2F2">
                         <tr> 
                          <td width="25%" height="23" class="t2"> <div align="right">分类排序:</div></td>
                          <td width="75%"><input name="fen_4" type="text" id="fen_4"> 
                            <span class="t2"><font color="#FF0000">请输入数字（数字最小最前）</font></span> 
                          </td>
                        </tr>
                      
                        <tr> 
                          <td class="t2"><div align="right">类别名称：</div></td>
                          <td><input name="fen_1" type="text" id="fen_1"> </td>
                        </tr>
                                     
                        <tr> 
                          <td height="67" class="t2"> <div align="right"></div></td>
                          <td> <table width="60%" border="0" align="center">
                              <tr> 
                                <td width="27%"><div align="center"> 
                                    <input type="submit" name="Submit3" value="添加">
                                  </div></td>
                                <td width="73%"><input name="Submit32" type="button" onClick="location='fen-1.asp'" value="返回上页"></td>
                              </tr>
                            </table></td>
                        </tr>
                      </table>
                      <p class="t2">&nbsp;</p></td>
                  </tr>
                </table>
                <input type="hidden" name="MM_insert" value="form1">
              </form>
            </div></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
        <td height="12" colspan="5" valign="top" width="1000"><img src="../images/back.jpg"></td>
  </tr>
</table>
</body>
</html>

<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%> 
<!--#include file="../conn.asp"-->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="chao"
MM_authFailedURL="../hou3.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (false Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
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
Recordset1.ActiveConnection = MM_chao_STRING
Recordset1.Source = "SELECT * FROM photo WHERE ID = " + Replace(Recordset1__MMColParam, "'", "''") + ""
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>
<%
Dim Recordset2__MMColParam
Recordset2__MMColParam = "5"
If (Request("MM_EmptyValue") <> "") Then 
  Recordset2__MMColParam = Request("MM_EmptyValue")
End If
%>
<%
Dim Recordset2
Dim Recordset2_numRows

Set Recordset2 = Server.CreateObject("ADODB.Recordset")
Recordset2.ActiveConnection = MM_chao_STRING
Recordset2.Source = "SELECT * FROM piao WHERE ID = " + Replace(Recordset2__MMColParam, "'", "''") + ""
Recordset2.CursorType = 0
Recordset2.CursorLocation = 2
Recordset2.LockType = 1
Recordset2.Open()

Recordset2_numRows = 0
%>
<%
Dim MM_paramName 
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title> 图片管理</title>
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
<table width="801" border="0" cellpadding="0" cellspacing="0">
  <!--DWLayoutTable-->
  <tr> 
    <td height="101" colspan="6" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <!--DWLayoutTable-->
        <tr> 
          <td width="1000" height="101"><!--#include file="..\top6.asp"--></td>
        </tr>
      </table></td>
    <td width="1"></td>
  </tr>
  <tr> 
    <td width="113" rowspan="3" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <!--DWLayoutTable-->
        <tr> 
          <td width="113" height="140"><img src="../images/index_4.jpg" width="113" height="140"></td>
        </tr>
      </table></td>
    <td width="27" rowspan="9" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <!--DWLayoutTable-->
        <tr> 
          <td width="27" height="113" valign="top"><img src="../images/index_5.jpg" width="27" height="113"></td>
        </tr>
        <tr> 
          <td height="278" valign="top"><img src="../images/index_11.jpg" width="27" height="278"></td>
        </tr>
        <tr> 
          <td height="197" valign="top"><img src="../images/index_16.jpg" width="27" height="197"></td>
        </tr>
        <tr> 
          <td height="121" valign="top"><img src="../images/index_19.jpg" width="27" height="121"></td>
        </tr>
      </table></td>
    <td height="75" colspan="2" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" background="../images/index_644.jpg">
        <!--DWLayoutTable-->
        <tr> 
          <td width="16" height="43"></td>
          <td width="68"></td>
          <td width="23"></td>
          <td width="68"></td>
          <td width="23"></td>
          <td width="128">&nbsp;</td>
          <td width="25"></td>
          <td width="68"></td>
          <td width="70"></td>
        </tr>
        <tr> 
          <td height="4"></td>
          <td></td>
          <td></td>
          <td></td>
          <td></td>
          <td rowspan="2" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <!--DWLayoutTable-->
              <tr> 
                <td width="128" height="32"><div align="center"><font color="#FF6600"><strong>查看图片</strong></font></div></td>
              </tr>
            </table></td>
          <td></td>
          <td></td>
          <td></td>
        </tr>
        <tr> 
          <td height="28"></td>
          <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <!--DWLayoutTable-->
              <tr> 
                <td width="68" height="28"><div align="center"><font color="#666666" size="2"><A HREF="fen-1.asp"><font color="#666666">分类管理</font></A></font></div></td>
              </tr>
            </table></td>
          <td></td>
          <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <!--DWLayoutTable-->
              <tr> 
                <td width="68" height="28"><div align="center"><font color="#666666" size="2"><A HREF="fen-2.asp"><font color="#666666">添加分类</font></A></font></div></td>
              </tr>
            </table></td>
          <td></td>
          <td></td>
          <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <!--DWLayoutTable-->
              <tr> 
                <td width="68" height="28"><div align="center"><font color="#666666" size="2"><A HREF="photo-1.asp"><font color="#666666">图片管理</font></A></font></div></td>
              </tr>
            </table></td>
          <td></td>
        </tr>
      </table></td>
    <td colspan="2" rowspan="2" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <!--DWLayoutTable-->
        <tr> 
          <td width="171" height="103"><img src="../images/index_7.jpg" width="171" height="103"></td>
        </tr>
      </table></td>
    <td></td>
  </tr>
  <tr> 
    <td width="110" height="28"></td>
    <td width="379"></td>
    <td></td>
  </tr>
  <tr> 
    <td colspan="3" rowspan="6" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <!--DWLayoutTable-->
        <tr> 
          <td width="610" height="581"><div align="center">
              <table width="98%" height="574" border="0" align="left">
                <tr> 
                  <td valign="top"><div align="center"> 
                      <form name="form1" method="post" action="">
                        <table width="97%" border="0">
                          <tr> 
                            <td height="26" colspan="4" class="a1">&nbsp;</td>
                          </tr>
                          <tr valign="middle" class="t3"> 
                            <td width="20%" height="17"> <div align="center"></div>
                              <div align="center">排序</div></td>
                            <td width="18%"><div align="center">图片名称</div></td>
                            <td width="16%"><div align="center">日期</div></td>
                            <td width="17%"><div align="center">图片编辑</div></td>
                          </tr>
                          <tr class="t2"> 
                            <td> <div align="center"><%=(Recordset1.Fields.Item("photo_ps").Value)%></div></td>
                            <td> <div align="center"><%=(Recordset1.Fields.Item("photo_mc").Value)%></div></td>
                            <td> <div align="center"><%=(Recordset1.Fields.Item("photo_rq").Value)%></div></td>
                            <td> <div align="center" class="t2">[<A HREF="photo-4.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ID=" & Recordset1.Fields.Item("ID").Value %>"><font color="#666666">修改</font></A>][<A HREF="photo-5.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ID=" & Recordset1.Fields.Item("ID").Value %>"><font color="#666666">删除</font></A>]</div></td>
                          </tr>
                          <tr> 
                            <td class="t2"><div align="center">图片显示：</div></td>
                            <td colspan="3" rowspan="5" class="t2"><img src="images/photo/<%=(Recordset1.Fields.Item("photo_st").Value)%>" width="169" height="143"></td>
                          </tr>
                          <tr> 
                            <td class="t2"><div align="right"></div></td>
                          </tr>
                          <tr> 
                            <td class="t2"><div align="right"></div></td>
                          </tr>
                          <tr> 
                            <td class="t2"><div align="center"><span class="t3"> 
                                <input name="Submit32" type="button" onClick="location='photo-1.asp'" value="返回上页">
                                </span></div></td>
                          </tr>
                          <tr> 
                            <td><div align="center" class="t3"></div></td>
                          </tr>
                        </table>
                      </form>
                      <p class="t2">&nbsp;</p>
                    </div></td>
                </tr>
              </table>
            </div></td>
        </tr>
      </table></td>
    <td width="50" rowspan="2" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <!--DWLayoutTable-->
        <tr> 
          <td width="50" height="87"><img src="../images/index_10.jpg" width="50" height="87"></td>
        </tr>
      </table></td>
    <td height="37"></td>
  </tr>
  <tr> 
    <td rowspan="2" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <!--DWLayoutTable-->
        <tr> 
          <td width="113" height="194"><img src="../images/index_12.jpg" width="113" height="194"></td>
        </tr>
      </table></td>
    <td height="50"></td>
  </tr>
  <tr> 
    <td height="144" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <!--DWLayoutTable-->
        <tr> 
          <td width="50" height="144"><img src="../images/index_13.jpg" width="50" height="144"></td>
        </tr>
      </table></td>
    <td></td>
  </tr>
  <tr> 
    <td height="186" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <!--DWLayoutTable-->
        <tr> 
          <td width="113" height="186"><img src="../images/index_14.jpg" width="113" height="186"></td>
        </tr>
      </table></td>
    <td rowspan="2" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <!--DWLayoutTable-->
        <tr> 
          <td width="50" height="229"><img src="../images/index_15.jpg" width="50" height="229"></td>
        </tr>
      </table></td>
    <td></td>
  </tr>
  <tr> 
    <td rowspan="4" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <!--DWLayoutTable-->
        <tr> 
          <td width="113" height="254"><img src="../images/index_17.jpg" width="113" height="254"></td>
        </tr>
      </table></td>
    <td height="43"></td>
  </tr>
  <tr> 
    <td rowspan="2" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <!--DWLayoutTable-->
        <tr> 
          <td width="50" height="146"><img src="../images/index_18.jpg" width="50" height="146"></td>
        </tr>
      </table></td>
    <td height="121"></td>
  </tr>
  <tr> 
    <td height="25" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <!--DWLayoutTable-->
        <tr> 
          <td width="110" height="25"><img src="../images/index_20.jpg" width="110" height="25"></td>
        </tr>
      </table></td>
    <td colspan="2" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" background="../images/index_21.jpg">
        <!--DWLayoutTable-->
        <tr> 
          <td width="500" height="25">&nbsp;</td>
        </tr>
      </table></td>
    <td></td>
  </tr>
  <tr> 
    <td height="65" colspan="5" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <!--DWLayoutTable-->
        <tr> 
          <td width="687" height="65"><img src="../images/index_24.jpg" width="687" height="65"></td>
        </tr>
      </table></td>
    <td></td>
  </tr>
  <tr> 
    <td height="1"></td>
    <td></td>
    <td></td>
    <td></td>
    <td width="121"></td>
    <td></td>
    <td></td>
  </tr>
</table>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
<%
Recordset2.Close()
Set Recordset2 = Nothing
%>

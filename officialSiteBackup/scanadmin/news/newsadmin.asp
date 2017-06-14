<%@LANGUAGE="VBSCRIPT"%>

<!--#include file="../../conn2.asp" -->
<%if session("admin")="" then
response.Redirect "../login.asp"
end if
%>
<%
set rs = Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = conn
rs.Source = "SELECT * FROM news ORDER BY id DESC"
rs.CursorType = 0
rs.CursorLocation = 2
rs.LockType = 3
rs.Open()
rs_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rs_numRows = rs_numRows + Repeat1__numRows
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then MM_removeList = MM_removeList & "&" & MM_paramName & "="
MM_keepURL="":MM_keepForm="":MM_keepBoth="":MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each Item In Request.QueryString
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & NextItem & Server.URLencode(Request.QueryString(Item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each Item In Request.Form
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & NextItem & Server.URLencode(Request.Form(Item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
if (MM_keepBoth <> "") Then MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
if (MM_keepURL <> "")  Then MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
if (MM_keepForm <> "") Then MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>新聞管理</title>
<link rel="stylesheet" href="../css.css" type="text/css">
<style type="text/css">
<!--
.style1 {color: #FFFFFF}
-->
</style>
<body background="images/bg_3.gif">
<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
var checkflag = "false";
function check(field) {
if (checkflag == "false") {
for (i = 0; i < field.length; i++) {
field[i].checked = true;}
checkflag = "true";
return "取消选择"; }
else {
for (i = 0; i < field.length; i++) {
field[i].checked = false; }
checkflag = "false";
return "全部选定"; }
}
//  End -->
</script> <table><tr>
  <td width="850" height="30" bgcolor="333333">&nbsp;&nbsp;<span class="style1">新聞管理</span></td>
</tr></table>
<form name="form1" method="post" action="del3.asp"> <table width="600" border="0" cellspacing="0" cellpadding="0" align="center"> 
<tr> <td width="140"><img src="../images/T_left.gif" width="140" height="21"></td><td background="../images/T_bg.gif">&nbsp;</td><td width="56"><img src="../images/T_right.gif" width="56" height="21"></td></tr> 
</table><table width="600" border="0" cellspacing="0" cellpadding="0" align="center" > 
<tr> <td height="28" align="center" background="../images/title_3.gif"><font color="#FFFFFF"><b></b></font></td>
</tr> 
<tr> <td  height="15"> <% 
While ((Repeat1__numRows <> 0) AND (NOT rs.EOF)) 
%> <table width="100%" border="0" cellspacing="1" cellpadding="0"> <tr onmouseover="this.style.background='#F6F6F6'" 
                    onmouseout="this.style.background='FFFFFF'"  background='FFFFFF'  > <td width="25" align="center" height="25"><img src="../images/tp009.gif" width="15" height="15"></td><td><A HREF="../../shownews.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "id=" & rs.Fields.Item("id").Value %>" target="_blank"><%=(rs.Fields.Item("new_A1").Value)%></A></td><td width="80" align="right"><a href="../newshow.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "id=" & rs.Fields.Item("id").Value %>" target="_blank"><%=(rs.Fields.Item("new_A2").Value)%></a></td><td width="25" align="center" bgcolor="#FFFFFF"> 
<input type="checkbox" name="checkbox" value="<%=(rs.Fields.Item("id").Value)%>"> 
</td><td width="30" align="center" bgcolor="#FFFFFF"><A HREF="newsedit.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "id=" & rs.Fields.Item("id").Value %>"><img src="../images/dnsedit.gif" width="30" height="23" border="0"></A></td></tr> 
</table><% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rs.MoveNext()
Wend
%> <% If rs.EOF And rs.BOF Then %> <TABLE WIDTH="400" BORDER="0" CELLSPACING="0" CELLPADDING="0" ALIGN="center" HEIGHT="100"> 
<TR> <TD ALIGN="center"><FONT COLOR="#FF0000"><B>沒有相關新聞！</B></FONT></TD>
</TR> 
</TABLE><% End If ' end rs.EOF And rs.BOF %> </td></tr> <tr> <td bgcolor="#ffffff" height="30" align="right" background="images/title_3.gif"> 
<input type=button value="全部選定" onClick="this.value=check(this.form.checkbox)" name="button"> 
<input type="submit" name="Submit" value="刪除所選"> </td></tr> <A NAME="StranLink"></A>
   </TABLE><SCRIPT LANGUAGE=Javascript SRC="../../images/big5style.js"></SCRIPT>
</form><%
rs.Close()
%> 
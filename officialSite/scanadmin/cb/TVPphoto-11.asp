<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<!--#include file="../Connections/chao.asp"-->

<%
dim range,action
range=request("range")
action=request("action")
scc=request("scc")
if action="looked" then
set rs4= Server.CreateObject("adodb.recordset")
sql="select * from c where id="&range
rs4.open sql,MM_chao_STRING,1,3
rs4("New")=1
rs4.update
rs4.close
set rs4=nothing
response.write"<script language=javascript>window.location.href='TVPphoto-11.asp?scc="&request("scc")&"&id="&request("id")&"&page="&request("page")&"';</script>"
response.end
end if

if action="nolooked" then
set rs4= Server.CreateObject("adodb.recordset")
sql="select * from c where id="&range
rs4.open sql,MM_chao_STRING,1,3
rs4("New")=0
rs4.update
rs4.close
set rs4=nothing
response.write"<script language=javascript>window.location.href='TVPphoto-11.asp?scc="&request("scc")&"&id="&request("id")&"&page="&request("page")&"';</script>"
response.end
end if %>

<%
Dim Recordset3__MMColParam
Recordset3__MMColParam = "1"
If (Request.QueryString("scc") <> "") Then 
  Recordset3__MMColParam = Request.QueryString("scc")
End If
%>
<%
Dim Recordset3
Dim Recordset3_numRows

Set Recordset3 = Server.CreateObject("ADODB.Recordset")
Recordset3.ActiveConnection = MM_chao_STRING
Recordset3.Source = "SELECT * FROM c WHERE ClassID= " + Replace(Recordset3__MMColParam, "'", "''") + " ORDER BY ID DESC"
Recordset3.CursorType = 0
Recordset3.CursorLocation = 2
Recordset3.LockType = 1
Recordset3.Open()

Recordset3_numRows = 0
%>
<%
Dim Recordset4__MMColParam
Recordset4__MMColParam = "1"
If (Request.QueryString("scc") <> "") Then 
  Recordset4__MMColParam = Request.QueryString("scc")
End If
%>
<%
Dim Recordset4
Dim Recordset4_numRows

Set Recordset4 = Server.CreateObject("ADODB.Recordset")
Recordset4.ActiveConnection = MM_chao_STRING
Recordset4.Source = "SELECT * FROM cc WHERE ClassID= " + Replace(Recordset4__MMColParam, "'", "''") + ""
Recordset4.CursorType = 0
Recordset4.CursorLocation = 2
Recordset4.LockType = 1
Recordset4.Open()

Recordset4_numRows = 0
%>
<%
Dim Recordset5
Dim Recordset5_numRows

Set Recordset5 = Server.CreateObject("ADODB.Recordset")
Recordset5.ActiveConnection = MM_chao_STRING
Recordset5.Source = "SELECT * FROM cc ORDER BY ClassID ASC"
Recordset5.CursorType = 0
Recordset5.CursorLocation = 2
Recordset5.LockType = 1
Recordset5.Open()

Recordset5_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
Recordset6_numRows = Recordset6_numRows + Repeat1__numRows
%>
<%
Dim HLooper1__numRows
HLooper1__numRows = 12
Dim HLooper1__index
HLooper1__index = 0
Recordset3_numRows = Recordset3_numRows + HLooper1__numRows
%>

<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim Recordset3_total
Dim Recordset3_first
Dim Recordset3_last

' set the record count
Recordset3_total = Recordset3.RecordCount

' set the number of rows displayed on this page
If (Recordset3_numRows < 0) Then
  Recordset3_numRows = Recordset3_total
Elseif (Recordset3_numRows = 0) Then
  Recordset3_numRows = 1
End If

' set the first and last displayed record
Recordset3_first = 1
Recordset3_last  = Recordset3_first + Recordset3_numRows - 1

' if we have the correct record count, check the other stats
If (Recordset3_total <> -1) Then
  If (Recordset3_first > Recordset3_total) Then
    Recordset3_first = Recordset3_total
  End If
  If (Recordset3_last > Recordset3_total) Then
    Recordset3_last = Recordset3_total
  End If
  If (Recordset3_numRows > Recordset3_total) Then
    Recordset3_numRows = Recordset3_total
  End If
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (Recordset3_total = -1) Then

  ' count the total records by iterating through the recordset
  Recordset3_total=0
  While (Not Recordset3.EOF)
    Recordset3_total = Recordset3_total + 1
    Recordset3.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (Recordset3.CursorType > 0) Then
    Recordset3.MoveFirst
  Else
    Recordset3.Requery
  End If

  ' set the number of rows displayed on this page
  If (Recordset3_numRows < 0 Or Recordset3_numRows > Recordset3_total) Then
    Recordset3_numRows = Recordset3_total
  End If

  ' set the first and last displayed record
  Recordset3_first = 1
  Recordset3_last = Recordset3_first + Recordset3_numRows - 1
  
  If (Recordset3_first > Recordset3_total) Then
    Recordset3_first = Recordset3_total
  End If
  If (Recordset3_last > Recordset3_total) Then
    Recordset3_last = Recordset3_total
  End If

End If
%>

<%
Dim MM_paramName 
%>
<%
' *** Move To Record and Go To Record: declare variables

Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined

Dim MM_param
Dim MM_index

Set MM_rs    = Recordset3
MM_rsCount   = Recordset3_total
MM_size      = Recordset3_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  MM_param = Request.QueryString("index")
  If (MM_param = "") Then
    MM_param = Request.QueryString("offset")
  End If
  If (MM_param <> "") Then
    MM_offset = Int(MM_param)
  End If

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
  If (MM_rs.EOF) Then 
    MM_offset = MM_index  ' set MM_offset to the last possible record
  End If

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  MM_index = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = MM_index
    If (MM_size < 0 Or MM_size > MM_rsCount) Then
      MM_size = MM_rsCount
    End If
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While (Not MM_rs.EOF And MM_index < MM_offset)
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
Recordset3_first = MM_offset + 1
Recordset3_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (Recordset3_first > MM_rsCount) Then
    Recordset3_first = MM_rsCount
  End If
  If (Recordset3_last > MM_rsCount) Then
    Recordset3_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
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
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev

Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 1) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    MM_paramList = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For MM_paramIndex = 0 To UBound(MM_paramList)
      MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
      If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then 
  MM_keepMove = MM_keepMove & "&"
End If

MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="

MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
  MM_movePrev = MM_urlStr & "0"
Else
  MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If
%>
<script language="javascript">
<!-- 
//图片按比例显示程序
var flag=false; 
function DrawImage(ImgD,Width,Height){ 
var image=new Image(); 
image.src=ImgD.src; 
if(image.width>0 && image.height>0){ 
flag=true; 
if(image.width/image.height>= Width/Height){ 
if(image.width>Width){ 
ImgD.width=Width; 
ImgD.height=(image.height*Width)/image.width; 
}else{ 
ImgD.width=image.width; 
ImgD.height=image.height; 
} 
} 
else{ 
if(image.height>Height){ 
ImgD.height=Height; 
ImgD.width=(image.width*Height)/image.height; 
}else{ 
ImgD.width=image.width; 
ImgD.height=image.height; 
} 
} 
} 
} 
//--> 
</script> 
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

<table width="850" border="0" cellpadding="0" cellspacing="0">
  <!--DWLayoutTable-->
  <tr bgcolor="f2f2f2"> 
    
    <td height="33" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" background="../images2/products_04.jpg">
        <!--DWLayoutTable-->
        <tr> 
          <td width="233" height="33" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <!--DWLayoutTable-->
              <tr> 
                <td width="233" height="33"><table width="88%" border="0" align="center">
                    <tr> 
                      <td>您的位置是：项目产品管理</td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
          <td width="567"><table width="60%" border="0" align="right">
              <tr bgcolor="f2f2f2"> 
                <td width="27%"><div align="center"><form name="form2" action="TVPphoto-11.asp">
                              <table width="50%" border="0" align="center">
                                <tr> 
                                  <td width="33%"><div align="center"> 
                                      <select name="scc" id="scc">
                                        <%
While (NOT Recordset5.EOF)
%>
                                        <option value="<%=(Recordset5.Fields.Item("ClassID").Value)%>" <%If (Not isNull((Recordset5.Fields.Item("ClassName").Value))) Then If (CStr(Recordset5.Fields.Item("ClassID").Value) = CStr((Recordset5.Fields.Item("ClassName").Value))) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(Recordset5.Fields.Item("ClassName").Value)%></option>
                                        <%
  Recordset5.MoveNext()
Wend
If (Recordset5.CursorType > 0) Then
  Recordset5.MoveFirst
Else
  Recordset5.Requery
End If
%>
                                      </select>
                                    </div></td>
                                  <td width="67%"> <div align="left"> 
                                      <input type="submit" name="Submit" value="  GO  ">
                                    </div></td>
                                </tr>
                              </table>
                            </form></div></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="1000" height="410"><table width="98%" height="574" border="0" align="left">
        <tr> 
          <td valign="top"><div align="center"> 
              <% If Not Recordset3.EOF Or Not Recordset3.BOF Then %>
              <form name="form1" method="post" action="">
                <table width="100%" border="0">
                  <tr> 
                    <td height="26" colspan="4" class="a1"><div align="center"> 
                        <table width="94%" border="0">
                          <tr bgcolor="#CCCCCC"> 
                            <td><font color="#000000" size="2">产品类别：<%=(Recordset4.Fields.Item("ClassName").Value)%></font></td><td>&nbsp;</td>
                          </tr>
                        </table>
                        <strong></strong></div></td>
                  </tr>
                  <tr> 
                    <td height="79" colspan="4" class="t2"> <table width="95%">
                        <%
startrw = 0
endrw = HLooper1__index
numberColumns = 4
numrows = 3
while((numrows <> 0) AND (Not Recordset3.EOF))
	startrw = endrw + 1
	endrw = endrw + numberColumns
 %>
                        <tr align="center" valign="top"> 
                          <%
While ((startrw <= endrw) AND (Not Recordset3.EOF))
%>
                          <td> <table width="95%" border="0">
                              <tr> 
                                <td height="26" class="t2"><div align="center"><a href="TVPproductshow.asp?ID=<%=(Recordset3.Fields.Item("ID").Value)%>" ><img src="images\photo\<%=(Recordset3.Fields.Item("ProImg").Value)%>" onload='DrawImage(this,100,100);' border="0"></a></div></td>
                         
                              </tr>
                              <tr> 
                                <td width="4%" height="26" class="t2"><div align="center"> 
                                    <table width="85%" border="0">
                                      <tr> 
                                        <td align="center" class="t1"><table><tr><td><div align="center"><%=(Recordset3.Fields.Item("ProName").Value)%>
										<%'if Recordset3.Fields.Item("New").Value=0 then%>
          &nbsp;
          <%'else%>
          <!--font color=red>[封面]</font-->
          <%'end if%>
										</div></td></tr></table></td>
                                      </tr>
                                    </table>
                                  </div></td>
                            
                              </tr>
                              <tr> 
                                <td class="t2"> <div align="center"> 
                                    <table width="100%" border="0">
                                      <tr> 
                                        <td bgcolor="e7e7e7"> <div align="left"><A HREF="photo-3.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ID=" & Recordset3.Fields.Item("ID").Value %>"></A></div>
                                          <div align="center" class="t2">[<a href="works_id_manage.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ID=" & Recordset3.Fields.Item("ID").Value %>"><font color="#666666">修改</font></a>][<a href="TVPphoto-5.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "ID=" & Recordset3.Fields.Item("ID").Value %>"><font color="#666666">删除</font></a>]
										   <!--<%'if Recordset3.Fields.Item("New").Value=0 then%>
          <a href="?action=looked&range=<%'=Recordset3.Fields.Item("id").Value%>&scc=<%'=Recordset3.Fields.Item("ClassID").Value%>&page=<%=page%>" title="设为封面">[封面]</a>
          <%'else%>
          <a href="?action=nolooked&range=<%'=Recordset3.Fields.Item("id").Value%>&scc=<%'=Recordset3.Fields.Item("ClassID").Value%>&page=<%=page%>" title="设为普通内页">[内页]</a>
          <%'end if%> --></div></td>
                                      </tr>
                                    </table>
                                  </div>
                                  <div align="center" class="t2"></div></td>
                       
                              </tr>
                              <tr> 
                                <td colspan="6">&nbsp;</td>
                              </tr>
                            </table></td>
                          <%
	startrw = startrw + 1
	Recordset3.MoveNext()
	Wend
	%>
                        </tr>
                        <%
 numrows=numrows-1
 Wend
 %>
                      </table></td>
                  </tr>
                  <tr> 
                    <td width="54%"><div align="center" class="t2"> 记录 <%=(Recordset3_first)%> 
                        到 <%=(Recordset3_last)%> (总共 <%=(Recordset3_total)%> 个记录) 
                      </div></td>
                    <td width="46%" colspan="2">&nbsp; <table width="50%" border="0" align="center" class="t2">
                        <tr> 
                          <td align="center"> 
                            <% If MM_offset <> 0 Then %>
                            <a href="<%=MM_moveFirst%>"><font color="#666666">第一页</font></a> 
                            <% End If ' end MM_offset <> 0 %>
                            <% If MM_offset <> 0 Then %>
                            <a href="<%=MM_movePrev%>"><font color="#666666">前一页</font></a> 
                            <% End If ' end MM_offset <> 0 %>
                            <% If Not MM_atTotal Then %>
                            <a href="<%=MM_moveNext%>"><font color="#666666">下一页</font></a> 
                            <% End If ' end Not MM_atTotal %>
                            <% If Not MM_atTotal Then %>
                            <a href="<%=MM_moveLast%>"><font color="#666666">最后一页</font></a> 
                            <% End If ' end Not MM_atTotal %>
                          </td>
                        </tr>
                      </table></td>
                  </tr>
                </table>
              </form>
              <% End If ' end Not Recordset3.EOF Or NOT Recordset3.BOF %>
              <% If Recordset3.EOF And Recordset3.BOF Then %>
              <table width="69%" border="0">
                <tr> 
                  <td><div align="center"><strong><font color="#FF0000">该分类没有图片</font></strong></div></td>
                </tr>
              </table>
              <% End If ' end Recordset3.EOF And Recordset3.BOF %>
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
Recordset3.Close()
Set Recordset3 = Nothing
%>
<%
Recordset4.Close()
Set Recordset4 = Nothing
%>


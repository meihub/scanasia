<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp" -->
<%
Dim Recordset3
Dim Recordset3_numRows

Set Recordset3 = Server.CreateObject("ADODB.Recordset")
Recordset3.ActiveConnection = conn
Recordset3.Source = "SELECT * FROM news ORDER BY ID DESC"
Recordset3.CursorType = 0
Recordset3.CursorLocation = 2
Recordset3.LockType = 1
Recordset3.Open()

Recordset3_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 100
Repeat1__index = 0
Recordset3_numRows = Recordset3_numRows + Repeat1__numRows
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
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString<>"" Then MM_LoginAction = MM_LoginAction + "?" + Request.QueryString
MM_valUsername=CStr(Request.Form("m_username"))
If MM_valUsername <> "" Then
  MM_fldUserAuthorization=""
  MM_redirectLoginSuccess="index.asp"
  MM_redirectLoginFailed="couwu.asp"
  MM_flag="ADODB.Recordset"
  set MM_rsUser = Server.CreateObject(MM_flag)
  MM_rsUser.ActiveConnection = MM_chao_STRING
  MM_rsUser.Source = "SELECT m_username, m_passwd"
  If MM_fldUserAuthorization <> "" Then MM_rsUser.Source = MM_rsUser.Source & "," & MM_fldUserAuthorization
  MM_rsUser.Source = MM_rsUser.Source & " FROM memberData WHERE m_username='" & Replace(MM_valUsername,"'","''") &"' AND m_passwd='" & Replace(Request.Form("m_passwd"),"'","''") & "'"
  MM_rsUser.CursorType = 0
  MM_rsUser.CursorLocation = 2
  MM_rsUser.LockType = 3
  MM_rsUser.Open
  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
    Else
      Session("MM_UserAuthorization") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And false Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MM_rsUser.Close
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!-- saved from url=(0027)http://www.gzlinyi.cn/scan/ -->
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD><TITLE>SCAN ASIA</TITLE>
<META http-equiv=Content-Type content="text/html; charset=utf-8">
<LINK href="images/css2.css" type=text/css rel=stylesheet>
<LINK href="js/default.css" type=text/css rel=stylesheet>
<SCRIPT src="images/ilank_flash.js" type=text/javascript></SCRIPT>
<LINK href="js/lightbox.css" type=text/css rel=stylesheet>
<SCRIPT src="js/jquery.js" type=text/javascript></SCRIPT>
<SCRIPT src="js/drag.js" type=text/javascript></SCRIPT>
<SCRIPT src="js/scroll.js" type=text/javascript></SCRIPT>
<SCRIPT src="js/about.js" type=text/javascript></SCRIPT>
<script type="text/javascript"> 
window.onload=function(){ 
for(var ii=0; ii<document.links.length; ii++) 
document.links[ii].onfocus=function(){this.blur()} 
} 
</script>
<META content="MSHTML 6.00.6000.21015" name=GENERATOR></HEAD>
<body style="background-color:transparent"><table ><tr>
<td width="612" id=content_scroll style="OVERFLOW: hidden">
<table id=media_content ><tr><td valign="top" > <% If Not Recordset3.EOF Or Not Recordset3.BOF Then %>
            <table   border="0" cellpadding="0" cellspacing="0"  >
              <!--DWLayoutTable-->
              <tr> 
                <td width="612" height="25" valign="top">
				<% 
While ((Repeat1__numRows <> 0) AND (NOT Recordset3.EOF)) 
%>
                  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="vip2">
                    <!--DWLayoutTable-->
                    <tr> 
                      <td width="610"> <table width="100%" border="0" >
                          <tr>  <td width="20" height="24" align="center"><img src="images/m_n.gif"></td>
                            <td width="581" align="left" id="news" ><a href="shownews.asp?<%= MM_keepURL & MM_joinChar(MM_keepURL) & "ID=" & Recordset3.Fields.Item("ID").Value %>" target="_blank" ><%=(Recordset3.Fields.Item("new_A1").Value)%></a>&nbsp;<span id="nt"><%=(Recordset3.Fields.Item("new_A2").Value)%></span></td>
                            <td ></td>
                          </tr>
						  <tr><td colspan="3" align="left"><table width="100%"><tr><td width="30px"></td><td id="nt"><% 
			if len(Recordset3.Fields.Item("new_A4").Value)>200 then
			response.Write(left(Recordset3.Fields.Item("new_A4").Value,200)&"…")
			else
			response.Write(Recordset3.Fields.Item("new_A4").Value)
			end if
			%></td></tr></table></td></tr>
                        </table></td>
                    </tr>
                  </table>
                  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset3.MoveNext()
Wend
%> </td>
              </tr>
              <tr> 
                <td height="50" valign="bottom"><!--DWLayoutEmptyCell-->&nbsp; </td>
              </tr>
            </table>
            <% End If ' end Not Recordset3.EOF Or NOT Recordset3.BOF %> </td></tr></table></td>
</tr></table>
<DIV id=scrollArea>
<DIV id=scroller></DIV></body>
<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../../conn2.asp" -->
<!--#include file="../ScriptLibrary/incPureUpload.asp" -->


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

  MM_editConnection = conn
  MM_editTable = "news"
  MM_editRedirectUrl = "newsadmin.asp"
  MM_fieldsStr  = "new_A1|value|new_A2|value|new_A3|value|news_content|value"
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
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = conn
Recordset1.Source = "SELECT * FROM news"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>


<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>添加新聞</title>

<link href="../css.css" rel="stylesheet" type="text/css">
<style type=text/css>
a:link {text-decoration: none}
a:hover {text-decoration:none}
a:visited {text-decoration: none}
.style1 {color: #FFFFFF}
</style>
</head>
<title>管理</title>
<link rel="stylesheet" href="css.css" type="text/css">
<script language="JavaScript">
<!--

function checkFileUpload(form,extensions,requireUpload,sizeLimit,minWidth,minHeight,maxWidth,maxHeight,saveWidth,saveHeight) { //v2.0
  document.MM_returnValue = true;
  if (extensions != '') var re = new RegExp("\.(" + extensions.replace(/,/gi,"|") + ")$","i");
  for (var i = 0; i<form.elements.length; i++) {
    field = form.elements[i];
    if (field.type.toUpperCase() != 'FILE') continue;
    if (field.value == '') {
      if (requireUpload) {alert('File is required!');document.MM_returnValue = false;field.focus();break;}
    } else {
      if(extensions != '' && !re.test(field.value)) {
        alert('This file type is not allowed for uploading.\nOnly the following file extensions are allowed: ' + extensions + '.\nPlease select another file and try again.');
        document.MM_returnValue = false;field.focus();break;
      }
    document.PU_uploadForm = form;
    re = new RegExp(".(gif|jpg|png|bmp|jpeg)$","i");
    if(re.test(field.value) && (sizeLimit != '' || minWidth != '' || minHeight != '' || maxWidth != '' || maxHeight != '' || saveWidth != '' || saveHeight != '')) {
      setTimeout('if (document.MM_returnValue) document.PU_uploadForm.submit()',500);
      checkImageDimensions(field.value,sizeLimit,minWidth,minHeight,maxWidth,maxHeight,saveWidth,saveHeight);
    } } }
}

function showImageDimensions() { //v2.0
  if ((this.minWidth != '' && this.width > this.minWidth) || (this.minHeight != '' && this.height < this.minHeight)) {
    alert('Uploaded Image is too small!\nShould be at least ' + this.minWidth + ' x ' + this.minHeight); return;}
  if ((this.maxWidth != '' && this.width > this.maxWidth) || (this.maxHeight != '' && this.height > this.maxHeight)) {
    alert('Uploaded Image is too big!\nShould be max ' + this.maxWidth + ' x ' + this.maxHeight); return;}
  if (this.sizeLimit != '' && this.fileSize > this.sizeLimit) {
    alert('Uploaded Image File Size is too big!\nShould be max ' + this.sizeLimit + ' bytes'); return;}
  if (this.saveWidth != '') document.PU_uploadForm[this.saveWidth].value = this.width;
  if (this.saveHeight != '') document.PU_uploadForm[this.saveHeight].value = this.height;
  document.MM_returnValue = true;
}

function checkImageDimensions(fileName,sizeLimit,minWidth,minHeight,maxWidth,maxHeight,saveWidth,saveHeight) { //v2.0
  document.MM_returnValue = false; var imgURL = 'file:///' + fileName, img = new Image();
  img.sizeLimit = sizeLimit; img.minWidth = minWidth; img.minHeight = minHeight; img.maxWidth = maxWidth; img.maxHeight = maxHeight;
  img.saveWidth = saveWidth; img.saveHeight = saveHeight;
  img.onload = showImageDimensions; img.src = imgURL;
}
//-->
</script>
<script language=javascript>
<!--
function showSending() {
	sending.style.visibility="visible";
	}
-->
</script>
<body background="images/bg_3.gif">
<br><br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
  <td width="100%" height="30" bgcolor="333333">&nbsp;&nbsp;<span class="style1">添加新聞</span></td>
</tr>
  <!--DWLayoutTable-->
  <tr>
    <td width="710" height="581" align="center"><div align="center">
        <form name="form1" method="POST" action="<%=MM_editAction%>">
          <table width="96%" border="0">
            <tr>
              <td><table width="100%" height="341" border="0" align="left" bgcolor="#F2F2F2">
                  <tr>
                    <td width="20%" height="47">&nbsp;</td>
                    <td width="80%">&nbsp;</td>
                  </tr>
                  <tr>
                    <td class="t2"><div align="right">新聞標題：</div></td>
                    <td><input name="new_A1" type="text" id="new_A12" size="30"></td>
                  </tr>
                  <tr>
                    <td class="t2"><div align="right">日期：</div></td>
                    <td><input name="new_A2" type="text" id="new_A22" value="<%=date()%>" size="30"></td>
                  </tr>
                  <tr>
                    <td class="t2"><div align="right">編輯人：</div></td>
                    <td><input name="new_A3" type="text" id="new_A32" size="30"></td>
                  </tr>

                  <tr>
                    <td class="t2"><div align="right">新聞内容：</div></td>
                    <td><table width="100%" height="42" border="0" align="center">
                        <tr>
                          <td>
                    
                            <input type=hidden name=news_content>
                            <iframe ID='eWebEditor1' src="../ewebeditor/EwebEditor.Htm?id=news_content"  frameborder="0"scrolling="No" width="550" height="400"></iframe>
	 </td>
                        </tr>
                    </table></td>
                  </tr>
                  <tr>
                    <td height="67" class="t2">
                      <div align="right"></div></td>
                    <td>
                      <table width="60%" border="0" align="center">
                        <tr>
                          <td width="27%"><div align="center">
                              <input type="submit" name="Submit3" value="添加新聞">
                          </td>
                          <td width="73%"><input name="Submit32" type="button" onClick="location='newsadmin.asp'" value="返回上頁"></td>
                        </tr>
                    </table></td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>
          <input type="hidden" name="MM_insert" value="form1">
        </form>    
        <p>&nbsp;</p>
        <p>&nbsp;</p>
        <p>&nbsp;</p></td>
  </tr>
<A NAME="StranLink"></A>
   </TABLE><SCRIPT LANGUAGE=Javascript SRC="../../images/big5style.js"></SCRIPT>
<br><br>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>

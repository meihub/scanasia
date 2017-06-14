<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%> 
<!--#include file="../Connections/chao.asp"-->
<%if session("admin")="" then
response.Redirect "../login.asp"
end if
%>
<!--#include file="../ScriptLibrary/incPureUpload1.asp" -->
<%
'*** Pure ASP File Upload -----------------------------------------------------
' Copyright (c) 2001-2002 George Petrov, www.UDzone.com
' Process the upload
' Version: 2.0.9
'------------------------------------------------------------------------------
'*** File Upload to: """images/photo""", Extensions: "", Form: form1, Redirect: "", "file", "", "uniq", "true", "", "" , "", "", "", "", "600", "", "300", "100"

Dim GP_redirectPage, RequestBin, UploadQueryString, GP_uploadAction, UploadRequest
PureUploadSetup

If (CStr(Request.QueryString("GP_upload")) <> "") Then
  on error resume next
  Dim reqPureUploadVersion, foundPureUploadVersion
  reqPureUploadVersion = 2.09
  foundPureUploadVersion = getPureUploadVersion()
  if err or reqPureUploadVersion > foundPureUploadVersion then
    Response.Write "<b>You don't have latest version of ScriptLibrary/incPureUpload.asp uploaded on the server.</b><br>"
    Response.Write "This library is required for the current page. It is fully backwards compatible so old pages will work as well.<br>"
    Response.End    
  end if
  on error goto 0
  GP_redirectPage = ""
  Server.ScriptTimeout = 600
  
  RequestBin = Request.BinaryRead(Request.TotalBytes)
  Set UploadRequest = CreateObject("Scripting.Dictionary")  
  BuildUploadRequest RequestBin, """../../ProImages""", "file", "", "uniq"
  
  If (GP_redirectPage <> "" and not (CStr(UploadFormRequest("MM_insert")) <> "" or CStr(UploadFormRequest("MM_update")) <> "")) Then
    If (InStr(1, GP_redirectPage, "?", vbTextCompare) = 0 And UploadQueryString <> "") Then
      GP_redirectPage = GP_redirectPage & "?" & UploadQueryString
    End If
    Response.Redirect(GP_redirectPage)  
  end if  
else
  if UploadQueryString <> "" then
    UploadQueryString = UploadQueryString & "&GP_upload=true"
  else  
    UploadQueryString = "GP_upload=true"
  end if  
end if
' End Pure Upload
'------------------------------------------------------------------------------
%>
<%
' *** Edit Operations: (Modified for File Upload) declare variables

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
If (UploadQueryString <> "") Then
  MM_editAction = MM_editAction & "?" & UploadQueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Insert Record: (Modified for File Upload) set variables

If (CStr(UploadFormRequest("MM_insert")) = "form1") Then

  MM_editConnection = MM_chao_STRING
  MM_editTable = "cc"
  MM_editColumn = "ClassID"
  MM_editRedirectUrl = "fenlei1.asp"
  MM_fieldsStr  = "file|value"
  MM_columnsStr = "Img|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(UploadFormRequest(MM_fields(MM_i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And UploadQueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And UploadQueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & UploadQueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & UploadQueryString
    End If
  End If

End If

%>
<%
' *** Insert Record: (Modified for File Upload) construct a sql insert statement and execute it

Dim MM_tableValues
Dim MM_dbValues

If (CStr(UploadFormRequest("MM_insert")) <> "") Then

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
Dim Recordset1__MMColParam
Recordset1__MMColParam = "2"
If (Request("MM_EmptyValue") <> "") Then 
  Recordset1__MMColParam = Request("MM_EmptyValue")
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
Dim Recordset2__MMColParam
Recordset2__MMColParam = "2"
If (Request("MM_EmptyValue") <> "") Then 
  Recordset2__MMColParam = Request("MM_EmptyValue")
End If
%>
<%
Dim Recordset3
Dim Recordset3_numRows

Set Recordset3 = Server.CreateObject("ADODB.Recordset")
Recordset3.ActiveConnection = MM_chao_STRING
Recordset3.Source = "SELECT * FROM cc"
Recordset3.CursorType = 0
Recordset3.CursorLocation = 2
Recordset3.LockType = 1
Recordset3.Open()

Recordset3_numRows = 0
%>
<html>
<head>
<script language="JavaScript">
<!--

function checkFileUpload(form,extensions,requireUpload,sizeLimit,minWidth,minHeight,maxWidth,maxHeight,saveWidth,saveHeight) { //v2.09
  document.MM_returnValue = true;
  for (var i = 0; i<form.elements.length; i++) {
    field = form.elements[i];
    if (field.type.toUpperCase() != 'FILE') continue;
    checkOneFileUpload(field,extensions,requireUpload,sizeLimit,minWidth,minHeight,maxWidth,maxHeight,saveWidth,saveHeight);
} }

function checkOneFileUpload(field,extensions,requireUpload,sizeLimit,minWidth,minHeight,maxWidth,maxHeight,saveWidth,saveHeight) { //v2.09
  document.MM_returnValue = true;
  if (extensions != '') var re = new RegExp("\.(" + extensions.replace(/,/gi,"|").replace(/\s/gi,"") + ")$","i");
    if (field.value == '') {
      if (requireUpload) {alert('�����ͼƬ·��');document.MM_returnValue = false;field.focus();return;}
    } else {
      if(extensions != '' && !re.test(field.value)) {
        alert('?��2????|��????t���衧��D��a2??��1��|??????��\n??��3?��|????|��????t���衧��D��a??��1����?��?��o ' + extensions + '.\n???????y����?��|��????t���衧��D��a??��');
        document.MM_returnValue = false;field.focus();return;
      }
    document.PU_uploadForm = field.form;
    re = new RegExp(".(gif|jpg|png|bmp|jpeg)$","i");
    if(re.test(field.value) && (sizeLimit != '' || minWidth != '' || minHeight != '' || maxWidth != '' || maxHeight != '' || saveWidth != '' || saveHeight != '')) {
      checkImageDimensions(field,sizeLimit,minWidth,minHeight,maxWidth,maxHeight,saveWidth,saveHeight);
    } }
}

function showImageDimensions(fieldImg) { //v2.09
  var isNS6 = (!document.all && document.getElementById ? true : false);
  var img = (fieldImg && !isNS6 ? fieldImg : this);
  if (img.width > 0 && img.height > 0) {
  if ((img.minWidth != '' && img.minWidth > img.width) || (img.minHeight != '' && img.minHeight > img.height)) {
    alert('��?��?��?��???3?��?��?D?��?\n��?D?��|???a��o ' + img.minWidth + ' x ' + img.minHeight); return;}
  if ((img.maxWidth != '' && img.width > img.maxWidth) || (img.maxHeight != '' && img.height > img.maxHeight)) {
    alert('��?��?��?��???3?��?��?�䨮��?\n��?�䨮��|???a��o ' + img.maxWidth + ' x ' + img.maxHeight); return;}
  if (img.sizeLimit != '' && img.fileSize > img.sizeLimit) {
    alert('��?��?��?��??????t��??y��?�䨮��?\n��?�䨮��|???a��o ' + (img.sizeLimit/1024) + ' KB'); return;}
  if (img.saveWidth != '') document.PU_uploadForm[img.saveWidth].value = img.width;
  if (img.saveHeight != '') document.PU_uploadForm[img.saveHeight].value = img.height;
  document.MM_returnValue = true;
} }

function checkImageDimensions(field,sizeL,minW,minH,maxW,maxH,saveW,saveH) { //v2.09
  if (!document.layers) {
    var isNS6 = (!document.all && document.getElementById ? true : false);
    document.MM_returnValue = false; var imgURL = 'file:///' + field.value.replace(/\\/gi,'/').replace(/:/gi,'|').replace(/"/gi,'').replace(/^\//,'');
    if (!field.gp_img || (field.gp_img && field.gp_img.src != imgURL) || isNS6) {field.gp_img = new Image();
		   with (field) {gp_img.sizeLimit = sizeL*1024; gp_img.minWidth = minW; gp_img.minHeight = minH; gp_img.maxWidth = maxW; gp_img.maxHeight = maxH;
  	   gp_img.saveWidth = saveW; gp_img.saveHeight = saveH; gp_img.onload = showImageDimensions; gp_img.src = imgURL; }
	 } else showImageDimensions(field.gp_img);}
}
//-->
</script><SCRIPT LANGUAGE="JavaScript">
function check()
{


        if(checkspace(document.form1.photo_mc.value)) {
	document.form1.photo_mc.focus();
    alert("����д��Ʒ���ƣ�");
	return false;
  }
  
   
  
              if(checkspace(document.form1.photo_lh.value)) {
	document.form1.photo_lh.focus();
    alert("��ѡ���Ʒ��Ŀ��");
	return false;
  }
  
  
                   if(checkspace(document.form1.smallid.value)) {
	document.form1.smallid.focus();
    alert("��ѡ���Ʒ�ͻ���С�ࣩ");
	return false;
  } 
  

        
 } 





function checkspace(checkstr) {
  var str = '';
  for(i = 0; i < checkstr.length; i++) {
    str = str + ' ';
  }
  return (str == checkstr);
}
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ͼƬ����</title>
<script language=javascript>
<!--
function showSending() {
	sending.style.visibility="visible";
	}
-->
</script>
<SCRIPT language=JavaScript type=text/JavaScript>
function deleteitem(){
	if (confirm('Delete this content item?')) self.location.href='search.asp?kill=55';
}
</SCRIPT>

<link href="../css.css" rel="stylesheet" type="text/css">
<style type=text/css>
a:link {text-decoration: none}
a:hover {text-decoration:none}
a:visited {text-decoration: none}
.style1 {color: #FF0000}
</style>
</head>

<body onload=javascript:LoadContent(); bgcolor="#FFFFFF" leftmargin="1" topmargin="1">
<table width="1000" border="0" cellpadding="0" cellspacing="0">
  <!--DWLayoutTable-->
  <tr bgcolor="f2f2f2"> 
 
    <td height="33" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" background="../images2/products_04.jpg">
        <!--DWLayoutTable-->
        <tr> 
          <td width="333" height="33" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <!--DWLayoutTable-->
              <tr > 
                <td width="333" height="33"><table width="88%" border="0" align="center">
                    <tr> 
                      <td>����λ���ǣ���Ӳ�ƷͼƬ 
                      </td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
          <td width="967"><table width="60%" border="0" align="right">
              <!--tr> 
                <td width="23%"><div align="right"><A HREF="TVPphoto-1.asp">ͼƬ����</A></div></td>
                <td width="24%"><div align="center"><A HREF="TVPphoto-2.asp">���ͼƬ</A></div></td>
              </tr-->
            </table></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="950" height="410"><table width="100%" height="710" border="0" align="center">
        <!--DWLayoutTable-->
        <tr> 
          <td height="940" valign="top"> <div align="center"> 
              <form action="<%=MM_editAction%>" method="POST" enctype="multipart/form-data" name="form1" onSubmit="checkFileUpload(this,'',true,'','','','','','','');return document.MM_returnValue">
                <table width="100%" border="0">
                  <tr> 
                    <td width="21%" class="t2"> <div align="right">��Ʒ��ͼ��</div></td>
                    <td width="77%"> <input name="file" type="file" onChange="checkOneFileUpload(this,'',true,'','','','','','','')">
                   <span class="style1"> (��Ʒ�ߴ�Ϊ91*75)</span>
                    </td>
                  </tr>
                  <tr> 
                    <td class="t2">&nbsp;</td>
                    <td>&nbsp;</td>
                  </tr>
				  <tr> 
                    <td class="t2">&nbsp;</td>
                    <td>&nbsp;</td>
                  </tr>
              <tr align="center">
      <td height="28" bgcolor="#FFFFFF">&nbsp;</td>
      <td height="28" colspan="2" align="left" bgcolor="#FFFFFF">&nbsp;</td>
    </tr>
                  <tr valign="top"> 
                    <td height="447" colspan="2" class="t2"> <div align="right"> 
                        <TABLE cellSpacing=1 cellPadding=2 width="89%" align=center border=0>
                          <TBODY>
                            <TR vAlign=top align=left> 
                              <TD bgColor=#666666 colSpan=2></TD>
                            </TR>
                            <!--TR vAlign=top align=left> 
                              <TD colspan="2" bgColor=#eeeeee> <div align="right"><font size="2" color="#FFFFFF"></font></div>
                                <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                                  <TBODY>
                                    <TR> 
                                      <TD><FONT size=2> 
                                        <INPUT 
            onclick=javascript:htmleditor(); type=radio CHECKED value=0 
            name=whicheditor>
                                        HTML ģʽ 
                                        <INPUT 
            onclick=javascript:texteditor(); type=radio value=1 
            name=whicheditor>
                                        �ı�ģʽ</FONT></TD>
                                    </TR>
                                    <TR id=thetexteditor style="DISPLAY: none"> 
                                      <TD> <FONT size=2><B> 
                                        <textarea id=textarea style="WIDTH: 100%; HEIGHT: 340px" name=content rows=12></textarea>
                                        <INPUT id=autoformat2 
            type=checkbox value=checked name=autoformat>
                                        ȡ���Զ��ֶθ�ʽ</B><BR>
                                        <FONT size=2
							>����ֱ�ӱ༭HTML����ʱ����ѡ�����. </FONT></FONT></TD>
                                    </TR>
                                    <TR id=thehtmleditor> 
                                      <TD> <SCRIPT language=JavaScript src="edit.files/yusasp_ace.js"></SCRIPT> 
                                        <SCRIPT language=JavaScript 
src="edit.files/yusasp_color.js"></SCRIPT> <SCRIPT>
	function SubmitForm()
		{
		//STEP 9: Before getting the edited content, the display mode of the editor 
		//must not in HTML view.
		if(obj1.displayMode == "HTML")
			{
			alert("Please uncheck HTML view")
			return ;
			}
			
		//STEP 10: Here we move the edited content & style into form fields.
		if (document.getElementById("thehtmleditor").style.display!='none'){
			form1.content.value = obj1.getContentBody() //move edited content into form field, to be submitted.
			}
		form1.submit()
		}
		
	function LoadContent()
		{
		obj1.putContent(form1.content.value) 
		}
		
	function transferContent()
		{
		form1.content.value=obj1.getContentBody();
		}
		
	function texteditor(){
		transferContent();
		document.getElementById("thetexteditor").style.display='';
		document.getElementById("thehtmleditor").style.display='none';
	}
	
	function htmleditor(){
		LoadContent();
		document.getElementById("thetexteditor").style.display='none';
		document.getElementById("thehtmleditor").style.display='';
	}
	</SCRIPT> <SCRIPT>
		var obj1 = new ACEditor("obj1") 
			obj1.width = "100%" //editor width
			obj1.height = 360 //editor height
			obj1.ImagePageURL = "selimage.asp" //specify Image library management page
			//obj1.PageStyle = "ace_document.css"; //apply external css to the document
			obj1.PageStylePath_RelativeTo_EditorPath = "style/"; //location of the external css (relative to the editor)
			obj1.isFullHTML = false //edit full HTML (not just BODY)s
			obj1.RUN() //show
	</SCRIPT> </TD>
                                    </TR>
                                  </TBODY>
                                </TABLE></TD>
                            </TR--><script language="JavaScript" type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
                            <TR vAlign=top align=left> 
                              <TD width="35%" align="left" bgColor=#efeff7> <INPUT type=hidden value=55 name=contentid></TD>
                              <TD width="65%" align="left" bgColor=#efeff7> <INPUT id=Button12  onClick="return check()" type="submit" name=Button1 value="��Ӳ�Ʒ"> 
                               &nbsp;<input type="reset" name="Submit4" value="�����趨"></TD>
                            </TR>
							
                            <TR vAlign=top align=left> 
                              <TD bgColor=#666666 colSpan=2></TD>
                            </TR>
                          </TBODY>
                        </TABLE>
                        <p><br>
                          <br>
                        </p>
                      </div></td>
                  </tr>
                </table>
                <input type="hidden" name="MM_insert" value="form1">
              </form>
              <p class="t2">&nbsp;</p>
            </div></td>
        </tr>
      </table></td>
  </tr>
  <tr> <td height="79" colspan="2" valign="middle" align="center" width="1000" bgcolor="#f7f7f7" > 
</td>
  </tr>
</table>
<script language="JavaScript">
<!--
var subcat = new Array();
<%
exec="select * from scc"
set ors=server.createobject("adodb.recordset")
ors.open exec,MM_chao_STRING,1,1
dim n
n=0
do while not ors.eof
%>
subcat[<%=n%>]=new Array('<%=trim(ors("ClassID"))%>','<%=trim(ors("name"))%>','<%=trim(ors("id"))%>')

<%
ors.movenext
n=n+1
loop
ors.close
set ors=nothing
%>


function changeselect(selectValue)
{
document.form1.smallid.length = 0;
document.form1.smallid.options[0] = new Option('��ѡ������˵�','');
for (i=0; i<subcat.length; i++)
{
if (subcat[i][0] == selectValue)
{
document.form1.smallid.options[document.form1.smallid.length] = new Option(subcat[i][1], subcat[i][2]);
}
}
}
-->
</script>

</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>

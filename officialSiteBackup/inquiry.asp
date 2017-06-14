<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp" -->
<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Session("MM_username") <> "") Then 
  Recordset1__MMColParam = Session("MM_username")
End If
%>
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = conn
Recordset1.Source = "SELECT * FROM memberData WHERE m_username = '" + Replace(Recordset1__MMColParam, "'", "''") + "'"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>

<%
Dim MM_paramName 
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
  MM_editTable = "liao"
  MM_editRedirectUrl = "digiBoard.asp"
  MM_fieldsStr  = "liao_mc|value|liao_dz|value|liao_dh|value|liao_tz|value|liao_nr|value"
  MM_columnsStr = "liao_mc|',none,''|liao_dz|',none,''|liao_dh|',none,''|liao_tz|',none,''|liao_nr|',none,''"

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
<LINK href="images/css.css" type=text/css rel=stylesheet>

<SCRIPT src="images/ilank_flash.js" type=text/javascript></SCRIPT>

<META content="MSHTML 6.00.6000.21015" name=GENERATOR></HEAD>
<BODY onLoad="eventLoader()"> 
<DIV id=main>
<DIV id=menu align="center"></DIV>
<DIV id=about2 ></DIV>
<DIV class=x_b align="center"><div class="inq"><table width="95%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
  <td width="200" align="left" valign="top"><table width="95%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td id="news" height="30px" align="left">&nbsp;&nbsp;INQUIRY</td>
    </tr>
    <tr>
      <td background="images/d2.gif" height="1"></td>
    </tr>
  </table></td>
   <td valign="top" align="left"><form action="mailto:scanasia@ms18.hinet.net" method="POST" name="form1" class="box_2">
              
              <table width="100%" height="261" border="0" align="left" cellpadding="0" cellspacing="2">
			  <tr><td valign="top"><table border="0" align="left" cellpadding="0" cellspacing="3">
                <tr> 
                  <td id="nt">
              <div align="left">Company:</div></td>
               </tr>
                <tr> 
                  <td><input name="liao_mc" type="text" id="liao_mc" size="38"></td>
                </tr>
                <tr> <td><table border="0" cellpadding="0" cellspacing="0"><tr>
                  <td width="46%" id="nt">
              <div align="left">First Name</div></td>                  
			  <td width="52%" id="nt">
               <div align="left">Last Name</div></td>
               </tr>
			   <tr> 
                  <td><input name="liao_ch" type="text" id="liao_ch"></td>
                  <td><input name="liao_cs" type="text" id="liao_cs"></td></tr></table></td>
                </tr>
                <tr> 
                  <td id="nt">
                 <div align="left">TEL</div></td>
				 </tr>
				 <tr>
                  <td><input name="liao_dz" type="text" id="liao_dz" size="55"></td>
                </tr>
                 <tr> 
                  <td id="nt">
                 <div align="left">E-MAIL</div></td>
				 </tr>
				 <tr>
                  <td><input name="liao_dz" type="text" id="liao_dz" size="55"></td>
                </tr>
				<tr> 
                  <td id="nt">
               <div align="left">Content</div></td>
			   </tr>
			   <tr>
                  <td ><textarea name="liao_nr" cols="44" rows="3" id="liao_nr"></textarea></td>
                <tr>
                  </table></td>
				<td valign="top"><table border="0" align="left" cellpadding="0" cellspacing="0"><tr><td id="nt"><div align="left">Inquiry Product:</div></td></tr>
				<tr>
				<td>
				<TABLE border=0 align=center cellPadding=0 cellSpacing=0 width="100%">
        <TBODY>
	        <TR>
          <TD><FONT face=Arial><FONT size=1>
		  <INPUT type=checkbox value="PANEL, " name=Inquiry> PANEL
		  <INPUT type=checkbox value="ALUMINIUM COMPOSITE PLYWOOD, " name=Inquiry> ALUMINIUM COMPOSITE PLYWOOD <BR> 
		  <INPUT type=checkbox value="WOODEN PARTITION, " name=Inquiry> WOODEN PARTITION
		  <INPUT type=checkbox value="ACRYLIC SHEET, " name=Inquiry> ACRYLIC SHEET
		    <INPUT type=checkbox value="PS SHEET, " name=Inquiryr> PS SHEET<BR> 
		  <INPUT type=checkbox value="POLYCARBONATE SHEET, " name=Inquiry> POLYCARBONATE SHEET  
		  <INPUT type=checkbox value="CURTAIN ACCESSORIES, " name=Inquiry> CURTAIN ACCESSORIES 
		  <INPUT type=checkbox value="DRAWER SLIDE, " name=Inquiry> DRAWER SLIDE 
          <INPUT type=checkbox value="CONCEALED HINGE, " name=Inquiry> CONCEALED HINGE 
		   <INPUT type=checkbox value="HANDLE, " name=Inquiry> HANDLE 
		    <INPUT type=checkbox value="DRAWER LOCK, " name=Inquiry> DRAWER LOCK 
			 <INPUT type=checkbox value="METAL Leg, " name=Inquiry> METAL Leg 
			 <INPUT type=checkbox value="CONNECTOR, " name=Inquiry> CONNECTOR 
			 <INPUT type=checkbox value="WARDROBE SERIES, " name=Inquiry> WARDROBE SERIES 
			 <INPUT type=checkbox value="CASTOR, " name=Inquiry> CASTOR 
			 <INPUT type=checkbox value="PVC EDGE BANDING & FAST CAP, " name=Inquiry> PVC EDGE BANDING & FAST CAP 
			 <INPUT type=checkbox value="SCREW, " name=Inquiry> SCREW 
			 <INPUT type=checkbox value="DOOR HINGES, " name=Inquiry> DOOR HINGES 
			 <INPUT type=checkbox value="DOOR HANDLE, " name=Inquiry> DOOR HANDLE
			 <INPUT type=checkbox value="BODY LOCK ,CYLINDER AND CYLINDRICAL LOCK, " name=Inquiry> BODY LOCK ,CYLINDER AND CYLINDRICAL LOCK
			 <INPUT type=checkbox value="ALUMINIUM ACCESSORIES, " name=Inquiry> ALUMINIUM ACCESSORIES
			 <INPUT type=checkbox value="FURNITURE DOOR CLOSER SERISE, " name=Inquiry> FURNITURE DOOR CLOSER SERISE
			 <INPUT type=checkbox value="OTHERS, " name=Inquiry> OTHERS
 
			 
            </FONT></FONT></TD></TR></TBODY></TABLE>
				</td>
				</tr>
				<tr> 
                  <td id="nt">
                 <div align="left">Model Number</div></td>
				 </tr>
				 <tr>
                  <td><input name="liao_dz" type="text" id="liao_dz" size="60"></td>
                </tr>
                <tr>
                  <td  height="50" align="right"><table width="50%" border="0">
                      <tr> 
                        <td width="50%" align="center"><input type="submit" name="Submit" value="Submit"></td>
                        <td width="50%"><input type="reset" name="Submit2" value="Reset"></td>
                      </tr>
                    </table>
                    </td>
                </tr></table></td></tr>
              </table>
              <input type="hidden" name="MM_insert" value="form1">
            </form> </td>  
  </tr>
</table></div>
</DIV>
<SCRIPT type=text/javascript>
var objFlash = new ilankFlash("menu/inquiry.swf","","960","157","9","","false","high");
objFlash.addParam("wmode","transparent");
objFlash.write("menu");
</SCRIPT>
<SCRIPT type=text/javascript>
var objFlash = new ilankFlash("about2.swf","","980","324","9","","false","high");
objFlash.addParam("wmode","transparent");
objFlash.write("about2");
</SCRIPT>


</DIV><DIV id=footer><SPAN class=f1>Copyright @Scan Asia Corporation. 2013. All Rights Reserved </SPAN> 
</DIV></BODY></HTML>

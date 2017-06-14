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
<BODY onload="eventLoader()"> 
<DIV id=main>
<DIV id=logo></DIV>
<!--DIV id=menu>
<UL>
  <LI><SPAN>Concealed hinge</SPAN><!--[if lte IE 6]><a href="#nogo"><table><tr><td><![endif]--> 
  <!--DL>
    <DT><A href="products_l.asp?ClassID=1">Concealed hinge</A> 
    <DD><A title="Mini concealed hinge" href="products_x.asp?scc=27">Mini 
    concealed hinge</A> 
    <DD><A title="concealed hinge for glass door" 
    href="products_x.asp?scc=32">concealed hinge for glass door</A> 
    <DD><A title="one way concealed hinge" 
    href="products_x.asp?scc=33">one way concealed hinge</A> 
    <DD><A title="two way concealed hinge" 
    href="products_x.asp?scc=36">two way concealed hinge</A> 
    <DD><A title="clip on concealed hinge" 
    href="products_x.asp?scc=37">clip on concealed hinge</A> 
    <DD><A title="175 degree concealed hinge" 
    href="products_x.asp?scc=38">175 degree concealed hinge</A> 
    <DD><A title="concealed hinge for metal material" 
    href="products_x.asp?scc=40">concealed hinge for metal material</A> 
    <DD><A title="hydraulic buffering concealed hinge" 
    href="products_x.asp?scc=46">hydraulic buffering concealed hinge</A> 
    <DD><A title="special concealed hinge" 
    href="products_x.asp?scc=47">special concealed hinge</A> </DD></DL><!--[if lte IE 6]></td></tr></table></a><![endif]--><!--/LI>
  <LI><SPAN>Drawer slide</SPAN><!--[if lte IE 6]><a href="#nogo"><table><tr><td><![endif]--> 
  <!--DL>
    <DT><A href="products_l.asp?ClassID=2">Drawer slide</A> 
    <DD><A title="self closing slide" href="products_x.asp?scc=68">self closing slide</A> 
    <DD><A title="Metal box slide" href="products_x.asp?scc=69">Metal box slide</A> 
    <DD><A title="Clamp on slide" href="products_x.asp?scc=70">Clamp on slide</A> 
    <DD><A title="17mm mini slide" href="products_x.asp?scc=71">17mm mini slide</A> 
    <DD><A title="27mm single extension slide" href="products_x.asp?scc=72">27mm single extension slide</A> 
    <DD><A title="27mm keyboard slide" href="products_x.asp?scc=73">27mm keyboard slide</A> 
    <DD><A title="27mm keyboard slide for PVC" 
    href="products_x.asp?scc=74">27mm keyboard slide for PVC</A> 
    <DD><A title="35mm flipper door slide" 
    href="products_x.asp?scc=75">35mm flipper door slide</A> 
    <DD><A title="35mm single extension slide" 
    href="products_x.asp?scc=76">35mm single extension slide</A> 
    <DD><A title="35mm full extension slide" 
    href="products_x.asp?scc=77">35mm full extension slide</A> 
    <DD><A title="45mm full extension slide" 
    href="products_x.asp?scc=78">45mm full extension slide</A> 
    <DD><A title="45mm bayonet mounting full extension slide" 
    href="products_x.asp?scc=79">45mm bayonet mounting full extension slide</A> 
    <DD><A title="Concealed under mount slide" 
    href="products_x.asp?scc=80">Concealed under mount slide</A> 
    <DD><A title="TV swivel slide" href="products_x.asp?scc=81">TV swivel slide</A> </DD></DL><!--[if lte IE 6]></td></tr></table></a><![endif]--><!--/LI>
  <!--LI><SPAN>Furniture fitting</SPAN><!--[if lte IE 6]><a href="#nogo"><table><tr><td><![endif]--> 
  <!--DL>
    <DT><A href="products_l.asp?ClassID=3">Furniture fitting</A> 
    <DD><A title="Magnetic catch and latch" 
    href="products_x.asp?scc=90">Magnetic catch and latch</A> 
    <DD><A title="Glass door fitting" href="products_x.asp?scc=91">Glass door fitting</A> 
    <DD><A title="Connecting fitting and connecting screw" 
    href="products_x.asp?scc=92">Connecting fitting and connecting screw</A> 
    <DD><A title="Knock-down Fitting" 
    href="products_x.asp?scc=93">Knock-down Fitting</A> 
    <DD><A title="Self support" href="products_x.asp?scc=94">Self support</A> 
    <DD><A title="Furniture glide" href="products_x.asp?scc=95">Furniture glide</A> 
    <DD><A title="Rail and rail support" 
    href="products_x.asp?scc=96">Rail and rail support</A> 
    <DD><A title="Drawer lock" href="products_x.asp?scc=97">Drawer lock</A> 
    <DD><A title="Frame, mirror and glass fitting" 
    href="products_x.asp?scc=98">Frame, mirror and glass fitting</A> 
    <DD><A title="Bracket and bed fitting" 
    href="products_x.asp?scc=99">Bracket and bed fitting</A> 
    <DD><A title="Furniture bolt and tower bolt" 
    href="products_x.asp?scc=100">Furniture bolt and tower bolt</A> 
    <DD><A title="Furniture door closer series" 
    href="products_x.asp?scc=101">Furniture door closer series</A> 
    <DD><A title="Wardrobe Series" href="products_x.asp?scc=102">Wardrobe Series</A> </DD></DL><!--[if lte IE 6]></td></tr></table></a><![endif]--><!--/LI>
  <!--LI><SPAN>Sofa &amp; table Leg</SPAN><!--[if lte IE 6]><a href="#nogo"><table><tr><td><![endif]--> 
  <!--DL>
    <DT><A href="products_l.asp?ClassID=4">Sofa &amp; table Leg</A> 
    <DD><A title="Zamac Leg" href="products_x.asp?scc=157">Zamac Leg</A> 
    <DD><A title="Iron Leg" href="products_x.asp?scc=158">Iron Leg</A> 
    <DD><A title="Aluminum Leg" href="products_x.asp?scc=159">Aluminum Leg</A> 
    <DD><A title="Wood Leg" href="products_x.asp?scc=160">Wood Leg</A> 
    <DD><A title="Stainless Steel Leg" 
    href="products_x.asp?scc=161">Stainless Steel Leg</A> </DD></DL><!--[if lte IE 6]></td></tr></table></a><![endif]--><!--/LI>
  <!--LI><SPAN>Furniture handle</SPAN><!--[if lte IE 6]><a href="#nogo"><table><tr><td><![endif]--> 
  <!--DL>
    <DT><A href="products_l.asp?ClassID=5">Furniture handle</A> 
    <DD><A title="Wood handle" href="products_x.asp?scc=119">Wood handle</A> 
    <DD><A title="Zinc alloy handle" href="products_x.asp?scc=120">Zinc alloy handle</A> 
    <DD><A title="Plastic handle" href="products_x.asp?scc=121">Plastic handle</A> 
    <DD><A title="Ceramic handle" href="products_x.asp?scc=122">Ceramic handle</A> 
    <DD><A title="Aluminum handle" href="products_x.asp?scc=123">Aluminum handle</A> 
    <DD><A title="Steel handle" href="products_x.asp?scc=124">Steel handle</A> 
    <DD><A title="Stainless Steel handle(new)" 
    href="products_x.asp?scc=125">Stainless Steel handle(new)</A> </DD></DL><!--[if lte IE 6]></td></tr></table></a><![endif]--><!--/LI>
  <!--LI><SPAN>Hardware</SPAN><!--[if lte IE 6]><a href="#nogo"><table><tr><td><![endif]--> 

  <!--DL>
    <DT><A href="products_l.asp?ClassID=6">Hardware</A> 
    <DD><A title="Door Hinge" href="products_x.asp?scc=126">Door Hinge</A> 

    <DD><A title="Door stopper" href="products_x.asp?scc=127">Door stopper</A> 
    <DD><A title=Caster href="products_x.asp?scc=128">Caster</A> 
    <DD><A title=Others href="products_x.asp?scc=129">Others</A> </DD></DL><!--[if lte IE 6]></td></tr></table></a><![endif]--><!--/LI>
  <LI><SPAN>Door Handle &amp; lock set</SPAN><!--[if lte IE 6]><a href="#nogo"><table><tr><td><![endif]--> 
  <!--DL>
    <DT><A href="products_l.asp?ClassID=7">Door Handle &amp; lock set</A> 
    <DD><A title="Zinc alloy, Aluminum handle" 
    href="products_x.asp?scc=130">Zinc alloy, Aluminum handle</A> 
    <DD><A title="steel handle" href="products_x.asp?scc=131">steel handle</A> 
    <DD><A title="stainless steel handle" 
    href="products_x.asp?scc=132">stainless steel handle</A> 
    <DD><A title="body lock" href="products_x.asp?scc=133">body lock</A> 
    <DD><A title=cylinder href="products_x.asp?scc=134">cylinder</A> 
    <DD><A title="cylindrical lock" 
    href="products_x.asp?scc=135">cylindrical lock</A> </DD></DL><!--[if lte IE 6]></td></tr></table></a><![endif]--><!--/LI>
  <!--LI><SPAN>Decorative product</SPAN><!--[if lte IE 6]><a href="#nogo"><table><tr><td><![endif]--> 
  <!--DL>
    <DT><A href="products_l.asp?ClassID=8">Decorative product</A> 
    <DD><A title="curtain rod" href="products_x.asp?scc=136">curtain rod</A> 
    <DD><A title="curtain track" href="products_x.asp?scc=137">curtain track</A> 
    <DD><A title="curtain hanger and accessory" 
    href="products_x.asp?scc=138">curtain hanger and accessory</A> 
    <DD><A title="PVC edge banding" href="products_x.asp?scc=139">PVC edge banding</A> 
    <DD><A title="Wooden veneer" href="products_x.asp?scc=140">Wooden veneer</A> 
    <DD><A title="PANEL" href="products_x.asp?scc=141">PANEL</A> 
    <DD><A title="Poly carbonate sheet" href="products_x.asp?scc=142">Poly carbonate sheet</A> 
    <DD><A title="Acrylic Sheet" href="products_x.asp?scc=143">Acrylic Sheet</A> </DD></DL><!--[if lte IE 6]></td></tr></table></a><![endif]--><!--/LI>
  <LI><SPAN>Fastener</SPAN><!--[if lte IE 6]><a href="#nogo"><table><tr><td><![endif]--> 

  <!--DL>
    <DT><A href="products_l.asp?ClassID=9">Fastener</A> 
    <DD><A title="Self Tapping Screw w/struss head" 
    href="products_x.asp?scc=144">Self Tapping Screw w/struss head</A> 
    <DD><A title="Dry Wall Screw" href="products_x.asp?scc=145">Dry Wall Screw</A> 
    <DD><A title="Dry Wall Screw" href="products_x.asp?scc=146">Clipboard screw</A> 
    <DD><A title="Wood Screw" href="products_x.asp?scc=147">Wood Screw</A> 

    <DD><A title="Machine Screw" href="products_x.asp?scc=148">Machine Screw</A> 
    <DD><A title="Confirmat Screw" href="products_x.asp?scc=149">Confirmat Screw</A> 
    <DD><A title="Break machine screw" href="products_x.asp?scc=150">Break machine screw</A> 
    <DD><A title="Toggle bolt" href="products_x.asp?scc=151">Toggle bolt</A> 
    <DD><A title="Frame Anchor w/nylon and metal" 
    href="products_x.asp?scc=152">Frame Anchor w/nylon and metal</A> 
    <DD><A title="Lag screw" href="products_x.asp?scc=153">Lag screw</A> 
    <DD><A title="Self Drilling Screw" href="products_x.asp?scc=154">Self Drilling Screw</A> 
    <DD><A title="Wall plug" href="products_x.asp?scc=155">Wall plug</A> 
    <DD><A title=Anchor href="products_x.asp?scc=156">Anchor</A> </DD></DL><!--[if lte IE 6]></td></tr></table></a><![endif]--><!--/LI></UL></DIV-->
<DIV class=menu_deputy>
<UL>
  <LI><a href="index.htm">HOME</a> 
  <LI><a href="index.htm">ABOUT US</a>  
  <LI><a href="news.asp">EVENTS</a>  
  <LI><a href="products.asp">PRODUCTS</a>  
  <LI><SPAN><a href="inquiry.asp" target="_blank">INQUIRY</a></SPAN>  
  <LI><a href="contact.htm">CONTACT US</a> </LI>
</UL></DIV>
<DIV id=about2 align="center"></DIV>
<DIV class=x_b align="center"><div class="inq"><table width="95%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
  <td width="250" align="left" valign="top"><table width="60%"  border="0" cellspacing="0" cellpadding="0">
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
   <td valign="top" align="left"><form action="<%=MM_editAction%>" method="POST" name="form1" class="box_2">
              
              <table width="50%"  border="0" cellspacing="0" cellpadding="0" >
          <tr>
            <td>&nbsp;</td>
          </tr>
          <tr>
                <td valign="top"> 
                  <div align="center">
                    <p><strong><font color="#333333">You have submited you informations to our database</font></strong></p>
                    <p><strong><font color="#333333"> we will reply ASAP! Thanks.</font></strong></p>
                  </div></td>
              </tr>
          
        </table>
              <input type="hidden" name="MM_insert" value="form1">
            </form> </td>  
  </tr>
</table></div>
</DIV>
<SCRIPT type=text/javascript>
var objFlash = new ilankFlash("about2.swf","","980","324","9","","false","high");
objFlash.addParam("wmode","transparent");
objFlash.write("about2");
</SCRIPT>


</DIV><DIV id=footer><SPAN class=f1>Copyright @Scan Asia Corporation. 2013. All Rights Reserved </SPAN> 
</DIV></BODY></HTML>

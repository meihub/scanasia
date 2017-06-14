<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!-- saved from url=(0058)http://www.xigla.com/absolutecr/xlaabsolutecr/selimage.asp -->
<HTML>
<HEAD>
<TITLE>Insert/Update Image</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<STYLE>BODY {
	FONT-SIZE: 9pt}
TABLE {
	FONT-SIZE: 9pt}
INPUT {
font-size: 9pt}
SELECT {
TOP: 2px; font-size: 9pt}
.bar {
	BORDER-TOP: #99ccff 1px solid; BACKGROUND: #336699; WIDTH: 100%; BORDER-BOTTOM: #000000 1px solid; HEIGHT: 20px
}
</STYLE>
<SCRIPT language=JavaScript type=text/JavaScript>
function alerta(){
	alert('This option is disabled on this demo');
}</SCRIPT>
<META content="MSHTML 6.00.2600.0" name=GENERATOR>
</HEAD>
<BODY bottomMargin=5 vLink=mediumslateblue aLink=mediumslateblue link=blue 
bgColor=#efeff7 leftMargin=5 topMargin=5 onload=checkImage() rightMargin=5>
<TABLE cellSpacing=0 cellPadding=0 border=0>
  <TBODY> 
  <TR> 
    <TD vAlign=top> 
      <!-- Content -->
      <TABLE cellSpacing=3 cellPadding=3 align=center border=0>
        <TBODY> 
        <TR> 
          <TD 
          style="BORDER-RIGHT: #336699 1px solid; BORDER-TOP: #336699 1px solid; BORDER-LEFT: #336699 1px solid; BORDER-BOTTOM: #336699 1px solid" 
          align=middle bgColor=white valign="top"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="5" align="center">
              <tr>
                <td>
                  <div align="center">图片预览</div>
                </td>
              </tr>
            </table>
            <DIV id=divImg 
            style="OVERFLOW: auto; WIDTH: 150px; HEIGHT: 170px"></DIV>
          </TD>
          <TD vAlign=top> 
            <TABLE cellSpacing=0 cellPadding=0 width=262 border=0>
              <TBODY> 
              <TR> 
                <TD> 
                  <DIV class=bar style="PADDING-LEFT: 5px"><FONT 
                  color=white size=2>文件名</FONT> </DIV>
                </TD>
              </TR>
              </TBODY> 
            </TABLE>
            <DIV 
            style="BORDER-RIGHT: lightsteelblue 1px solid; OVERFLOW: auto; BORDER-LEFT: #316ac5 1px solid; WIDTH: 26
						0px; BORDER-BOTTOM: lightsteelblue 1px solid; HEIGHT: 150px"> 
              <TABLE cellSpacing=0 cellPadding=3 width=260 border=0>
                <TBODY> 
                <TR bgColor=gainsboro> 
                  <TD vAlign=top width="157">1111.jpg</TD>
                  <TD vAlign=top width="35">8 kb</TD>
                  <TD style="CURSOR: hand" 
                onclick="selectImage('content/1111.jpg')" 
                  vAlign=top width="36"><U><FONT color=blue>选择</FONT></U></TD>
                  <TD style="CURSOR: hand" onclick=alerta(); vAlign=top width="44"><U><FONT color=blue>删除</FONT></U></TD>
                </TR>
                <TR bgColor=gainsboro> 
                  <TD vAlign=top width="157">1112.jpg</TD>
                  <TD vAlign=top width="35">9 kb</TD>
                  <TD style="CURSOR: hand" 
                onclick="selectImage('content/1112.jpg')" 
                  vAlign=top width="36"><U><FONT color=blue>选择</FONT></U></TD>
                  <TD style="CURSOR: hand" onclick=alerta(); vAlign=top width="44"><U><FONT color=blue>删除</FONT></U></TD>
                </TR>
                <TR bgColor=gainsboro> 
                  <TD vAlign=top width="157">1113.jpg</TD>
                  <TD vAlign=top width="35">3 kb</TD>
                  <TD style="CURSOR: hand" 
                onclick="selectImage('content/1113.jpg')" 
                vAlign=top width="36"><U><FONT color=blue>选择</FONT></U></TD>
                  <TD style="CURSOR: hand" onclick=alerta(); vAlign=top width="44"><U><FONT color=blue>删除</FONT></U></TD>
                </TR>
                <TR bgColor=gainsboro> 
                  <TD vAlign=top width="157">1114.jpg</TD>
                  <TD vAlign=top width="35">3 kb</TD>
                  <TD style="CURSOR: hand" 
                onclick="selectImage('content/1114.jpg')" 
                vAlign=top width="36"><U><FONT color=blue>选择</FONT></U></TD>
                  <TD style="CURSOR: hand" onclick=alerta(); vAlign=top width="44"><U><FONT color=blue>删除</FONT></U></TD>
                </TR>
                <TR bgColor=gainsboro> 
                  <TD vAlign=top width="157">1115.jpg</TD>
                  <TD vAlign=top width="35">3 kb</TD>
                  <TD style="CURSOR: hand" 
                onclick="selectImage('content/1115.jpg')" 
                vAlign=top width="36"><U><FONT color=blue>选择</FONT></U></TD>
                  <TD style="CURSOR: hand" onclick=alerta(); vAlign=top width="44"><U><FONT color=blue>删除</FONT></U></TD>
                </TR>
                <TR bgColor=gainsboro> 
                  <TD vAlign=top width="157">1116.jpg</TD>
                  <TD vAlign=top width="35">3 kb</TD>
                  <TD style="CURSOR: hand" 
                onclick="selectImage('content/1116.jpg')" 
                vAlign=top width="36"><U><FONT color=blue>选择</FONT></U></TD>
                  <TD style="CURSOR: hand" onclick=alerta(); vAlign=top width="44"><U><FONT color=blue>删除</FONT></U></TD>
                </TR>
                <TR bgColor=gainsboro> 
                  <TD vAlign=top width="157">1117.jpg</TD>
                  <TD vAlign=top width="35">9 kb</TD>
                  <TD style="CURSOR: hand" 
                onclick="selectImage('content/1117.jpg')" 
                  vAlign=top width="36"><U><FONT color=blue>选择</FONT></U></TD>
                  <TD style="CURSOR: hand" onclick=alerta(); vAlign=top width="44"><U><FONT color=blue>删除</FONT></U></TD>
                </TR>
                </TBODY> 
              </TABLE>
            </DIV>
            <FORM id=form1 name=form1 action=selimage.asp?action=upload 
            method=post encType=multipart/form-data>
              上传图片: <BR>
              <INPUT 
            id=inpFile style="FONT: 9pt verdana,arial,sans-serif" type=file 
            size=22 name=inpFile>
              <INPUT onclick=＃; type=submit value=上传 name="Submit">
            </FORM>
          </TD>
        </TR>
        <TR> 
          <TD colSpan=2> 
            <HR>
            <TABLE cellSpacing=1 cellPadding=3 width=340 border=0>
              <TBODY> 
              <TR> 
                <TD width="66">图片来源:</TD>
                <TD colSpan=3> 
                  <INPUT id=inpImgURL size=39 name=inpImgURL>
                  <!--<font color=red>(you can type your own image path here)</font>-->
                </TD>
              </TR>
              <TR> 
                <TD width="66">提示信息:</TD>
                <TD colSpan=3> 
                  <INPUT id=inpImgAlt size=39 
name=inpImgAlt>
                </TD>
              </TR>
              <TR> 
                <TD width="66">对齐方式</TD>
                <TD width="86"> 
                  <SELECT id=inpImgAlign name=inpImgAlign>
                    <OPTION 
                    value="" selected>&lt;Not Set&gt;</OPTION>
                    <OPTION 
                    value=absBottom>absBottom</OPTION>
                    <OPTION 
                    value=absMiddle>absMiddle</OPTION>
                    <OPTION 
                    value=baseline>baseline</OPTION>
                    <OPTION 
                    value=bottom>bottom</OPTION>
                    <OPTION 
                    value=left>left</OPTION>
                    <OPTION 
                    value=middle>middle</OPTION>
                    <OPTION 
                    value=right>right</OPTION>
                    <OPTION 
                    value=textTop>textTop</OPTION>
                    <OPTION 
                  value=top>top</OPTION>
                  </SELECT>
                </TD>
                <TD width="62">图像边框:</TD>
                <TD width="111"> 
                  <SELECT id=inpImgBorder name=inpImgBorder>
                    <OPTION 
                    value=0 selected>0</OPTION>
                    <OPTION value=1>1</OPTION>
                    <OPTION value=2>2</OPTION>
                    <OPTION value=3>3</OPTION>
                    <OPTION value=4>4</OPTION>
                    <OPTION value=5>5</OPTION>
                  </SELECT>
                  像素 </TD>
              </TR>
              <TR> 
                <TD width="66">图片宽度:</TD>
                <TD width="86"> 
                  <INPUT id=inpImgWidth size=2 name=inpImgWidth>
                  像素 </TD>
                <TD width="62">水平间距:</TD>
                <TD width="111"> 
                  <INPUT id=inpHSpace size=2 name=inpHSpace>
                  像素 </TD>
              </TR>
              <TR> 
                <TD width="66">图片高度:</TD>
                <TD width="86"> 
                  <INPUT id=inpImgHeight size=2 name=inpImgHeight>
                  像素 </TD>
                <TD width="62">垂直间距:</TD>
                <TD width="111"> 
                  <INPUT id=inpVSpace size=2 
            name=inpVSpace>
                  像素 </TD>
              </TR>
              </TBODY> 
            </TABLE>
          </TD>
        </TR>
        <TR> 
          <TD align=middle colSpan=2> 
            <TABLE cellSpacing=0 cellPadding=0 align=center>
              <TBODY> 
              <TR> 
                <TD> 
                  <INPUT id=Button1 style="FONT: 8pt verdana,arial,sans-serif; HEIGHT: 22px" onclick=self.close(); type=button value=取消 name=Button1>
                </TD>
                <TD><SPAN id=btnImgInsert style="DISPLAY: none"> 
                  <INPUT id=Button2 style="FONT: 8pt verdana,arial,sans-serif; HEIGHT: 22px" onclick=InsertImage();self.close(); type=button value=嵌入 name=Button2>
                  </SPAN><SPAN id=btnImgUpdate style="DISPLAY: none"> 
                  <INPUT id=Button3 style="FONT: 8pt verdana,arial,sans-serif; HEIGHT: 22px" onclick=UpdateImage();self.close(); type=button value=更新 name=Button3>
                  </SPAN></TD>
              </TR>
              </TBODY> 
            </TABLE>
          </TD>
        </TR>
        </TBODY> 
      </TABLE>
      <!-- /Content -->
      <BR>
    </TD>
  </TR>
  </TBODY> 
</TABLE>
<SCRIPT language=JavaScript>
function deleteImage(sURL)
	{
	if (confirm("Delete this document ?") == true) 
		{
		window.navigate("selimage.asp?action=del&file="+sURL+"&catid="+form2.catid.value);
		}
	}
function selectImage(sURL)
	{
	inpImgURL.value = sURL;
	
	divImg.style.visibility = "hidden"
	divImg.innerHTML = "<img id='idImg' src='" + sURL + "'>";
	

	var width = idImg.width
	var height = idImg.height 
	var resizedWidth = 150;
	var resizedHeight = 170;

	var Ratio1 = resizedWidth/resizedHeight;
	var Ratio2 = width/height;

	if(Ratio2 > Ratio1)
		{
		if(width*1>resizedWidth*1)
			idImg.width=resizedWidth;
		else
			idImg.width=width;
		}
	else
		{
		if(height*1>resizedHeight*1)
			idImg.height=resizedHeight;
		else
			idImg.height=height;
		}
	
	divImg.style.visibility = "visible"
	}

/***************************************************
	If you'd like to use your own Image Library :
	- use InsertImage() method to insert image
		Params : url,alt,align,border,width,height,hspace,vspace
	- use UpdateImage() method to update image
		Params : url,alt,align,border,width,height,hspace,vspace
	- use these methods to get selected image properties :
		imgSrc()
		imgAlt()
		imgAlign()
		imgBorder()
		imgWidth()
		imgHeight()
		imgHspace()
		imgVspace()
		
	Sample uses :
		window.opener.obj1.InsertImage(...[params]...)
		window.opener.obj1.UpdateImage(...[params]...)
		inpImgURL.value = window.opener.obj1.imgSrc()
	
	Note: obj1 is the editor object.
	We use window.opener since we access the object from the new opened window.
	If we implement more than 1 editor, we need to get first the current 
	active editor. This can be done using :
	
		oName=window.opener.oUtil.oName // return "obj1" (for example)
		obj = eval("window.opener."+oName) //get the editor object
		
	then we can use :
		obj.InsertImage(...[params]...)
		obj.UpdateImage(...[params]...)
		inpImgURL.value = obj.imgSrc()
		
***************************************************/	
function checkImage()
	{
	oName=window.opener.oUtil.oName
	obj = eval("window.opener."+oName)
	
	if (obj.imgSrc()!="") selectImage(obj.imgSrc())//preview image
	inpImgURL.value = obj.imgSrc()
	inpImgAlt.value = obj.imgAlt()
	inpImgAlign.value = obj.imgAlign()
	inpImgBorder.value = obj.imgBorder()
	inpImgWidth.value = obj.imgWidth()
	inpImgHeight.value = obj.imgHeight()
	inpHSpace.value = obj.imgHspace()
	inpVSpace.value = obj.imgVspace()

	if (obj.imgSrc()!="") //If image is selected 
		btnImgUpdate.style.display="block";
	else
		btnImgInsert.style.display="block";
	}
function UpdateImage()
	{
	oName=window.opener.oUtil.oName
	eval("window.opener."+oName).UpdateImage(inpImgURL.value,inpImgAlt.value,inpImgAlign.value,inpImgBorder.value,inpImgWidth.value,inpImgHeight.value,inpHSpace.value,inpVSpace.value);	
	}
function InsertImage()
	{
	oName=window.opener.oUtil.oName
	eval("window.opener."+oName).InsertImage(inpImgURL.value,inpImgAlt.value,inpImgAlign.value,inpImgBorder.value,inpImgWidth.value,inpImgHeight.value,inpHSpace.value,inpVSpace.value);
	}	
/***************************************************/
</SCRIPT>
<INPUT id=inpActiveEditor contentEditable=true style="DISPLAY: none" 
name=inpActiveEditor> </BODY></HTML>

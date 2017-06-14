/**********************************************************
Advanced Content Editor - Lite
Copyright © 2001-2003, Yusuf Wiryonoputro. All rights reserved.
***********************************************************/
var oUtil = new ACEUtilities();
function ACEUtilities()
	{
	this.ACEObjects = "";
	this.oName;
	this.onloadFunction = "";
	}

/*** B O X  D R A W ***/
function boxDrawDrop(width,content)
	{
	var sHTML = "" +
		"<style>" +
		"		body {border:lightgrey 0px solid;background: white;}" +
		"		td {font:8pt verdana,arial,sans-serif}" +
		"		.dropdown {cursor:hand}" +
		"</style>" +
		"<body topmargin=0 leftmargin=0 rightmargin=0 bottommargin=0 onselectstart=\"return event.srcElement.tagName=='INPUT'\" oncontextmenu='return false'>" +
		"		<table border=0 cellpadding=1 cellspacing=0 style='table-layout: fixed;border-right:#c3c3c3 1 solid;border-bottom:#c3c3c3 1 solid;border-left:#aeaeae 1 solid;border-top:#aeaeae 1 solid;' ID=tblPopup>" +
		"		<col width="+width+">" +
		"		<tr>" +
		"		<td>" +
		content +
		"		</td>" +
		"		</tr>" +
		"		</table>" +
		"<input type=text style='display:none;' id='inpActiveEditor' name='inpActiveEditor' contentEditable=true>" +
		"</body>";
	return sHTML;
	}

function boxDrawPop(width,title,content)
	{
	var sHTML = "" +
	"<style>" +
	"		select{height: 22px; top:2;	font:8pt verdana,arial,sans-serif}" +
	"		body {border:lightgrey 0px solid;background: #ece9d8;filter: progid:DXImageTransform.Microsoft.gradient(startColorstr=white, endColorstr=#e3e3e3);}" +
	"		td {font:8pt verdana,arial,sans-serif}" +
	"		.dropdown {cursor:hand}" +
	"		.bar{padding-left: 5px;border-top: #99ccff 1px solid; background: #004684;filter: progid:DXImageTransform.Microsoft.gradient(startColorstr=#004684, endColorstr=#7189b7); WIDTH: 100%; border-bottom: #004684 1px solid;height: 20px}" +
	"		.bar2{border-top: #99ccff 1px solid; background: #004684;filter: progid:DXImageTransform.Microsoft.gradient(startColorstr=#004684, endColorstr=#7189b7); WIDTH: 100%; border-bottom: #004684 1px solid;height: 20px}" +
	"		div	{	font:10pt tahoma,arial,sans-serif}" +
	"</style>" +
	"<body onload='' topmargin=0 leftmargin=0 rightmargin=0 bottommargin=0 onselectstart=\"return event.srcElement.tagName=='INPUT'\" oncontextmenu='return false'>" +
	"<table border=0 cellpadding=0 cellspacing=0 style='table-layout: fixed' ID=tblPopup>" +
	"<col width="+width+"><col width=13>" +
	"<tr>" +
	"<td>" +
	"		<div id='popup_BarArea' class=bar>" +
	"		<font size=2 face=tahoma color=white><b>"+title+"</b></font>" +
	"		</div>" +
	"</td>" +
	"<td style='cursor:hand' onclick=\"eval('parent.'+inpActiveEditor.value).boxHide()\">" +
	"		<div id='popup_BarArea_close' class=bar2>" +
	"		<font size=2 face=tahoma color=white><b>X</b></font>" +
	"		</div>" +
	"</td>" +
	"</tr>" +
	"<tr>" +
	"<td id='popup_ContentBorder' colspan=2 style='border-left: #336699 1px solid;border-right: #336699 1px solid;border-bottom: #336699 1px solid;' valign=top>" +
	"		<br>" +
	"		<div id=divPopup align=center>" + content +
	"		</div>" +
	"		<br>" +
	"</td>" +
	"</tr>" +
	"</table>" +
	"<input type=text style='display:none;' id='inpActiveEditor' name='inpActiveEditor' contentEditable=true>" +
	"</body>";
	return sHTML;
	}

/*** B O X   S H A R E D   F U N C T I O N S ***/
function saveSelection()
	{
	var oEditor = eval("idContent"+this.oName);
	this.Sel = oEditor.document.selection.createRange();
	this.SelType = oEditor.document.selection.type;
	}

function boxDimension(boxName)
	{
	var tblPopup = document.frames(boxName).document.body.document.all("tblPopup");
	eval("document.all."+boxName).style.width = tblPopup.offsetWidth;
	eval("document.all."+boxName).style.height = tblPopup.offsetHeight;
	}

function boxPosition(boxName)
	{
	var tblPopup = document.frames(boxName).document.body.document.all("tblPopup");
	var oArea = eval("idArea" + this.oName);

	myTop = 0; stmp = "";
	while(eval("idArea"+ this.oName + stmp).tagName!="BODY")
		{
		myTop += eval("idArea"+ this.oName + stmp).offsetTop;
		stmp += ".offsetParent";
		}

	myLeft = 0;	stmp = "";
	while(eval("idArea"+ this.oName + stmp).tagName!="BODY")
		{
		myLeft += eval("idArea"+ this.oName + stmp).offsetLeft;
		stmp += ".offsetParent";
		}

	if(oArea.offsetHeight-tblPopup.offsetHeight > 0) 
		eval("document.all."+boxName).style.pixelTop=(myTop + (oArea.offsetHeight-tblPopup.offsetHeight)/2);	
	else 
		eval("document.all."+boxName).style.pixelTop=(myTop + (oArea.offsetHeight-tblPopup.offsetHeight)/2);

	if(oArea.offsetWidth-tblPopup.offsetWidth > 0)
		eval("document.all."+boxName).style.pixelLeft=(myLeft + (oArea.offsetWidth-tblPopup.offsetWidth)/2);
	else
		eval("document.all."+boxName).style.pixelLeft=myLeft;
	}

function setPosition(idImg,boxName)//for dropdown
	{
	myTop = 0; stmp = "";
	while(eval("idImg" + stmp).tagName!="BODY")
		{stmp += ".offsetParent"; myTop += eval("idImg" + stmp).offsetTop;}
	myTop = myTop + 28;

	myLeft = 0;	stmp = "";
	while(eval("idImg" + stmp).tagName!="BODY")
		{stmp += ".offsetParent"; myLeft += eval("idImg" + stmp).offsetLeft;}
	myLeft = myLeft + 0;

	eval("document.all."+boxName).style.pixelLeft = myLeft;
	eval("document.all."+boxName).style.pixelTop = myTop;
	}

function boxHide()
	{
	if(this.useParagraph) document.all.boxParagraph.style.visibility = "hidden";
	if(this.useFontName) document.all.boxFont.style.visibility = "hidden";
	if(this.useSize) document.all.boxSize.style.visibility = "hidden";
	if(this.useForeColor) document.all.boxForecolor.style.visibility = "hidden";
	if(this.useBackColor) document.all.boxBackcolor.style.visibility = "hidden";
	if(this.useExternalLink) document.all.boxLink.style.visibility = "hidden";
	}

function selOver(el)
	{
	el.style.background=this.selectionColor;
	el.style.color="white";
	}
function selOut(el)
	{
	el.style.background="";
	el.style.color="black";
	}

/**** B O X ( G E N E R I C ) ***/
function boxShowPop(boxName)
	{
	this.boxHide();
	this.saveSelection();

	eval(boxName).document.body.document.all.inpActiveEditor.innerText = this.oName;

	eval(boxName).document.body.document.all.popup_ContentBorder.style.cssText= this.style_popup_ContentBorder;
	eval(boxName).document.body.document.all.popup_BarArea.style.cssText= "width: 100%; height: 20px;padding-left: 5px;" + this.style_popup_BarArea;
	eval(boxName).document.body.document.all.popup_BarArea_close.style.cssText= "width: 100%; height: 20px;" + this.style_popup_BarArea;
	eval(boxName).document.body.style.cssText= "border:lightgrey 0px solid;" + this.style_popup_Body;

	this.boxDimension(boxName);
	this.boxPosition(boxName);
	eval("document.all."+boxName).style.zIndex = 2;
	eval("document.all."+boxName).style.visibility = "";
	eval("document.all."+boxName).focus();//editor lost focus
	}

function boxShowDrop(boxName)
	{
	this.boxHide();
	this.saveSelection();

	eval(boxName).document.body.document.all.inpActiveEditor.innerText = this.oName;

	this.boxDimension(boxName);
	this.boxPosition(boxName);
	eval("document.all."+boxName).style.zIndex = 2;
	eval("document.all."+boxName).style.visibility = "";
	eval("document.all."+boxName).focus();//editor lost focus
	}

/*** B O X ( G E N E R I C ) =>  P A R A G R A P H  ***/
function applyPara(val)
	{
	for(var i=0;i<parent.frames.length;i++)
		{
		if(parent.frames(i).name == window.name)
			{//frame
			eval("parent.frames."+window.name+"." + boxParagraph.document.body.document.all.inpActiveEditor.value).doCmd2("FormatBlock",val);
			return;
			}
		}
	eval("parent." + boxParagraph.document.body.document.all.inpActiveEditor.value).doCmd2("FormatBlock",val);
	}

/*** B O X ( G E N E R I C ) =>  F O N T  N A M E ***/
function applyFont(val)
	{
	for(var i=0;i<parent.frames.length;i++)
		{
		if(parent.frames(i).name == window.name)
			{//frame
			eval("parent.frames."+window.name+"." + boxFont.document.body.document.all.inpActiveEditor.value).doCmd2("fontname",val);
			return;
			}
		}
	eval("parent."+ boxFont.document.body.document.all.inpActiveEditor.value).doCmd2("fontname",val);
	}

/*** B O X ( G E N E R I C ) =>  F O N T  S I Z E ***/
function applySize(val)
	{
	for(var i=0;i<parent.frames.length;i++)
		{
		if(parent.frames(i).name == window.name)
			{//frame
			eval("parent.frames."+window.name+"." + boxSize.document.body.document.all.inpActiveEditor.value).doCmd2("fontsize",val);
			return;
			}
		}
	eval("parent." + boxSize.document.body.document.all.inpActiveEditor.value).doCmd2("fontsize",val);
	}

/*** B O X ( G E N E R I C ) =>  F O N T  C O L O R ***/
function applyForecolor(val)
	{
	for(var i=0;i<parent.frames.length;i++)
		{
		if(parent.frames(i).name == window.name)
			{//frame
			eval("parent.frames."+window.name+"." + boxForecolor.document.body.document.all.inpActiveEditor.value).doCmd2("ForeColor",val);
			return;
			}
		}
	eval("parent." + boxForecolor.document.body.document.all.inpActiveEditor.value).doCmd2("ForeColor",val);
	}

/*** B O X ( G E N E R I C ) =>  B A C K  C O L O R ***/
function applyBackcolor(val)
	{
	for(var i=0;i<parent.frames.length;i++)
		{
		if(parent.frames(i).name == window.name)
			{//frame
			eval("parent.frames."+window.name+"." + boxBackcolor.document.body.document.all.inpActiveEditor.value).doCmd2("BackColor",val);
			return;
			}
		}
	eval("parent." + boxBackcolor.document.body.document.all.inpActiveEditor.value).doCmd2("BackColor",val);
	}

/*** B O X ( C U S T O M ) => E X T E R N A L  L I N K ***/
function boxShow_ExternalLink()
	{
	this.boxHide();
	this.saveSelection();
	eval("boxLink").document.body.document.all.inpActiveEditor.innerText = this.oName;

	boxLink.document.body.document.all.popup_ContentBorder.style.cssText= this.style_popup_ContentBorder;
	boxLink.document.body.document.all.popup_BarArea.style.cssText= "width: 100%; height: 20px;padding-left: 5px;" + this.style_popup_BarArea;
	boxLink.document.body.document.all.popup_BarArea_close.style.cssText= "width: 100%; height: 20px;" + this.style_popup_BarArea;
	boxLink.document.body.style.cssText= "border:lightgrey 0px solid;" + this.style_popup_Body;

	//Additional
	var oSel	= this.Sel;
	var sType = this.SelType;

	if (oSel.parentElement)//If text is selected on a layer
		{
		this.SelTextFrame = GetElement(oSel.parentElement(),"DIV");//Store the active layer, so we can activate it after editing links
		if(this.SelTextFrame)this.SelTextFrameActive = true;
		else this.SelTextFrameActive = false;
		}
	else//If control is selected
		this.SelTextFrameActive = false;

	//idLinkImage & idLinkImageBorder
	if (oSel.parentElement)//If text is selected
		{
		oEl = GetElement(oSel.parentElement(),"A");//Get A element if any
		boxLink.document.body.document.all.idLinkImage.style.display = "none";//do not display Image features on the Link Dialog
		}
	else //If control is selected
		{
		oEl = GetElement(oSel.item(0),"A");//Get A element if any
		if ((oSel.item) && (oSel.item(0).tagName=="IMG"))//If an image
			{
			boxLink.document.body.document.all.idLinkImage.style.display = "block"; //display Image features on the Link Dialog
			boxLink.document.body.document.all.idLinkImageBorder.value = oSel.item(0).border; //get image border - dropdown
			}
		}

	//idLinkTarget & idLinkType & idLinkURL & idLinkName
	//Is there an A element ?
	if (oEl)//If Yes
		{
		boxLink.document.body.document.all.btnLinkInsert.style.display = "none";
		boxLink.document.body.document.all.btnLinkUpdate.style.display = "block";

		sURL = oEl.href; //get image url
		boxLink.document.body.document.all.idLinkTarget.value = oEl.target;//get image target
		if(oEl.name!=null) boxLink.document.body.document.all.idLinkName.innerText = oEl.name;
		if(oEl.NAME!=null) boxLink.document.body.document.all.idLinkName.innerText = oEl.NAME;
		if(sURL.indexOf(":")!=-1)
			{
			switch(sURL.split(":")[0])
				{
				case "http":
					boxLink.document.body.document.all.idLinkType.value = "http://";
					boxLink.document.body.document.all.idLinkURL.innerText = sURL.substr(7);
					break;
				case "https":
					boxLink.document.body.document.all.idLinkType.value = "https://";
					boxLink.document.body.document.all.idLinkURL.innerText = sURL.substr(8);
					break;
				case "mailto":
					boxLink.document.body.document.all.idLinkType.value = "mailto:";
					boxLink.document.body.document.all.idLinkURL.innerText = sURL.split(":")[1];
					break;
				case "ftp":
					boxLink.document.body.document.all.idLinkType.value = "ftp://";
					boxLink.document.body.document.all.idLinkURL.innerText = sURL.substr(6);
					break;
				case "news":
					boxLink.document.body.document.all.idLinkType.value = "news:";
					boxLink.document.body.document.all.idLinkURL.innerText = sURL.split(":")[1];
					break;
				}
			}
		else
			{
			boxLink.document.body.document.all.idLinkType.value = "";
			boxLink.document.body.document.all.idLinkURL.innerText = sURL;
			}
		}
	else //If No A element
		{
		boxLink.document.body.document.all.btnLinkInsert.style.display = "block";
		boxLink.document.body.document.all.btnLinkUpdate.style.display = "none";	

		boxLink.document.body.document.all.idLinkTarget.value = "";
		boxLink.document.body.document.all.idLinkName.innerText = "";
		boxLink.document.body.document.all.idLinkType.value = "";
		boxLink.document.body.document.all.idLinkURL.innerText = "";
		boxLink.document.body.document.all.idLinkImageBorder.value = 0;
		}

	this.boxDimension("boxLink");
	this.boxPosition("boxLink");
	document.all.boxLink.style.zIndex = 2;
	document.all.boxLink.style.visibility = "";
	document.all.boxLink.focus();
	}

function applyLink()
	{
	var inpURL = boxLink.document.body.document.all.idLinkURL.value;
	var inpURLType = boxLink.document.body.document.all.idLinkType.value;
	var sURL = inpURLType + inpURL;
	var inpAnchor = boxLink.document.body.document.all.idLinkName.value;

	var oSel	= eval(boxLink.document.body.document.all.inpActiveEditor.value).Sel;
	var sType = eval(boxLink.document.body.document.all.inpActiveEditor.value).SelType;

	//a must
	oSel = fixSel(boxLink.document.body.document.all.inpActiveEditor.value,oSel);
	sType = fixSelType(boxLink.document.body.document.all.inpActiveEditor.value,oSel,sType);

	//idLinkImageBorder
	if ((oSel.item) && (oSel.item(0).tagName=="IMG"))//If image is selected
		{
		oSel.item(0).width = oSel.item(0).offsetWidth;//give width attribute
		oSel.item(0).height = oSel.item(0).offsetHeight;//give height attribute
		oSel.item(0).border = boxLink.document.body.document.all.idLinkImageBorder.value;
		}

	if(inpURL!="")
		{
		if (oSel.parentElement)
			{
			if(boxLink.document.body.document.all.btnLinkInsert.style.display == "block")
				{
				if(oSel.text=="")//If no (text) selection, then build selection using the typed URL
					{
					var oSelTmp = oSel.duplicate();
					oSel.text = sURL;
					oSel.setEndPoint("StartToStart",oSelTmp);
					oSel.select();
					sType="Text";
					}
				}
			}

		//idLinkType & idLinkURL (sURL)
		oSel.execCommand("CreateLink",false,sURL);

		//After A element created, then add the Target
		if (oSel.parentElement)
			oEl = GetElement(oSel.parentElement(),"A");
		else
			oEl = GetElement(oSel.item(0),"A");
		if(oEl)
			{
			if(boxLink.document.body.document.all.idLinkTarget.value=="")
				oEl.removeAttribute("target",0);
			else
				oEl.target = boxLink.document.body.document.all.idLinkTarget.value;

			oEl.removeAttribute("NAME",0);
			oEl.removeAttribute("name",0);
			oEl.NAME = boxLink.document.body.document.all.idLinkName.value;

			//~~~ ANCHOR ~~~
			if(inpAnchor!="")
				{
				isImgAnchorExist=false;
				for(var j=0;j<oEl.childNodes.length;j++)
					{
					if(oEl.childNodes[j].tagName=="IMG")
						{
						sSrc = oEl.childNodes[j].src;
						if(sSrc.indexOf("yusasp_anchor.gif")!=-1) isImgAnchorExist=true;
						}
					}
				if(isImgAnchorExist!=true)
					oEl.innerHTML = "<IMG src='include/yusasp_anchor.gif' align=middle border=0>" + oEl.innerHTML;
				}
			else
				{
				for(var j=0;j<oEl.childNodes.length;j++)
					{
					if(oEl.childNodes[j].tagName=="IMG")
						{
						sSrc = oEl.childNodes[j].src;
						if(sSrc.indexOf("yusasp_anchor.gif")!=-1)
							oEl.removeChild(oEl.childNodes[j]);
						}
					}
				}
			//~~~
			}

		eval("idContent"+boxLink.document.body.document.all.inpActiveEditor.value).focus();
		oSel.select();//tambahan
		}
	else if(inpAnchor!="")
		{
		if (oSel.parentElement)
			{
			if(boxLink.document.body.document.all.btnLinkInsert.style.display == "block")
				{
				if(oSel.text=="")//If no (text) selection, then build selection using the typed URL
					{
					var oSelTmp = oSel.duplicate();
					oSel.text = inpAnchor;
					oSel.setEndPoint("StartToStart",oSelTmp);
					oSel.select();
					sType="Text";
					}
				}
			}

		oSel.execCommand("CreateLink",false,"");

		if (oSel.parentElement)
			oEl = GetElement(oSel.parentElement(),"A");
		else
			oEl = GetElement(oSel.item(0),"A");

		if(oEl)
			{
			//~~~ ANCHOR ~~~
			for(var j=0;j<oEl.childNodes.length;j++)
				{
				if(oEl.childNodes[j].tagName=="IMG")
					{
					sSrc = oEl.childNodes[j].src;
					if(sSrc.indexOf("yusasp_anchor.gif")!=-1)
						oEl.removeChild(oEl.childNodes[j]);
					}
				}
			//~~~

			oEl.removeAttribute("target",0);
			oEl.removeAttribute("href",0);

			oEl.removeAttribute("NAME",0);
			oEl.removeAttribute("name",0);
			oEl.NAME = inpAnchor;

			//~~~ ANCHOR ~~~
			oEl.innerHTML = "<IMG src='include/yusasp_anchor.gif' align=middle border=0>" + oEl.innerHTML;
			//~~~
			}
		eval("idContent"+boxLink.document.body.document.all.inpActiveEditor.value).focus();
		try
			{
			oSel.select();//tambahan	(unsp.err if no test in anchor)
			}
		catch(e){}
		}
	else
		{
		//~~~ ANCHOR ~~~
		if (oSel.parentElement)
			oEl = GetElement(oSel.parentElement(),"A");
		else
			oEl = GetElement(oSel.item(0),"A");

		if(oEl)
			{
			for(var j=0;j<oEl.childNodes.length;j++)
				{
				if(oEl.childNodes[j].tagName=="IMG")
					{
					sSrc = oEl.childNodes[j].src;
					if(sSrc.indexOf("yusasp_anchor.gif")!=-1)
						oEl.removeChild(oEl.childNodes[j]);
					}
				}
			}
		//~~~

		oSel.execCommand("unlink");//unlink
		}

	//Activate layer again
	if(eval(boxLink.document.body.document.all.inpActiveEditor.value).SelTextFrameActive) eval(boxLink.document.body.document.all.inpActiveEditor.value).SelTextFrame.setActive();

	eval(boxLink.document.body.document.all.inpActiveEditor.value).boxHide();
	}
	
/*** B O X ( C U S T O M ) => I M A G E S ***/
function boxShow_Image()
	{
	var oEditor = eval("idContent"+this.oName);
	oEditor.focus();//Why ?
	//Because:focus akan men-set oUtil.oNamem shg bisa dipakai di default_Asset.htm

	this.boxHide();
	this.saveSelection();
	if(this.ImagePageURL=="")return;
	var popleft=((document.body.clientWidth - this.ImagePageWidth) / 2)+window.screenLeft;
	var poptop=(((document.body.clientHeight - this.ImagePageHeight) / 2))+window.screenTop-40;	
	window.open(this.ImagePageURL,"NewWindow","scrollbars=NO,width="+this.ImagePageWidth+",height="+this.ImagePageHeight+",left="+popleft+",top="+poptop);
	}

function imgSrc()
	{
	oEditor=eval("idContent"+this.oName);
	var oSel = oEditor.document.selection.createRange();
	if ((oSel.item) && (oSel.item(0).tagName=="IMG"))
		return oSel.item(0).src;
	else
		return "";
	}
function imgAlt()
	{
	oEditor=eval("idContent"+this.oName);
	var oSel = oEditor.document.selection.createRange();
	if ((oSel.item) && (oSel.item(0).tagName=="IMG"))
		return oSel.item(0).alt;
	else
		return "";
	}
function imgAlign()
	{
	oEditor=eval("idContent"+this.oName);
	var oSel = oEditor.document.selection.createRange();
	if ((oSel.item) && (oSel.item(0).tagName=="IMG"))
		return oSel.item(0).align;
	else
		return "";
	}
function imgBorder()
	{
	oEditor=eval("idContent"+this.oName);
	var oSel = oEditor.document.selection.createRange();
	if ((oSel.item) && (oSel.item(0).tagName=="IMG"))
		return oSel.item(0).border;
	else
		return "";
	}
function imgWidth()
	{
	oEditor=eval("idContent"+this.oName);
	var oSel = oEditor.document.selection.createRange();
	if ((oSel.item) && (oSel.item(0).tagName=="IMG"))
		return oSel.item(0).width;
	else
		return "";
	}
function imgHeight()
	{
	oEditor=eval("idContent"+this.oName);
	var oSel = oEditor.document.selection.createRange();
	if ((oSel.item) && (oSel.item(0).tagName=="IMG"))
		return oSel.item(0).height;
	else
		return "";
	}
function imgHspace()
	{
	oEditor=eval("idContent"+this.oName);
	var oSel = oEditor.document.selection.createRange();
	if ((oSel.item) && (oSel.item(0).tagName=="IMG"))
		return oSel.item(0).hspace;
	else
		return "";
	}
function imgVspace()
	{
	oEditor=eval("idContent"+this.oName);
	var oSel = oEditor.document.selection.createRange();
	if ((oSel.item) && (oSel.item(0).tagName=="IMG"))
		return oSel.item(0).vspace;
	else
		return "";
	}

function UpdateImage(inpImgURL,inpImgAlt,inpImgAlign,inpImgBorder,inpImgWidth,inpImgHeight,inpHSpace,inpVSpace)
	{
	var oEditor = eval("idContent"+this.oName);
	var oSel = oEditor.document.selection.createRange();
	var sType = oEditor.document.selection.type;
	if ((oSel.item) && (oSel.item(0).tagName=="IMG"))
		{
		oSel.item(0).style.width="";
		oSel.item(0).style.height="";

		oSel.item(0).src = this.getEditorPreviewPath(inpImgURL);
		if(inpImgAlt!="")
			oSel.item(0).alt = inpImgAlt;
		else
			oSel.item(0).removeAttribute("alt",0);
		oSel.item(0).align = inpImgAlign;
		oSel.item(0).border = inpImgBorder;
		if(inpImgWidth!="")
			oSel.item(0).width = inpImgWidth;
		else
			oSel.item(0).removeAttribute("width",0);	
		if(inpImgHeight!="")
			oSel.item(0).height = inpImgHeight;
		else
			oSel.item(0).removeAttribute("height",0);		
		if(inpHSpace!="")
			oSel.item(0).hspace = inpHSpace;
		else
			oSel.item(0).removeAttribute("hspace",0);		
		if(inpVSpace!="")
			oSel.item(0).vspace = inpVSpace;
		else
			oSel.item(0).removeAttribute("vspace",0);	
		}
	}

function InsertImage(inpImgURL,inpImgAlt,inpImgAlign,inpImgBorder,inpImgWidth,inpImgHeight,inpHSpace,inpVSpace)
	{
	this.doCmd("InsertImage",this.getEditorPreviewPath(inpImgURL));
	var oEditor = eval("idContent"+this.oName);
	var oSel = oEditor.document.selection.createRange();
	var sType = oEditor.document.selection.type;
	if ((oSel.item) && (oSel.item(0).tagName=="IMG"))
		{
		if(inpImgAlt!="")
			oSel.item(0).alt = inpImgAlt;
		oSel.item(0).align = inpImgAlign;
		oSel.item(0).border = inpImgBorder;
		if(inpImgWidth!="")
			oSel.item(0).width = inpImgWidth;
		if(inpImgHeight!="")
			oSel.item(0).height = inpImgHeight;
		if(inpHSpace!="")
			oSel.item(0).hspace = inpHSpace;
		if(inpVSpace!="")
			oSel.item(0).vspace = inpVSpace;
		}
	}

/*** M I S C.  F U N C T I O N S ***/
function setDisplayMode()
	{
	this.boxHide();
	this.saveSelection();

	var oEditor = eval("idContent"+this.oName);

	if(this.displayMode=='RICH')
		{
		//~~~~~
		var sTmp = this.getContent();

		if(!this.isFullHTML)
			{
			sTmp=sTmp.substr(sTmp.indexOf("<BODY")+1);
			sTmp=sTmp.substr(sTmp.indexOf(">")+1);//">\r\n"

			if(sTmp.substr(0,2)=="\r\n") sTmp=sTmp.substr(2);

			var arrTmp = sTmp.split("</BODY>");
			sTmp = arrTmp[0];
			}

		oEditor.document.body.innerText = sTmp;//this.getContent();
		//~~~~~

		//RED color for additional codes
		var sContent = oEditor.document.body.innerHTML;

		var arrTmp = sContent.split('&lt;LINK href="' + this.getEditorPreviewPath(this.PageStylePath_RelativeTo_EditorPath + this.PageStyle) + '" type=text/css rel=stylesheet&gt;');
		if (arrTmp.length > 1) sContent = arrTmp.join('<br><font color=green>&lt;link href="' + this.getEditorPreviewPath(this.PageStylePath_RelativeTo_EditorPath + this.PageStyle) + '" type=text/css rel=stylesheet&gt;</font>');

		oEditor.document.body.innerHTML = sContent;

		//HTML view style
		oEditor.document.body.clearAttributes;
		oEditor.document.body.style.fontFamily = 'Courier New';
		oEditor.document.body.style.fontSize = '13px';
		oEditor.document.body.style.color = 'black';
		oEditor.document.body.style.background = 'white';

		//change status (displayMode)
		this.displayMode = 'HTML';
		eval("idToolbar"+this.oName).style.display = "none";
		}
	else
		{
		this.putContent(oEditor.document.body.innerText);

		//change status (displayMode)
		this.displayMode = 'RICH';
		eval("idToolbar"+this.oName).style.display = "block";

		document.body.focus();
		}
	}

function GetElement(oElement,sMatchTag)
	{
	while (oElement!=null && oElement.tagName!=sMatchTag)
		{
		if(oElement.id=="idContent"+this.oName) return null;
		oElement = oElement.parentElement;
		}
	return oElement;
	}

function doOnKeyPress()
	{
	var oName = oUtil.oName;
	var oEditor = eval("idContent"+oName);
	if(document.getElementById("chkBreakMode"+oName).checked)
		{
		if (oEditor.event.keyCode==13)
			{
			if (oEditor.event.shiftKey==false)
				{
				var oSel = oEditor.document.selection.createRange();	
				oSel.pasteHTML('<br>');
				oEditor.event.cancelBubble = true;
				oEditor.event.returnValue = false;
				oSel.select();
				oSel.moveEnd("character", 1);
				oSel.moveStart("character", 1);
				oSel.collapse(false);
				return false;
				}
			else
				return event.keyCode=13;
			}
		}
	}

/*** G E T & P U T  C O N T E N T ***/
function getContentBody()
	{
	var sTmp = this.getContent();
	if(!this.isFullHTML)
		{
		sTmp=sTmp.substr(sTmp.indexOf("<BODY")+1);
		sTmp=sTmp.substr(sTmp.indexOf(">")+1);//">\r\n"
		var arrTmp = sTmp.split("</BODY>");
		sTmp = arrTmp[0];
		}
	return sTmp;
	}
function getContent()
	{
	if(this.displayMode != "RICH")
		{
		alert("Please uncheck HTML view");
		return;
		}

	var oEditor = eval("idContent"+this.oName);
	var oEditorTmp = eval("idContentTmp"+this.oName);

	var oDoc = oEditorTmp.document.open("text/html", "replace");
	oDoc.write(oEditor.document.documentElement.outerHTML);
	oDoc.close();

	//REMOVE UNUSED
	oEditorTmp.document.body.style.border = "";
	oEditorTmp.document.body.removeAttribute("contentEditable",0);

	//~~~ ANCHOR ~~~
	for(var i=0;i<oEditorTmp.document.all.length;i++)
		{
		if(oEditorTmp.document.all[i].tagName=="A")
			{
			for(var j=0;j<oEditorTmp.document.all[i].childNodes.length;j++)
				{
				if(oEditorTmp.document.all[i].childNodes[j].tagName=="IMG")
					{
					sSrc = oEditorTmp.document.all[i].childNodes[j].src;
					if(sSrc.indexOf("yusasp_anchor.gif")!=-1)
						oEditorTmp.document.all[i].removeChild(oEditorTmp.document.all[i].childNodes[j]);
					}
				}
			}
		}
	//~~~

	var str = oEditorTmp.document.documentElement.outerHTML;

	var arrTmp = str.split("<BASE href="+this.base+">");
	if (arrTmp.length > 1) str = arrTmp.join("");
	
	var arrTmp = str.split('<PARAM NAME="Play" VALUE="0">');
	if (arrTmp.length > 1) str = arrTmp.join('<PARAM NAME="Play" VALUE="-1">');		

	str = str.replace("<STYLE><\/STYLE>","");
	str = str.replace("BORDER-RIGHT: medium none; ","");
	str = str.replace("BORDER-TOP: medium none; ","");
	str = str.replace("BORDER-BOTTOM: medium none; ","");
	str = str.replace("BORDER-LEFT: medium none; ","");
	str = str.replace("BORDER-RIGHT: medium none","");
	str = str.replace("BORDER-TOP: medium none","");
	str = str.replace("BORDER-BOTTOM: medium none","");
	str = str.replace("BORDER-LEFT: medium none","");
	str = str.replace(" style=\"\"","");

	//Fix & Bug
	myregexp = /&amp;/g;
	str = str.replace(myregexp,"&");

	return str;
	}

function putContent(sContent)
	{
	var oEditor = eval("idContent"+this.oName);
	var oEditorTmp = eval("idContentTmp"+this.oName);

	//LOAD CONTENT ***/
	var oDoc = oEditorTmp.document.open("text/html", "replace");
	oDoc.write(sContent);
	oDoc.close();

	var oDoc = oEditor.document.open("text/html", "replace");
	if(this.base!="")
		oDoc.write("<BASE HREF=\""+this.base+"\"/>" + oEditorTmp.document.documentElement.outerHTML);
	else
		oDoc.write(oEditorTmp.document.documentElement.outerHTML);
	oDoc.close();

	if(this.PageStyle!="")
		{
		//Apply external css to the document, if there is no existing..
		sHREF = this.getEditorPreviewPath(this.PageStylePath_RelativeTo_EditorPath + this.PageStyle);
		var bPageStyleExist = false;
		for(var i=0;i<oEditor.document.styleSheets.length;i++)
			{
			if(oEditor.document.styleSheets(i).href==sHREF)
				bPageStyleExist = true;
			}
		if(bPageStyleExist==false)
			{
			oElement = oEditor.document.createElement("<link rel='stylesheet' href='" + this.getEditorPreviewPath(this.PageStylePath_RelativeTo_EditorPath + this.PageStyle) + "' type='text/css'>");
			var oEl = oEditor.document.childNodes[0].childNodes[0].appendChild(oElement);
			}
		}

	//Additional ***/
	oEditor.document.body.style.border = this.style_editor_Border;
	oEditor.document.body.contentEditable = true;
	oEditor.document.body.oncontextmenu = retFalse;

	//ADDITIONAL SETTINGS ***/
	oEditor.document.execCommand("2D-Position", true, true);
	oEditor.document.execCommand("MultipleSelection", true, true);
	oEditor.document.execCommand("LiveResize", true, true);
	oEditor.document.onkeypress = doOnKeyPress;

	//~~~ ANCHOR ~~~
	for(var i=0;i<oEditor.document.all.length;i++)
		{
		if(oEditor.document.all[i].tagName=="A")
			{
			if(oEditor.document.all[i].name!="")
				oEditor.document.all[i].innerHTML = "<IMG src='include/yusasp_anchor.gif' align=middle border=0>" + oEditor.document.all[i].innerHTML;
			}
		}
	//~~~~~~~~~~~~~
	}

function getTitle()
	{
	var oEditor = eval("idContent"+this.oName);
	return oEditor.document.title;
	}

function setTitle(sTitle)
	{
	var oEditor = eval("idContent"+this.oName);
	oEditor.document.title = sTitle;
	}

/*** D O  C O M M A N D ***/
function doCmd(sCmd,sOption)
	{
	this.boxHide();

	var oEditor = eval("idContent"+this.oName);
	var oSel = oEditor.document.selection.createRange();
	var sType = oEditor.document.selection.type;

	var oName = this.oName;//fail kalo pake oUtil.oName, krn ada delay..mesti tunggu

	//a must
	oSel = fixSel(oName,oSel);
	sType = fixSelType(oName,oSel,sType);

	var oTarget = (sType == "None" ? oEditor.document : oSel);
	oTarget.execCommand(sCmd, false, sOption);
	}
function doCmd2(sCmd,sOption)//For BOX
	{
	this.boxHide();

	var oEditor = eval("idContent"+this.oName);
	var oSel = this.Sel;
	var sType = this.SelType;

	var oName = this.oName;

	//a must
	oSel = fixSel(oName,oSel);
	sType = fixSelType(oName,oSel,sType);

	var oTarget = (sType == "None" ? oEditor.document : oSel);
	oSel.select();//tambahan
	oTarget.execCommand(sCmd, false, sOption);
	}

/*** I N D E P E N D E N T (a must) ***/
function fixSel(oName,oSel)//oName=string, oSel=object
	{
	var oEditor = eval("idContent"+oName);

	if(oSel.parentElement != null)
		{
		if(!IsInsideEditor(oSel.parentElement()))
			{
			oEditor.focus();
			var oSel = oEditor.document.selection.createRange();
			}
		}
	else
		{
		if(!IsInsideEditor(oSel.item(0)))
			{
			oEditor.focus();
			var oSel = oEditor.document.selection.createRange();
			}
		}
	return oSel;
	}
function fixSelType(oName,oSel,sType)//oName=string, oSel=object
	{
	var oEditor = eval("idContent"+oName);
	if(oSel.parentElement != null)
		{
		if(!IsInsideEditor(oSel.parentElement()))
			{
			oEditor.focus();
			var sType = oEditor.document.selection.type;
			}
		}
	else
		{
		if(!IsInsideEditor(oSel.item(0)))
			{
			oEditor.focus();
			var sType = oEditor.document.selection.type;
			}
		}
	return sType;
	}
function IsInsideEditor(oElement)
	{
	while (oElement!=null)
		{
		if(oElement.tagName=="BODY" && oElement.contentEditable=="true")return true;
		oElement = oElement.parentElement;
		}
	return false;
	}
function retFalse()
	{
	return false;
	}

/*** T O O L B A R   F U N C T I O N S ***/
function drawIcon(block,nTop,width,id,title,command,isToggle)
	{
	document.write("<span unselectable='on' id='span_"+id+"' style=\"margin-left:2;margin-right:0;margin-bottom:3;width: "+width+";height:26;\"><span unselectable='on' style=\"position:absolute;clip: rect(0 "+width+" 26 0)\">");
	switch(block)
		{
		case 1:
			document.write("<img unselectable='on' id='"+id+"' btntoggle='"+isToggle+"' title=\""+title+"\" src=\""+ this.iconFolder +"icon1.gif\" style=\"position:absolute;top:-"+nTop+";width:58\" onmousedown=\"button_down(this,"+nTop+")\" onmouseup=\"button_up(this,"+nTop+");"+command+"\">");
			break;
		case 2:
			document.write("<img unselectable='on' id='"+id+"' btntoggle='"+isToggle+"' title=\""+title+"\" src=\""+ this.iconFolder +"icon2.gif\" style=\"position:absolute;top:-"+nTop+";width:75\" onmousedown=\"button_down(this,"+nTop+")\" onmouseup=\"button_up(this,"+nTop+");"+command+"\">");
			break;
		case 3:
			document.write("<img unselectable='on' id='"+id+"' btntoggle='"+isToggle+"' title=\""+title+"\" src=\""+ this.iconFolder +"icon3.gif\" style=\"position:absolute;top:-"+nTop+";width:27\" onmousedown=\"button_down(this,"+nTop+")\" onmouseup=\"button_up(this,"+nTop+");"+command+"\">");
			break;
		case 4:
			document.write("<img unselectable='on' id='"+id+"' btntoggle='"+isToggle+"' title=\""+title+"\" src=\""+ this.iconFolder +"icon4.gif\" style=\"position:absolute;top:-"+nTop+";width:27\" onmousedown=\"button_down(this,"+nTop+")\" onmouseup=\"button_up(this,"+nTop+");"+command+"\">");
			break;
		case 5:
			document.write("<img unselectable='on' id='"+id+"' btntoggle='"+isToggle+"' title=\""+title+"\" src=\""+ this.iconFolder +"icon5.gif\" style=\"position:absolute;top:-"+nTop+";width:27\" onmousedown=\"button_down(this,"+nTop+")\" onmouseup=\"button_up(this,"+nTop+");"+command+"\">");
			break;
		case 6:
			document.write("<img unselectable='on' id='"+id+"' btntoggle='"+isToggle+"' title=\""+title+"\" src=\""+ this.iconFolder +"icon6.gif\" style=\"position:absolute;top:-"+nTop+";width:27\" onmousedown=\"button_down(this,"+nTop+")\" onmouseup=\"button_up(this,"+nTop+");"+command+"\">");
			break;
		case 7:
			document.write("<img unselectable='on' id='"+id+"' btntoggle='"+isToggle+"' title=\""+title+"\" src=\""+ this.iconFolder +"icon7.gif\" style=\"position:absolute;top:-"+nTop+";width:27\" onmousedown=\"button_down(this,"+nTop+")\" onmouseup=\"button_up(this,"+nTop+");"+command+"\">");
			break;
		case 8:
			document.write("<img unselectable='on' id='"+id+"' btntoggle='"+isToggle+"' title=\""+title+"\" src=\""+ this.iconFolder +"icon8.gif\" style=\"position:absolute;top:-"+nTop+";width:27\" onmousedown=\"button_down(this,"+nTop+")\" onmouseup=\"button_up(this,"+nTop+");"+command+"\">");
			break;
		case 9:
			document.write("<img unselectable='on' id='"+id+"' btntoggle='"+isToggle+"' title=\""+title+"\" src=\""+ this.iconFolder +"icon9.gif\" style=\"position:absolute;top:-"+nTop+";width:27\" onmousedown=\"button_down(this,"+nTop+")\" onmouseup=\"button_up(this,"+nTop+");"+command+"\">");
			break;
		case 10:
			document.write("<img unselectable='on' id='"+id+"' btntoggle='"+isToggle+"' title=\""+title+"\" src=\""+ this.iconFolder +"icon10.gif\" style=\"position:absolute;top:-"+nTop+";width:27\" onmousedown=\"button_down(this,"+nTop+")\" onmouseup=\"button_up(this,"+nTop+");"+command+"\">");
			break;
		case 11:
			document.write("<img unselectable='on' id='"+id+"' btntoggle='"+isToggle+"' title=\""+title+"\" src=\""+ this.iconFolder +"icon11.gif\" style=\"position:absolute;top:-"+nTop+";width:27\" onmousedown=\"button_down(this,"+nTop+")\" onmouseup=\"button_up(this,"+nTop+");"+command+"\">");
			break;
		}
	document.write("</span></span>");
	}
function button_up(eButton, nTop)
	{
	eButton.style.top = -nTop;
	}
function button_down(eButton, nTop)
	{
	eButton.style.top = -nTop - 26;
	}

/*** E D I T O R   C L A S S ***/
function RUN()
	{
	oName = this.oName;
	obj = eval(oName);
	if(this.base=="") this.base = this.baseEditor;

	//PREPARE ICONS ***/
	document.write("<IMG SRC='" +this.iconFolder+ "icon1.gif' width=1 height=1 style='display:none'>");
	document.write("<IMG SRC='" +this.iconFolder+ "icon2.gif' width=1 height=1 style='display:none'>");
	document.write("<IMG SRC='" +this.iconFolder+ "icon3.gif' width=1 height=1 style='display:none'>");
	document.write("<IMG SRC='" +this.iconFolder+ "icon4.gif' width=1 height=1 style='display:none'>");
	document.write("<IMG SRC='" +this.iconFolder+ "icon5.gif' width=1 height=1 style='display:none'>");
	document.write("<IMG SRC='" +this.iconFolder+ "icon6.gif' width=1 height=1 style='display:none'>");
	document.write("<IMG SRC='" +this.iconFolder+ "icon7.gif' width=1 height=1 style='display:none'>");
	document.write("<IMG SRC='" +this.iconFolder+ "icon8.gif' width=1 height=1 style='display:none'>");
	document.write("<IMG SRC='" +this.iconFolder+ "icon9.gif' width=1 height=1 style='display:none'>");
	document.write("<IMG SRC='" +this.iconFolder+ "icon10.gif' width=1 height=1 style='display:none'>");
	document.write("<IMG SRC='" +this.iconFolder+ "icon11.gif' width=1 height=1 style='display:none'>");

	document.write("<table width='"+this.width+"' height='"+this.height+"' border=0 cellpadding=0 cellspacing=0><tr><td>")
	
	//PREPARE TOOLBAR ***/
	document.write("<table style='display:none' id='idToolbar"+oName+"' cellpadding=0 cellspacing=0 border=0 width='100%' valign=bottom style=\""+this.style_editor_Toolbar+"\"><tr><td height=4></td></tr><tr><td>");
	
	if(this.useSave)obj.drawIcon(1,0,27,"btnSave"+oName,"±£´æ",oName+".onSave()",false);
	if(this.usePrint)obj.drawIcon(1,52,27,"btnPrint"+oName,"´òÓ¡",oName+".doCmd('Print')",false);
	
	if(this.useParagraph)obj.drawIcon(2,0,75,"btnParagraph"+oName,"¶ÎÂä",oName+".boxShowDrop('boxParagraph');setPosition(this,'boxParagraph')",false);
	if(this.useFontName)obj.drawIcon(2,52,75,"btnFontName"+oName,"×ÖÌå",oName+".boxShowDrop('boxFont');setPosition(this,'boxFont')",false);
	if(this.useSize)obj.drawIcon(2,104,55,"btnSize"+oName,"³ß´ç",oName+".boxShowDrop('boxSize');setPosition(this,'boxSize')",false);

	if(this.useSelectAll)obj.drawIcon(3,0,27,"btnSelectAll"+oName,"Ñ¡ÔñÈ«²¿",oName+".doCmd('SelectAll')",false);
	if(this.useCut)obj.drawIcon(3,52,27,"btnCut"+oName,"¼ôÇÐ",oName+".doCmd('Cut')",false);
	if(this.useCopy)obj.drawIcon(3,104,27,"btnCopy"+oName,"¸´ÖÆ",oName+".doCmd('Copy')",false);
	if(this.usePaste)obj.drawIcon(3,156,27,"btnPaste"+oName,"Õ³Ìù",oName+".doCmd('Paste')",false);
	
	if(this.useUndo)obj.drawIcon(4,0,27,"btnUndo"+oName,"ÖØ×ö",oName+".doCmd('Undo')",false);
	if(this.useRedo)obj.drawIcon(4,52,27,"btnRedo"+oName,"Redo",oName+".doCmd('Redo')",false);
	if(this.useBold)obj.drawIcon(4,104,27,"btnBold"+oName,"´ÖÌå",oName+".doCmd('Bold')",false);
	if(this.useItalic)obj.drawIcon(4,156,27,"btnItalic"+oName,"Ð±Ìå",oName+".doCmd('Italic')",false);

	if(this.useUnderline)obj.drawIcon(5,0,27,"btnUnderline"+oName,"ÏÂ»®Ïß",oName+".doCmd('Underline')",false);
	if(this.useStrikethrough)obj.drawIcon(5,52,27,"btnStrikethrough"+oName,"Strikethrough",oName+".doCmd('Strikethrough')",false);
	if(this.useSuperscript)obj.drawIcon(5,104,27,"btnSuperscript"+oName,"ÉÏ±ê",oName+".doCmd('Superscript')",false);
	if(this.useSubscript)obj.drawIcon(5,156,27,"btnSubscript"+oName,"ÏÂ±ê",oName+".doCmd('Subscript')",false);
	
	if(this.useJustifyLeft)obj.drawIcon(6,52,27,"btnJustifyLeft"+oName,"×ó¶ÔÆë",oName+".doCmd('JustifyLeft')",false);
	if(this.useJustifyCenter)obj.drawIcon(6,104,27,"btnJustifyCenter"+oName,"ÖÐÑë¶ÔÆë",oName+".doCmd('JustifyCenter')",false);
	if(this.useJustifyRight)obj.drawIcon(6,156,27,"btnJustifyRight"+oName,"ÓÒ¶ÔÆë",oName+".doCmd('JustifyRight')",false);

	if(this.useJustifyFull)obj.drawIcon(7,0,27,"btnJustifyFull"+oName,"¶ÔÆë",oName+".doCmd('JustifyFull')",false);
	if(this.useNumbering)obj.drawIcon(7,52,27,"btnNumbering"+oName,"·ûºÅÏî",oName+".doCmd('InsertOrderedList')",false);
	if(this.useBullets)obj.drawIcon(7,104,27,"btnBullets"+oName,"Bullets",oName+".doCmd('InsertUnorderedList')",false);
	if(this.useIndent)obj.drawIcon(7,156,27,"btnIndent"+oName,"Ëõ½ø",oName+".doCmd('Indent')",false);
	
	if(this.useOutdent)obj.drawIcon(8,0,27,"btnOutdent"+oName,"Outdent",oName+".doCmd('Outdent')",false);
	if(this.useImage)obj.drawIcon(8,52,27,"btnImage"+oName,"Ç¶ÈëÍ¼Æ¬",oName+".boxShow_Image()",false);
	if(this.useForeColor)obj.drawIcon(8,104,27,"btnForeColor"+oName,"Ç°¾°É«",oName+".boxShowPop('boxForecolor')",false);
	if(this.useBackColor)obj.drawIcon(8,156,27,"btnBackColor"+oName,"±³¾°É«",oName+".boxShowPop('boxBackcolor')",false);

	if(this.useExternalLink)obj.drawIcon(9,0,27,"btnExternalLink"+oName,"³¬Á´½Ó",oName+".boxShow_ExternalLink()",false);

	if(this.useClean)obj.drawIcon(10,104,27,"btnClean"+oName,"Çå³ý",oName+".doCmd('RemoveFormat')",false);
	if(this.useLine)obj.drawIcon(10,156,27,"btnLine"+oName,"Ë®Æ½Ïß",oName+".doCmd('InsertHorizontalRule')",false);
	
	document.write("</td></tr>")
	document.write("</table>");
	
	document.write("</td></tr><tr><td height=100%>")

	//PREPARE EDITOR ***/
	document.write("	<table name='idArea"+oName+"' id='idArea"+oName+"' width=100% height=100% cellpadding=0 cellspacing=0>");
	document.write("	<tr><td>");
	document.write("	<iframe style=\"width:1;height:1;\" name='idContent"+oName+"' contentEditable=true id='idContent"+oName+"' noborder onfocus=\""+oName+".boxHide();oUtil.oName='"+oName+"'\"></iframe>");
	document.write("	</td></tr>");
	document.write("	</table>");

	document.write("</td></tr><tr><td>")
	
	//PREPARE EDITOR FOOTER ***/
	document.write("	<input type=checkbox onclick='"+oName+".setDisplayMode()' id=chkDisplayMode name=chkDisplayMode><font face=verdana size=1> HTML &nbsp;</font>");
	document.write("	<input type='checkbox' id=chkBreakMode"+oName+" name=chkBreakMode"+oName+"><font face=verdana size=1> BR &nbsp;</font>");

	document.write("</td></tr></table>")

	var oEditor = eval("idContent"+oName);
	IEver=0;
	if (navigator.appVersion.indexOf("MSIE")!=-1) IEver=parseFloat(navigator.appVersion.split("MSIE")[1]);
	if (IEver<5.5) oEditor.document.designMode = "on";//spy jln di IE50 tapi gak bisa insert object

	//LOAD PREBUILT CONTENT ***/
	var oDoc = oEditor.document.open("text/html", "replace");
	if(this.base!="")
		oDoc.write("<BASE HREF=\""+this.base+"\"/>"+this.prebuiltHTML);
	else
		oDoc.write(this.prebuiltHTML);
	oDoc.close();


	if(this.PageStyle!="")
		{
		//apply external css to the document
		oElement = oEditor.document.createElement("<link rel='stylesheet' href='" + this.getEditorPreviewPath(this.PageStylePath_RelativeTo_EditorPath + this.PageStyle) + "' type='text/css'>");
		var oEl = oEditor.document.childNodes[0].childNodes[0].appendChild(oElement);
		}


	//ADDITIONAL STYLES (Remove on getContent) ***/
	oEditor.document.body.style.border = this.style_editor_Border;
	oEditor.document.body.contentEditable = true;
	oEditor.document.body.oncontextmenu = retFalse;

	//ADDITIONAL SETTINGS ***/
	oEditor.document.execCommand("2D-Position", true, true);
	oEditor.document.execCommand("MultipleSelection", true, true);
	oEditor.document.execCommand("LiveResize", true, true);
	oEditor.document.onkeypress = doOnKeyPress;

	//ENABLE/SHOW EDITOR ***/
	eval("document.all.idContent"+this.oName).style.width = "100%";
	eval("document.all.idContent"+this.oName).style.height = "100%";
	eval("document.all.idToolbar"+this.oName).style.display="block";

	//PREPARE TEMP AREA ***/
	document.write("<iframe name='idContentTmp"+oName+"' contentEditable=true id='idContentTmp"+oName+"' style='width:1;height:1;overflow:auto;' noborder onfocus='"+oName+".boxHide()'></iframe>");
	var oEditor2 = eval("idContentTmp"+oName);
	oEditor2.document.designMode = "on";
	oEditor2.document.open("text/html","replace");
	oEditor2.document.write("<html><head></head><body><div></div></body></html>");
	oEditor2.document.close();

	/*** S H A R E D  B O X ***/
	if(this.useParagraph && !document.getElementById("boxParagraph"))//boxParagraph
		{
		sHTML = "<table align=center border=0 width=100% height=100% cellpadding=0 cellspacing=0>" +
			"<tr><td height=3></td></tr>" +
			"<tr><td onclick=\"parent.applyPara('<H1>')\"  align=center onmouseover='parent."+this.oName+".selOver(this)' onmouseout='parent."+this.oName+".selOut(this)' class=dropdown style='height:45'><h1 style='margin:0'>Header 1</h1></td></tr>" +
			"<tr><td onclick=\"parent.applyPara('<H2>')\"  align=center onmouseover='parent."+this.oName+".selOver(this)' onmouseout='parent."+this.oName+".selOut(this)' class=dropdown style='height:40'><h2 style='margin:0'>Header 2</h2></td></tr>" +
			"<tr><td onclick=\"parent.applyPara('<H3>')\"  align=center onmouseover='parent."+this.oName+".selOver(this)' onmouseout='parent."+this.oName+".selOut(this)' class=dropdown style='height:35'><h3 style='margin:0'>Header 3</h3></td></tr>" +
			"<tr><td onclick=\"parent.applyPara('<H4>')\"  align=center onmouseover='parent."+this.oName+".selOver(this)' onmouseout='parent."+this.oName+".selOut(this)' class=dropdown style='height:30'><h4 style='margin:0'>Header 4</h4></td></tr>" +
			"<tr><td onclick=\"parent.applyPara('<H5>')\"  align=center onmouseover='parent."+this.oName+".selOver(this)' onmouseout='parent."+this.oName+".selOut(this)' class=dropdown style='height:25'><h5 style='margin:0'>Header 5</h5></td></tr>" +
			"<tr><td onclick=\"parent.applyPara('<H6>')\"  align=center onmouseover='parent."+this.oName+".selOver(this)' onmouseout='parent."+this.oName+".selOut(this)' class=dropdown style='height:20'><h6 style='margin:0'>Header 6</h6></td></tr>" +
			"<tr><td onclick=\"parent.applyPara('<PRE>')\" align=center onmouseover='parent."+this.oName+".selOver(this)' onmouseout='parent."+this.oName+".selOut(this)' class=dropdown style='height:20'><pre style='margin:0'>Preformatted</pre></td></tr>" +
			"<tr><td onclick=\"parent.applyPara('<P>')\"   align=center onmouseover='parent."+this.oName+".selOver(this)' onmouseout='parent."+this.oName+".selOut(this)' class=dropdown style='height:20'><p style='margin:0'>Normal</p></td></tr>" +
			"<tr><td height=3></td></tr>" +
			"</table>";
		sHTML = boxDrawDrop(200,sHTML);
		document.write("<iframe id=boxParagraph name=boxParagraph style='position: absolute; visibility: hidden; z-index: -1;width:1;height:1;'></iframe>");
		boxParagraph.document.open("text/html","replace");
		boxParagraph.document.write(sHTML);
		boxParagraph.document.close();
		}

	if(this.useFontName && !document.getElementById("boxFont"))//boxFont
		{
		sHTML = "<table align=center border=0 width=100% height=100% cellpadding=0 cellspacing=0>" +
			"<tr><td height=3></td></tr>" +
			"<tr><td onclick=\"parent.applyFont('Arial')\"  align=center onmouseover='parent."+this.oName+".selOver(this)' onmouseout='parent."+this.oName+".selOut(this)' class=dropdown style='height:17'><font face='Arial' size=1>Arial</font></td></tr>" +
			"<tr><td onclick=\"parent.applyFont('Arial Black')\"  align=center onmouseover='parent."+this.oName+".selOver(this)' onmouseout='parent."+this.oName+".selOut(this)' class=dropdown style='height:17'><font face='Arial Black' size=1>Arial Black</font></td></tr>" +
			"<tr><td onclick=\"parent.applyFont('Arial Narrow')\"  align=center onmouseover='parent."+this.oName+".selOver(this)' onmouseout='parent."+this.oName+".selOut(this)' class=dropdown style='height:17'><font face='Arial Narrow' size=1>Arial Narrow</font></td></tr>" +
			"<tr><td onclick=\"parent.applyFont('Comic Sans MS')\"  align=center onmouseover='parent."+this.oName+".selOver(this)' onmouseout='parent."+this.oName+".selOut(this)' class=dropdown style='height:17'><font face='Comic Sans MS' size=1>Comic Sans MS</font></td></tr>" +
			"<tr><td onclick=\"parent.applyFont('Courier New')\"  align=center onmouseover='parent."+this.oName+".selOver(this)' onmouseout='parent."+this.oName+".selOut(this)' class=dropdown style='height:17'><font face='Courier New' size=1>Courier New</font></td></tr>" +
			"<tr><td onclick=\"parent.applyFont('System')\"  align=center onmouseover='parent."+this.oName+".selOver(this)' onmouseout='parent."+this.oName+".selOut(this)' class=dropdown style='height:19'><font face='System' size=1>System</font></td></tr>" +
			"<tr><td onclick=\"parent.applyFont('Tahoma')\" align=center onmouseover='parent."+this.oName+".selOver(this)' onmouseout='parent."+this.oName+".selOut(this)' class=dropdown style='height:17'><font face='Tahoma' size=1>Tahoma</font></td></tr>" +
			"<tr><td onclick=\"parent.applyFont('Times New Roman')\"   align=center onmouseover='parent."+this.oName+".selOver(this)' onmouseout='parent."+this.oName+".selOut(this)' class=dropdown style='height:17'><font face='Times New Roman' size=1>Times New Roman</font></td></tr>" +
			"<tr><td onclick=\"parent.applyFont('Verdana')\" align=center onmouseover='parent."+this.oName+".selOver(this)' onmouseout='parent."+this.oName+".selOut(this)' class=dropdown style='height:17'><font face='Verdana' size=1>Verdana</font></td></tr>" +
			"<tr><td onclick=\"parent.applyFont('Wingdings')\"   align=center onmouseover='parent."+this.oName+".selOver(this)' onmouseout='parent."+this.oName+".selOut(this)' class=dropdown style='height:17'><font face='Wingdings' size=1>Wingdings</font></td></tr>" +
			"<tr><td height=3></td></tr>" +
			"</table>";
		sHTML = boxDrawDrop(150,sHTML);
		document.write("<iframe id=boxFont name=boxFont style='position: absolute; visibility: hidden; z-index: -1;width:1;height:1;'></iframe>");
		boxFont.document.open("text/html","replace");
		boxFont.document.write(sHTML);
		boxFont.document.close();
		}

	if(this.useSize && !document.getElementById("boxSize"))//boxSize
		{
		sHTML = "<table align=center border=0 width=100% height=100% cellpadding=0 cellspacing=0>" +
			"<tr><td height=3></td></tr>" +
			"<tr><td onclick=\"parent.applySize('1')\"  align=center onmouseover='parent."+this.oName+".selOver(this)' onmouseout='parent."+this.oName+".selOut(this)' class=dropdown style='height:20'><font size=1>Size 1</font></td></tr>" +
			"<tr><td onclick=\"parent.applySize('2')\"  align=center onmouseover='parent."+this.oName+".selOver(this)' onmouseout='parent."+this.oName+".selOut(this)' class=dropdown style='height:20'><font size=2>Size 2</font></td></tr>" +
			"<tr><td onclick=\"parent.applySize('3')\"  align=center onmouseover='parent."+this.oName+".selOver(this)' onmouseout='parent."+this.oName+".selOut(this)' class=dropdown style='height:22'><font size=3>Size 3</font></td></tr>" +
			"<tr><td onclick=\"parent.applySize('4')\"  align=center onmouseover='parent."+this.oName+".selOver(this)' onmouseout='parent."+this.oName+".selOut(this)' class=dropdown style='height:27'><font size=4>Size 4</font></td></tr>" +
			"<tr><td onclick=\"parent.applySize('5')\"  align=center onmouseover='parent."+this.oName+".selOver(this)' onmouseout='parent."+this.oName+".selOut(this)' class=dropdown style='height:33'><font size=5>Size 5</font></td></tr>" +
			"<tr><td onclick=\"parent.applySize('6')\"  align=center onmouseover='parent."+this.oName+".selOver(this)' onmouseout='parent."+this.oName+".selOut(this)' class=dropdown style='height:40'><font size=6>Size 6</font></td></tr>" +
			"<tr><td onclick=\"parent.applySize('7')\" align=center onmouseover='parent."+this.oName+".selOver(this)' onmouseout='parent."+this.oName+".selOut(this)' class=dropdown style='height:65'><font size=7>Size 7</font></td></tr>" +
			"<tr><td height=3></td></tr>" +
			"</table>";
		sHTML = boxDrawDrop(200,sHTML);
		document.write("<iframe id=boxSize name=boxSize style='position: absolute; visibility: hidden; z-index: -1;width:1;height:1;'></iframe>");
		boxSize.document.open("text/html","replace");
		boxSize.document.write(sHTML);
		boxSize.document.close();
		}

	if(this.useForeColor && !document.getElementById("boxForecolor"))//boxForecolor
		{
		var objColor1 = new YusASPColor("objColor1");
		sHTML = objColor1.drawColor +
			"<table border=0 width=365 align=center cellpadding=0 cellspacing=0>" +
			"<tr>" +
			"<td align=right width=50%><INPUT type='button' onclick=\"eval('parent.'+inpActiveEditor.value).boxHide()\"  value='Cancel'></td>" +
			"<td width=50%><INPUT type='button' onclick=\"parent.applyForecolor(parent.boxForecolor.document.body.document.all."+objColor1.getElementColorText+".value)\" value='&nbsp;&nbsp;OK&nbsp;&nbsp;'></td>" +
			"</tr></table>";
		sHTML = boxDrawPop(390,"Forecolor",sHTML);
		document.write("<iframe id=boxForecolor name=boxForecolor style='position: absolute; visibility: hidden; z-index: -1;width:1;height:1;'></iframe>");
		boxForecolor.document.open("text/html","replace");
		boxForecolor.document.write(sHTML);
		boxForecolor.document.close();
		}

	if(this.useBackColor && !document.getElementById("boxBackcolor"))//boxBackcolor
		{
		var objColor2 = new YusASPColor("objColor2");
		sHTML = objColor2.drawColor +
			"<table border=0 width=365 align=center cellpadding=0 cellspacing=0>" +
			"<tr>" +
			"<td align=right width=50%><INPUT type='button' onclick=\"eval('parent.'+inpActiveEditor.value).boxHide()\"  value='Cancel'></td>" +
			"<td width=50%><INPUT type='button' onclick=\"parent.applyBackcolor(parent.boxBackcolor.document.body.document.all."+objColor2.getElementColorText+".value)\" value='&nbsp;&nbsp;OK&nbsp;&nbsp;'></td>" +
			"</tr></table>";
		sHTML = boxDrawPop(390,"Backcolor",sHTML);
		document.write("<iframe id=boxBackcolor name=boxBackcolor style='position: absolute; visibility: hidden; z-index: -1;width:1;height:1;'></iframe>");
		boxBackcolor.document.open("text/html","replace");
		boxBackcolor.document.write(sHTML);
		boxBackcolor.document.close();
		}

	if(this.useExternalLink && !document.getElementById("boxLink"))//boxLink
		{
		sHTML = "<table cellpadding=0 cellspacing=3 align=center>" +
			"<tr>" +
			"<td>" +
			"		&nbsp;" +
			"		<select id=idLinkType NAME=idLinkType>" +
			"		<option value='' selected></option>" +
			"		<option value='http://' selected>http://</option>" +
			"		<option value='https://'>https://</option>" +
			"		<option value='mailto:'>mailto:</option>" +
			"		<option value='ftp://'>ftp://</option>" +
			"		<option value='news:'>news:</option>" +
			"		</select>" +
			"		<input type=text size=30 id=idLinkURL value='' style='height: 19px;font:8pt verdana,arial,sans-serif' NAME=idLinkURL contentEditable=true>" +
			"</td>" +
			"</tr>" +
			"<tr>" +
			"<td>" +
			"		<table cellpadding=0 cellspacing=0 style='TABLE-LAYOUT: fixed'>" +
			"		<col width=100><col width=220>" +
			"		<tr><td>" +
			"			&nbsp;&nbsp;Target:" +
			"		</td><td>" +
			"			<select id=idLinkTarget NAME=idLinkTarget>" +
			"			<option value='' selected></option>" +
			"			<option value='_self' selected>self</option>" +
			"			<option value='_blank'>blank</option>" +
			"			<option value='_parent'>parent</option>" +
			"			</select>" +
			"		</td></tr>" +
			"		<tr id=idLinkImage style='display:none'><td>" +
			"			&nbsp;&nbsp;Image Border:" +
			"		</td><td>" +
			"			<select id=idLinkImageBorder NAME=idLinkImageBorder>" +
			"			<option value=0 selected>0</option>" +
			"			<option value=1>1</option>" +
			"			<option value=2>2</option>" +
			"			<option value=3>3</option>" +
			"			<option value=4>4</option>" +
			"			<option value=5>5</option>" +
			"			</select>" +
			"		</td></tr>" +
			"		<tr>" +
			"		<td>" +
			"			&nbsp;&nbsp;Name:" +
			"		</td><td>" +
			"			<input type=text size=30 id=idLinkName value='' style='height: 19px;font:8pt verdana,arial,sans-serif' NAME=idLinkName contentEditable=true>" +
			"		</td>" +
			"		</tr>" +
			"		</table>" +
			"</td>" +
			"</tr>" +
			"<tr>" +
			"<td align=center><br>" +
			"		<table cellpadding=0 cellspacing=0><tr>" +
			"		<td>" +
			"		<input type=button value=Cancel onclick=\"eval('parent.'+inpActiveEditor.value).boxHide()\" style='height: 22px;font:8pt verdana,arial,sans-serif'>" +
			"		</td>" +
			"		<td>" +
			"		<input type=button id=btnLinkInsert name=btnLinkInsert value=Insert onclick='parent.applyLink()' style='display:none;height: 22px;font:8pt verdana,arial,sans-serif'>" +
			"		<input type=button id=btnLinkUpdate name=btnLinkUpdate value=Update onclick='parent.applyLink()' style='display:none;height: 22px;font:8pt verdana,arial,sans-serif'>" +
			"		</td>" +
			"		</tr></table>" +
			"		</td>" +
			"</tr>" +
			"</table>";
		sHTML = boxDrawPop(310,"Hyperlink",sHTML);
		document.write("<iframe id=boxLink name=boxLink style='position: absolute; visibility: hidden; z-index: -1;width:1;height:1;'></iframe>");
		boxLink.document.open("text/html","replace");
		boxLink.document.write(sHTML);
		boxLink.document.close();
		}
	}

function ACEditor(oName)
	{
	//Editor Look & Style
	this.width = "100%";
	this.height = "70%";

	this.drawIcon = drawIcon;
	this.iconFolder = "editorimages/";

	this.style_editor_Border = "1px #c3c3c3 solid";
	this.style_editor_Toolbar = "border-left:#d9d9d9 1 solid;border-top:#d9d9d9 1 solid;border-right:#e0e0e0 1 solid;background:#efeff7;";

	this.style_popup_ContentBorder = "border-left: #336699 1px solid;border-right: #336699 1px solid;border-bottom: #336699 1px solid;";
	this.style_popup_BarArea =	"background: #004684;" +
								"filter:progid:DXImageTransform.Microsoft.gradient(startColorstr=#004684,endColorstr=#7189b7);" +
								"border-top: #99ccff 1px solid; border-bottom: #004684 1px solid;";
	this.style_popup_Body =		"background: #ece9d8;" +
								"filter: progid:DXImageTransform.Microsoft.gradient(startColorstr=white, endColorstr=#e3e3e3);";
	this.selectionColor = "#5579aa";

	//Display Mode
	this.displayMode = "RICH";
	this.setDisplayMode = setDisplayMode;
	this.tmp = "";

	//Pop & Drop
	this.boxHide = boxHide;
	this.boxShowPop = boxShowPop;
	this.boxShowDrop = boxShowDrop;
	this.boxDimension = boxDimension;
	this.boxPosition = boxPosition;
	this.selOver = selOver;
	this.selOut = selOut;

	//Common
	this.oName = oName;
	this.doCmd = doCmd;
	this.doCmd2 = doCmd2;

	this.Sel;
	this.SelType;
	this.saveSelection = saveSelection;

	this.SelTextFrame;
	this.SelTextFrameActive;

	//Image
	this.boxShow_Image = boxShow_Image;
	this.ImagePageURL = "";
	this.ImagePageWidth = 443;
	this.ImagePageHeight = 470;
	
	this.UpdateImage = UpdateImage;
	this.InsertImage = InsertImage;	
	this.imgSrc = imgSrc;
	this.imgAlt = imgAlt;
	this.imgAlign = imgAlign;
	this.imgBorder = imgBorder;
	this.imgWidth = imgWidth;
	this.imgHeight = imgHeight;
	this.imgHspace = imgHspace;
	this.imgVspace = imgVspace;
	
	//External Link
	this.boxShow_ExternalLink  = boxShow_ExternalLink;

	this.GetElement = GetElement;
	
	//Misc
	this.getContent = getContent;
	this.getContentBody = getContentBody;
	this.putContent = putContent;
	this.isFullHTML = false;
	this.setTitle = setTitle;
	this.getTitle = getTitle;

	this.onSave = function(){return true;};
	
	//Buttons
	this.useSave = false;

	this.usePrint = true;this.useParagraph = true;
	this.useFontName = true; this.useSize = true;
	this.useSelectAll = true; this.useCut = true;
	this.useCopy = true; this.usePaste = true;
	this.useUndo = true; this.useRedo = true;
	this.useBold = true; this.useItalic = true; this.useUnderline = true;
	this.useStrikethrough = true; this.useSuperscript = true; this.useSubscript = true;
	this.useJustifyLeft = true;	this.useJustifyCenter = true; this.useJustifyRight = true;
	this.useJustifyFull = true; this.useNumbering = true;
	this.useBullets = true;	this.useIndent = true;
	this.useOutdent = true;	this.useImage = true;
	this.useForeColor = true; this.useBackColor = true;
	this.useExternalLink = true;
	this.useClean = true;
	this.useLine = true;

	//Run
	this.prebuiltHTML = "<html><head></head><body><div></div></body></html>";
	this.RUN = RUN;

	//Path Stuff
	this.base = "";
	this.baseEditor = "";
	this.getEditorPreviewPath = getEditorPreviewPath;

	//Stylesheet Stuff
	this.PageStyle = "";
	this.PageStylePath_RelativeTo_EditorPath = "";
	}

function getEditorPreviewPath(APPLIED_LINK)//APPLIED_LINK=> Relative
	{
	//kalo absolute (ada http://, https://... tdk diproses
	//diasumsikan applied link absolute yg dikehendaki
	//todo: buat fungsi yg memproses APPLIED_LINK=> Absolute
	if(APPLIED_LINK.indexOf(":")!=-1) return APPLIED_LINK;

	//kalo base & baseEditor tdk di=specify, dianggap editing & publishing sama (tdk perlu diproses).
	if(this.base=="") return APPLIED_LINK;

	base = this.base;
	baseEditor = this.baseEditor;

	//STEP1
	nDeep1=APPLIED_LINK.split("../").length-1;

	//STEP2
	var sTmp=baseEditor.substring(0,baseEditor.length-1);//remove last "/"
	for(var i=0;i<nDeep1;i++)
		{
		sTmp = sTmp.substring(0,sTmp.lastIndexOf("/"));
		}
	base1=sTmp+"/";//add last "/"

	//STEP3
	str=APPLIED_LINK;
	var arrTmp = str.split("../");
	if (arrTmp.length > 1) str = arrTmp.join("");
	ext1 = str;

	//STEP4
	APPLIED_LINK_ABSOLUTE = base1 + ext1;

	//STEP5
	var arr1 = base.split("/");
	var arr2 = APPLIED_LINK_ABSOLUTE.split("/");
	var sSamePart = "";
	var sTmp = "";
	if(arr1.length>arr2.length)
		{
		for(var i=0;i<arr1.length;i++)
			{
			if(arr2[i]==arr1[i])
				{
				if(i==arr1.length-1)sSamePart=sTmp;
				sTmp+=arr2[i]+"/";
				}
			else
				sSamePart=sTmp;
			}
		//alert(sSamePart)
		}
	else
		{
		for(var i=0;i<arr2.length;i++)
			{
			if(arr2[i]==arr1[i])
				{
				if(i==arr2.length-1)sSamePart=sTmp;
				sTmp+=arr2[i]+"/";
				}
			else
				sSamePart=sTmp;
			}
		//alert(sSamePart)
		}

	//STEP6
	str=base;
	var arrTmp = str.split(sSamePart);
	if (arrTmp.length > 1) str = arrTmp.join("");
	nDeep2=str.split("/").length-1;

	//STEP7
	base2="";
	for(var j=0;j<nDeep2;j++)
		{
		base2+="../";
		}

	//STEP8
	str=APPLIED_LINK_ABSOLUTE;
	var arrTmp = str.split(sSamePart);
	if (arrTmp.length > 1) str = arrTmp.join("");
	ext2=str;

	//STEP9
	return base2 + ext2;
	}

<script>
//判断是否有选择上传的文件,无则返回,有则通过获取文件后缀检查上传文件的格式///////////////////
function Check_FileType()
{
 var str=document.form0.src.value;
 
 if(str=="")
	{
	alert("上传文件不能为空,请选择你要上传的文件!");
	document.form0.src.focus();
	return false;
	}
else
{
      var pos = str.lastIndexOf(".");
      var lastname = str.substring(pos,str.length);
	 if (lastname.toLowerCase()!=".jpg" && lastname.toLowerCase()!=".gif")
	 {
		 alert("您上传的文件类型为"+lastname+"，图片必须为.jpg /.gif类型");
		 document.form0.src.value="";
		 document.form0.src.focus();
		 return false;
	 }
}

}
//////////////////////////////////////////////////////////////////////////////////////////////



//////////////浏览时预览图片/////////////////////////
function play()
{
      var str=document.form0.src.value;
      var pos = str.lastIndexOf(".");
      var lastname = str.substring(pos,str.length);
	  
if (lastname.toLowerCase()!=".jpg" && lastname.toLowerCase()!=".gif")

	 {
     document.getElementById("zi0").style.display="";
	 document.getElementById("zi").style.display="none";
	 document.getElementById("img").style.display="none";
	 }
	 
else

	{
	document.getElementById("zi").style.display="none";
	document.getElementById("zi0").style.display="none";
	document.form0.img.style.display="";
	document.form0.img.src=document.form0.src.value;
	}
}
////////////////////////////////////////////////////


////////////////////////////检查空值///////////////////////////
function checkspace(checkstr) {
  var str = '';
  for(i = 0; i < checkstr.length; i++) {
    str = str + ' ';
  }
  return (str == checkstr);
}
////////////////////////////////////////////////////////////////



////图片按比例缩小//////////////////////////////////////////
function zoom(it)
{
if(it.height>300||it.width>300)
it.style.zoom=300/(it.height>it.width?it.height:it.width);
}
/////////////////////////////////////////////////////////////

</script>

<script>
//�ж��Ƿ���ѡ���ϴ����ļ�,���򷵻�,����ͨ����ȡ�ļ���׺����ϴ��ļ��ĸ�ʽ///////////////////
function Check_FileType()
{
 var str=document.form0.src.value;
 
 if(str=="")
	{
	alert("�ϴ��ļ�����Ϊ��,��ѡ����Ҫ�ϴ����ļ�!");
	document.form0.src.focus();
	return false;
	}
else
{
      var pos = str.lastIndexOf(".");
      var lastname = str.substring(pos,str.length);
	 if (lastname.toLowerCase()!=".jpg" && lastname.toLowerCase()!=".gif")
	 {
		 alert("���ϴ����ļ�����Ϊ"+lastname+"��ͼƬ����Ϊ.jpg /.gif����");
		 document.form0.src.value="";
		 document.form0.src.focus();
		 return false;
	 }
}

}
//////////////////////////////////////////////////////////////////////////////////////////////



//////////////���ʱԤ��ͼƬ/////////////////////////
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


////////////////////////////����ֵ///////////////////////////
function checkspace(checkstr) {
  var str = '';
  for(i = 0; i < checkstr.length; i++) {
    str = str + ' ';
  }
  return (str == checkstr);
}
////////////////////////////////////////////////////////////////



////ͼƬ��������С//////////////////////////////////////////
function zoom(it)
{
if(it.height>300||it.width>300)
it.style.zoom=300/(it.height>it.width?it.height:it.width);
}
/////////////////////////////////////////////////////////////

</script>
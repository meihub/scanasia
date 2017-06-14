<%
'--------定义部份------------------
Dim Fy_Post,Fy_Get,Fy_In,Fy_Inf,Fy_Xh,Fy_db,Fy_dbstr
'自定义需要过滤的字串,用 "蔡" 分隔
Fy_In = "'蔡;蔡and蔡exec蔡insert蔡select蔡delete蔡update蔡count蔡*蔡%蔡chr蔡mid蔡master蔡truncate蔡char蔡declare"
'----------------------------------
%>

<%
Fy_Inf = split(Fy_In,"蔡")
'--------POST部份------------------
If Request.Form<>"" Then
For Each Fy_Post In Request.Form

For Fy_Xh=0 To Ubound(Fy_Inf)
If Instr(LCase(Request.Form(Fy_Post)),Fy_Inf(Fy_Xh))<>0 Then
'--------写入数据库----------头-----
set Fy_db=Server.Createobject("adodb.connection")
Fy_db.open "Provider=Microsoft.Jet.OLEDB.4.0;"& "Data Source=" & Server.MapPath("cainiao\SqlIn.mdb")
Fy_db.Execute("insert into SqlIn(Sqlin_IP,SqlIn_Web,SqlIn_FS,SqlIn_CS,SqlIn_SJ) values('"&Request.ServerVariables("REMOTE_ADDR")&"','"&Request.ServerVariables("URL")&"','POST','"&Fy_Post&"','"&replace(Request.Form(Fy_Post),"'","''")&"')")
Fy_db.close
Set Fy_db = Nothing
'--------写入数据库----------尾-----
Response.Write "<Script Language=JavaScript>alert('SQL通用防注入系统提示↓\n\n请不要在参数中包含非法字符尝试注入！\n\n www.caiku.com');</Script>"
Response.Write "非法操作！系统做了如下记录↓<br>"
Response.Write "操作ＩＰ："&Request.ServerVariables("REMOTE_ADDR")&"<br>"
Response.Write "操作时间："&Now&"<br>"
Response.Write "操作页面："&Request.ServerVariables("URL")&"<br>"
Response.Write "提交方式：ＰＯＳＴ<br>"
Response.Write "提交参数："&Fy_Post&"<br>"
Response.Write "提交数据："&Request.Form(Fy_Post)
Response.End
End If
Next

Next
End If
'----------------------------------

'--------GET部份-------------------
If Request.QueryString<>"" Then
For Each Fy_Get In Request.QueryString

For Fy_Xh=0 To Ubound(Fy_Inf)
If Instr(LCase(Request.QueryString(Fy_Get)),Fy_Inf(Fy_Xh))<>0 Then
'--------写入数据库----------头-----
set Fy_db=Server.Createobject("adodb.connection")
Fy_db.open "Provider=Microsoft.Jet.OLEDB.4.0;"& "Data Source=" & Server.MapPath("cainiao\SqlIn.mdb")
Fy_db.Execute("insert into SqlIn(Sqlin_IP,SqlIn_Web,SqlIn_FS,SqlIn_CS,SqlIn_SJ) values('"&Request.ServerVariables("REMOTE_ADDR")&"','"&Request.ServerVariables("URL")&"','GET','"&Fy_Get&"','"&replace(Request.QueryString(Fy_Get),"'","''")&"')")
Fy_db.close
Set Fy_db = Nothing
'--------写入数据库----------尾-----

Response.Write "<Script Language=JavaScript>alert('蔡鸟SQL防注入系统提示\n\n请不要在参数中包含非法字符尝试注入！\n\n www.caiku.com ');</Script>"
Response.Write "非法操作！系统做了如下记录↓<br>"
Response.Write "操作ＩＰ："&Request.ServerVariables("REMOTE_ADDR")&"<br>"
Response.Write "操作时间："&Now&"<br>"
Response.Write "操作页面："&Request.ServerVariables("URL")&"<br>"
Response.Write "提交方式：ＧＥＴ<br>"
Response.Write "提交参数："&Fy_Get&"<br>"
Response.Write "提交数据："&Request.QueryString(Fy_Get)
Response.End
End If
Next
Next
End If
%>
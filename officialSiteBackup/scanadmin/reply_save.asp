<!--#include file="conn.asp"-->
<!--#include file="check.asp"-->
<%
action=request.querystring("action")
reply_id=request("reply_id")
guest_id=request("guest_id")
reply_name=request("reply_name")
reply_time=request("add_time")
reply_content=request("content")
select case action
'//�޸�����
case "edit" 
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from reply where reply_id=" & reply_id,conn,1,3
rs("reply_time")=reply_time
rs("reply_name")=reply_name
rs("reply_content")=reply_content
guest_id=rs("guest_id")
rs.Update
rs.Close
set rs=nothing
   response.write("<SCRIPT language=JavaScript>alert('�����޸ĳɹ���');")
   response.write("window.location.href='guest_more.asp?guest_id=" & guest_id & "'")
   response.Write("</SCRIPT>")
'ɾ������
case "del"
conn.execute ("delete from reply where reply_id=" & reply_id)
response.write("<SCRIPT language=JavaScript>alert('����ɾ���ɹ���');")
   response.write("window.location.href='guest_more.asp?guest_id=" & guest_id & "'")
   response.Write("</SCRIPT>")
end select
%>

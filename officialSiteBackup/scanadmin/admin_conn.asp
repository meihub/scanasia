<!--#include file="../data.asp" -->
<% 
db="../DataBase/"&dbname
	on error resume next
	set conn=server.createobject("adodb.connection")
	conn.open "driver={microsoft access driver (*.mdb)};dbq="&server.mappath(db)
	if err then
		err.clear
		set conn = Nothing
		response.write "���ݿ�����ά�������Ժ���ʡ�"
		response.end
	end if
	function CloseDB
		Conn.Close
		set Conn=Nothing
	End Function
%>
<% 
	dim badword
	badword="'|and|select|update|chr|delete|%20from|;|insert|mid|master.|set|chr(37)|="
    if request.QueryString<>"" then
		chk=split(badword,"|")
		for each query_name in request.querystring
			for i=0 to ubound(chk)
				if instr(lcase(request.querystring(query_name)),chk(i))<>0 then
					response.write "<script language=javascript>alert('���δ��󣡲��� "&query_name&" ��ֵ�а����Ƿ��ַ�����\n\n');location='"&request.ServerVariables("HTTP_REFERER")&"'</Script>"
					response.end
				end if
		    next
	    next
    end if
%>


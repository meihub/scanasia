

<%


' ִ��ÿ��ֻ�账��һ�ε��¼�
Call BrandNewDay()

' ��ʼ�����ݿ�����
Call DBConnBegin()

' ���ñ���
Dim sAction, sPosition
sAction = UCase(Trim(Request.QueryString("action")))



' ********************************************
' ����Ϊҳ�湫��������
' ********************************************
' ============================================
' ���ÿҳ���õĶ�������
' ============================================
Sub Header()
	Response.Write "<html><head>"
	
	' ��� meta ���
	Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & _
		"<meta name='Author' content='Andy.m'>" & _
		"<link rev=MADE href='mailto:webmaster@webasp.net'>"
	
	' �������
	Response.Write "<title>eWebEditor - eWebSoft�����ı��༭�� - ��̨����</title>"
	
    ' ���ÿҳ��ʹ�õĻ�����ʽ��
	Response.Write "<link rel='stylesheet' type='text/css' href='admin/style.css'>"

	' ���ÿҳ��ʹ�õĻ����ͻ��˽ű�
	Response.Write "<script language='javaScript' SRC='admin/private.js'></SCRIPT>"
	
	Response.Write "</head>"

	Response.Write "<body topmargin=0 leftmargin=0 bgcolor=#ffffff>"
	Response.Write "<a name=top></a>"
	





End Sub

' ============================================
' ���ÿҳ���õĵײ�����
' ============================================
Sub Footer()
	' �ͷ��������Ӷ���
	Call DBConnEnd()

	Response.Write "</td>"
	Response.Write "<td width=7></td><td width=1 class=rightbg></td>"
	Response.Write "</tr></table>"
	

	Response.Write "</body></html>"
End Sub



' ===============================================
' ��ʼ��������
'	s_FieldName	: ���ص���������	
'	a_Name		: ��ֵ������
'	a_Value		: ��ֵֵ����
'	v_InitValue	: ��ʼֵ
'	s_Sql		: �����ݿ���ȡֵʱ,select name,value from table
'	s_AllName	: ��ֵ������,��:"ȫ��","����","Ĭ��"
' ===============================================
Function InitSelect(s_FieldName, a_Name, a_Value, v_InitValue, s_Sql, s_AllName)
	Dim i
	InitSelect = "<select name='" & s_FieldName & "' size=1>"
	If s_AllName <> "" Then
		InitSelect = InitSelect & "<option value=''>" & s_AllName & "</option>"
	End If
	If s_Sql <> "" Then
		oRs.Open s_Sql, oConn, 0, 1
		Do While Not oRs.Eof
			InitSelect = InitSelect & "<option value=""" & inHTML(oRs(1)) & """"
			If oRs(1) = v_InitValue Then
				InitSelect = InitSelect & " selected"
			End If
			InitSelect = InitSelect & ">" & outHTML(oRs(0)) & "</option>"
			oRs.MoveNext
		Loop
		oRs.Close
	Else
		For i = 0 To UBound(a_Name)
			InitSelect = InitSelect & "<option value=""" & inHTML(a_Value(i)) & """"
			If a_Value(i) = v_InitValue Then
				InitSelect = InitSelect & " selected"
			End If
			InitSelect = InitSelect & ">" & outHTML(a_Name(i)) & "</option>"
		Next
	End If
	InitSelect = InitSelect & "</select>"
End Function


%>


<SCRIPT LANGUAGE="VBSCRIPT" RUNAT="SERVER">
'*** Pure ASP File Upload Version: 2.09+    ---------------------------------
' ����:
'  * ��Ҫ ADO 2.5 �����ϰ汾
'  * �����ļ�����ʽ����װ�����������Ӵ�����
'  * ����UltraDev ��MX�ķ�������Ϊ��չ
'  * �ṩ�ϴ����������ļ���С���Ƽ�飬�ļ����ͼ�飬�ļ��Ƿ��Ѿ����ڼ��
'  * ֧�� UltraDev��DW MX ����/���·�������Ϊ 
'  * ֧���ϴ����浽��ͬ���ļ���
'  * ֧������˫�ֽڱ���
'  * ʹ��ʱ��������Ϊ�ļ�����ͬ��֧�������ļ����ļ��ϴ�
'  * �����޸ģ�Keith
'  * �����ʼ���fastpace@163.com
'  * ����޸ģ�2003-09-24
'------------------------------------------------------------------------------
Function getPureUploadVersion()
  getPureUploadVersion = 2.09
End Function

Sub BuildUploadRequest(RequestBin,UploadDirectory,storeType,sizeLimit,nameConflict)

  Dim PosBeg, PosEnd, checkADOConn, AdoVersion, Length, boundary, boundaryPos, Pos
  Dim PosFile, Name, PosBound, FileName, ContentType, Value, ValueBeg, ValueEnd, ValueLen
  
  'Get the boundary
  PosBeg = 1
  PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
  if PosEnd = 0 then
    Response.Write "<b>�ύ������ ENCTYPE=""multipart/form-data""����</b><br>"
    Response.Write "���޸ĺ󰴷��� <A HREF=""javascript:history.back(1)"">����</a>"    
    Response.End
  end if
  'Check ADO Version
	set checkADOConn = Server.CreateObject("ADODB.Connection")
  on error resume next
	adoVersion = CSng(checkADOConn.Version)
	if err then 
		adoVersion = Replace(checkADOConn.Version,".",",")  
		adoVersion = CSng(adoVersion)
	end if	
  	err.clear
  on error goto 0	
	set checkADOConn = Nothing
	if adoVersion < 2.5 then
    Response.Write "<b>��Ļ�����û�а�װ ADO 2.5�����ϵİ汾</b><br>"
    Response.Write "�ϴ��ļ� ��Ҫ ADO 2.5 �����ϰ汾��������<br>"
    Response.Write "����Դ�<a href=""www.microsoft.com/data"">www.microsoft.com/data</a>���������µ�MDAC (ADO is included) �汾<br>"
    Response.End
	end if		
  'Check content length if needed    �����Ҫ������ݳ���
	Length = CLng(Request.ServerVariables("HTTP_Content_Length")) 'Get Content-Length header
	If "" & sizeLimit <> "" Then
    sizeLimit = CLng(sizeLimit) * 1024
    If Length > sizeLimit Then
      Request.BinaryRead (Length)
      Response.Write "�ϴ��ļ���С " & FormatNumber(Length, 0) & "B ���������Ƶ�" & FormatNumber(sizeLimit, 0) & "B"
      Response.Write "������󣬰����˰�ť��<A HREF=""javascript:history.back(1)"">���ԣ�</a>"      
      Response.End
    End If
  End If
  boundary = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
  boundaryPos = InstrB(1,RequestBin,boundary)    
  'Get all data inside the boundaries
  Do until (boundaryPos=InstrB(RequestBin,boundary & getByteString("--")))
    'Members variable of objects are put in a dictionary object
    Dim UploadControl
    Set UploadControl = CreateObject("Scripting.Dictionary")
    'Get an object name 
    Pos = InstrB(BoundaryPos,RequestBin,getByteString("Content-Disposition"))
    Pos = InstrB(Pos,RequestBin,getByteString("name="))
    PosBeg = Pos+6
    PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(34)))
    Name = LCase(getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg)))
    PosFile = InstrB(BoundaryPos,RequestBin,getByteString("filename="))
    PosBound = InstrB(PosEnd,RequestBin,boundary)
    'Test if object is of file type
    If  PosFile<>0 AND (PosFile<PosBound) Then
      'Get Filename, content-type and content of file 
      PosBeg = PosFile + 10
      PosEnd =  InstrB(PosBeg,RequestBin,getByteString(chr(34)))
      FileName = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
	  FileName = Mid(FileName,InStrRev(FileName,"\")+1)
     '��ȡ�ļ������ļ�����
      FileNameArray=Split(FileName,".")
      FileNameArray(0)=makefilename(FileNameArray(0))
      FileName=FileNameArray(0)&"."&FileNameArray(1)
	  'Add filename to dictionary object
      UploadControl.Add "FileName", FileName
      Pos = InstrB(PosEnd,RequestBin,getByteString("Content-Type:"))
      PosBeg = Pos+14
      PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
      'Add content-type to dictionary object
      ContentType = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
      UploadControl.Add "ContentType",ContentType
      'Get content of object  
      PosBeg = PosEnd+4
      PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
      Value = FileName
      ValueBeg = PosBeg-1
      ValueLen = PosEnd-Posbeg
    Else
      'Get content of object  
      Pos = InstrB(Pos,RequestBin,getByteString(chr(13)))
      PosBeg = Pos+4
      PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
      Value = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
      ValueBeg = 0
      ValueEnd = 0
    End If
    'Add content to dictionary object
    UploadControl.Add "Value" , Value	
    UploadControl.Add "ValueBeg" , ValueBeg
    UploadControl.Add "ValueLen" , ValueLen	
    'Add dictionary object to main dictionary
    if UploadRequest.Exists(name) then
      UploadRequest(name).Item("Value") = UploadRequest(name).Item("Value") & "," & Value
    else
      UploadRequest.Add name, UploadControl 
    end if    
    BoundaryPos=InstrB(BoundaryPos+LenB(boundary),RequestBin,boundary)
  Loop

  Dim GP_keys, GP_i, GP_curKey, GP_value, GP_valueBeg, GP_valueLen, GP_curPath, GP_FullPath
  Dim GP_CurFileName, GP_FullFileName, fso, GP_BegFolder, GP_RelFolder, GP_FileExist, Begin_Name_Num
  Dim orgUploadDirectory
    
  if InStr(UploadDirectory,"""") > 0 then 
    on error resume next
    orgUploadDirectory = UploadDirectory
    UploadDirectory = eval(UploadDirectory)  
    if err then
      Response.Write "<B>�ϴ��ļ�����Ч</B><br><br>"      
      Response.Write "�ϴ��ļ���: " & Trim(orgUploadDirectory) & "<br>"
      Response.Write "������󣬰����˰�ť��<A HREF=""javascript:history.back(1)"">���ԣ�</a>"
		  err.clear
 	    response.End
    end if    
    on error goto 0
  end if  
  
  GP_keys = UploadRequest.Keys
  for GP_i = 0 to UploadRequest.Count - 1
    GP_curKey = GP_keys(GP_i)
    if UploadRequest.Item(GP_curKey).Item("FileName") <> "" then
      GP_value = UploadRequest.Item(GP_curKey).Item("Value")
      GP_valueBeg = UploadRequest.Item(GP_curKey).Item("ValueBeg")
      GP_valueLen = UploadRequest.Item(GP_curKey).Item("ValueLen")

      if InStr(UploadDirectory,"\") > 0 then
        GP_curPath = UploadDirectory
        if Mid(GP_curPath,Len(GP_curPath),1)  <> "\" then
          GP_curPath = GP_curPath & "\"
        end if         
        GP_FullPath = GP_curPath
      else
        GP_curPath = Request.ServerVariables("PATH_INFO")
        GP_curPath = Trim(Mid(GP_curPath,1,InStrRev(GP_curPath,"/")) & UploadDirectory)
        if Mid(GP_curPath,Len(GP_curPath),1)  <> "/" then
          GP_curPath = GP_curPath & "/"
        end if 
        GP_FullPath = Trim(Server.mappath(GP_curPath))
      end if
      
      
      if GP_valueLen = 0 then
        Response.Write "<B>�����ϴ��ļ�ʱ����һ������!</B><br><br>"
        Response.Write "�ļ���: " & Trim(GP_curPath) & UploadRequest.Item(GP_curKey).Item("FileName") & "<br>"
        Response.Write "�ļ������ڻ���Ϊ��<br>"
        Response.Write "������󣬰����˰�ť��<A HREF=""javascript:history.back(1)"">���ԣ�</a>"
	  	  response.End
	    end if
      
      'Create a Stream instance
      Dim GP_strm1, GP_strm2
      Set GP_strm1 = Server.CreateObject("ADODB.Stream")
      Set GP_strm2 = Server.CreateObject("ADODB.Stream")
      
      'Open the stream
      GP_strm1.Open
      GP_strm1.Type = 1 'Binary
      GP_strm2.Open
      GP_strm2.Type = 1 'Binary
        
      GP_strm1.Write RequestBin
      GP_strm1.Position = GP_ValueBeg
      GP_strm1.CopyTo GP_strm2,GP_ValueLen
    
      'Create and Write to a File
      GP_CurFileName = UploadRequest.Item(GP_curKey).Item("FileName")      
      GP_FullFileName = GP_FullPath & "\" & GP_CurFileName
      Set fso = CreateObject("Scripting.FileSystemObject")
      'Check if the folder exist
      If NOT fso.FolderExists(GP_FullPath) Then
        GP_BegFolder = InStr(GP_FullPath,"\")
        while GP_begFolder > 0 
          GP_RelFolder = Mid(GP_FullPath,1,GP_BegFolder-1)
          If NOT fso.FolderExists(GP_RelFolder) Then  
            fso.CreateFolder(GP_RelFolder)
          end if          
          GP_BegFolder = InStr(GP_BegFolder+1,GP_FullPath,"\")          
        wend
        If NOT fso.FolderExists(GP_FullPath) Then        
          fso.CreateFolder(GP_FullPath)        
        end if  
      end if
      GP_FileExist = false
      If fso.FileExists(GP_FullFileName) Then
        GP_FileExist = true
      End If      
      if nameConflict = "error" and GP_FileExist then
        Response.Write "<B>�ļ��Ѿ�����!</B><br><br>"
        Response.Write "������󣬰����˰�ť��<A HREF=""javascript:history.back(1)"">���ԣ�</a>"
				GP_strm1.Close
				GP_strm2.Close
	  	  response.End
      end if
      if ((nameConflict = "over" or nameConflict = "uniq") and GP_FileExist) or (NOT GP_FileExist) then
        if nameConflict = "uniq" and GP_FileExist then
          Begin_Name_Num = 0
          while GP_FileExist    
            Begin_Name_Num = Begin_Name_Num + 1
            GP_FullFileName = Trim(GP_FullPath)& "\" & fso.GetBaseName(GP_CurFileName) & "_" & Begin_Name_Num & "." & fso.GetExtensionName(GP_CurFileName)
            GP_FileExist = fso.FileExists(GP_FullFileName)
          wend  
          UploadRequest.Item(GP_curKey).Item("FileName") = fso.GetBaseName(GP_CurFileName) & "_" & Begin_Name_Num & "." & fso.GetExtensionName(GP_CurFileName)
					UploadRequest.Item(GP_curKey).Item("Value") = UploadRequest.Item(GP_curKey).Item("FileName")
        end if
        on error resume next
        GP_strm2.SaveToFile GP_FullFileName,2
        if err then
          Response.Write "<B>�����ϴ��ļ�ʱ����һ������!</B><br><br>"
          Response.Write "�ļ���: " & Trim(GP_curPath) & UploadRequest.Item(GP_curKey).Item("FileName") & "<br>"
          Response.Write "����Ŀ��Ŀ¼������, ������û��д��Ȩ����<br>"
          Response.Write "������󣬰����˰�ť��<A HREF=""javascript:history.back(1)"">���ԣ�</a>"
    		  err.clear
  				GP_strm1.Close
  				GP_strm2.Close
  	  	  response.End
  	    end if
  			GP_strm1.Close
  			GP_strm2.Close
  			if storeType = "path" then
  				UploadRequest.Item(GP_curKey).Item("Value") = GP_curPath & UploadRequest.Item(GP_curKey).Item("Value")
  			end if
        on error goto 0
      end if
    end if
  next

End Sub

'String to byte string conversion 
Function getByteString(StringStr)
  getByteString = ""
  Dim i, char
  For i = 1 to Len(StringStr)
 	  varchar = Mid(StringStr,i,1)
	    varasc = Asc(varchar)
       If varasc<0 Then
          varasc = varasc + 65535
       End If
       If varasc>255 Then
          vardovialow = Left(Hex(Asc(varchar)),2)
          vardoviahigh = right(Hex(Asc(varchar)),2)
          getByteString = getByteString & chrB("&H" & vardovialow) & chrB("&H" & vardoviahigh)
       Else
	  getByteString = getByteString & chrB(AscB(varchar))
       End If
  Next
End Function

'Byte string to string conversion (with double-byte support now)
Function getString(StringBin)
  Dim doviaskipflag,get1Byte,vardovialen,vardoviastring
  vardovialen =""
   doviaskipflag=0
 If Not IsNull(StringBin) Then 
      vardovialen = LenB(StringBin) 
  For i = 1 to vardovialen
  If doviaskipflag=0 Then
    get1Byte = MidB(StringBin,i,1)
    If AscB(get1Byte) > 127 Then 
    vardoviastring =vardoviastring & Chr(AscW(MidB(StringBin,i+1,1) & get1Byte))
                doviaskipflag=1
             Else
       vardoviastring = vardoviastring & chr(AscB(get1Byte)) 
             End If
          Else
             doviaskipflag=0
	          End If
  Next
  end if
  getstring = vardoviastring
End Function

Function UploadFormRequest(name)
 Dim keyName
  keyName = LCase(name)                 
  if IsObject(UploadRequest) then      
    if UploadRequest.Exists(keyName) then
      if UploadRequest.Item(keyName).Exists("Value") then
        UploadFormRequest = UploadRequest.Item(keyName).Item("Value")
      end if  
    end if  
  end if  
End Function

'�ѵ�ǰʱ����Ϊ�ļ���
Function makefilename(fname)
  fname = now()
  fname = replace(fname,"-","")
  fname = replace(fname," ","") 
  fname = replace(fname,":","")
  fname = replace(fname,"PM","")
  fname = replace(fname,"AM","")
  fname = replace(fname,"����","")
  fname = replace(fname,"����","")
  makefilename=fname
End Function 

Function RemoveInvalidChars(str)
  Dim newStr, ci, curChar
  for ci = 1 to Len(str)  
    curChar = Asc(LCase(Mid(str,ci,1)))  
    if curChar = 95 or curChar = 45 or curChar = 46 or (curChar >= 97 and curChar <= 122) or (curChar >= 48 and curChar <= 57) then
      newStr = newStr & Mid(str,ci,1)
    end if
  next
 RemoveInvalidChars = newStr
End Function

Sub PureUploadSetup()
  UploadQueryString = Replace(Request.QueryString,"GP_upload=true","")
  if mid(UploadQueryString,1,1) = "&" then
  	UploadQueryString = Mid(UploadQueryString,2)
  end if
  GP_uploadAction = CStr(Request.ServerVariables("URL")) & "?GP_upload=true"
  If (Request.QueryString <> "") Then  
    if UploadQueryString <> "" then
  	  GP_uploadAction = GP_uploadAction & "&" & UploadQueryString
    end if 
  End If
End Sub

Function FixFieldsForUpload(GP_fieldsStr, GP_columnsStr)
  Dim GP_counter, GP_Fields, GP_Columns, GP_FieldName, GP_FieldValue, GP_CurFileName, GP_CurContentType

  GP_Fields = Split(GP_fieldsStr, "|")
  GP_Columns = Split(GP_columnsStr, "|") 
  GP_fieldsStr = ""
  ' ȡ�ñ�ֵ
  For GP_counter = LBound(GP_Fields) To UBound(GP_Fields) Step 2
    GP_FieldName = LCase(GP_Fields(GP_counter))
    GP_FieldValue = GP_Fields(GP_counter+1)
  	if UploadRequest.Exists(GP_FieldName) then
      GP_CurFileName = UploadRequest.Item(GP_FieldName).Item("FileName")
      GP_CurContentType = UploadRequest.Item(GP_FieldName).Item("ContentType")
  	else  
	    GP_CurFileName = ""
	    GP_CurContentType = ""
  	end if	
    if (GP_CurFileName = "" and GP_CurContentType = "") or (GP_CurFileName <> "" and GP_CurContentType <> "") then
      GP_fieldsStr = GP_fieldsStr & GP_FieldName & "|" & GP_FieldValue & "|"
    end if 
  Next
  if GP_fieldsStr <> "" then
    GP_fieldsStr = Mid(GP_fieldsStr,1,Len(GP_fieldsStr)-1)
  else  
    Response.Write "<B>���¼�¼ʱ������һ������!</B><br><br>"
    Response.Write "δ�ҵ������ֶ� ...<br>"
    Response.Write "���ϴ��ļ��ֶν�Ϊ���ֶΣ����������ݿ��д�ʼ����ֶΡ�<br>"
    Response.Write "�������������˰�ť��<A HREF=""javascript:history.back(1)"">���ԣ�</a>"
    Response.End
  end if
  
  FixFieldsForUpload = GP_fieldsStr    
End Function

Function FixColumnsForUpload(GP_fieldsStr, GP_columnsStr)
  Dim GP_counter, GP_Fields, GP_Columns, GP_FieldName, GP_ColumnName, GP_ColumnValue,GP_CurFileName, GP_CurContentType

  GP_Fields = Split(GP_fieldsStr, "|")
  GP_Columns = Split(GP_columnsStr, "|") 
  GP_columnsStr = "" 
  ' ȡ�ñ�ֵ
  For GP_counter = LBound(GP_Fields) To UBound(GP_Fields) Step 2
    GP_FieldName = LCase(GP_Fields(GP_counter))  
    GP_ColumnName = GP_Columns(GP_counter)  
    GP_ColumnValue = GP_Columns(GP_counter+1)
  	if UploadRequest.Exists(GP_FieldName) then
  	  GP_CurFileName = UploadRequest.Item(GP_FieldName).Item("FileName")
  	  GP_CurContentType = UploadRequest.Item(GP_FieldName).Item("ContentType")	  
  	else  
  	  GP_CurFileName = ""
  	  GP_CurContentType = ""
  	end if  
    if (GP_CurFileName = "" and GP_CurContentType = "") or (GP_CurFileName <> "" and GP_CurContentType <> "") then
      GP_columnsStr = GP_columnsStr & GP_ColumnName & "|" & GP_ColumnValue & "|"
    end if 
  Next
  if GP_columnsStr <> "" then
    GP_columnsStr = Mid(GP_columnsStr,1,Len(GP_columnsStr)-1)    
  end if
  FixColumnsForUpload = GP_columnsStr
End Function

</SCRIPT>

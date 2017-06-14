<SCRIPT LANGUAGE="VBSCRIPT" RUNAT="SERVER">
'*** Pure ASP File Upload -----------------------------------------------------
' Copyright 2000 (c) George Petrov, www.UDzone.com
'
' Script partially based on code from Philippe Collignon 
'              (http://www.asptoday.com/articles/20000316.htm)
'
' New features from GP:
'  * Fast file save with ADO 2.5 stream object
'  * new wrapper functions, extra error checking
'  * UltraDev Server Behavior extension

' Copyright 2001-2002 (c) George Petrov and Modify by xPilot
' *** Date: 12/15/2001 ***
' *** 支持所有双字节文件名，而且修复了原函数中遇到空格也会自动截断文件名的错误！ ***
' *** 保证百分百以原文件名保存上传文件！***
' *** Welcome to visite pilothome.yeah.net or mail xpilot@21cn.com to me！***
'
' Version: 2.0.2 Beta for UTF-8,BIG5,Japan,Korea ...
'*** Pure ASP File Upload -----------------------------------------------------
Sub BuildUploadRequest(RequestBin,UploadDirectory,storeType,sizeLimit,nameConflict)
  'Get the boundary
  PosBeg = 1
  PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
  if PosEnd = 0 then
    Response.Write "<b>Form was submitted with no ENCTYPE=""multipart/form-data""</b><br>"
    Response.Write "Please correct and <A HREF=""javascript:history.back(1)"">try again</a>"    
    Response.End
  end if
  'Check ADO Version
	set checkADOConn = Server.CreateObject("ADODB.Connection")
	adoVersion = CSng(checkADOConn.Version)
	set checkADOConn = Nothing
	if adoVersion < 2.5 then
    Response.Write "<b>You don't have ADO 2.5 installed on the server.</b><br>"
    Response.Write "The File Upload extension needs ADO 2.5 or greater to run properly.<br>"
    Response.Write "You can download the latest MDAC (ADO is included) from <a href=""www.microsoft.com/data"">www.microsoft.com/data</a><br>"
    Response.End
	end if		
  'Check content length if needed
	Length = CLng(Request.ServerVariables("HTTP_Content_Length")) 'Get Content-Length header
	If "" & sizeLimit <> "" Then
    sizeLimit = CLng(sizeLimit)
    If Length > sizeLimit Then
      Request.BinaryRead (Length)
      Response.Write "Upload size " & FormatNumber(Length, 0) & "B exceeds limit of " & FormatNumber(sizeLimit, 0) & "B"
      Response.Write "Please correct and <A HREF=""javascript:history.back(1)"">try again</a>"      
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
    Name = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
    PosFile = InstrB(BoundaryPos,RequestBin,getByteString("filename="))
    PosBound = InstrB(PosEnd,RequestBin,boundary)
    'Test if object is of file type
    If  PosFile<>0 AND (PosFile<PosBound) Then
      'Get Filename, content-type and content of file
      PosBeg = PosFile + 10
      PosEnd =  InstrB(PosBeg,RequestBin,getByteString(chr(34)))
      FileName = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
      FileName = Mid(FileName,InStrRev(FileName,"\")+1)
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
    UploadRequest.Add name, UploadControl	
    'Loop to next object
    BoundaryPos=InstrB(BoundaryPos+LenB(boundary),RequestBin,boundary)
  Loop

  GP_keys = UploadRequest.Keys
  for GP_i = 0 to UploadRequest.Count - 1
    GP_curKey = GP_keys(GP_i)
    'Save all uploaded files
    if UploadRequest.Item(GP_curKey).Item("FileName") <> "" then
      GP_value = UploadRequest.Item(GP_curKey).Item("Value")
      GP_valueBeg = UploadRequest.Item(GP_curKey).Item("ValueBeg")
      GP_valueLen = UploadRequest.Item(GP_curKey).Item("ValueLen")

      if GP_valueLen = 0 then
        Response.Write "<B>An error has occured saving uploaded file!</B><br><br>"
        Response.Write "Filename: " & Trim(GP_curPath) & UploadRequest.Item(GP_curKey).Item("FileName") & "<br>"
        Response.Write "File does not exists or is empty.<br>"
        Response.Write "Please correct and <A HREF=""javascript:history.back(1)"">try again</a>"
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
      GP_curPath = Request.ServerVariables("PATH_INFO")
      GP_curPath = Trim(Mid(GP_curPath,1,InStrRev(GP_curPath,"/")) & UploadDirectory)
      if Mid(GP_curPath,Len(GP_curPath),1)  <> "/" then
        GP_curPath = GP_curPath & "/"
      end if 
      GP_CurFileName = UploadRequest.Item(GP_curKey).Item("FileName")
      GP_FullPath = Trim(Server.mappath(GP_curPath))
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
      'Check if the file already exist
      GP_FileExist = false
      If fso.FileExists(GP_FullFileName) Then
        GP_FileExist = true
      End If      
      if nameConflict = "error" and GP_FileExist then
        Response.Write "<B>File already exists!</B><br><br>"
        Response.Write "Please correct and <A HREF=""javascript:history.back(1)"">try again</a>"
				GP_strm1.Close
				GP_strm2.Close
	  	  response.End
      end if
      if ((nameConflict = "over" or nameConflict = "uniq") and GP_FileExist) or (NOT GP_FileExist) then
        if nameConflict = "uniq" and GP_FileExist then
          Begin_Name_Num = 0
          while GP_FileExist    
            Begin_Name_Num = Begin_Name_Num + 1
            GP_FullFileName = Trim(Server.mappath(GP_curPath))& "\" & fso.GetBaseName(GP_CurFileName) & "_" & Begin_Name_Num & "." & fso.GetExtensionName(GP_CurFileName)
            GP_FileExist = fso.FileExists(GP_FullFileName)
          wend  
          UploadRequest.Item(GP_curKey).Item("FileName") = fso.GetBaseName(GP_CurFileName) & "_" & Begin_Name_Num & "." & fso.GetExtensionName(GP_CurFileName)
					UploadRequest.Item(GP_curKey).Item("Value") = UploadRequest.Item(GP_curKey).Item("FileName")
        end if
        on error resume next
        GP_strm2.SaveToFile GP_FullFileName,2
        if err then
          Response.Write "<B>An error has occured saving uploaded file!</B><br><br>"
          Response.Write "Filename: " & Trim(GP_curPath) & UploadRequest.Item(GP_curKey).Item("FileName") & "<br>"
          Response.Write "Maybe the destination directory does not exist, or you don't have write permission.<br>"
          Response.Write "Please correct and <A HREF=""javascript:history.back(1)"">try again</a>"
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

'把普通字符串转成二进制字符串函数
Function getByteString(StringStr)
    getByteString=""
  For i = 1 To Len(StringStr) 
    XP_varchar = mid(StringStr,i,1)
    XP_varasc = Asc(XP_varchar) 
    If XP_varasc < 0 Then 
       XP_varasc = XP_varasc + 65535 
    End If 

    If XP_varasc > 255 Then 
       XP_varlow = Left(Hex(Asc(XP_varchar)),2) 
       XP_varhigh = right(Hex(Asc(XP_varchar)),2) 
       getByteString = getByteString & chrB("&H" & XP_varlow) & chrB("&H" & XP_varhigh) 
    Else 
       getByteString = getByteString & chrB(AscB(XP_varchar)) 
    End If 
  Next 
End Function

'把二进制字符串转换成普通字符串函数 
Function getString(StringBin)
   getString =""
   Dim XP_varlen,XP_vargetstr,XP_string,XP_skip
   XP_skip = 0 
   XP_string = "" 
 If Not IsNull(StringBin) Then 
      XP_varlen = LenB(StringBin) 
    For i = 1 To XP_varlen 
      If XP_skip = 0 Then
         XP_vargetstr = MidB(StringBin,i,1) 
         If AscB(XP_vargetstr) > 127 Then 
           XP_string = XP_string & Chr(AscW(MidB(StringBin,i+1,1) & XP_vargetstr)) 
           XP_skip = 1 
         Else 
           XP_string = XP_string & Chr(AscB(XP_vargetstr)) 
         End If 
      Else 
      XP_skip = 0
   End If 
    Next 
 End If 
      getString = XP_string 
End Function 

Function UploadFormRequest(name)
  on error resume next
  if UploadRequest.Item(name) then
    UploadFormRequest = UploadRequest.Item(name).Item("Value")
  end if  
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
</SCRIPT>


function isObj(obj)
  if (not IsObject(obj)) then    '' (typeof(obj) != "object")
    isObj = false
    exit function
  elseif (obj Is Nothing) then   '' (obj == null)
    isObj = false
    exit function
  end if
  isObj = true                   '' isNotNullObj
end function

function getURL(url)
  Dim oHTTP
  ''set oHTTP = window.XMLHttpRequest
  set oHTTP = CreateObject("Microsoft.XMLHTTP")
  On Error Resume Next

  oHTTP.Open "GET", url, false
  if Err.Number<>0 then
    getURL = "oHTTP.Open::Err#" & Err.Number & ": " & Err.Description & vbCrLf & "[" & url & "]"
    exit function
  end if

  oHTTP.Send
  if Err.Number<>0 then
    getURL = "oHTTP.Send::Err#" & Err.Number & ": " & Err.Description & vbCrLf & "[" & url & "]"
    exit function
  end if

  getURL = oHTTP.responseText
  Set oHTTP = Nothing
end function

function StrB(str)
  ''covert double byte unicode string to single byte string
  Dim s
  s = ""
  for i = 1 to Len(str)
    ''UNICODE :: c == ChrB(AscB(c)) & ChrB(0)
    s = s & ChrB(AscB(Mid(str,i,1)))
  next
  StrB = s
end function

function getChrVal(str, idx, opt)
  select case opt
   case "ASCB" : getChrVal=AscB(Mid(str, idx, 1))
   case "ASCW" : getChrVal=AscW(Mid(str, idx, 1))
   case  else  : getChrVal=Asc(Mid(str, idx, 1))
  end select
end function

function getSurrogatePairVal(charcode1, charcode2)
  '' "astral characters" are represented by 2 virtual characters
  '' this is the way how UTF-16 works; chrw(132523|x205AB) == chrw(55361|xD841) & chrw(56747|xDDAB)
  '' http://en.wikipedia.org/wiki/Mapping_of_Unicode_characters#Surrogates
  if (charcode1>55295 and charcode1<56320 and charcode2>56319 and charcode2<57344) then
    '' u = &H10000 + (h-&HD800)*&H0400 + (l-&HDC00)
    getSurrogatePairVal = 65536 + (charcode1-55296)*1024 + (charcode2-56320)
  else
    '' Invalid Surrogate Pair !!
  end if
end function

function ChrU(charcode)
  Dim c, c1, c2
  c = charcode
  if c <= 65535 then                   '' first plane, plane 0, Basic Multilingual Plane (BMP)
    ChrU = ChrW(c)
  elseif c <= 2097151 then             '' astral character
    '' not supported by ChrW directly, need to break into 2 virtual char [http://en.wikipedia.org/wiki/UTF-16]
    c  = c - 65536                     '' subtract 0x10000 from the code point, leaving a 20 bit number [0..0xFFFFF]
    c1 = 55296 + (c \ 2^10)            '' high/lead surrogate == top 10 bits + 0xD800
    c2 = 56320 + (c AND (2^10)-1)      '' low/trail surrogate == low 10 bits + 0xDC00
    ChrU  = ChrW(c1) & ChrW(c2)        '' surrogate pair [UTF16 standard]
  else                                 '' illegal unicode character !!
    ChrU = "[x" & c & "]&#" & c & ";"
  end if
end function

function decodeUTF8(str, opt)
  Dim objStream, singleByteString
  On Error Resume Next
    ''HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\ActiveX Compatibility\{00000566-0000-0010-8000-00AA006D2EA4}
    ''MS has disabled use of ado stream object at the client due to security issues. There were just too many security issues with rogue web pages attempting to read and write files on the user's machine.
    Set objStream = CreateObject("ADODB.Stream")
  On Error Goto 0
  if isObj(objStream) then
    ''str = "objStream:" & str
    with objStream
      .Open : .Type=2         '' adTypeBinary=1 adTypeText=2
      .CharSet="ISO-8859-1"   '' do not use X-ANSI/Windows-1252; default="Unicode"
      .WriteText str
      '---read it back and decode the data using a new CharSet rule----'
      .Position=0 : .Type=2 : .CharSet="UTF-8"
      decodeUTF8 = .ReadText
    end with
    exit function
  end if

  Dim l, i, c, n, n2, n3, n4, s

  s="" : l=len(str)
  i=1  : do while (i <= l)
    c = Mid(str, i, 1)
    n = getChrVal(str, i, opt) '' AscW(c)
    if n < 128 then
      ''valid single byte UTF-8 character encoding [%00 : %7F]
    elseif (n > 191 and n < 224) then
      ''double byte UTF-8 character encoding [%C0%xx : %DF%xx]
      n2 = getChrVal(str, i+1, opt)
      if (n2 > 127 and n2 < 192) then
        n = ((n and &H1F) * 2^6) or (n2 and &H3F)
        i = i + 1
      else
        ''illegal 2nd byte of UTF-8 !!
      end if
    elseif (n > 223 and n < 240) and (i+2 <= l) then
      ''triple byte UTF-8 character encoding [%E0%xx%xx : %EF%xx%xx]
      n2 = getChrVal(str, i+1, opt)
      n3 = getChrVal(str, i+2, opt)
      if (n2 > 127 and n2 < 192) and (n3 > 127 and n2 < 192) then
        n = ((n and &H0F) * 2^12) or ((n2 and &H3F) * 2^6) or (n3 and &H3F)
        i = i + 2
      else
        ''illegal 2nd/3rd byte of UTF-8 !!
      end if
    elseif (n > 239) and (i+3 <= l) then
      ''quad byte UTF-8 character encoding [%F0%xx%xx%xx : %FF%xx%xx%xx]
      n2 = getChrVal(str, i+1, opt)
      n3 = getChrVal(str, i+2, opt)
      n4 = getChrVal(str, i+3, opt)
      if (n2 > 127 and n2 < 192) and (n3 > 127 and n2 < 192) _
       and (n4 > 127 and n4 < 192) then
        n = ((n and &H07) * 2^18) or ((n2 and &H3F) * 2^12) _
         or ((n3 and &H3F) * 2^6) or (n4 and &H3F)
        i = i + 3
      else
        ''illegal 2nd/3rd/4th byte of UTF-8 !!
      end if
    elseif (n > 127 and n < 192) then
      ''illegal single byte UTF-8 character encoding! [%80 : %BF]
    end if
    s = s & ChrU(n)
    i = i + 1
  loop
  decodeUTF8 = s
end function

function bin2str(binData)
  Dim rs, binDataSz, strData, idx
  binDataSz = LenB(binData)
  if binDataSz > 0 Then
    ' -- This is a less efficient method (??) ---
    'for idx = 1 to binDataSz
    '  strData = strData & Chr(AscB(MidB( binData, idx, 1)))
    'next
    ' -- Below is another method by using the ADO object ---
    set rs = CreateObject("ADODB.Recordset")
    rs.fields.append "data", 201, binDataSz   '' adLongVarChar = 201
    rs.Open : rs.AddNew
    rs.fields(0).AppendChunk binData
    rs.Update
    strData = rs.fields(0).value
    rs.close
  end if
  bin2str = strData
end function

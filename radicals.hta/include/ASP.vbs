Dim Session, Request, Response
'** Note: Maximum URL length is 2083 characters in Internet Explorer **'

function HTMLEncode(str)
  Dim l, i, c, c0, s
  l = Len(str)
  For i = 1 To l
      c = getChrVal(str, i, "ASCW")                    '' [ -32768..[0]..65535 ]
      if (c < 0) Then c = 65536 + c '' 0xFFF '' http://support.microsoft.com/kb/272138
      If c = 0 Then 
        ' skip
      ElseIf ((c > 55295) and (c < 56320)) Then        '' [%D8%xx : %DB%xx] high/lead surrogates
        c0 = c  '' [keep for processsing on next char]
      ElseIf ((c > 56319) and (c < 57344)) Then        '' [%DC%xx : %DF%xx] low/trail surrogates
        '' [c0==high/lead surrogates]
        '' c0 = ascw(mid(str, i-1, 1)) : if c0 < 0 then c0 = 65536 + c0
        c = getSurrogatePairVal(c0, c)                 '' [%F0%xx%xx%xx : %FF%xx%xx%xx] astral characters
        s = s & "&#x" & Hex(c) & ";"
      ElseIf (c<48 Or c>126) Or (c>56 And c<=64) Then  '' 
        ''s = s & "&#" & c & ";"
        s = s & "&#x" & Hex(c) & ";"
      Else
        s = s & ChrW(c)
      End If
  Next
  HTMLEncode = s
end function

function HTMLDecode(str)
  Dim l, i, c, c0, j, n, nn, s
  s="" : l=len(str)
  i=1  : do while (i <= l)
    c = Mid(str, i, 1)
    if c = "&" then
      if i+2 <= l then
        j = i + 1
        if Mid(str,j,1) = "#" then
          j = j + 1
          n = Mid(str, j, 1)
          if (n="x" or n="X") then
            nn = "&H"
          elseif IsNumeric(n) then
            nn = n
          else
            nn = ""
          end if
          if (nn<>"") then 
            do while (j <= l)
              j = j + 1
              n = Mid(str, j, 1)
              if n=";" then exit do
              n = nn & n
              if not IsNumeric(n) then
                j = j - 1
                exit do
              end if
              nn = n
            loop
            if not (nn="" or nn="&H") then
              c = ChrU(CLng(nn))
              i = j
            end if
          end if
        end if
      end if
    end if
    s = s & c
    i = i + 1
  loop
  HTMLDecode = s
end function

Function URLEncode(ByVal str)
'  On Error Resume Next
'    '' just use the Javascript function
'    URLEncode = escape(str)
'    if Err.Number=0 then
'      exit function
'    end if
'  On Error Goto 0

  '' adapted from javascript at http://www.webtoolkit.info/javascript-url-decode-encode.html
  Dim i, c, c0, s
  s = ""
  for i = 1 to len(str)
    c = getChrVal(str, i, "ASCW")              '' [ -32768..[0]..65535 ]
    if c < 0 then c = 65536 + c '' http://support.microsoft.com/kb/272138
    if (c < 128) then                          '' [%00 : %7F]
      c = Hex(c) : if len(c) = 1 then c = "0" & c
      s = s & "%" & c
    elseif ((c > 127) and (c < 2048)) then     '' [%C0%xx : %DF%xx]
      s = s & "%" & Hex((c\2^6) OR 192) & _
              "%" & Hex((c AND 63) OR 128)
    elseif ((c > 55295) and (c < 56320)) then  '' [%D8%xx : %DB%xx] high/lead surrogates
      c0 = c  '' [keep for processsing on next char]
    elseif ((c > 56319) and (c < 57344)) then  '' [%DC%xx : %DF%xx] low/trail surrogates
      '' [c0==high/lead surrogates]
      '' c0 = ascw(mid(str, i-1, 1)) : if c0 < 0 then c0 = 65536 + c0
      c = getSurrogatePairVal(c0, c)           '' [%F0%xx%xx%xx : %FF%xx%xx%xx] astral characters
      s = s & "%" & Hex((c\2^18) OR 240) & _
              "%" & Hex(((c\2^12) AND 4095) OR 128) & _
              "%" & Hex(((c\2^6) AND 63) OR 128) & _
              "%" & Hex((c AND 63) OR 128)
    elseif (c <= 65535) then                   '' [%E0%xx%xx : %EF%xx%xx]
      s = s & "%" & Hex((c\2^12) OR 224) & _
              "%" & Hex(((c\2^6) AND 63) OR 128) & _
              "%" & Hex((c AND 63) OR 128)
    else
      '' ascw will not return anything above 65535, in fact vbscript
      '' will break those "astral characters" into 2 virtual characters
      '' this is the way how UTF-16 works
      '' chrw(132523|x205AB) == chrw(55361|xD841) & chrw(56747|xDDAB)
      ''       "high/lead surrogates" [D8xx-DBxx] & "low/trail surrogates" [DCxx-DFxx]
      '' http://en.wikipedia.org/wiki/Mapping_of_Unicode_characters#Surrogates
      '' u = &H10000 + (h-&HD800)*&H0400 + (l-&HDC00)
      '' u = 65536 + (h-55296)*1024 + (l-56320)
      s = s & "%" & c
    end if
  next

  URLEncode = s
End Function

function URLDecode(str)
  On Error Resume Next
    '' just use the Javascript function
    URLDecode = decodeUTF8(replace(unescape(str),"+"," "), "ASCW")
    if Err.Number=0 then
      exit function
    end if
  On Error Goto 0

  Dim l, i, c, s
  s="" : l=len(str)
  i=1  : do while (i <= l)
    c = Mid(str, i, 1)
    if c = "+" then
      c = " "
    elseif c = "%" then
      if i+2 <= l then
        '' use 
        c = ChrW(CLng("&H" & Mid(str,i+1,2)))
        i = i + 2
      end if
    end if
    s = s & c
    i = i + 1
  loop
  URLDecode = decodeUTF8(s, "ASCW")
end function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Class requestClass
  Public ServerVariables, Query, QueryString
  Private qryParam

  Public Default Property Get getVal(key)
    getVal = Query(key)
  End Property

  Private Sub Class_Initialize()
    Set ServerVariables = CreateObject("Scripting.Dictionary")
    ServerVariables.CompareMode = 1
    ServerVariables("PATH_INFO") = window.location.href

    Set qryParam = CreateObject("Scripting.Dictionary")
    qryParam.CompareMode = 1

    Dim s, l, p
    s = ServerVariables("PATH_INFO")
    ''s = "file://D:/a.htm"
    ''s = "http://localhost:81/"
    l = len(s)
    if l > 1 then
      p = InStrRev(s, "?")
      if p>0 then
        if p<>l then ServerVariables("QUERY_STRING") = right(s, l-p)
        l = p-1  :  s = left(s, l)
      end if
    end if

    if l > 1 then
      p = InStr(s, "//")
      if p>0 then
        ServerVariables("SERVER_PROTOCOL") = left(s, p+1)
        if p<>l then
          l = l-(p+1)  :  s=right(s, l)
          p = InStr(s, "/")
          if p=0 then
            if instr(ServerVariables("SERVER_PROTOCOL"), "file")=1 then
             ServerVariables("SCRIPT_NAME") = s
            else
             ServerVariables("SERVER_NAME") = s
            end if
          elseif p=1 then
            ServerVariables("SERVER_NAME") = ""
            ServerVariables("SCRIPT_NAME") = right(s, l-p)
          else
            ServerVariables("SERVER_NAME") = left(s, p-1)
            l = l-(p-1)  :  s=right(s, l)
            ServerVariables("SCRIPT_NAME") = s
          end if
        end if
      end if
    end if

    s = ServerVariables("SERVER_NAME")
    l = len(s)
    p = instr(s, ":")
    if p>1 and instr(ServerVariables("SERVER_PROTOCOL"), "file")<>1 then
      ServerVariables("SERVER_NAME") = left(s, p-1)
      ServerVariables("SERVER_PORT") = right(s, l-p)
    end if

    if ServerVariables.Exists("QUERY_STRING") then
      QueryString = ServerVariables("QUERY_STRING")
      for each s in split(QueryString, "&")
        l = len(s)
        p = InStr(s, "=")
        if p>1 then
          qryParam(URLDecode(left(s,p-1))) = URLDecode(right(s,l-p))
        end if
      next
    end if

    set Query = qryParam
    ServerVariables("QUERY_STRING") = QueryString
  End Sub

  Private Sub Class_Terminate()
    ''
  End Sub 

End Class

Class responseClass
  Public Function Write(x)
    document.write x
  End Function
  Public Function Quit
    document.write "</script></table></form></BODY></HTML>"
    document.execCommand "Stop"
  End Function
End Class

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

function MapPath(path)
  Dim root
  root = Request.ServerVariables("SCRIPT_NAME")
  if root<>"" then root=left(root,InStrRev(root, "/"))
  if path="." then
    if InStrRev(root, "/") = len(root) then
      root = left(root, len(root)-1)
    end if
    path=""
  end if
  MapPath = URLDecode(root & path)
end function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub dumpRequestData
 On Error Resume Next
  Dim i, keys, vals
  response.write "<br><b>Request.Query</b>"
  response.write "<form method=post name=dumpRequestData>"
  response.write "<table>"
  with Request.Query
    keys = .Keys  :  vals = .Items
    for i = 0 to ubound(keys)
      response.write "<tr><td>" & HTMLEncode(keys(i)) & "<td>" & HTMLEncode(vals(i))
      response.write "<input type=hidden name='" & HTMLEncode(keys(i)) & "' value='" & HTMLEncode(vals(i)) & "'>"
    next
    keys = ""  : vals = ""
  end with
  response.write "</table>"

  response.write "<br><b>Request.Form</b><br>"
  response.write "<table>"
  with Request.Form
    keys = .Keys  :  vals = .Items
    for i = 0 to ubound(keys)
      response.write "<tr><td>" & HTMLEncode(keys(i)) & "<td>" & HTMLEncode(vals(i))
      response.write "<input type=hidden name='" & HTMLEncode(keys(i)) & "' value='" & HTMLEncode(vals(i)) & "'>"
    next
    keys = ""  : vals = ""
  end with
  response.write "</table>"

  response.write "<br><b>RequestAll</b><br>"
  response.write "<table>"
  for each item in RequestAll
    response.write "<tr><td>" & HTMLEncode(item) & "<td>" & HTMLEncode(RequestAll(item))
    response.write "<input type=hidden name='" & HTMLEncode(item) & "' value='" & HTMLEncode(RequestAll(item)) & "'>"
  next
  response.write "</table>"
  response.write "</form>"
  On Error Goto 0
End Sub

'---------init-----------'

Set Request=New requestClass
Set Response=New responseClass
Set Session=Nothing

set top_ = window
On Error Resume Next
  while not(top_.parent is top_.self)
    set top_ = top_.parent
    if IsObject(top_.Session) then set Session = top_.Session
  wend
On Error Goto 0
if not(isObj(Session)) and not(top_.self Is window.self) then
  ''alert "Error: session object not defined in the top window!"
end if
if (Session Is Nothing) then
  Set Session = CreateObject("Scripting.Dictionary")
  Session.CompareMode = 1
  ''if (top_.self Is window.self) then
    Randomize  :  Session("SessionID") = CStr(CLng(Rnd * 10^9))
    Session("PATH_INFO") = Request.ServerVariables("PATH_INFO")
  ''end if

  if (top_.self Is window.self) then
    '' !!! not work well !!!
    '' these objects are to be created in top window only !!
    'Set Session("FSO") = CreateObject("Scripting.FileSystemObject")
    'Set Session("ADODB.Stream") = CreateObject("ADODB.Stream")
  end if
end if

Request.ServerVariables("HTTP_REFERER") = Session("LAST_HTTP_REFERER")
Session("LAST_HTTP_REFERER") = Request.ServerVariables("PATH_INFO")

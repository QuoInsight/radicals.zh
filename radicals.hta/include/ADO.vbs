Sub HtmlStopAlert(msg)
   msg = replace(msg, "'", "\'")
   msg = replace(msg, vbCR, "\n")
   msg = replace(msg, vbLF, "")
   Response.Write "<SCRIPT>alert('ERROR !!!\n" & msg & "');history.back();</SCRIPT>"
   Response.Quit
End Sub

Function connectDB
  if ( isObj(session("cn")) ) then
    set connectDB = session("cn")
    exit function
  end if

  Dim cn, connstr
  set cn = CreateObject("ADODB.Connection")

  connstr="Provider=MSDAORA;Data Source=(DESCRIPTION=" & _
          "(ADDRESS=(PROTOCOL=TCP)(Host=asoprdb2.ap.mot.com" & _
          ")(Port=1526))(CONNECT_DATA=(SID=ASOASP02" & _
          ")));User ID=motc;Password=motc123"

  connstr="Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=""" & MapPath("..") & "/radicals.mdb"""

  cn.Open connstr
  set session("cn") = cn
  set connectDB = cn
End Function

Function getConnectionInfo(cn)
 On Error Resume Next
  Dim info
  if ( isObj(cn) ) then
    select case cn.State
      case &H00000000 : info = "Closed"
      case &H00000001 : info = "Connected"
      case &H00000002 : info = "Connecting"
      case &H00000004 : info = "Executing"
      case &H00000008 : info = "Fetching"
    end select

    info = info & "! " & vbCrLf & _
      "[ADO Version] " & cn.Version & vbCrLf & _
      "[Provider] " & cn.Provider & "; " & cn.Properties("Provider Version") & _
       "; " & cn.Properties("Provider Name") & vbCrLf & _ 
      "[DB Version Info] " & cn.Properties("DBMS Name") & "; " & cn.Properties("DBMS Version") & vbCrLf & _
      "[Connection Mode] " & cn.Mode & vbCrLf
  end if
  if info="" then info="Invalid ADODB.Connection!"
  getConnectionInfo = info
End Function

Function runQuery(cn, sqlqry)
  Dim rs
  On Error Resume Next
    Set rs = cn.Execute(sqlqry)
    If err.Number <> 0 Then
      Call HtmlStopAlert(err.Description & ": " & sqlqry)
      Set rs = Nothing
    End If
  On Error Goto 0
  Set runQuery = rs
End Function

Function openRs(cn, sqlqry)
  Dim rs
  On Error Resume Next
    set rs = CreateObject("ADODB.RecordSet")
    rs.MaxRecords = 1000 '' JET does not support this !!! [http://support.microsoft.com/kb/186267]
    rs.CursorLocation = 3 '' adUseClient
    rs.Open sqlqry, cn, 3, 1
    If err.Number <> 0 Then
      Call HtmlStopAlert(err.Description & ": " & sqlqry)
      Set rs = Nothing
    End If
  On Error Goto 0
  Set openRs = rs
End Function

Function newAdoCmd(cn, cmd)
  Dim adocmd
  set adocmd = CreateObject("ADODB.Command")
  with adocmd
    .ActiveConnection = cn
    .CommandText = cmd
  end with
  set newAdoCmd = adocmd
End Function

Function execAdoCmd(adocmd)
  Dim rs, param, param_data
  On Error Resume Next
    Set rs = adocmd.Execute
    If err.Number <> 0 Then
      param_data = ""
      for each param in adocmd.Parameters
        param_data = param_data & param.name & "=" & param.value & ";"
      next
      Call HtmlStopAlert(err.Description & ": " & adocmd.CommandText & "[" & param_data & "]")
    End If
  On Error Goto 0
  Set execAdoCmd = rs
End Function

Sub writeRsTable(rs)
  Dim x, i
  Response.Write "<TABLE class=table1 width=100% border=1 style='border-collapse: collapse'><TR>" & vbLf
  On Error Resume Next
  if not rs.BOF then rs.MoveFirst
  if (rs.BOF and rs.EOF) then
    select case err.Number
      case    0:
        Response.Write "<TD>No Data</TD>" & vbLf
      case 3704:
        Response.Write "<TD>Recordset Closed</TD>" & vbLf
      case else:
        Response.Write "<TD>" & Err.Number & ": " & Err.Description & "</TD>" & vbLf
    end select
    Response.Write "</TR></TABLE>" & vbLf
    exit sub
  end if
  On Error Goto 0

  Response.Write "<TH>&nbsp;</TH>" & vbLf
  For x = 0 To (rs.Fields.Count - 1)
    Response.Write "<TH>" & rs.Fields.Item(x).Name & "</TH>" & vbLf
  Next
  Response.Write "</TR>" & vbLf

  While Not rs.EOF
    Response.Write "<TR>"
    i = i + 1
    Response.Write "<TD>" & i & "</TD>" & vbLf
    For X = 0 To (rs.Fields.Count - 1)
      Select Case rs.Fields.Item(X).Type
       Case 205 ' adLongVarBinary
         Response.Write "<TD align=center>" & rs.Fields.Item(X).ActualSize & " bytes</TD>" & vbLf
       Case Else
         Response.Write "<TD>" & rs.Fields.Item(X).Value & "</TD>" & vbLf
      End Select
    Next
    rs.MoveNext
    Response.Write "</TR>" & vbLf
  WEnd

  Response.Write "</TABLE>" & vbLf
End Sub

  Sub writeRStable2(rs)       '''Recordset Dump in HTML table for debugging
    Dim x, i
    Response.Write("<TABLE border=1 style='margin:3;border-collapse:collapse'>" & vbLf)
    If (rs.BOF and rs.EOF) Then
      Response.Write("<TD>No Data</TD>" & vbLf)
      Response.Write("</TR></TABLE>" & vbLf)
      Exit Sub
    End If
    For x = 0 To (rs.Fields.Count - 1)
      Response.Write("<TR align=left><TH><font face=times size=2>" & rs.Fields.Item(x).Name & "</font></TH>" & vbLf)
      If not rs.BOF then rs.MoveFirst
      While Not rs.EOF
        Response.write("<TD>" & rs.Fields.Item(X).Value & "</TD>" & vbLf)
        rs.MoveNext
      WEnd
    Next
    Response.Write("</TABLE>" & vbLf)
  End Sub

Sub WriteSelectOptions(opt_value,opt_text,opt_selected)
  Dim i, selected
  Dim opt_value_asc, opt_text_asc, opt_selected_asc

  selected = False
  opt_selected_asc=Replace(opt_selected,"'","&#39;")
  For i = LBound(opt_value) To UBound(opt_value)
    opt_value_asc=Replace(opt_value(i),"'","&#39;")
    opt_text_asc=Replace(opt_text(i),"'","&#39;")
    Response.Write "<OPTION VALUE='" & opt_value_asc & "'"
    If ( opt_selected_asc & "." = opt_value_asc & "." ) Then 
      Response.Write " SELECTED"
      selected = True
    End If
    Response.Write ">" & opt_text_asc & vbLf
  Next
  If (Not selected) And (opt_selected&""<>"") Then
    Response.Write "<OPTION VALUE='" & opt_selected_asc & "' SELECTED>" & opt_selected_asc & vbLf
  End If
End Sub

Sub WriteSelectRsOptions(rs, opt_selected)
  If Not rs.BOF Then rs.MoveFirst 
  If rs.EOF Then Exit Sub
  Dim rsFieldsCount, opt_value_asc, opt_text_asc, opt_selected_asc
  rsFieldsCount = rs.Fields.Count
  opt_selected_asc=Replace(opt_selected & "","'","&#39;")
  Do Until rs.EOF
    opt_value_asc=Replace(rs(0),"'","&#39;")
    Response.Write "<OPTION VALUE='" & opt_value_asc & "'"
    If rsFieldsCount > 1 Then 
      opt_text_asc = Replace(rs(1),"'","&#39;")
    Else 
      opt_text_asc = opt_value_asc
    End If
    If ( opt_selected_asc & "." = opt_value_asc & "." ) _
     or ( opt_selected_asc & "." = opt_text_asc & "." ) _
    Then Response.Write " SELECTED"
    Response.Write ">" & opt_text_asc
    rs.MoveNext
  Loop
End Sub

Function quote(str)
  Dim tmpstr
  If str&"." = "." Then
   tmpstr = "NULL"
  Else
   tmpstr = Replace(str, "'", "''")
   tmpstr = "'" & tmpstr & "'"
  End If
  quote = tmpstr
End Function

Function nvl(val1, val2)
  if IsNull(val1) or IsEmpty(val1) then
    nvl = val2
  else
    nvl = val1
  end if
End Function

Function pad(str, n, c)
  Dim i
  i = abs(n) - len(str & "")
  if n=0 then
    pad = ""
  elseif i=0 then
    pad = str
  elseif i < 0 then
    if n > 0 then
      pad = left(str, abs(n))
    else
      pad = right(str, abs(n))
    end if
  elseif i > 0 then
    if n > 0 then
      pad = String(i, c) & str
    else
      pad = str & String(i, c)
    end if
  end if
End Function

function toOraDate(datetime)
  if not IsDate(datetime) then
    exit function
  end if
  Dim dd, mm, yyyy, monArr
  monArr = Array("JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC")
  datetime = CDate(datetime)
  dd = Day(datetime)
  mm = Month(datetime)
  yyyy = Year(datetime)
  toOraDate = dd & "-" & monArr(mm-1) & "-" & right(yyyy,2)
end function

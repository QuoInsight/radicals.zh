
  Set cn = CreateObject("ADODB.Connection")
  set rs = CreateObject("ADODB.Recordset")
  rs.CursorLocation = 3 '' adUseClient
  rs.LockType = 3 '' adLockOptimistic

  cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
          "Data Source=D:\admin tools\txt\zh\radical\radicals.mdb"

  rs.open "select top 100000 * from characters where strokes is null", cn

  while not rs.EOF
    wscript.echo rs("character")
    set rs2 = cn.Execute("select strokes from data_pm where fUTF8='" & rs("character") & "'")
    if not rs2.EOF then
      rs("strokes")= rs2("strokes")
      rs.update
    end if
    rs.MoveNext
  wend


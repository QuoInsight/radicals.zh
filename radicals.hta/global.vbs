  '<SCRIPT LANGUAGE=VBScript SRC="./include/stdlib.vbs"></SCRIPT>
  '<SCRIPT LANGUAGE=VBScript SRC="./include/ASP.vbs"></SCRIPT>
  '<SCRIPT LANGUAGE=VBScript SRC="./include/ADO.vbs"></SCRIPT>

  vbs_src = Array( _
    "./include/stdlib.vbs", _
    "./include/ASP.vbs", _
    "./include/ADO.vbs" _
  )
  set oHTTP = CreateObject("Microsoft.XMLHTTP")
  for each src in vbs_src
    oHTTP.Open "GET", src, false
    oHTTP.Send
    ExecuteGlobal oHTTP.responseText
  next
  Set oHTTP = Nothing

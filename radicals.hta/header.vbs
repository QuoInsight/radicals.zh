﻿
function header(active_menu)
  Dim top_menu_arr
  top_menu_arr=Array("部首", "radicals.htm", _
                     "文字", "char.htm", _
                     "繁↔简","simp.htm", _
                     "SQL","sql.htm", _
                     "SQL2","sql2.htm", _
                     "國語辭典","http://dict.revised.moe.edu.tw/cbdic/search.htm", _
                     "新華字典","http://tool.httpcn.com/zi/", _
                     "林語堂字典","http://humanum.arts.cuhk.edu.hk/Lexis/Lindict/", _
                     "搜詞尋字","http://words.sinica.edu.tw/sou/dictionary.html", _
                     "爱词霸在线词典","http://cd.iciba.com/", _
                     "粵語","http://humanum.arts.cuhk.edu.hk/Lexis/lexi-can/",_
                     "indic","indic.htm", _
                     "OpenFolder", "./")

  Response.Write "<TABLE CLASS=TOPMENU WIDTH='100%' HEIGHT=28 CELLPADDING=0 CELLSPACING=0>"
  Response.Write "<TR><TD NOWRAP>"
  For i=0 To (UBound(top_menu_arr)-1)/2
    If i=0 Then
      Response.Write "&nbsp; "
    Else
      Response.Write "&nbsp; | &nbsp; "
    End If
    If top_menu_arr(i*2+1)="" Then
      Response.Write "" & top_menu_arr(i*2) & "" & CHR(13)
    Else
      If top_menu_arr(i*2)="OpenFolder" or instr(top_menu_arr(i*2+1), "http:")<>0 Then
        Response.Write "<A STYLE='color:#003366;' onclick=""try{window.open('" & top_menu_arr(i*2+1) & "')}catch(e){}"">" & top_menu_arr(i*2) & "</A>" & vbLf
      ElseIf top_menu_arr(i*2)=active_menu Then
        Response.Write "<A STYLE='color:#000000;' onclick=""window.navigate('" & top_menu_arr(i*2+1) & "')""><B>" & top_menu_arr(i*2) & "</B></A>" & vbLf
      Else
        Response.Write "<A STYLE='color:#003366;' onclick=""window.navigate('" & top_menu_arr(i*2+1) & "')"">" & top_menu_arr(i*2) & "</A>" & vbLf
      End If
    End If
  Next
  Response.Write "</TD></TR>"
  Response.Write "</TABLE>"
end function

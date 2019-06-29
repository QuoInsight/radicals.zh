
  function dumpExcelFileData(filePath)
    Dim objExcel, objBook, objSheet

    On Error Resume Next
      set objExcel = CreateObject("Excel.Application")
      ''objExcel.Visible = True
      set objBook = objExcel.Workbooks.Open(filePath,0,1) '' don't update links and open as read-only
  
      if (err.number<>0 or (not IsObject(objBook))) then
        response.write "Failed to open Excel Workbook "
        response.write "ERR#" & err.number & ": " & err.description
      else
        for each objSheet in objBook.WorkSheets
          response.write objSheet.Name & vbCrLf
          dumpExcelSheet objSheet
        next
        objBook.Close 0
      end if
  
      objExcel.Quit
      set objExcel = nothing
    On Error Resume Next
  end function

  function dumpExcelSheet(objSheet)
    Dim maxrow, maxcol, i, j
    maxrow = objSheet.Cells.Find("*",,,,1,2).Row    ''xlByRows=1,xlPrevious=2
    maxcol = objSheet.Cells.Find("*",,,,2,2).Column ''xlByColumns=2,xlPrevious=2
    response.write "<table border=1>"
    for i = 1 to maxrow
      response.write "<tr>"
      for j = 1 to maxcol
        response.write "<td>"
        ''response.write i & "," & j & ": " 
        response.write HTMLEncode(objSheet.cells(i,j).value) & vbCrLf
      next
    next
    response.write "</table>"
  end function

  function getExcelDataDict(objSheet)
    Dim objDict, maxrow, maxcol, i, j
    set objDict = CreateObject("Scripting.Dictionary")
    maxrow = objSheet.Cells.Find("*",,,,1,2).Row    ''xlByRows=1,xlPrevious=2
    maxcol = objSheet.Cells.Find("*",,,,2,2).Column ''xlByColumns=2,xlPrevious=2
    for i = 1 to maxrow
      for j = 1 to maxcol
        objDict(i & ":" & j) = CStr(objSheet.cells(i,j).value)
      next
    next
    set getExcelDataDict = objDict
  end function

  function getExcelDataArray(objSheet)
    Dim arrData(), maxrow, maxcol, i, j
    maxrow = objSheet.Cells.Find("*",,,,1,2).Row    ''xlByRows=1,xlPrevious=2
    maxcol = objSheet.Cells.Find("*",,,,2,2).Column ''xlByColumns=2,xlPrevious=2
    redim arrData(maxrow, maxcol)                   ''elements of 0,j & i,0 will be reserved
    for i = 1 to maxrow
      for j = 1 to maxcol
        arrData(i, j) = CStr(objSheet.cells(i,j).value)
      next
    next
    getExcelDataArray = arrData
  end function

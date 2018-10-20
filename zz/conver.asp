<%
Dim xlApp, workBook1, workBook2,aSheets, fileName, aInfo2,aInfo1,oExcel

  Const xlDelimited = 1
  Const xlTextQualifierDoubleQuote = 1
  Const xlTextFormat = 2
  Const xlGeneralFormat = 1



  Set xlApp = CreateObject("Excel.Application")  

xlApp.WorkBooks.OpenText "C:\inetpub\wwwroot\excel\gz\file\order.csv", , , xlDelimited, xlTextQualifierDoubleQuote, true, false, false, true, false, true, "CRLF", Array(Array (1,2),Array (2,2),Array (3,2),Array (4,1),Array (5,2),Array (6,1),Array (7,1),Array (8,1),Array (9,1),Array (10,1),Array (11,1)), true, false

  Set workBook1 = xlApp.ActiveWorkBook
  xlApp.Visible = true

  workBook1.Save "C:\inetpub\wwwroot\excel\gz\file\020data.xlsx", xlNormal
  workBook1.Close

%>
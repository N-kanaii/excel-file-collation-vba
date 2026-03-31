Attribute VB_Name = "Module1"
Option Explicit

Sub collatedata()
Dim BK_A As Workbook, Sheet_A As Worksheet
Dim rng As Range, sfname As String
Dim lastRまとめ As Long
Dim bookname As String
Dim chk As CheckBox
Dim wsまとめ As Worksheet
Set wsまとめ = ThisWorkbook.Worksheets("まとめ") ' ← 実シート名


    For Each chk In wsまとめ.CheckBoxes
        chk.Delete
    Next
    
wsまとめ.Range("A1") = Format(DateAdd("m", -1, Date), "ggge年mm月分")
 '和暦●年２桁月のひと月前

     bookname = ThisWorkbook.Path
   
sfname = Dir(bookname & "\" & "keyword_*.xlsx")


If sfname = "" Then Exit Sub
wsまとめ.Range("A2", Cells(wsまとめ.Rows.Count, 1).End(xlUp)).Offset(1, 0).EntireRow.Delete
 '集約先のファイルが二行目から始まるため

Do
    Set BK_A = Workbooks.Open(bookname & "\" & sfname)
    Set Sheet_A = BK_A.Worksheets("Sheet_Input")
    Sheet_A.Rows(1).Delete '集約させたいファイル二行分不要なため
    Sheet_A.Rows(1).Delete
    Set rng = Sheet_A.UsedRange
    lastRまとめ = wsまとめ.Cells(wsまとめ.Rows.Count, 1).End(xlUp).Offset(1, 0).Row - 1
    lastRまとめ = lastRまとめ + 1
    rng.Copy Destination:=wsまとめ.Cells(lastRまとめ, 1)
    BK_A.Close SaveChanges:=False
    sfname = Dir()
       
Loop While sfname <> ""

End Sub

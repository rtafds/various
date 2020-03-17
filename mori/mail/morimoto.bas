Sub mailer()
'
' mail Macro
' mail ログ抽出
'
'
    Dim starttime As Double
    Dim endtime As Double
    
    Dim mails As String
    Dim spm As Variant
    Dim i, j, IsInNotDaiwa As Long
    Dim sheetname As String


    Sheet(1).Activate
    starttime = Timer

    ' E列にdaiwaメール以外のメールがあるかを判定した結果を入れる。
    Cells(1, 5).Value = "IsInNotDaiwa"
    n = Cells(Rows.count, "D").End(xlUp).Row
    For i = 2 To n
        mails = Cells(i, 4).Value
        spm = Split(mails, ";")
        
        IsInNotDaiwa = 0
        For j = LBound(spm) To UBound(spm)
            If Not InStr(spm(j), "@daiwa.co.jp") > 0 Then
                IsInNotDaiwa = 1
                Exit For
            End If
        Next
        Cells(i, 5).Value = IsInNotDaiwa
    Next


    ' 別シートに貼り付け
    sheetname = "NotDaiwaMail"
    Worksheets.Add.Name = sheetname


    Worksheets("Sheet1").Range("A1").AutoFilter Field:=5, Criteria1:="=1"
    Worksheets("Sheet1").Range("A1").CurrentRegion.SpecialCells(xlVisible).Copy Worksheets(sheetname).Range("A1")
    Worksheets("Sheet1").AutoFilter


    With Worksheets("Sheet1").Range("A1")
        .AutoFilter Field:=5, Criteria1:="=1"
        .CurrentRegion.SpecialCells(xlVisible).Copy Worksheets(sheetname).Range("A1")
        .AutoFilter
    End With


    endtime = Timer
    processtime = enttime - starttime

    Application.CutCopyMode = False

    MsgBox "処理時間" & processtime & "(秒)"

End Sub

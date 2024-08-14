Option Explicit


Sub RowSplitSample(numRows As Long, _
    Optional startRow As Long = 2)
' 行　分割保存　サンプル

    Application.ScreenUpdating = False

    ' タイムスタンプ
    Dim ts As String
    ts = Format(Now(), "yyyymmdd-hhmmss")

    ' 現在のワークシートを設定
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets(1)
    Debug.Print "分割元シート名: " & ws.Name

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Debug.Print "lastRow: " & lastRow

    ' ヘッダー　開始行
    Dim headerBeginRow As Long
    headerBeginRow = 1

    ' ヘッダー　終了行
    Dim headerEndRow As Long
    headerEndRow = startRow - 1

    Dim partCounter As Integer
    partCounter = 1

    ' ファイル保存先
    Dim saveFileHeader As String
    saveFileHeader = ActiveWorkbook.Path & _
        "\Part-" & ts & "-"

    Do While startRow <= lastRow
        ' 新しいファイル　保存先
        Dim saveFileName As String
        saveFileName = saveFileHeader & partCounter & ".xls"
        Debug.Print "◆保存先: " & saveFileName

        ' 新しいブックを作成
        Dim newWorkbook As Workbook
        Set newWorkbook = Workbooks.Add

        Dim destSheet As Worksheet
        Set destSheet = newWorkbook.Sheets(1)

        ' ヘッダー行をコピペ
        ws.Rows(headerBeginRow & ":" & headerEndRow).Copy _
            destSheet.Rows(headerBeginRow)

        ' endRowを設定
        Dim endRow As Long
        endRow = WorksheetFunction.min(startRow + numRows - 1, lastRow)

        Debug.Print "startRow: " & startRow
        Debug.Print "endRow: " & endRow

        ' データをコピペ
        ws.Rows(startRow & ":" & endRow).Copy destSheet.Rows(headerEndRow + 1)

        ' 保存
        Application.DisplayAlerts = False
        newWorkbook.SaveAs saveFileName
        Application.DisplayAlerts = True
        '
        newWorkbook.Close False

        ' 次の開始行を設定
        startRow = endRow + 1
        partCounter = partCounter + 1
    Loop

    Application.ScreenUpdating = True

    Debug.Print "分割保存が完了しました。"
End Sub


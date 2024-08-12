Option Explicit


Sub ColumnSortSample( _
    orderAry() As Variant, _
    Optional headerRowIdx As Long = 1)
' 列　並び替え　サンプル

    ' 挿入位置　列番号
    Dim destCol As Long
    destCol = 1

    ' ヘッダー範囲
    Dim headerRng As Range
    Set headerRng = Rows(headerRowIdx)

    ' 移動する列の項目名
    Dim col As Variant
    For Each col In orderAry
        '
        On Error GoTo SkipToHere

        ' 移動する列の列番号
        Dim foundColIdx As Long
        foundColIdx = Application.Match(col, headerRng, 0)
        '
        If foundColIdx <> destCol Then
            ' 切り取り
            Cells(1, foundColIdx).EntireColumn.Cut
            ' 挿入
            Cells(1, destCol).EntireColumn.Insert Shift:=xlToRight
        End If
        GoTo IncreDc
        '
SkipToHere:
        Debug.Print "Not Found: " & col
IncreDc:
        destCol = destCol + 1
    Next col

    Debug.Print "Sort End"
End Sub


Sub EmptyColumnDeleteSample(Optional headerRow as Long = 1)
' 空列削除　サンプル

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim tableRange As Range
    Set tableRange = ws.Range("A1:E11")  ' 削除対象の範囲

    dim deleteCount as Long
    deleteCount = 0

    ' 列を右から左にループして削除を行う
    Dim i As Long
    For i = tableRange.Columns.Count To 1 Step -1
        Dim isEmpty As Boolean
        isEmpty = True

        Dim col As Range
        Set col = tableRange.Columns(i)

        ' 項目名以外の行をチェック
        Dim cell As Range
        For Each cell In col.Cells
            ' 項目名よりも下の行　かつ　空白でない　なら
            If cell.row > headerRow And cell.Value <> "" Then
                ' この列は空列ではない
                isEmpty = False
                Exit For
            End If
        Next cell

        ' 項目名以外の行がすべて空白なら列を削除
        If isEmpty Then
            Debug.Print "Delete column: " _
                & col.Cells(1, 1).Value & " at " & col.Column
            ' 列　削除
            col.Delete
            ' 削除数　カウント＋１
            deleteCount = deleteCount + 1
        End If
    Next i

    Debug.Print "Deleted " & deleteCount & " Empty Columns"
End Sub


Sub ColumnDeleteSample( _
    columnNames As Variant, _
    Optional headerRange As Range = Nothing)
' 列削除　サンプル
    Dim ws As Worksheet
    Set ws = ActiveSheet

    If headerRange Is Nothing Then
        ' ヘッダーを1行目から取得
        Set headerRange = Range("A1", Range("A1").End(xlToRight).Address)
    End If

    ' ヘッダー開始セル　取得
    Dim headerBegin As Range
    Set headerBegin = headerRange.Cells(1, 1)
    Debug.Print "header begin at: " & headerBegin.Address

    ' 列名でループする
    Dim colName As Variant
    Dim delCount As Long  ' 削除した列数をカウント
    delCount = 0
    '
    For Each colName In columnNames
        ' ヘッダー行を検索して列を削除
        Dim cell As Range
        For Each cell In headerRange.Cells

            If cell.Value = colName Then

                Debug.Print "Delete column by name: " & cell.Value
                ' 列　削除
                ws.Columns(cell.Column).Delete
                ' 削除　カウント＋１
                delCount = delCount + 1

                ' ヘッダー　再取得
                Set headerRange = Range(headerBegin, _
                    headerBegin.End(xlToRight).Address)

                Exit For ' 列を削除したらループを抜ける
            End If
        Next cell
    Next colName

    Debug.Print "Deleted " & delCount & " columns"
End Sub


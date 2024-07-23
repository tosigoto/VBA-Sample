Option Explicit

Sub CsvImportSample(filePath As String, _
    ws As Worksheet, _
    Optional addr As String = "A1", _
    Optional isCsv As Boolean = True, _
    Optional mjCode As Long = 932 _
    )
' CSV　インポート　サンプル

    'CSVを開く
    With ws.QueryTables.Add( _
        Connection:="TEXT;" & filePath, _
        Destination:=ws.Range(addr))

        'CSVの列を文字で区切る
        .TextFileParseType = xlDelimited
        
        '区切り文字＝カンマか？
        If isCsv Then
            .TextFileCommaDelimiter = True
        Else  '区切り文字＝tabとして扱う
            .TextFileTabDelimiter = True
        End If

        'CSVの文字コード
        .TextFilePlatform = mjCode  'Shift-Jis:932 / UTF-8:65001

        ' 以下は状況に応じて使う
        '.TextFileConsecutiveDelimiter = False
        '.TextFileSemicolonDelimiter = False
        '.TextFileSpaceDelimiter = False
        '.TextFileOtherDelimiter = False
        '.TextFileColumnDataTypes = Array(1)
        '.TextFileTrailingMinusNumbers = True

        '表示
        .Refresh BackgroundQuery:=False
    End With
End Sub


Sub CsvExportSample_ADODB_Bulk(filePath As String, _
    ws As Worksheet, _
    Optional addr As String = "A1", _
    Optional isDq As Boolean = True, _
    Optional dlm As String = ",", _
    Optional mojiCode As String = "Shift-Jis", _
    Optional kaigyo As String = vbCrLf _
    )
'CSV　エクスポート　サンプル by ADODB　全行を一括で書き込み

    ' buf
    Dim bufAry() As String
    ReDim bufAry(0 To ws.Range(addr).CurrentRegion.Rows.Count - 1)

    ' buf index init
    Dim bufIdx As Integer
    bufIdx = 0
    ' CSV内容を生成
    Dim row As Range
    For Each row In ws.Range(addr).CurrentRegion.Rows
        Dim line As String
        line = ""

        Dim cell As Range
        Dim vv As String
        For Each cell In row.Cells
            If isDq Then
                ' "で囲む ＆ "をエスケープする
                vv = """" & Replace(cell.Value, """", """""") & """"
            Else
                ' 囲まない
                vv = cell.Value
            End If

            line = line & vv & dlm
        Next cell

        ' 行の最後の区切り文字列を削除して改行を追加
        line = Left(line, Len(line) - Len(dlm)) & kaigyo

        ' buf hold
        bufAry(bufIdx) = line
        bufIdx = bufIdx + 1
    Next row

    With CreateObject("ADODB.Stream")
        .Charset = mojiCode
        .Open
        .Position = 0  ' 書き込み位置：ファイルの先頭
        .WriteText Join(bufAry, ""), 0
        .SaveToFile filePath, 2  ' 上書き
        .Close
    End With
End Sub


Sub CsvExportSample_ADODB_BatchLine(filePath As String, _
    ws As Worksheet, _
    Optional addr As String = "A1", _
    Optional isDq As Boolean = True, _
    Optional dlm As String = ",", _
    Optional mojiCode As String = "Shift-Jis", _
    Optional kaigyo As String = vbCrLf, _
    Optional batch As Long = 500 _
    )
'CSV　エクスポート　サンプル by ADODB　batch行ずつ書き込み

    ' 書き込みオブジェクト　生成
    Dim adobj As Object
    Set adobj = CreateObject("ADODB.Stream")
    ' 書き込みファイルを初期化
    With adobj
        .Charset = mojiCode
        .Open
        .Position = 0  ' 書き込み位置：ファイルの最後
        .WriteText "", 0
        .SaveToFile filePath, 2  ' 上書き
        .Close
    End With

    ' CSV内容を生成
    Dim batchCount As Long
    batchCount = 0
    '
    Dim batchBuf() As String
    ReDim batchBuf(0 To batch - 1)
    '
    Dim row As Range
    For Each row In ws.Range(addr).CurrentRegion.Rows
        Dim line As String
        line = ""

        Dim cell As Range
        Dim vv As String
        For Each cell In row.Cells
            If isDq Then
                ' "で囲む ＆ "をエスケープする
                vv = """" & Replace(cell.Value, """", """""") & """"
            Else
                ' 囲まない
                vv = cell.Value
            End If

            line = line & vv & dlm
        Next cell

        ' 行の最後の区切り文字列を削除して改行を追加
        line = Left(line, Len(line) - Len(dlm)) & kaigyo
        
        batchBuf(batchCount) = line
        
        batchCount = batchCount + 1
        
        If batchCount = batch Then
            ' 書き込みファイルに追記
            With adobj
                .Charset = mojiCode
                .Open
                .LoadFromFile filePath  ' ファイル指定
                .Position = .Size  ' 書き込み位置：ファイルの最後
                .WriteText Join(batchBuf, ""), 0
                .SaveToFile filePath, 2  ' 上書き
                .Close
            End With
            
            batchCount = 0
            ReDim batchBuf(0 To batch - 1)
        End If

    Next row

    If batchCount > 0 Then
        ' 書き込みファイルに追記
        With adobj
            .Charset = mojiCode
            .Open
            .LoadFromFile filePath  ' ファイル指定
            .Position = .Size  ' 書き込み位置：ファイルの最後
            .WriteText Join(batchBuf, ""), 0
            .SaveToFile filePath, 2  ' 上書き
            .Close
        End With
        
        batchCount = 0
        Erase batchBuf
    End If

    Set adobj = Nothing
    
    Debug.Print "Done"
End Sub

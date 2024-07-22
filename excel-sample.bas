Option Explicit

Sub CollectionSample()
'コレクション　サンプル
    Dim buf As New Collection

    'シート名　取得
    For Each Sh In Worksheets
        buf.Add Sh.name
    Next Sh

    'シート名　表示
    For Each vv In buf
        Debug.Print vv
    Next vv

    'コレクション　長さ　表示
    Debug.Print buf.Count
End Sub


Sub ArraySample()
'配列　サンプル　ForEach
    Dim ary() As String
    ReDim ary(0 To Worksheets.Count - 1)
    Dim i As Integer

    'シート名　取得
    i = 0
    For Each Sh In Worksheets
        ary(i) = Sh.name
        i = i + 1
    Next Sh

    'シート名　表示
    Debug.Print Join(ary, vbNewLine)
    
    '配列　長さ　表示
    Debug.Print UBound(ary) - LBound(ary) + 1
End Sub


Sub ArraySample2()
'配列　サンプル　For
    Dim ary() As String
    ReDim ary(0 To Worksheets.Count - 1)
    Dim i As Integer

    'シート名　取得
    For i = 0 To Worksheets.Count - 1
        ary(i) = Worksheets(i + 1).name
    Next

    'シート名　表示
    Debug.Print Join(ary, vbNewLine)

    '配列　長さ　表示
    Debug.Print UBound(ary) - LBound(ary) + 1
End Sub


Sub CellAllSample()
'すべてのセルを扱う
    Cells.Select

    With Selection
        .Font.Bold = True
    End With
End Sub


Sub RangeSample()
'Range　指定　サンプル
    'A2
    Range("A2").Value = 1

    'B2
    [B2].Value = "abcd"

    'C2
    Cells(2, 3).Value = Date
End Sub


Sub FormatSample()
'フォーマット　サンプル
    '56.79%
    Debug.Print Format(0.56789, Format:="Percent")
    '00.57
    Debug.Print Format(0.56789, Format:="00.00")
    '日付　時刻
    Debug.Print Format(Date, Format:="yyyy/mm/dd")
    Debug.Print Format(Now, Format:="yyyy/mm/dd hh:mm:ss")
    '12,134
    Debug.Print Format(12345, Format:="#,###")
End Sub


Sub WorksheetFunctionSample()
'関数使用　サンプル
    Dim mySum As Double
    Dim myAvg As Double
    Dim myMax As Double

    '範囲　取得
    Dim myRange As Range
    Set myRange = Range("A1:A5")

    '範囲　計算
    mySum = WorksheetFunction.Sum(myRange)
    myAvg = WorksheetFunction.Average(myRange)
    myMax = WorksheetFunction.Max(myRange)

    Debug.Print _
        "sum: " & mySum & vbNewLine _
        & "avg: " & myAvg & vbNewLine _
        & "max: " & myMax
End Sub


Sub DateSample()
'日数　計算　サンプル
    Dim td As Date
    Dim ud As Date
    
    td = Date
    ud = DateAdd("yyyy", 1, td)

    '365
    Debug.Print ud - td
End Sub




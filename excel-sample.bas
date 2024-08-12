Option Explicit

Sub CollectionSample()
' コレクション　サンプル
    Dim buf As New Collection

    ' シート名　取得
    For Each Sh In Worksheets
        buf.Add Sh.name
    Next Sh

    ' シート名　表示
    For Each vv In buf
        Debug.Print vv
    Next vv

    ' コレクション　長さ　表示
    Debug.Print buf.Count
End Sub


Sub ArraySample()
' 配列　サンプル　ForEach
    Dim ary() As String
    ReDim ary(0 To Worksheets.Count - 1)

    ' シート名　取得
    Dim i As Integer
    i = 0
    For Each Sh In Worksheets
        ary(i) = Sh.name
        i = i + 1
    Next Sh

    ' シート名　表示
    Debug.Print Join(ary, vbNewLine)
    
    ' 配列　長さ　表示
    Debug.Print UBound(ary) - LBound(ary) + 1
End Sub


Sub ArraySample2()
' 配列　サンプル　For
    Dim ary() As String
    ReDim ary(0 To Worksheets.Count - 1)

    ' シート名　取得
    Dim i As Integer
    For i = 0 To Worksheets.Count - 1
        ary(i) = Worksheets(i + 1).name
    Next i

    ' シート名　表示
    Debug.Print Join(ary, vbNewLine)

    ' 配列　長さ　表示
    Debug.Print UBound(ary) - LBound(ary) + 1
End Sub


Sub CellAllSample()
' すべてのセルを扱う
    Cells.Select

    With Selection
        .Font.Bold = True
    End With
End Sub


Sub RangeSample()
' Range　指定　サンプル
    ' A2
    Range("A2").Value = 1

    ' B2
    [B2].Value = "abcd"

    ' C2
    Cells(2, 3).Value = Date
End Sub


Sub FormatSample()
' フォーマット　サンプル
    ' 56.79%
    Debug.Print Format(0.56789, Format:="Percent")
    ' 00.57
    Debug.Print Format(0.56789, Format:="00.00")
    ' 日付　時刻
    Debug.Print Format(Date, Format:="yyyy/mm/dd")
    Debug.Print Format(Now, Format:="yyyy/mm/dd hh:mm:ss")
    ' 12,134
    Debug.Print Format(12345, Format:="#,###")
End Sub


Sub WorksheetFunctionSample()
' 関数使用　サンプル
    Dim mySum As Double
    Dim myAvg As Double
    Dim myMax As Double

    ' 範囲　取得
    Dim myRange As Range
    Set myRange = Range("A1:A5")

    ' 範囲　計算
    mySum = WorksheetFunction.Sum(myRange)
    myAvg = WorksheetFunction.Average(myRange)
    myMax = WorksheetFunction.Max(myRange)

    Debug.Print _
        "sum: " & mySum & vbNewLine _
        & "avg: " & myAvg & vbNewLine _
        & "max: " & myMax
End Sub


Sub DateSample()
' 日数　計算　サンプル
    Dim td As Date
    Dim ud As Date
    
    td = Date
    ud = DateAdd("yyyy", 1, td)

    ' 365
    Debug.Print ud - td
End Sub


Sub SelectCaseSample()
' Select-Case　サンプル
    Dim ce As Range
    For Each ce In Range(Cells(2, 1), Cells(2, 1).End(xlDown).Address)

        Dim tgt As Range
        Set tgt = ce.Offset(0, 1)
        ' 文字　中央寄せ
        tgt.HorizontalAlignment = xlCenter

        Select Case ce.Value
            Case Is < 40
                tgt.Value = "u-40"
            Case Is < 60
                tgt.Value = "u-60"
            Case Is < 80
                tgt.Value = "u-80"
            Case Else
                tgt.Value = "me-80"
        End Select
    Next ce
End Sub


Sub DoWhileSample()
' Do-While　サンプル
    Range("A1").Activate

    Do While ActiveCell.Value <> ""
        Debug.Print ActiveCell.Value
        ActiveCell.Offset(1, 0).Activate
    Loop
End Sub


Sub DictionarySample()
' ディクショナリー　サンプル
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")

    dic.Add "山", "200"
    dic.Add "川", "400"
    dic.Add "海", "1000"

    Dim kk As Variant
    For Each kk In dic
        Debug.Print kk & ": " & dic(kk)
    Next kk
End Sub


Sub ErrorGotoSample()
' Error　サンプル
    Dim num As Variant

    On Error GoTo myErr

    num = InputBox("分母を入力して下さい。")
    If num = "" Then
        GoTo myEnd
    End If

    ActiveCell.Value = ActiveCell.Offset(0, -1).Value / num
    GoTo myEnd

myErr:
    MsgBox "入力値でエラー発生。", vbCritical

myEnd:
End Sub


Sub ClearSample()
' ClearContents: 値のみ消去 / Clear: 書式なども消去
    Rows(1).ClearContents

    Rows(2).Clear
End Sub


Sub AppInputSample()
' Application.InputBox　サンプル
    Range("a2").Value = _
        Application.InputBox("なんでも入力")

    ' Typeで型指定
    Range("b2").Value = _
        Application.InputBox("数値を入力", Type:=1)
End Sub


Sub ResumeSample()
' Resume　サンプル
    Dim num As Variant

myResume:
    On Error GoTo myErr

    num = InputBox("enter a value to divide by")

    If num = "" Then
        Exit Sub
    End If

    ActiveCell.Value = ActiveCell.Offset(0, -1).Value / num
    Exit Sub

myErr:
    Dim msg As String
    msg = "エラー発生。" & vbNewLine & vbNewLine & "やり直しますか？"

    Dim answer As Integer
    answer = MsgBox(msg, vbYesNo + vbCritical)

    If answer = vbYes Then
        Resume myResume
    Else
        Exit Sub
    End If
End Sub


Sub InsertNewSheetSample()
' 新規シート追加　サンプル
    Worksheets.Add
    Range("A1").Value = Now
End Sub


Function FunctionSample(num As Integer) As Boolean
' Function　サンプル
    If num Mod 2 = 0 Then
        FunctionSample = True
    Else
        FunctionSample = False
    End If
End Function


Sub FuncDescriptionSample()
' ユーザー定義関数にDescriptionを設定　サンプル
    Application.MacroOptions Macro:="FunctionSample", _
        Description:="Return True if num is even number"

    ' disabled in XP
    'Application.MacroOptions Macro:="FunctionSample", _
    '     ArgumentDescriptions:=Array( _
    '     "number")
End Sub


Sub FilterOnSample()
' オートフィルター　ON　サンプル
    With Range("A1")
        ' 1列目の条件
        .AutoFilter Field:=1, _
            Criteria1:="<>b", Operator:=xlAnd, _
            Criteria2:="<>c"
        ' 2列目の条件
        .AutoFilter Field:=2, _
            Criteria1:="<>1"
    End With
End Sub


Sub FilterOffSample()
' オートフィルター　OFF　サンプル
    Range("A1").AutoFilter
End Sub


Sub VlookupSample()
' Vlookup　サンプル

    ' 参照範囲
    Dim refRng As Range
    Set refRng = Workbooks("reference.xls").Worksheets("Sheet1") _
        .Range("A1").CurrentRegion

    Debug.Print "refRng.Address: " & refRng.Address

    ' vlookup検索開始セル
    Dim baseCell As Range
    Set baseCell = Range("A2")

    Dim cl As Range
    For Each cl In Range(baseCell, baseCell.End(xlDown))
        ' vlookup結果を保持
        Dim vlValue As Variant

        ' エラー発生時に飛ぶ
        On Error GoTo ErrHandle

        'get vlookup
        vlValue = WorksheetFunction.VLookup( _
            cl.Value, refRng, 2, 0)

        ' セルに値を設定
        GoTo GoSetValue

ErrHandle:  ' エラー発生時
        vlValue = "※なし"

        ' セルに値を設定
        GoTo GoSetValue

GoSetValue:  ' セルにvlookup結果を設定
        cl.Offset(0, 2).Value = vlValue
    Next cl

End Sub


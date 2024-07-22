Option Explicit

Sub PivotMake()
'ピボットテーブルを作成する
'
    'ピボットテーブル名
    Const pvtName As String = "myPvt"
    'data sourceシート名
    Const dsName As String = "data"

    'data range address
    '行(R)：1〜max
    '列(C)：1〜max
    Dim drAddr As String
    drAddr = dsName & "!R1C1:R" _
        & Worksheets(dsName).Range("a1").End(xlDown).Row _
        & "C" _
        & Worksheets(dsName).Range("a1").End(xlToRight).Column

    'ピボットテーブル　作成
    ActiveWorkbook.PivotCaches.Add( _
        SourceType:=xlDatabase, _
        SourceData:=drAddr) _
        .CreatePivotTable _
            TableDestination:="", _
            TableName:=pvtName, _
            DefaultVersion:=xlPivotTableVersion10

    'ピボットテーブル　設置＠左上
    ActiveSheet.PivotTableWizard TableDestination:=ActiveSheet.Cells(1, 1)

    '行　項目　設定
    With ActiveSheet.PivotTables(pvtName).PivotFields("shop")
        .Orientation = xlRowField
        .Position = 1
    End With

    '列　項目　設定
    With ActiveSheet.PivotTables(pvtName).PivotFields("item")
        .Orientation = xlColumnField
        .Position = 1
    End With

    'データ　項目　設定
    ActiveSheet.PivotTables(pvtName).AddDataField _
        ActiveSheet.PivotTables(pvtName).PivotFields("count"), _
        "合計 / count", _
        xlSum

    'ツールダイアログ類　非表示
    ActiveWorkbook.ShowPivotTableFieldList = False
    Application.CommandBars("PivotTable").Visible = False

End Sub

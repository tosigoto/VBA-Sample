Option Explicit

Sub CreateWordDoc()

    'Dim wdApp As Word.Application
    Dim wdApp As Object
    Set wdApp = CreateObject("Word.Application")

    With wdApp
        .Visible = True
        .Activate

        '新規ドキュメント　開く
        .Documents.Add

        'テーブル　コピー in excel
        Worksheets("data") _
            .Range("a1", Range("a2").End(xlDown).End(xlToRight)).Copy

        'テーブル　貼り付け in word
        .Selection.Paste

        'セレクション　解除
        Application.CutCopyMode = False

        'save doc file
        Dim SaveAsName As String
        SaveAsName = Environ("UserProfile") _
            & "\○○○\DataReport" _
            & Format(Now, "yyyy-mm-dd hh-mm-ss") & ".doc"
        '
        .Activedocument.SaveAs Filename:=SaveAsName

        'close doc
        .Activedocument.Close
        .Quit
    End With

End Sub


Sub CreateWordDocFromTemplate()
'テンプレートを使用する場合
    Const bookMarkName As String = "PutTableHere"

    'Dim wdApp As Word.Application
    Dim wdApp As Object
    Set wdApp = CreateObject("Word.Application")

    With wdApp
        .Visible = True
        .Activate

        'テンプレートを開く
        .Documents.Add "D:\○○○\DataReportTemplate.dot"

        'テーブルをコピー
        Worksheets("data") _
            .Range("a1", Range("a2").End(xlDown).End(xlToRight)).Copy

        'ブックマークの存在を確認
        Debug.Print "bm exists=" & wdApp.Activedocument.Bookmarks.Exists(bookMarkName)
        'ブックマークにジャンプ
        'OK for XP
        .Selection.Goto Name:=bookMarkName
        '
        '-1 OK for XP
        '.Selection.Goto What:=-1, Name:=bookMarkName
        '
        'wdGoToBookmark OK for NOT-XP
        '.Selection.Goto What:=wdGoToBookmark, Name:=bookMarkName

        'テーブルを貼り付け
        .Selection.Paste

        'セレクション　解除
        Application.CutCopyMode = False

        'save doc file
        Dim SaveAsName As String
        SaveAsName = Environ("UserProfile") _
            & "\○○○\DataReport" _
            & Format(Now, "yyyy-mm-dd hh-mm-ss") & ".doc"
        '
        .Activedocument.SaveAs Filename:=SaveAsName

        'close doc
        .Activedocument.Close
        .Quit
    End With

End Sub

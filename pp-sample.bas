Option Explicit

Sub CreatePPointSlides()

    'Dim ppApp As PowerPoint.Application
    Dim ppApp As Object
    Set ppApp = CreateObject("PowerPoint.Application")

    With ppApp
        .Visible = True
        .Activate

        Dim ppPrsnt As Object
        Dim ppSlide As Object

        '新規プレゼンテーション　開く
        Set ppPrsnt = .Presentations.Add

        'slide 1
        'XP disabled ppLayoutTitle
        Set ppSlide = ppPrsnt.Slides.Add(1, 1)
        ppSlide.Select

        'show slide variables type
        Call ShowSlideType(ppSlide)

        'set text
        ppSlide.Shapes(1).TextFrame.TextRange = _
            "Data Information Slide"
        'set text
        ppSlide.Shapes(2).TextFrame.TextRange = _
            "by ■■"

        'slide 2
        'XP disabled ppLayoutBlank
        Set ppSlide = ppPrsnt.Slides.Add(2, 12)
        ppSlide.Select

        'テーブル　コピー in excel
        Worksheets("data").Range("a1").CurrentRegion.Copy
        'テーブル　貼り付け in pp
        ppSlide.Shapes.Paste

        'セレクション　解除
        Application.CutCopyMode = False

        'pp file name to save
        Dim SaveAsName As String
        SaveAsName = Environ("UserProfile") _
            & "\○○○\DataPP" _
            & Format(Now, "yyyy-mm-dd hh-mm-ss") & ".ppt"

        Debug.Print "pp save to: " & SaveAsName

        'save
        ppPrsnt.SaveAs Filename:=SaveAsName

        'close
        ppPrsnt.Close
        .Quit

    End With
End Sub


Sub ShowSlideType(ByRef ppSlide As Object)
    Debug.Print "ppSlide: " & vbNewLine & vbTab & TypeName(ppSlide)
    Debug.Print "ppSlide.Shapes(1): " & vbNewLine & vbTab & TypeName(ppSlide.Shapes(1))
    '
    Debug.Print "ppSlide.Shapes(1).TextFrame: " & vbNewLine & vbTab & TypeName(ppSlide.Shapes(1).TextFrame)
    Debug.Print "ppSlide.Shapes(1).TextFrame.TextRange: " & vbNewLine & vbTab & TypeName(ppSlide.Shapes(1).TextFrame.textRange)
End Sub

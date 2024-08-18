Option Explicit


Function GetDictionaryFromTableSample( _
    Optional ws As Worksheet = Nothing, _
    Optional headerRow As Long = 1, _
    Optional headerStartColumn As Long = 1, _
    Optional dataStartRow As Long = 2) As Object
' 表からDictionary作成　サンプル

    If ws Is Nothing Then
       Set ws = ActiveSheet
    End If

    Debug.Print "Dictionary作成元：表.Name: " & ws.Name

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, headerStartColumn).End(xlUp).Row

    Dim lastColumn As Long
    lastColumn = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    Debug.Print "lastRow, lastColumn = " & lastRow & ", " & lastColumn

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim y As Long
    For y = dataStartRow To lastRow
        Dim key As Variant
        key = ws.Cells(y, headerStartColumn).value

        Dim value As Object
        Set value = CreateObject("Scripting.Dictionary")

        Dim x As Long
        For x = headerStartColumn + 1 To lastColumn
            value(ws.Cells(headerRow, x)) = ws.Cells(y, x).value
        Next x

        dict.Add key, value
    Next y

    Set GetDictionaryFromTableSample = dict
End Function


Sub TEST_GetDictionaryFromTableSample()
' テストコード：表からDictionary作成　サンプル

    Dim dict As Object
    ' ActiveSheet
    Set dict = GetDictionaryFromTableSample( _
        ws:=ActiveSheet, _
        headerRow:=1, _
        headerStartColumn:=2, _
        dataStartRow:=4 _
         )

    ' 連想配列の内容を出力
    Dim subDict As Object
    Set subDict = CreateObject("Scripting.Dictionary")
    '
    Dim id As Variant
    For Each id In dict.keys
        Set subDict = dict(id)
        Dim kx As Variant

        Debug.Print "id: " & id
        For Each kx In subDict.keys
            Debug.Print vbTab & kx & ": " & subDict(kx)
        Next kx
    Next id
End Sub


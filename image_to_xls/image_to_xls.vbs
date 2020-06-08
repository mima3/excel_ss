Option Explicit
const MARGIN_NEXT_ROW = 5

' 画像オブジェクトから次に張り付ける位置を取得する
Private Function GetNextRow(ByVal sht, ByVal row, ByVal p)
    Dim pos
    pos = p.Top + p.Height + MARGIN_NEXT_ROW
    While sht.Cells(row, 1).Top < pos
        row = row + 1
    Wend
    GetNextRow = row
End Function

' シートに画像の貼り付けを行う
' もし、横幅がlimitWidthの範囲を超える場合は、その大きさを上限にサイズを変更する
' この関数は次に貼り付ける行を返す
Private Function PasteImage(ByVal imgpath, ByVal sht, Byval targetCol, ByVal targetRow, ByVal limitWidth)
    sht.Cells(targetRow, targetCol).Activate
    Dim pictObj
    Set pictObj = sht.Pictures.Insert(imgpath)
    ' 上限の列を超えた画像を張り付けた場合は、画像の大きさを調整する
    If limitWidth < (pictObj.Left + pictObj.Width) Then
        pictObj.Width = limitWidth - pictObj.Left
    End If

    PasteImage = GetNextRow(sht, targetRow, pictObj)
End Function


Private Sub Main()
    Dim args
    Set args = WScript.Arguments
    If args.Count <> 6 Then
        WScript.Echo "CScript image_to_xls.vbs テンプレートのEXCELへのパス シート名 貼り付け開始セルの位置 貼り付け終了の列 入力ファイル  出力パス"
        WScript.Echo "例："
        WScript.Echo "CScript image_to_xls.vbs C:\dev\excel_ss\image_to_xls\test\001\template.xlsx Sheet1 B2 L test\001\input.txt C:\dev\excel_ss\image_to_xls\test\001\out.xlsx"
        Call WScript.Quit()
    End If

    Dim tmplatePath
    Dim sheetName
    Dim offsetCell
    Dim limitCol
    Dim inputPath
    Dim outPath

    tmplatePath = args(0)
    sheetName = args(1)
    offsetCell = args(2)
    limitCol = args(3)
    inputPath = args(4)
    outPath = args(5)

    Dim app
    Dim book
    Dim sht
    Dim targetRow
    Dim targetCol
    Dim limitWidth

    Set app = createobject("Excel.Application")
    app.visible=True
    Set book = app.Workbooks.Open(tmplatePath)
    Set sht = book.Sheets(sheetName)
    sht.Select
    sht.Range(offsetCell).Activate
    targetRow = app.ActiveCell.Row
    targetCol = app.ActiveCell.Column
    limitWidth = sht.Cells.Range(limitCol&"1").Left + sht.Cells.Range(limitCol&"1").Width


    Dim fso
    Dim objInputFile
    Set fso = createObject("Scripting.FileSystemObject")
    ' iomode: ReadOnly
    ' create: False
    ' format: TristateTrue
    Set objInputFile = fso.OpenTextFile(inputPath,1, False, -1)
    Do While objInputFile.AtEndOfStream <> True
        targetRow = PasteImage(objInputFile.ReadLine(), sht, targetCol, targetRow, limitWidth)
    Loop
    objInputFile.Close
    call book.SaveAs(outPath)


    book.Close
    app.Quit
    Set app = Nothing
End Sub


Call Main()
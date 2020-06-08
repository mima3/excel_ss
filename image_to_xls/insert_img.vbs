Option Explicit
const MARGIN_NEXT_ROW = 5

' �摜�I�u�W�F�N�g���玟�ɒ���t����ʒu���擾����
Private Function GetNextRow(ByVal sht, ByVal row, ByVal p)
    Dim pos
    pos = p.Top + p.Height + MARGIN_NEXT_ROW
    While sht.Cells(row, 1).Top < pos
        row = row + 1
    Wend
    GetNextRow = row
End Function


Private Sub Main()
    Dim args
    Set args = WScript.Arguments
    If args.Count <> 2 Then
        WScript.Echo "�t�@�C���p�X�ƑΏۃu�b�N���i������v�j���w�肳��Ă��܂���B"
        Call WScript.Quit()
    End If

    Dim imagePath
    Dim targetName

    imagePath = args(0)
    targetName = args(1)

    Dim app
    Dim bookTmp
    Dim book
    Set book = Nothing
    Dim targetRow
    Dim targetCol

On Error Resume Next
    Set app = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
        WScript.Echo "�\��t�����Excel���N�����Ă��������B���u" & targetName & "�v���܂ރu�b�N���J���Ă��������B"
        Call WScript.Quit()
    End If
On Error GoTo 0
    For Each bookTmp In app.WorkBooks
        If InStr(bookTmp.Name, targetName) > 0 Then
            Set book = bookTmp
        End If
    Next
    If book is Nothing Then
        WScript.Echo "�\��t�����Excel���N�����Ă��������B���u" & targetName & "�v���܂ރu�b�N���J���Ă��������B"
        Call WScript.Quit()
    End If
    
    targetRow = app.ActiveCell.Row
    targetCol = app.ActiveCell.Column

    Dim pictObj
    Set pictObj = app.ActiveSheet.Pictures.Insert(imagePath)
    targetRow = GetNextRow(app.ActiveSheet, targetRow, pictObj)
    app.ActiveSheet.Cells(targetRow, targetCol).Activate
End Sub


Call Main()


# image_to_xls
指定のフォルダ中の画像をExcelに張り付けます。  
この際、最初に張り付けるセルの位置と、画像の幅の上限の列を指定します。
これにより、大きすぎる画像は縮小してExcelに張り付けられます。


## 実行例

```
Get-ChildItem -LiteralPath "test\001\" -Filter *.jpg | Sort-Object -Property LastWriteTime |  % { $_.FullName } > test\001\input.txt
CScript image_to_xls.vbs C:\dev\excel_ss\image_to_xls\test\001\template.xlsx Sheet1 B2 L test\001\input.txt C:\dev\excel_ss\image_to_xls\test\001\out.xlsx
```

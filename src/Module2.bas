Option Explicit

'============================================================
' Module2
' 補助関数
'
' ■内容
' ・最終行取得
' ・最終列取得
'
' ■用途
' Module1 から呼び出して、
' データ範囲の終端判定に使う
'============================================================
'========================
' 最終行取得
'========================
Function LastRow(ColLong As Long) As Long
    LastRow = Cells(Rows.Count, ColLong).End(xlUp).Row
End Function

'========================
' 最終列取得
'========================
Function LastColumn(RowLong As Long) As Long
    LastColumn = Cells(RowLong, Columns.Count).End(xlToLeft).Column
End Function

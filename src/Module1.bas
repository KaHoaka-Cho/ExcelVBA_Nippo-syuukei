Option Explicit

'============================================================
' Module1
' 日報集計メイン処理
'
' ■目的
' 複数のExcelファイル内の「日報」シートを集約し、
' DB形式のデータとしてまとめる
'
' ■主な処理
' ・対象フォルダ選択
' ・各Excelファイルを順番に開く
' ・「日報」シートをコピー
' ・不要列、不要行、合計行を削除
' ・配列を使ってDB形式へ整形
' ・集計ブックとして保存
'
' ■前提
' ・対象シート名は「日報」
' ・元データのレイアウトがある程度共通
'============================================================
Option Explicit

Dim Book As Workbook '日報集計.xlsx
Dim NDB() As Variant '日報DB配列

'========================
' メイン呼び出し
'========================
Public Sub BE()

    Application.ScreenUpdating = False

    Set Book = Workbooks.Add
    Call Main

    Application.ScreenUpdating = True

End Sub

'========================
' メイン処理
'========================
Sub Main()

    Application.ScreenUpdating = False

    '========================
    ' 変数定義
    '========================
    Const cnsDIR = "\*.xls*"
    Dim FilePath As String
    Dim strFileName As String
    Dim i As Long

    Dim App As Excel.Application
    Dim Sheet As Worksheet
    Dim BookName As String

    Dim Book2 As Workbook
    Dim Sheet2 As Worksheet

    '========================
    ' 保存ファイル名
    '========================
    BookName = "日報集計.xlsx"

    '========================
    ' フォルダ選択
    '========================
    FilePath = FolderSelect()

    If FilePath = "" Then
        MsgBox "処理キャンセル"
        End
    End If

    '========================
    ' 初期シート作成
    '========================
    Book.Sheets.Add
    Book.ActiveSheet.Name = "DB用日報"

    With Book.ActiveSheet
        .Cells(1, 1) = "年月日"
        .Cells(1, 2) = "営業日"
        .Cells(1, 3) = "作業分類"
        .Cells(1, 4) = "業務項目"
        .Cells(1, 5) = "カテゴリ"
        .Cells(1, 6) = "担当者"
        .Cells(1, 7) = "作業時間"
    End With

    '書式設定
    Book.Sheets("DB用日報").Range("A:A").NumberFormat = "yyyy/mm/dd"
    Book.Sheets("DB用日報").Range("G:G").NumberFormat = "[hh]:mm"

    Application.DisplayAlerts = False

    '========================
    ' ファイルループ
    '========================
    strFileName = Dir(FilePath & cnsDIR, vbNormal)

    Do While strFileName <> ""

        Set Book2 = Workbooks.Open(FilePath & "\" & strFileName)

        Book2.Worksheets("日報").Select

        'コピー
        Book2.Worksheets("日報").Copy after:=Book.Worksheets(1)

        '仮シート作成
        Book.ActiveSheet.Copy after:=Book.Worksheets(1)
        ActiveSheet.Name = "日報DB"

        'フィルタ解除
        If ActiveSheet.AutoFilterMode = True Then Range("B4").AutoFilter

        '不要列削除
        Dim CC As Long
        CC = LastColumn(3)
        Columns(CC).Select
        Selection.Delete Shift:=xlToLeft

        '========================
        ' データ整形（行削除）
        '========================
        Dim RC As Long

        '合計削除
        RC = LastRow(3)
        Rows(RC + 1 & ":" & RC + 3).Delete

        'マニュアル合計削除
        RC = 特定行検索("C:C", "マニュアル合計")
        Rows(RC).Delete

        'トレース合計削除
        RC = 特定行検索("C:C", "トレース合計")
        Rows(RC).Delete

        '実務合計削除
        RC = 特定行検索("C:C", "実務合計")
        Rows(RC).Delete

        '不要行削除
        Call 不要行削除(2, 3, 5, "")
        Call 不要行削除(2, 3, 5, 0)

        '========================
        ' 配列処理
        '========================
        Call 配列処理実体

        'シート名変更
        Book.ActiveSheet.Name = Trim(Replace(Book2.Name, ".xlsx", ""))

        Book2.Close False

        strFileName = Dir()

    Loop

    '========================
    ' 保存
    '========================
    Book.SaveAs FilePath & "\" & BookName

    MsgBox "処理終了"

End Sub

'========================
' フォルダ選択
'========================
Function FolderSelect() As String

    Dim objFileDialog As Object
    Dim strPath As String

    Set objFileDialog = Application.FileDialog(msoFileDialogFolderPicker)

    With objFileDialog
        .Title = "フォルダを選択してください"
        If .Show = False Then Exit Function
        strPath = .SelectedItems(1)
    End With

    FolderSelect = strPath

End Function

'========================
' 行検索
'========================
Function 特定行検索(Col As String, SearchText As String) As Long

    Dim r As Range

    Set r = Range(Col).Find(what:=SearchText)

    If r Is Nothing Then
        特定行検索 = 0
    Else
        特定行検索 = r.Row
    End If

End Function

'========================
' 不要行削除
'========================
Sub 不要行削除(SearchRow As Long, SearchColumn As Long, CountEnd As Long, SearchValue As Variant)

    Dim i As Long
    Dim RC As Long
    Dim CC As Long
    Dim Vr As Variant

    RC = LastRow(SearchRow)
    CC = LastColumn(SearchColumn)

    For i = RC To CountEnd Step -1
        Vr = Cells(i, CC)
        If Vr = SearchValue Then
            Rows(i).Delete
        End If
    Next i

End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*********************************************************
' 関数名：フォルダ選択処理
' 概  要：フォルダ選択ダイアログを開き、フォルダパスを取得する
' 引  数：初期表示パス
'         未指定の場合は、ドキュメントフォルダ
' 戻り値：選択したフォルダパス
'*********************************************************
Public Function selectFolder(aDefault As String) As String
    Dim res As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If Not IsEmpty(aDefault) Then
            .InitialFileName = aDefault
        End If
        If .Show = True Then
            res = .SelectedItems(1)
            If res <> "" And Right(res, 1) <> Application.PathSeparator Then
                res = res & Application.PathSeparator
            End If
        End If
    End With
    selectFolder = res
End Function

'*********************************************************
' 関数名：拡張子除外処理
' 概  要：ファイル名から拡張子を除外する
' 引  数：ファイル名(拡張子有)
' 戻り値：ファイル名(拡張子無)
'*********************************************************
Public Function excludeExtension(aFileName As String) As String
    Dim res As String
    If InStrRev(aFileName, ".") <> 0 Then
        res = Left(aFileName, InStrRev(aFileName, ".") - 1)
    Else
        res = aFileName
    End If
    excludeExtension = res
End Function

'*********************************************************
' 関数名：フォルダ作成処理
' 概  要：引数に与えられたフォルダを作成する
' 引  数：フォルダパス
'         既存フォルダ対処(True：削除後作成/False：削除しない)
' 戻り値：処理結果(True：成功/False：失敗)
'*********************************************************
Public Function createFolder(aPath As String, aReCreated As Boolean)
    Dim res As Boolean
    Dim result As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If isFileExist(aPath) And aReCreated Then
        fso.DeleteFolder (aPath)
        fso.createFolder (aPath)
    ElseIf Not isFileExist(aPath) Then
        fso.createFolder (aPath)
    End If
    If Err = 0 Then
        res = True
    End If
    createFolder = res
End Function

'*********************************************************
' 関数名：ファイル(フォルダ)存在チェック処理
' 概  要：引数に与えられたファイル(フォルダ)の存在を確認する
' 引  数：ファイル(フォルダ)パス
' 戻り値：処理結果(True：存在する/False：存在しない)
'*********************************************************
Public Function isFileExist(aPath As String) As Boolean
    Dim res As Boolean
    If Dir(aPath) <> "" Then
        res = True
    End If
    isFileExist = res
End Function

'*********************************************************
' 関数名：シート存在チェック処理
' 概  要：引数に指定されたシートの存在を確認する
' 引  数：シート名
' 戻り値：処理結果(True：存在する/False：存在しない)
'*********************************************************
Public Function isSheetExist(aSheetName As String) As Boolean
    Dim res As Boolean
    Dim sh As Sheet
    For Each sh In Sheets
        If sh.Name = aSheetName Then
            res = True
            Exit For
        End If
    Next sh
    isSheetExist = res
End Function

'*********************************************************
' 関数名：アドレス参照型変換処理
' 概  要：A1形式の列とR1C1形式の列を変換する
' 引  数：列名(A1形式列またはR1C1形式列)
' 戻り値：列名(R1C1形式列またはA1形式列)
'*********************************************************
Public Function convAdd(aCol As Variant) As Variant
    Dim res As Variant
    If IsNumeric(aCol) Then
        res = Replace(Cells(Rows.Count, aCol).Address(False, False), Rows.Count, "")
    Else
        res = Range(aCol & Rows.Count).Column
    End If
    convAdd = res
End Function

'*********************************************************
' 関数名：A1形式列名取得処理
' 概  要：A1形式の列名を取得する
' 引  数：A1形式アドレス
' 戻り値：A1形式列名
'*********************************************************
Public Function getA1Col(anAdd As String) As String
    Dim res As String
    res = Split(Range(anAdd).Address, "$")(1)
    getA1Col = res
End Function

'*********************************************************
' 関数名：シート一覧取得処理
' 概  要：対象のブック内のシート一覧を取得する
' 引  数：対象ブック
' 戻り値：シート一覧
'*********************************************************
Public Function getSheetList(aBook As String) As String()

    Dim res() As String
    Dim i As Integer
    Dim tBook As Workbook
    Dim tSheet As Worksheet
    Workbooks.Open aBook, True
    Set tBook = Workbooks(Dir(aBook))
    For Each tSheet In tBook.Worksheets
        ReDim Preserve res(i)
        res(i) = tSheet.Name
        i = i + 1
    Next tSheet
    tBook.Close False
    getSheetList = res
End Function

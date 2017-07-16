Attribute VB_Name = "ComUtil"
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
Public Function createFolder(aPath As String, aReCreated as Boolean)
    Dim res As Boolean
    Dim result As String
    Dim fso = CreateObject("Scripting.FileSystemObject")
    If isFileExist(aPath) And aReCreated Then
        fso.DeleteFolder(aPath)
        fso.CreateFolder(aPath)
    ElseIf Not isFileExist(aPath) Then
        fso.CreateFolder(aPath)
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
            res = true
            Exit For
        End If
    Next sh
    isSheetExist = res
End Function
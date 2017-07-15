Attribute VB_Name = "ComUtil"
'/////////////////////////////////////////////////////////
'=========================================================
'*********************************************************
'#########################################################
' 関数名：フォルダ選択処理
' 概  要：フォルダ選択ダイアログを開き、フォルダパスを取得する
' 引  数：初期表示パス(文字列)
'         未指定の場合は、ドキュメントフォルダ
' 戻り値：選択したフォルダパス
'
Public Function selectFolder(aDefault As String) As String
    Dim res As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If Not IsEmpty(aDefault) Then
            .InitialFileName = aDefault
        End If
        If .Show = True Then
            res = .SelectedItems(1)
            If res <> "" And Right(res, 1) <> "\" Then
                res = res & Application.PathSeparator
            End If
        End If
    End With
    selectFolder = res
End Function

Private Sub ForDebug()
    MsgBox selectFolder("")
End Sub


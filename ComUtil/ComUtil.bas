Attribute VB_Name = "ComUtil"
'/////////////////////////////////////////////////////////
'=========================================================
'*********************************************************
'#########################################################
' �֐����F�t�H���_�I������
' �T  �v�F�t�H���_�I���_�C�A���O���J���A�t�H���_�p�X���擾����
' ��  ���F�����\���p�X(������)
'         ���w��̏ꍇ�́A�h�L�������g�t�H���_
' �߂�l�F�I�������t�H���_�p�X
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


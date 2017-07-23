Attribute VB_Name = "ComUtil"
'*********************************************************
' �֐����F�t�H���_�I������
' �T  �v�F�t�H���_�I���_�C�A���O���J���A�t�H���_�p�X���擾����
' ��  ���F�����\���p�X
'         ���w��̏ꍇ�́A�h�L�������g�t�H���_
' �߂�l�F�I�������t�H���_�p�X
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
' �֐����F�g���q���O����
' �T  �v�F�t�@�C��������g���q�����O����
' ��  ���F�t�@�C����(�g���q�L)
' �߂�l�F�t�@�C����(�g���q��)
'*********************************************************
Public Function excludeExtension(aFileName As String) As String
    Dim res As String
    If InStrRev(aFileName, ".") <> 0 Then
        res = Left(aFileName, InStrRev(aFileName, ".") - 1)
    End If
    excludeExtension = res
End Function

'*********************************************************
' �֐����F�t�H���_�쐬����
' �T  �v�F�����ɗ^����ꂽ�t�H���_���쐬����
' ��  ���F�t�H���_�p�X
'         �����t�H���_�Ώ�(True�F�폜��쐬/False�F�폜���Ȃ�)
' �߂�l�F��������(True�F����/False�F���s)
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
' �֐����F�t�@�C��(�t�H���_)���݃`�F�b�N����
' �T  �v�F�����ɗ^����ꂽ�t�@�C��(�t�H���_)�̑��݂��m�F����
' ��  ���F�t�@�C��(�t�H���_)�p�X
' �߂�l�F��������(True�F���݂���/False�F���݂��Ȃ�)
'*********************************************************
Public Function isFileExist(aPath As String) As Boolean
    Dim res As Boolean
    If Dir(aPath) <> "" Then
        res = True
    End If
    isFileExist = res
End Function

'*********************************************************
' �֐����F�V�[�g���݃`�F�b�N����
' �T  �v�F�����Ɏw�肳�ꂽ�V�[�g�̑��݂��m�F����
' ��  ���F�V�[�g��
' �߂�l�F��������(True�F���݂���/False�F���݂��Ȃ�)
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
' �֐����F�A�h���X�Q�ƌ^�ϊ�����
' �T  �v�FA1�`���̗��R1C1�`���̗��ϊ�����
' ��  ���F��(A1�`����܂���R1C1�`����)
' �߂�l�F��(R1C1�`����܂���A1�`����)
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
' �֐����FA1�`���񖼎擾����
' �T  �v�FA1�`���̗񖼂��擾����
' ��  ���FA1�`���A�h���X
' �߂�l�FA1�`����
'*********************************************************
Public Function getA1Col(anAdd As String) As String
    Dim res As String
    res = Split(Range(anAdd).Address, "$")(1)
    getA1Col = res
End Function

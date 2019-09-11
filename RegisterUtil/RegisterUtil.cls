VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegisterUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' アプリケーション名
'   アプリケーション毎に設定する。
Private Const APP_NAME As String = ""


Public Function regist(sectionName As String, itemKey As String, itemValue As String) As Boolean
    If APP_NAME = "" _
            Or sectionName = "" _
            Or itemKey = "" _
            Then
        regist = False
        Exit Function
    End If
    SaveSetting APP_NAME, sectionName, itemKey, itemValue
    regist = True
End Function

Public Function selectByKey(sectionName As String, itemKey As String, defaultValue As String) As String
    selectByKey = GetSetting(APP_NAME, sectionName, itemKey, defaultValue)
End Function

Public Function selectBySection(sectionName As String) As Object
    Dim tmpMap As Variant
    Dim i As Long
    Dim resultMap As Object
    
    Set resultMap = CreateObject("Scripting.Dictionary")
    tmpMap = GetAllSettings(APP_NAME, sectionName)
    For i = 0 To UBound(tmpMap)
        resultMap.Add buf(i, 0), buf(i, 1)
    Next i
    selectBySection = resultMap
End Function

Public Function delete(Optional sectionName As String = "", Optional itemKey As String = "")
    If itemKey = "" Then
        DeleteSetting APP_NAME, sectionName
    ElseIf sectionName = "" Then
        DeleteSetting APP_NAME
    Else
        DeleteSetting APP_NAME, sectionName, itemKey
    End If
End Function

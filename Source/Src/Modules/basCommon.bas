Attribute VB_Name = "basCommon"
'-----------------------------------------------------------------------------------------------------
'
' [RelaxTools-Addin] v4
'
' Copyright (c) 2009 Yasuhiro Watanabe
' https://github.com/RelaxTools/RelaxTools-Addin
' author:relaxtools@opensquare.net
'
' The MIT License (MIT)
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'
'-----------------------------------------------------------------------------------------------------
Option Explicit

'--------------------------------------------------------------
'　フォルダ選択
'--------------------------------------------------------------
Public Function rlxSelectFolder() As String
 
    Dim objShell As Object
    Dim objPath As Object
    Dim WS As Object
    Dim strFolder As String
    
    Set objShell = CreateObject("Shell.Application")
    Set objPath = objShell.BrowseForFolder(&O0, "フォルダを選んでください", &H1 + &H10, "")
    If Not objPath Is Nothing Then
    
        'なぜか「デスクトップ」のパスが取得できない
        If objPath = "デスクトップ" Then
            Set WS = CreateObject("WScript.Shell")
            rlxSelectFolder = WS.SpecialFolders("Desktop")
        Else
            rlxSelectFolder = objPath.Items.Item.Path
        
        End If
    Else
        rlxSelectFolder = ""
    End If
    
End Function
'--------------------------------------------------------------
'　ファイルセパレータ付加
'--------------------------------------------------------------
Public Function rlxAddFileSeparator(ByVal strFile As String) As String
    If Right(strFile, 1) = "\" Then
        rlxAddFileSeparator = strFile
    Else
        rlxAddFileSeparator = strFile & "\"
    End If
End Function
'--------------------------------------------------------------
'　ファイル名取得
'--------------------------------------------------------------
Public Function rlxGetFullpathFromFileName(ByVal strPath As String) As String

    Dim lngCnt As Long
    Dim lngMax As Long
    Dim strResult As String
    
    strResult = strPath
    
    lngMax = Len(strPath)
    
    For lngCnt = lngMax To 1 Step -1
    
        Select Case Mid$(strPath, lngCnt, 1)
            Case "\", "/"
                If lngCnt = lngMax Then
                Else
                    strResult = Mid$(strPath, lngCnt + 1)
                End If
                Exit For
        End Select
    
    Next

    rlxGetFullpathFromFileName = strResult

End Function
'--------------------------------------------------------------
'　ファイル数カウント
'--------------------------------------------------------------
Public Sub rlxGetFilesCount(ByRef objFs As Object, ByVal strPath As String, ByRef lngFCnt As Long, ByVal blnFile As Boolean, ByVal blnFolder As Boolean, ByVal blnSubFolder As Boolean)

    Dim objfld As Object
    Dim objSub As Object

    Set objfld = objFs.GetFolder(strPath)
    
    If blnFile Then
        lngFCnt = lngFCnt + objfld.files.count
    End If
    
    If blnFolder Then
        lngFCnt = lngFCnt + objfld.SubFolders.count
    End If
    
        'フォルダ取得あり
    If blnSubFolder Then
        For Each objSub In objfld.SubFolders
            DoEvents
            rlxGetFilesCount objFs, objSub.Path, lngFCnt, blnFile, blnFolder, blnSubFolder
        Next
        
    End If
End Sub
'--------------------------------------------------------------
'　アプリケーションフォルダ取得
'--------------------------------------------------------------
Public Function rlxGetAppDataFolder() As String

    On Error Resume Next
    
    Dim strFolder As String
    
    rlxGetAppDataFolder = ""
    
    With CreateObject("Scripting.FileSystemObject")
    
        strFolder = .BuildPath(CreateObject("Wscript.Shell").SpecialFolders("AppData"), C_TITLE)
        
        If .FolderExists(strFolder) Then
        Else
            .createFolder strFolder
        End If
        
        rlxGetAppDataFolder = .BuildPath(strFolder, "\")
        
    End With
    

End Function
'--------------------------------------------------------------
'  フォルダの作成
'--------------------------------------------------------------
Sub rlxCreateFolder(ByVal strPath As String)

    Dim v As Variant
    Dim s As Variant
    
    Dim f As String
    
    v = Split(strPath, "\")

    On Error Resume Next
    For Each s In v
    
        If f = "" Then
            f = s
            MkDir f & "\"
        Else
            f = f & "\" & s
            MkDir f
        End If
    
    Next

End Sub
'--------------------------------------------------------------
'　エラーメッセージ表示
'--------------------------------------------------------------
Sub rlxErrMsg(ByRef objErr As Object)

    Select Case objErr.Number
        Case 0
        Case 1004
            MsgBox "エラーです。シート保護などを確認してください。", vbCritical + vbOKOnly, C_TITLE
        Case Else
            MsgBox objErr.Description & "(" & Err.Number & ")", vbCritical + vbOKOnly, C_TITLE
    End Select

End Sub

Attribute VB_Name = "basRibbon"
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
#If VBA7 And Win64 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
#Else
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
#End If

Private mIR As IRibbonUI

Private Const C_START_ROW As Long = 2
Private Const C_COL_NO As Long = 1
Private Const C_COL_ID As Long = 2
Private Const C_COL_IMAGE As Long = 3
Private Const C_COL_LABEL As Long = 4
Private Const C_COL_SUPERTIP As Long = 5

'メニュー
Public mObjMenu As Object

Public Function C_TITLE() As String
    C_TITLE = ThisWorkbook.BuiltinDocumentProperties("Title").value
End Function

'--------------------------------------------------------------------
' リボンロード時イベント
'--------------------------------------------------------------------
Public Sub ribbonLoadedSub(ByRef IR As IRibbonUI)

    On Error Resume Next
    Set mIR = IR
    SaveSetting C_TITLE, "Ribbon", "Address", CStr(ObjPtr(IR))

End Sub
'--------------------------------------------------------------------
' リボンのリフレッシュ
'--------------------------------------------------------------------
Public Sub RefreshRibbon(Optional control As IRibbonControl)

    Dim strBuf As String
    
    On Error GoTo e
    
    'グローバル変数がクリアされたしまった場合、レジストリから復帰
    If mIR Is Nothing Then
        
        strBuf = GetSetting(C_TITLE, "Ribbon", "Address", 0)
        Set mIR = getObjectFromAddres(strBuf)
        
    End If
    
    If mIR Is Nothing Then
    Else
        If control Is Nothing Then
            mIR.Invalidate
        Else
            mIR.InvalidateControl control.ID
        End If
    End If

e:

End Sub
'--------------------------------------------------------------------
'リボンより受け取ったIDをそのままマクロ名として実行するラッパー関数
'--------------------------------------------------------------------
Public Sub OnActionSub(control As IRibbonControl)

    Dim lngPos As Long
    Dim strBuf As String
    
    On Error GoTo e
    
    strBuf = getMacroName(control)
    
    '開始ログ
    Logger.LogBegin strBuf
    
    '文字列のマクロ名を実行する。
    Application.Run strBuf
    
    
    Call RefreshRibbon(control)

    Dim strLabel As String
    strLabel = getSheetItem(control, C_COL_LABEL)
    Application.OnRepeat strLabel, strBuf
    
    '終了ログ
    Logger.LogFinish strBuf
    
    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
'--------------------------------------------------------------------
' ヘルプ内容を表示する。customUIから使用
'--------------------------------------------------------------------
Public Sub GetSupertipSub(control As IRibbonControl, ByRef value)

    On Error GoTo e
    
    value = getSheetItem(control, C_COL_SUPERTIP)

    Call RefreshRibbon

    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
'--------------------------------------------------------------------
' メニュー表示内容を表示する。customUIから使用
'--------------------------------------------------------------------
Public Sub GetImageSub(control As IRibbonControl, ByRef value)

    On Error GoTo e

    value = getSheetItem(control, C_COL_IMAGE)

    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
'--------------------------------------------------------------------
' ラベルを表示する。customUIから使用
'--------------------------------------------------------------------
Public Sub GetLabelSub(control As IRibbonControl, ByRef value)

    On Error GoTo e
    
    value = getSheetItem(control, C_COL_LABEL)
    
    Exit Sub
e:
    Call rlxErrMsg(Err)
End Sub
'--------------------------------------------------------------------
' マクロ名取得
'--------------------------------------------------------------------
Private Function getMacroName(control As IRibbonControl) As String
    
    Dim lngPos As Long
    
    '同じマクロを複数登録可能とするためにドット以降の文字を削除
    lngPos = InStr(control.ID, ".")

    If lngPos = 0 Then
        getMacroName = control.ID
    Else
        getMacroName = Mid$(control.ID, 1, lngPos - 1)
    End If

End Function
'--------------------------------------------------------------------
' シートから指定項目を取得する
'--------------------------------------------------------------------
Private Function getSheetItem(control As IRibbonControl, lngItem As Long) As String

    Dim lngPos As Long
    Dim strBuf As String
    Dim i As Long
    Dim m As MenuDTO
    Dim Key As String
    
    getSheetItem = ""
    
    strBuf = getMacroName(control)
    
    If mObjMenu Is Nothing Then
    
        Set mObjMenu = CreateObject("Scripting.Dictionary")
    
        i = C_START_ROW
        
        Do Until ThisWorkbook.Worksheets("HELP").Cells(i, C_COL_NO).value = ""
            
            Set m = New MenuDTO
            m.ID = ThisWorkbook.Worksheets("HELP").Cells(i, C_COL_ID).value
            m.IMAGE = ThisWorkbook.Worksheets("HELP").Cells(i, C_COL_IMAGE).value
            m.LABEL = ThisWorkbook.Worksheets("HELP").Cells(i, C_COL_LABEL).value
            m.SUPERTIP = ThisWorkbook.Worksheets("HELP").Cells(i, C_COL_SUPERTIP).value
            
            If Not mObjMenu.Exists(m.ID) Then
                mObjMenu.Add m.ID, m
            Else
                MsgBox "メニューのマクロ名が重複しています。" & strBuf
            End If
            i = i + 1
        Loop
        
    End If
    
    If mObjMenu.Exists(strBuf) Then
        Select Case lngItem
            Case C_COL_IMAGE
                getSheetItem = mObjMenu.Item(strBuf).IMAGE
            Case C_COL_LABEL
                getSheetItem = mObjMenu.Item(strBuf).LABEL
            Case C_COL_SUPERTIP
                getSheetItem = mObjMenu.Item(strBuf).SUPERTIP
        End Select
    Else
        getSheetItem = ""
    End If
End Function
'--------------------------------------------------------------
'　アドレス文字列からオブジェクトに変換
'--------------------------------------------------------------
Private Function getObjectFromAddres(ByVal strAddress As String) As Object

    Dim obj As Object

    #If VBA7 And Win64 Then
        Dim p As LongPtr
        p = CLngPtr(strAddress)
    #Else
        Dim p As Long
        p = CLng(strAddress)
    #End If
  
    CopyMemory obj, p, LenB(p)
    
    Set getObjectFromAddres = obj

End Function


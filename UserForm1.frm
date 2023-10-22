VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   1344
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5832
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'ドロップダウンリストを作成・表示
Private Sub ComboBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim 終行 As Long, 行 As Long
    Dim 配列()
    With Sheets("郵便番号一覧")
        終行 = .Cells(Rows.Count, 1).End(xlUp).Row
        配列 = .Cells(2, 1).Resize(終行 - 1, 1).Value
    End With
    With ComboBox1
        .MatchEntry = fmMatchEntryNone
        Select Case KeyCode
            Case 28, 29, vbKeyBack, vbKeySpace, vbKeyDelete, vbKeyA To vbKeyZ, vbKey0 To vbKey9, vbKeyNumpad0 To vbKeyNumpad9
                .Clear
                For 行 = 1 To 終行 - 1
                    If 配列(行, 1) Like ComboBox1.Text & "*" Then .AddItem 配列(行, 1)
                Next
        End Select
    End With
    ComboBox1.DropDown
End Sub
'郵便番号を検索代入
Private Sub ComboBox1_Change()
    Dim 終行 As Long, 行 As Long
    Dim 配列()
    With Sheets("郵便番号一覧")
        終行 = .Cells(Rows.Count, 1).End(xlUp).Row
        配列 = .Cells(2, 1).Resize(終行 - 1, 2).Value
    End With
    With TextBox1
        For 行 = 1 To 終行 - 1
            If ComboBox1.Text = 配列(行, 1) Then .Text = 配列(行, 2)
        Next
    End With
End Sub
'入力フォームへ代入
Private Sub CommandButton1_Click()
    If TextBox1.Text = "" Then
        MsgBox "不正な住所です"
        Exit Sub
    End If
    With Sheets("入力フォーム")
        .Range("郵便番号") = TextBox1.Text
        Select Case InStr(ComboBox1.Text, "（")
            Case Is > 0: .Range("住所上段") = Left(ComboBox1.Text, InStr(ComboBox1.Text, "（") - 1)
            Case Else: .Range("住所上段") = ComboBox1.Text
        End Select
        Unload UserForm1
    End With
End Sub

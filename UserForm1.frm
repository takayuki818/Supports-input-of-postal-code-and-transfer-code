VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3120
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   6216
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'フォーム初期化設定
Private Sub UserForm_Initialize()
    With Sheets("入力フォーム")
        .Range("郵便番号").ClearContents
    End With
    With Me
        .Height = 80
    End With
    With CommandButton1
        .Height = 20.25
    End With
    With TextBox1
        .SetFocus
        .IMEMode = fmIMEModeHiragana
        .ControlSource = "入力フォーム!住所上段"
    End With
    With ListBox1
        .ColumnCount = 2
        .TextColumn = 2
        .ColumnWidths = "70 pt;"
        .Height = 103
        .Visible = False
    End With
End Sub
'リストボックスを作成・表示
Private Sub TextBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim 終行 As Long, 行 As Long, 添字 As Long
    Dim 配列()
    With Sheets("郵便番号一覧")
        終行 = .Cells(Rows.Count, 1).End(xlUp).Row
        配列 = .Cells(2, 1).Resize(終行 - 1, 2).Value
    End With
    With ListBox1
        Select Case KeyCode
            '28:変換キー、29:無変換キー（定数なし）
            Case 28, 29, vbKeyBack, vbKeySpace, vbKeyDelete, vbKeyA To vbKeyZ, vbKey0 To vbKey9, vbKeyNumpad0 To vbKeyNumpad9
                .Clear
                For 行 = 1 To 終行 - 1
                    If 配列(行, 2) Like "*" & TextBox1.Text & "*" Then
                        添字 = 添字 + 1
                        .AddItem ""
                        .List(添字 - 1, 0) = 配列(行, 1)
                        .List(添字 - 1, 1) = 配列(行, 2)
                    End If
                Next
                .Visible = True
                Me.Height = 180
                CommandButton1.Height = 120
        End Select
    End With
End Sub
'リストボックス選択→テキストボックス1に代入
Private Sub ListBox1_Click()
    TextBox1.Text = ListBox1.Text
End Sub
'入力フォームへ代入
Private Sub CommandButton1_Click()
    With Sheets("入力フォーム")
        ListBox1.TextColumn = 1
        .Range("郵便番号") = ListBox1.Text
        Select Case InStr(TextBox1.Text, "（")
            Case Is > 0: .Range("住所上段") = Left(TextBox1.Text, InStr(TextBox1.Text, "（") - 1)
            Case Else: .Range("住所上段") = TextBox1.Text
        End Select
        Unload Me
    End With
End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   3120
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7800
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub UserForm_Initialize()
    With Sheets("入力フォーム")
        .Range("本店コード").ClearContents
        .Range("支店コード").ClearContents
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
        .ControlSource = "入力フォーム!本店名"
    End With
    With TextBox2
        .IMEMode = fmIMEModeHiragana
        .ControlSource = "入力フォーム!支店名"
    End With
    With ListBox1
        Select Case TextBox1.Text
            Case "": .Visible = False
            Case Else
                .Visible = True
                Call 本店リスト格納
                Call 支店リスト格納
        End Select
        .ColumnCount = 2
        .ColumnWidths = "38 pt;"
        .TextColumn = 2
        .Height = 98
    End With
    With ListBox2
        Select Case TextBox2.Text
            Case "": .Visible = False
            Case Else
                .Visible = True
                Call 支店抜粋リスト格納
        End Select
        .ColumnCount = 2
        .ColumnWidths = "33 pt;"
        .TextColumn = 2
        .Height = 98
    End With
End Sub
Private Sub TextBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    With ListBox1
        Select Case KeyCode
            '28:変換キー、29:無変換キー（定数なし）
            Case 28, 29, vbKeyBack, vbKeySpace, vbKeyDelete, vbKeyA To vbKeyZ, vbKey0 To vbKey9, vbKeyNumpad0 To vbKeyNumpad9
                Call 本店リスト格納
        End Select
    End With
End Sub
Private Sub ListBox1_Click()
    TextBox1.Text = ListBox1.Text
    TextBox2.SetFocus
    Call 支店リスト格納
End Sub
Private Sub TextBox2_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    With ListBox2
        Select Case KeyCode
            '28:変換キー、29:無変換キー（定数なし）
            Case 28, 29, vbKeyBack, vbKeySpace, vbKeyDelete, vbKeyA To vbKeyZ, vbKey0 To vbKey9, vbKeyNumpad0 To vbKeyNumpad9
            Call 支店抜粋リスト格納
        End Select
    End With
End Sub
Private Sub ListBox2_Click()
    TextBox2.Text = ListBox2.Text
End Sub
'入力フォームへ代入
Private Sub CommandButton1_Click()
    With Sheets("入力フォーム")
        ListBox1.TextColumn = 1
        .Range("本店コード") = ListBox1.Text
        ListBox2.TextColumn = 1
        .Range("支店コード") = ListBox2.Text
        Unload UserForm2
    End With
End Sub
Sub 本店リスト格納()
    Dim 終行 As Long, 行 As Long, 添字 As Long
    Dim 配列()
    With Sheets("金融機関コード")
        終行 = .Cells(Rows.Count, 1).End(xlUp).Row
        配列 = .Cells(2, 1).Resize(終行 - 1, 2).Value
    End With
    With ListBox1
        .Clear
        '1行目処理
        If 配列(1, 2) Like "*" & TextBox1.Text & "*" Then
            添字 = 添字 + 1
            .AddItem ""
            .List(0, 0) = 配列(1, 1)
            .List(0, 1) = 配列(1, 2)
        End If
        '2行目以降処理
        For 行 = 2 To 終行 - 1
            If 配列(行, 1) <> 配列(行 - 1, 1) Then
                If 配列(行, 2) Like "*" & TextBox1.Text & "*" Then
                    添字 = 添字 + 1
                    .AddItem ""
                    .List(添字 - 1, 0) = 配列(行, 1)
                    .List(添字 - 1, 1) = 配列(行, 2)
                End If
            End If
        Next
        .Visible = True
        Me.Height = 170
        CommandButton1.Height = 115
    End With
End Sub
Sub 支店リスト格納()
    Dim 終行 As Long, 行 As Long, 添字 As Long
    Dim 配列()
    With Sheets("金融機関コード")
        終行 = .Cells(Rows.Count, 2).End(xlUp).Row
        配列 = .Cells(2, 2).Resize(終行 - 1, 3).Value
    End With
    With ListBox2
        .Clear
        For 行 = 1 To 終行 - 1
            If 配列(行, 1) = TextBox1.Text Then
                添字 = 添字 + 1
                .AddItem ""
                .List(添字 - 1, 0) = 配列(行, 2)
                .List(添字 - 1, 1) = 配列(行, 3)
            End If
        Next
        .Visible = True
        Me.Height = 170
        CommandButton1.Height = 115
    End With
End Sub
Sub 支店抜粋リスト格納()
    Dim 終行 As Long, 行 As Long, 添字 As Long
    Dim 配列()
    With Sheets("金融機関コード")
        終行 = .Cells(Rows.Count, 2).End(xlUp).Row
        配列 = .Cells(2, 2).Resize(終行 - 1, 3).Value
    End With
    With ListBox2
        .Clear
        For 行 = 1 To 終行 - 1
            If 配列(行, 1) = TextBox1.Text Then
                If 配列(行, 3) Like "*" & TextBox2.Text & "*" Then
                    添字 = 添字 + 1
                    .AddItem ""
                    .List(添字 - 1, 0) = 配列(行, 2)
                    .List(添字 - 1, 1) = 配列(行, 3)
                End If
            End If
        Next
        .Visible = True
        Me.Height = 170
        CommandButton1.Height = 115
    End With
End Sub

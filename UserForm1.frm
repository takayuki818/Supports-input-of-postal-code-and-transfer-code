VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3120
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   6216
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'�t�H�[���������ݒ�
Private Sub UserForm_Initialize()
    With Sheets("���̓t�H�[��")
        .Range("�X�֔ԍ�").ClearContents
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
        .ControlSource = "���̓t�H�[��!�Z����i"
    End With
    With ListBox1
        .ColumnCount = 2
        .TextColumn = 2
        .ColumnWidths = "70 pt;"
        .Height = 103
        .Visible = False
    End With
End Sub
'���X�g�{�b�N�X���쐬�E�\��
Private Sub TextBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim �I�s As Long, �s As Long, �Y�� As Long
    Dim �z��()
    With Sheets("�X�֔ԍ��ꗗ")
        �I�s = .Cells(Rows.Count, 1).End(xlUp).Row
        �z�� = .Cells(2, 1).Resize(�I�s - 1, 2).Value
    End With
    With ListBox1
        Select Case KeyCode
            '28:�ϊ��L�[�A29:���ϊ��L�[�i�萔�Ȃ��j
            Case 28, 29, vbKeyBack, vbKeySpace, vbKeyDelete, vbKeyA To vbKeyZ, vbKey0 To vbKey9, vbKeyNumpad0 To vbKeyNumpad9
                .Clear
                For �s = 1 To �I�s - 1
                    If �z��(�s, 2) Like "*" & TextBox1.Text & "*" Then
                        �Y�� = �Y�� + 1
                        .AddItem ""
                        .List(�Y�� - 1, 0) = �z��(�s, 1)
                        .List(�Y�� - 1, 1) = �z��(�s, 2)
                    End If
                Next
                .Visible = True
                Me.Height = 180
                CommandButton1.Height = 120
        End Select
    End With
End Sub
'���X�g�{�b�N�X�I�����e�L�X�g�{�b�N�X1�ɑ��
Private Sub ListBox1_Click()
    TextBox1.Text = ListBox1.Text
End Sub
'���̓t�H�[���֑��
Private Sub CommandButton1_Click()
    With Sheets("���̓t�H�[��")
        ListBox1.TextColumn = 1
        .Range("�X�֔ԍ�") = ListBox1.Text
        Select Case InStr(TextBox1.Text, "�i")
            Case Is > 0: .Range("�Z����i") = Left(TextBox1.Text, InStr(TextBox1.Text, "�i") - 1)
            Case Else: .Range("�Z����i") = TextBox1.Text
        End Select
        Unload Me
    End With
End Sub

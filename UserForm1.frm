VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   1344
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5832
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'�h���b�v�_�E�����X�g���쐬�E�\��
Private Sub ComboBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim �I�s As Long, �s As Long
    Dim �z��()
    With Sheets("�X�֔ԍ��ꗗ")
        �I�s = .Cells(Rows.Count, 1).End(xlUp).Row
        �z�� = .Cells(2, 1).Resize(�I�s - 1, 1).Value
    End With
    With ComboBox1
        .MatchEntry = fmMatchEntryNone
        Select Case KeyCode
            Case 28, 29, vbKeyBack, vbKeySpace, vbKeyDelete, vbKeyA To vbKeyZ, vbKey0 To vbKey9, vbKeyNumpad0 To vbKeyNumpad9
                .Clear
                For �s = 1 To �I�s - 1
                    If �z��(�s, 1) Like ComboBox1.Text & "*" Then .AddItem �z��(�s, 1)
                Next
        End Select
    End With
    ComboBox1.DropDown
End Sub
'�X�֔ԍ����������
Private Sub ComboBox1_Change()
    Dim �I�s As Long, �s As Long
    Dim �z��()
    With Sheets("�X�֔ԍ��ꗗ")
        �I�s = .Cells(Rows.Count, 1).End(xlUp).Row
        �z�� = .Cells(2, 1).Resize(�I�s - 1, 2).Value
    End With
    With TextBox1
        For �s = 1 To �I�s - 1
            If ComboBox1.Text = �z��(�s, 1) Then .Text = �z��(�s, 2)
        Next
    End With
End Sub
'���̓t�H�[���֑��
Private Sub CommandButton1_Click()
    If TextBox1.Text = "" Then
        MsgBox "�s���ȏZ���ł�"
        Exit Sub
    End If
    With Sheets("���̓t�H�[��")
        .Range("�X�֔ԍ�") = TextBox1.Text
        Select Case InStr(ComboBox1.Text, "�i")
            Case Is > 0: .Range("�Z����i") = Left(ComboBox1.Text, InStr(ComboBox1.Text, "�i") - 1)
            Case Else: .Range("�Z����i") = ComboBox1.Text
        End Select
        Unload UserForm1
    End With
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()

End Sub

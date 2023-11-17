VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �U������� 
   Caption         =   "UserForm2"
   ClientHeight    =   3120
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7800
   OleObjectBlob   =   "�U�������.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�U�������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub UserForm_Initialize()
    With Sheets("���̓t�H�[��")
        .Range("�{�X�R�[�h").ClearContents
        .Range("�x�X�R�[�h").ClearContents
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
        .ControlSource = "���̓t�H�[��!�{�X��"
    End With
    With TextBox2
        .IMEMode = fmIMEModeHiragana
        .ControlSource = "���̓t�H�[��!�x�X��"
    End With
    With ListBox1
        Select Case TextBox1.Text
            Case "": .Visible = False
            Case Else
                .Visible = True
                Call �{�X���X�g�i�[
                Call �x�X���X�g�i�[
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
                Call �x�X�������X�g�i�[
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
            '28:�ϊ��L�[�A29:���ϊ��L�[�i�萔�Ȃ��j
            Case vbKeyReturn, 28, 29, vbKeyBack, vbKeySpace, vbKeyDelete, vbKeyA To vbKeyZ, vbKey0 To vbKey9, vbKeyNumpad0 To vbKeyNumpad9
                Call �{�X���X�g�i�[
        End Select
    End With
End Sub
Private Sub ListBox1_Click()
    TextBox1.Text = ListBox1.Text
    TextBox2.SetFocus
    Call �x�X���X�g�i�[
End Sub
Private Sub TextBox2_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    With ListBox2
        Select Case KeyCode
            '28:�ϊ��L�[�A29:���ϊ��L�[�i�萔�Ȃ��j
            Case vbKeyReturn, 28, 29, vbKeyBack, vbKeySpace, vbKeyDelete, vbKeyA To vbKeyZ, vbKey0 To vbKey9, vbKeyNumpad0 To vbKeyNumpad9
            Call �x�X�������X�g�i�[
        End Select
    End With
End Sub
Private Sub ListBox2_Click()
    TextBox2.Text = ListBox2.Text
End Sub
'���̓t�H�[���֑��
Private Sub CommandButton1_Click()
    With Sheets("���̓t�H�[��")
        ListBox1.TextColumn = 1
        .Range("�{�X�R�[�h") = ListBox1.Text
        ListBox2.TextColumn = 1
        .Range("�x�X�R�[�h") = ListBox2.Text
        Unload Me
    End With
End Sub
Sub �{�X���X�g�i�[()
    Dim �I�s As Long, �s As Long, �Y�� As Long
    Dim �z��()
    With Sheets("���Z�@�փR�[�h")
        �I�s = .Cells(Rows.Count, 1).End(xlUp).Row
        �z�� = .Cells(2, 1).Resize(�I�s - 1, 2).Value
    End With
    With ListBox1
        .Clear
        '1�s�ڏ���
        If �z��(1, 2) Like "*" & TextBox1.Text & "*" Then
            �Y�� = �Y�� + 1
            .AddItem ""
            .List(0, 0) = �z��(1, 1)
            .List(0, 1) = �z��(1, 2)
        End If
        '2�s�ڈȍ~����
        For �s = 2 To �I�s - 1
            If �z��(�s, 1) <> �z��(�s - 1, 1) Then
                If �z��(�s, 2) Like "*" & TextBox1.Text & "*" Then
                    �Y�� = �Y�� + 1
                    .AddItem ""
                    .List(�Y�� - 1, 0) = �z��(�s, 1)
                    .List(�Y�� - 1, 1) = �z��(�s, 2)
                End If
            End If
        Next
        .Visible = True
        Me.Height = 170
        CommandButton1.Height = 115
    End With
End Sub
Sub �x�X���X�g�i�[()
    Dim �I�s As Long, �s As Long, �Y�� As Long
    Dim �z��()
    With Sheets("���Z�@�փR�[�h")
        �I�s = .Cells(Rows.Count, 2).End(xlUp).Row
        �z�� = .Cells(2, 2).Resize(�I�s - 1, 3).Value
    End With
    With ListBox2
        .Clear
        For �s = 1 To �I�s - 1
            If �z��(�s, 1) = TextBox1.Text Then
                �Y�� = �Y�� + 1
                .AddItem ""
                .List(�Y�� - 1, 0) = �z��(�s, 2)
                .List(�Y�� - 1, 1) = �z��(�s, 3)
            End If
        Next
        .Visible = True
        Me.Height = 170
        CommandButton1.Height = 115
    End With
End Sub
Sub �x�X�������X�g�i�[()
    Dim �I�s As Long, �s As Long, �Y�� As Long
    Dim �z��()
    With Sheets("���Z�@�փR�[�h")
        �I�s = .Cells(Rows.Count, 2).End(xlUp).Row
        �z�� = .Cells(2, 2).Resize(�I�s - 1, 3).Value
    End With
    With ListBox2
        .Clear
        For �s = 1 To �I�s - 1
            If �z��(�s, 1) = TextBox1.Text Then
                If �z��(�s, 3) Like "*" & TextBox2.Text & "*" Then
                    �Y�� = �Y�� + 1
                    .AddItem ""
                    .List(�Y�� - 1, 0) = �z��(�s, 2)
                    .List(�Y�� - 1, 1) = �z��(�s, 3)
                End If
            End If
        Next
        .Visible = True
        Me.Height = 170
        CommandButton1.Height = 115
    End With
End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ��_�ٷΰ����ư�����1 
   Caption         =   "�ٷΰ��� ��ư ����"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5060
   OleObjectBlob   =   "��_�ٷΰ����ư�����1.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "��_�ٷΰ����ư�����1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    ����Ʈ_�����̵�.Clear
    ����Ʈ_����.Clear
End Sub

Private Sub ����Ʈ_�����̵�_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    'With ����Ʈ_�����̵�
        '.ListCount�� ���� ������ ������ ���ϴ� ����� �ƴϰ�, _
            �Ʒ� ����� �˻��ؼ� ã�Ƴ´µ�, ������ ����غ��� ���� ������ ����
        'If .ItemsSelected.Count > 0 Then
            �����̵�_����
        'End If
    'End With
End Sub

Private Sub ��ư_����_Click()
   �����̵�_����
End Sub

Private Sub �����̵�_����()
    '���õ� �׸�鿡 ���ؼ��� �ݺ����� ������ ������, _
        ���� �������� ���� ����� ������ �Ǿ����� �ʾ� _
        ��� �׸� ���� ���ÿ��θ� Ȯ���ϸ� ����
    
    Dim i As Integer
    For i = 0 To ����Ʈ_�����̵�.ListCount - 1
        If ����Ʈ_�����̵�.Selected(i) Then
            With ����Ʈ_����
                .AddItem (����Ʈ_�����̵�.List(i, 0))
                .List(.ListCount - 1, 1) = ����Ʈ_�����̵�.List(i, 1)
            End With
            ����Ʈ_�����̵�.Selected(i) = False
        End If
    Next
End Sub

Private Sub ��ư_����_Click()
    With ����Ʈ_����
        '�̹� ó���̰ų� ���õ� �׸��� ���� ��
        If .ListIndex = 0 Or .ListIndex < 0 Then
            Exit Sub
        
        Else
            '.ListIndex-1 �ڸ��� �׸� �߰� -> .ListIndex�� 1 �þ, �� �׸�� index 2��ŭ ���� ��
            .AddItem .List(.ListIndex, 0), (.ListIndex - 1)
            
            .List((.ListIndex - 2), 1) = .List(.ListIndex, 1)
            .ListIndex = .ListIndex - 2
            .RemoveItem .ListIndex + 2
        End If
    End With
End Sub
Private Sub ��ư_�Ʒ���_Click()
    With ����Ʈ_����
        '�̹� �������̰ų� ���õ� �׸��� ���� ��
        If .ListIndex = .ListCount - 1 Or .ListIndex < 0 Then
            Exit Sub
        
        Else
            '.ListIndex+2 �ڸ��� �׸� �߰� -> .ListIndex ���� ����, �� �׸�� index 2��ŭ ���� ��
            .AddItem .List(.ListIndex, 0), (.ListIndex + 2)
            
            .List((.ListIndex + 2), 1) = .List(.ListIndex, 1)
            .ListIndex = .ListIndex + 2
            .RemoveItem .ListIndex - 2
        End If
    End With
End Sub
Private Sub ��ư_����_Click()
    With ����Ʈ_����
        'MultiSelect�� Single�� ��� _
            ���� �� .ListIndex (Ȥ�� .ListIndex-1)�� ���õǾ��ִ� ���� _
            ���߻����� �����ϱ� ������ _
            ������ �и���
        If .MultiSelect = fmMultiSelectSingle Then
            If .ListIndex >= 0 Then
                .RemoveItem (.ListIndex)
            End If
            
        Else
            For i = .ListCount - 1 To 0 Step -1
                If .Selected(i) Then
                    .RemoveItem i
                End If
            Next
        End If
    End With
End Sub

Private Sub ��ư_����_Click()
    If Me.����Ʈ_����.ListCount < 1 Then
        'If MsgBox("�ƹ� �͵� ���õ��� �ʾҽ��ϴ�. �����Ͻðڽ��ϱ�?", vbYesNo + vbDefaultButton2 + vbExclamation) = vbYes Then
        '    Unload Me
        'End If
        
        MsgBox "�����̵带 �������ּ���.", vbOKOnly + vbExclamation
        Exit Sub
    End If
    
    Load ��_�ٷΰ����ư�����2
    With ��_�ٷΰ����ư�����2
        With .����Ʈ_����
            For i = 0 To Me.����Ʈ_����.ListCount - 1
                .AddItem (Me.����Ʈ_����.List(i, 0))
                .List(i, 1) = Me.����Ʈ_����.List(i, 1)
                .List(i, 2) = ��Ʈ��ȯ(.List(i, 1))
            Next
        End With
        .���庯��_��󱸿��ε���.Caption = Me.���庯��_��󱸿��ε���.Caption
        
        Me.Hide
        .Show vbModal
    End With
    Unload Me
End Sub
Private Sub ��ư_���_Click()
    If MsgBox("������ ����Ͻðڽ��ϱ�?", vbOKCancel) = vbOK Then
        Unload Me
    End If
End Sub

Private Function ��Ʈ��ȯ(ByVal s As String)
    'ù �ܾ� �̴ϼ� ����
    If InStr(1, s, "Verse") Then
        s = Replace(s, "Verse", "v")
    ElseIf InStr(1, s, "Pre-chorus") Then
        s = Replace(s, "Pre-chorus", "p")
    ElseIf InStr(1, s, "Chorus") Then
        s = Replace(s, "Chorus", "c")
    ElseIf InStr(1, s, "Bridge") Then
        s = Replace(s, "Bridge", "b")
    ElseIf InStr(1, s, "fin") Then
        s = Replace(s, "fin", "f")
    End If
    
    '"(1/2)" ���� ������ ������ �����ϱ� ���� "(" �ڷδ� ����
    If InStr(1, s, "(") Then
        s = Split(s, "(", 2)(0)
    End If
    
    s = Replace(s, " ", "")
    
    ��Ʈ��ȯ = s
    Exit Function
End Function

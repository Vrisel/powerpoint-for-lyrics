Attribute VB_Name = "ImportText"
Sub �ؽ�Ʈto�����̵�()
    '�ؽ�Ʈ������ ���� �о���̸� �������� ���� ����� �� ã�Ƽ�;;
    '�׳� UserForm���ٰ� �����ڱ�... -_-;;;
    Load ufImportText
    With ufImportText
        .StoredParam.Caption = ""
        .Show vbModal
    End With
End Sub
Sub InputLyrics(oSl As Slide, s As String)
    Dim SplText() As String
    SplText() = Split(s, vbCrLf & "&&" & vbCrLf)
    
    '�������̶� Ư�� ��¼���� ���
    If Left(SplText(0), 1) = "[" Then
        '
    '�Ϲ����� ���
    Else
        '���� �Է�
        oSl.Shapes(2).TextFrame.TextRange.Text = SplText(0)
        '��Ʈ �Է� (������)
        If UBound(SplText) = 1 Then
            Dim oSh As Shape
            For Each oSh In oSl.NotesPage.Shapes
                '�Ʒ� If�� �� �ʿ����� �𸣰����� ������ ������.. ��ư ���� �ε���3 �ε�
                If oSh.PlaceholderFormat.Type = ppPlaceholderBody Then
                    oSh.TextFrame.TextRange.Text = SplText(1)
                    Exit Sub
                End If
            Next oSh
        End If
    End If
End Sub

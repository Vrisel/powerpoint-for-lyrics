Attribute VB_Name = "���系������"
Sub ���系������()
    Set ���罽���̵� = ActivePresentation.Windows.Item(1).Selection.SlideRange.Item(1)
    Dim �����ε���, ����ù�����̵�, ��������, �����������̵� As Integer
    
    �����ε��� = ���罽���̵�.sectionIndex
    ����ù�����̵� = ActivePresentation.SectionProperties.FirstSlide(�����ε���)
    �������� = ActivePresentation.SectionProperties.SlidesCount(�����ε���) - 1
        '������ -1�� �󽽶��̵忡 ���� ��
    �����������̵� = ����ù�����̵� + �������� - 1
    
    Dim ����, �ؽ�Ʈ, �����̵屸��, ��Ʈ���� As String
        ���� = ActivePresentation.SectionProperties.Name(�����ε���)
        �����̵屸�� = vbCrLf & "//" & vbCrLf
        ��Ʈ���� = vbCrLf & Chr(38) & Chr(38) & vbCrLf
    Dim i As Integer
    Dim �����̵� As Slide
    Dim ���� As Shape
    For i = ����ù�����̵� To �����������̵�
        Set �����̵� = ActivePresentation.Slides(i)
        
        Dim ���� As String
        ���� = ""
        For Each ���� In �����̵�.Shapes
            With ����
                If .Type = msoPlaceholder _
                And .PlaceholderFormat.Type = ppPlaceholderObject _
                And .TextFrame.HasText Then
                    ���� = .TextFrame.TextRange.Text
                    Exit For
                End If
            End With
        Next ����
        
        Dim ��Ʈ As String
        ��Ʈ = ""
        For Each ���� In �����̵�.NotesPage.Shapes
            With ����
                If .Type = msoPlaceholder _
                And .PlaceholderFormat.Type = ppPlaceholderBody _
                And .TextFrame.HasText Then
                    ��Ʈ = .TextFrame.TextRange.Text
                    Exit For
                End If
            End With
        Next ����
        
        If i > ����ù�����̵� Then
            �ؽ�Ʈ = �ؽ�Ʈ & �����̵屸��
        End If
        
        �ؽ�Ʈ = �ؽ�Ʈ & ����
        
        If ��Ʈ <> "" Then
            �ؽ�Ʈ = �ؽ�Ʈ & ��Ʈ���� & Split(��Ʈ, (vbCrLf & "["), 2)(0)
        End If
        
        If i = �����������̵� _
        And InStr(��Ʈ, (vbCrLf & "[")) Then
            �ؽ�Ʈ = �ؽ�Ʈ & ��Ʈ���� & "[" & Split(��Ʈ, (vbCrLf & "["), 2)(1)
        End If
    Next
    
    Dim ���ϰ�� As String
	'msoFileDialogSaveAs�δ� txt������ ������ ���� ��� ��ȸ��
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "������ ������ ������ �������ּ���."
        '.Filters.Clear
        '.Filters.Add "�ؽ�Ʈ ����", "*.txt"
        '.InitialFileName = ���� & ".txt"
        
        If .Show = True Then
            ���ϰ�� = .SelectedItems(1)
            
            Dim �����̸� As String
�����̸��Է�:
            �����̸� = InputBox("������ ���� �̸��� �Է����ּ���.", , ����)
            If �����̸� = "" Or LCase(�����̸�) = ".txt" Then
                Select Case MsgBox(("�����̸��� �Էµ��� �ʾҽ��ϴ�." & vbCrLf & "�ٽ� �Է��Ͻðڽ��ϱ�?"), vbYesNo)
                    Case vbYes
                        GoTo �����̸��Է�
                    Case vbNo
                        Exit Sub
                End Select
            End If
            
            If LCase(Right(�����̸�, 4)) <> ".txt" Then
                �����̸� = �����̸� & ".txt"
            End If
            
        Else
            Exit Sub
        End If
    End With
    
    With CreateObject("ADODB.Stream")
        .Open
        .Type = 2 'adTypeText
        .Charset = "UTF-8"
        .LineSeparator = -1 'adCrLf
        .WriteText �ؽ�Ʈ
        .SaveToFile (���ϰ�� & Chr(92) & �����̸�)
        .Close
    End With
End Sub

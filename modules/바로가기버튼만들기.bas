Attribute VB_Name = "�ٷΰ����ư�����"
Sub �ٷΰ����ư�����()
'���������� �۵��ߴ� ����Ű �Ҵ� ����� �Ʒ��� ����
'Sub Form_KeyDown(KeyCode As Integer, Ctrl As Integer, Alt As Integer)

    Set ���罽���̵� = ActivePresentation.Windows.Item(1).Selection.SlideRange.Item(1)
    Dim �����ε��� As Long
    Dim �������� As Long
    
    �����ε��� = ActivePresentation.SectionProperties.FirstSlide(���罽���̵�.sectionIndex)
    �������� = ActivePresentation.SectionProperties.SlidesCount(���罽���̵�.sectionIndex) - 1
    
    'â ����
    Load ��_�ٷΰ����ư�����1
    With ��_�ٷΰ����ư�����1
        '���� �� �����̵帶��
        For i = 0 To �������� - 1
            '��Ͽ� �׸� �߰�
            .����Ʈ_�����̵�.AddItem i + 1
            
            '�� �׸� �����̵��Ʈ ���� �߰�
            Dim ���� As Shape
            For Each ���� In ActivePresentation.Slides(�����ε��� + i).NotesPage.Shapes
                If ����.PlaceholderFormat.Type = ppPlaceholderBody Then
                    If ����.TextFrame.HasText Then
                        .����Ʈ_�����̵�.List(i, 1) = Split(����.TextFrame.TextRange.Text, vbCr, 2)(0)
                    Else
                        .����Ʈ_�����̵�.List(i, 1) = ""
                    End If
                End If
            Next ����
        Next
        
        '�Ű����� ���� (...)
        .���庯��_��󱸿��ε���.Caption = ���罽���̵�.sectionIndex
        
        .Show vbModal
    End With
End Sub

Sub �ٷΰ���(���� As Shape)
    Dim �����ؽ�Ʈ, �����ε���, ����ù�����̵�, Ÿ�� As Integer
    
    �����ؽ�Ʈ = CInt(����.TextFrame2.TextRange)
    �����ε��� = SlideShowWindows(1).View.Slide.sectionIndex
    ����ù�����̵� = ActivePresentation.SectionProperties.FirstSlide(�����ε���)
    Ÿ�� = ����ù�����̵� + �����ؽ�Ʈ - 1
    
    SlideShowWindows(1).View.GotoSlide Ÿ��
End Sub

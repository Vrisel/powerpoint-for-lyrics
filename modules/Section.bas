Attribute VB_Name = "Section"
Sub SectionNthClick(oShape As Shape)
    '���� �����̵�� â����.�����̵� �̵�: (���� ù �����̵��� �ε���(���� �����̵尡 ���Ե� ������ �ε���) + ������ ���ִ� ���� - 1)
    SlideShowWindows(1).View.GotoSlide (ActivePresentation.SectionProperties.FirstSlide(SlideShowWindows(1).View.Slide.sectionIndex) + oShape.TextFrame2.TextRange - 1)
End Sub

Sub ��ư�����()
'Sub Form_KeyDown(KeyCode As Integer, Ctrl As Integer, Alt As Integer)
    Set currentSlide = ActivePresentation.Windows.Item(1).Selection.SlideRange.Item(1)
    Dim secLBound As Long
    Dim secLen As Long
    
    secLBound = ActivePresentation.SectionProperties.FirstSlide(currentSlide.sectionIndex)
    secLen = ActivePresentation.SectionProperties.SlidesCount(currentSlide.sectionIndex) - 1
    
    Load UserForm1
    With UserForm1
        For i = 0 To secLen - 1
            .ReadListBox.AddItem i + 1
            
            Dim oSh As Shape
            For Each oSh In ActivePresentation.Slides(secLBound + i).NotesPage.Shapes
                If oSh.PlaceholderFormat.Type = ppPlaceholderBody Then
                    If oSh.TextFrame.HasText Then
                        .ReadListBox.List(i, 1) = Split(oSh.TextFrame.TextRange.Text, vbCr, 2)(0)
                    Else
                        .ReadListBox.List(i, 1) = ""
                    End If
                End If
                'If oSh.PlaceholderFormat.Type = ppPlaceholderBody Then
                '    MsgBox oSh.TextFrame.TextRange.Text
                'End If
            Next oSh
        Next
        .StoredParam.Caption = currentSlide.sectionIndex
        
        .Show vbModal
    End With
End Sub

Sub iCBWN(ByVal oBu As Boolean, ByVal oSl As Slide)
    '(�����) �ʱ�ȭ
    'Ȥ�� ���� ������ ���¸� �ľ��ؼ� ����� ���ڳİ� �ǹ��簡..
    '�� �����ϸ�, �̹� �ִ� ���¿��� ����..�� �����ڱ���. ���� �ٸ��� ��¿..
    If Not oBu Then
        Exit Sub
    End If
    
    Dim oSh As Shape
DelAgain:
    For Each oSh In oSl.Shapes
        If oSh.Type <> 14 And oSh.Type <> 17 Then '14�� ����/����, 17�� �ؽ�Ʈ����
            oSh.Delete
        End If
    Next oSh
    
    For Each oSh In oSl.Shapes
        If oSh.Type <> 14 And oSh.Type <> 17 Then '14�� ����/����, 17�� �ؽ�Ʈ����
            GoTo DelAgain
        End If
    Next oSh
    
    For Each oSh In oSl.NotesPage.Shapes
        If oSh.PlaceholderFormat.Type = ppPlaceholderBody Then
            With oSh.TextFrame.TextRange
                If .Parent.HasText Then
                    .Text = Split(.Text, "[", 2)(0) 'RTrim �� �ʿ䰡..?
                End If
            End With
        End If
    Next oSh
End Sub
Sub CreatButtons(ByVal n As Integer, ByRef s, ByVal oSl As Slide)
    With oSl
        Dim w As Single
        Dim h As Single
        w = .Parent.SlideMaster.width
        h = .Parent.SlideMaster.height
        For j = 0 To n - 1
            With .Shapes.AddShape(msoShapeRectangle, j * w / n, h / 2, w / n, h / 2 + 30)
                .Fill.Transparency = 1
                .Line.Visible = msoFalse
                .TextFrame2.VerticalAnchor = msoAnchorBottom
                .TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
                .TextFrame.TextRange.Text = s(j)
                .ActionSettings(ppMouseClick).Action = ppActionRunMacro
                .ActionSettings(ppMouseClick).Run = "SectionNthClick"
            End With
        Next
    End With
End Sub
Sub EditNote(ByVal s As String, ByVal oSl As Slide)
    Dim oSh As Shape
    For Each oSh In oSl.NotesPage.Shapes
        '�Ʒ� If�� �� �ʿ����� �𸣰����� ������ ������.. ��ư ���� �ε���3 �ε�
        If oSh.PlaceholderFormat.Type = ppPlaceholderBody Then
            With oSh.TextFrame.TextRange
                If .Parent.HasText Then
                    If Not Right(.Text, 1) = vbCr Then
                        .InsertAfter vbCr
                    End If
                End If
                .InsertAfter s
            End With
            Exit Sub
        End If
    Next oSh
End Sub

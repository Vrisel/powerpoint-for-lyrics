VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ��_�ٷΰ����ư�����2 
   Caption         =   "��ư ���� ����"
   ClientHeight    =   3020
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   4560
   OleObjectBlob   =   "��_�ٷΰ����ư�����2.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "��_�ٷΰ����ư�����2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    ����Ʈ_����.Clear
    �ؽ�Ʈ_����.Text = ""
End Sub
Private Sub ��ư_����_Click()
    With ����Ʈ_����
        If .ListIndex < 0 Then
            MsgBox "������ ������ �����̵带 �������ּ���.", vbOKOnly + vbExclamation
        Else
            .List(.ListIndex, 2) = �ؽ�Ʈ_����.Text
            .ListIndex = -1
            �ؽ�Ʈ_����.Text = ""
        End If
        
        .SetFocus
    End With
End Sub
Private Sub ��ư_Ȯ��_Click()
'���� ���� ����
    Dim ����ù�ε��� As Long
    Dim �������ε��� As Long
    Dim �������� As Long
    
    ����ù�ε��� = ActivePresentation.SectionProperties.FirstSlide(���庯��_��󱸿��ε���.Caption)
    �������ε��� = ActivePresentation.SectionProperties.SlidesCount(���庯��_��󱸿��ε���.Caption) + ����ù�ε��� - 2
    �������� = �������ε��� - ����ù�ε��� + 1
    
'��ư ����
    Dim ��ư���� As Integer
    ��ư���� = ����Ʈ_����.ListCount
    
'index�� array��
    Dim �ٷΰ��ε����迭() As String
    ReDim �ٷΰ��ε����迭(��ư���� - 1)
    For i = 0 To UBound(�ٷΰ��ε����迭)
        �ٷΰ��ε����迭(i) = ����Ʈ_����.List(i, 0)
    Next
    
'String ����
    Dim �����̵��Ʈ As String
    For i = 1 To ��ư����
        �����̵��Ʈ = �����̵��Ʈ & "[" & ����Ʈ_����.List(i - 1, 2) & "]"
        If i <> ��ư���� Then
            �����̵��Ʈ = �����̵��Ʈ & " "
        End If
    Next
    
'Ȯ�� ����
    If MsgBox("�ش� ���� (�����̵� " & ����ù�ε��� & "���� " & �������ε��� & "����)��" & vbCr _
            & "�� " & ��ư���� & "���� ��ư�� �����ϰ�" & vbCr _
            & "�����̵� ��Ʈ�� �Ʒ� ������ �߰��Ͻðڽ��ϱ�?" & vbCr _
            & ": " & �����̵��Ʈ, vbOKCancel + vbQuestion) = vbCancel Then
        Exit Sub
    End If
    
'���� �� �� Slide�� ���� ��ư�� ��Ʈ ����
    'Slides.Range���� Ȱ���� �迭 ����
    Dim �����̵�迭() As Integer
    ReDim �����̵�迭(1 To ��������)
    For i = 1 To ��������
        �����̵�迭(i) = ActivePresentation.Slides(����ù�ε��� + i - 1).SlideIndex
    Next
    '�ʱ�ȭ
    Dim �ʱ�ȭ���� As Boolean
    If MsgBox("�ʱ�ȭ �� �����Ͻðڽ��ϱ�?" & vbCr _
            & "(������: �ؽ�Ʈ ���ڸ� ������ ��� ������ ������ϴ�.)" _
            , vbYesNo + vbQuestion) = vbYes Then
        �ʱ�ȭ���� = True
    Else
        �ʱ�ȭ���� = False
    End If
    
'����
    For Each �����̵� In ActivePresentation.Slides.Range(�����̵�迭)
        ��ư�������ʱ�ȭ �ʱ�ȭ����, �����̵�
        ��ư���� ��ư����, �ٷΰ��ε����迭(), �����̵�
        �����̵��Ʈ���� �����̵��Ʈ, �����̵�
    Next
    
'��
    MsgBox ("���������� �Ϸ��Ͽ����ϴ�.")
    Unload Me
End Sub
Private Sub ��ư_���_Click()
    Unload Me
End Sub

Private Sub ��ư�������ʱ�ȭ(ByVal �ʱ�ȭ���� As Boolean, ByVal �����̵� As Slide)
    '(�����) �ʱ�ȭ
    'Ȥ�� ���� ������ ���¸� �ľ��ؼ� ����� ���ڳİ� �ǹ��簡..
    '�� �����ϸ�, �̹� �ִ� ���¿��� ����..�� �����ڱ���. ���� �ٸ��� ��¿..
    If Not �ʱ�ȭ���� Then
        Exit Sub
    End If
    
    Dim ���� As Shape
DelAgain:
    For Each ���� In �����̵�.Shapes
        If ����.Type <> 14 And ����.Type <> 17 Then '14�� ����/����, 17�� �ؽ�Ʈ����
            ����.Delete
        End If
    Next ����
    
    For Each ���� In �����̵�.Shapes
        If ����.Type <> 14 And ����.Type <> 17 Then '14�� ����/����, 17�� �ؽ�Ʈ����
            GoTo DelAgain
        End If
    Next ����
    
    For Each ���� In �����̵�.NotesPage.Shapes
        If ����.PlaceholderFormat.Type = ppPlaceholderBody Then
            With ����.TextFrame.TextRange
                If .Parent.HasText Then
                    .Text = Split(.Text, "[", 2)(0) 'RTrim �� �ʿ䰡..?
                End If
            End With
        End If
    Next ����
End Sub
Private Sub ��ư����(ByVal n As Integer, ByRef �ٷΰ��ε����迭() As String, ByVal �����̵� As Slide)
    With �����̵�
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
                .TextFrame.TextRange.Text = �ٷΰ��ε����迭(j)
                .ActionSettings(ppMouseClick).Action = ppActionRunMacro
                .ActionSettings(ppMouseClick).Run = "�ٷΰ���"
            End With
        Next
    End With
End Sub
Private Sub �����̵��Ʈ����(ByVal ��Ʈ As String, ByVal �����̵� As Slide)
    Dim ���� As Shape
    '�Ʒ�ó�� For Each - If�� �ۿ� ������ �� ����. �ε���3 ���δ� ���� �Ұ�.
    For Each ���� In �����̵�.NotesPage.Shapes
        If ����.PlaceholderFormat.Type = ppPlaceholderBody Then
            With ����.TextFrame.TextRange
                If .Parent.HasText Then
                    If Not Right(.Text, 1) = vbCr Then
                        .InsertAfter vbCr
                    End If
                End If
                
                .InsertAfter ��Ʈ
            End With
            Exit Sub
            
        End If
    Next ����
End Sub


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
    �ؽ�Ʈ_����.text = ""
End Sub
Private Sub ��ư_����_Click()
    With ����Ʈ_����
        If .ListIndex < 0 Then
            MsgBox "������ ������ �����̵带 �������ּ���.", vbOKOnly + vbExclamation
        Else
            .List(.ListIndex, 2) = �ؽ�Ʈ_����.text
            .ListIndex = -1
            �ؽ�Ʈ_����.text = ""
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
    
    '�����̵� ��Ʈ�� �߰��� ���� ����
    Dim �����̵��Ʈ As String
    For i = 0 To ��ư���� - 1
        �����̵��Ʈ = �����̵��Ʈ & "[" & ����Ʈ_����.List(i, 2) & "]"
        If i < ��ư���� Then
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
    
    'Slides.Range���� Ȱ���� �迭 ����
    Dim �����̵�迭() As Integer
    ReDim �����̵�迭(��������)
    For i = 0 To �������� - 1
        �����̵�迭(i) = ActivePresentation.Slides(����ù�ε��� + i).SlideIndex
    Next
    
    ��ư�׳�Ʈ���� �����̵�迭(), �ٷΰ��ε����迭(), �����̵��Ʈ
End Sub
Private Sub ��ư_���_Click()
    If MsgBox("������ ����Ͻðڽ��ϱ�?", vbOKCancel) = vbOK Then
        Unload Me
    End If
End Sub

Private Sub ��ư�׳�Ʈ����(�����̵�迭() As Integer, ByRef �ٷΰ��ε����迭() As String, ���Թ��� As String)
    Dim �ʱ�ȭ���� As Boolean
    Select Case MsgBox("�ʱ�ȭ �� �����Ͻðڽ��ϱ�?" & vbCr _
            & "(������: �ؽ�Ʈ ���ڸ� ������ ��� ������ ������ϴ�.)" _
            , vbYesNoCancel + vbQuestion)
    Case vbCancel
        Exit Sub
    Case vbYes
        �ʱ�ȭ���� = True
    Case vbNo
        �ʱ�ȭ���� = False
    End Select
    
    'For�� ������ ��ư���� ��Ʈ����
    Dim �����̵� As Slide
    For Each �����̵� In ActivePresentation.Slides.Range(�����̵�迭)
        ��ư���� �����̵�, �ʱ�ȭ����, �ٷΰ��ε����迭()
        �����̵��Ʈ���� �����̵�, �ʱ�ȭ����, ���Թ���
    Next
    
    MsgBox ("���������� �Ϸ��Ͽ����ϴ�.")
    Unload Me
End Sub

Private Sub ��ư����(�����̵� As Slide, �ʱ�ȭ���� As Boolean, ByRef �ٷΰ��ε����迭() As String)
    '�ʱ�ȭ
    If �ʱ�ȭ���� = True Then
        Dim i As Integer
        For i = �����̵�.Shapes.Count To 1 Step -1
            With �����̵�.Shapes(i)
                If .Type <> 14 And .Type <> 17 Then '14�� ����/����, 17�� �ؽ�Ʈ����
                    .Delete
                End If
            End With
        Next i
    End If
    
    '��ư ����
    Dim ��ư���� As Integer
    ��ư���� = UBound(�ٷΰ��ε����迭) - LBound(�ٷΰ��ε����迭) + 1
    
    With �����̵�.Parent.SlideMaster
        Dim �����̵�w, �����̵�h As Single
        �����̵�w = .width
        �����̵�h = .height
    End With
    
    Dim ��ưw, ��ưh, ��ưt As Single
    ��ưw = �����̵�w / ��ư����
    ��ưh = �����̵�h / 2 + 30
    ��ưt = �����̵�h / 2
    
    For i = 0 To ��ư���� - 1
        With �����̵�.Shapes.AddShape(msoShapeRectangle, ��ưw * i, ��ưt, ��ưw, ��ưh)
            .Fill.Transparency = 1
            .Line.Visible = msoFalse
            .TextFrame2.VerticalAnchor = msoAnchorBottom
            .TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
            .TextFrame.TextRange.text = �ٷΰ��ε����迭(i)
            .ActionSettings(ppMouseClick).Action = ppActionRunMacro
            .ActionSettings(ppMouseClick).Run = "�ٷΰ���"
        End With
    Next
End Sub
Private Sub �����̵��Ʈ����(�����̵� As Slide, �ʱ�ȭ���� As Boolean, ���Թ��� As String)
    Dim ���� As Shape
    '�Ʒ�ó�� For Each - If�� �ۿ� ������ �� ����. �ε���3 ���δ� ���� �Ұ�.
    For Each ���� In �����̵�.NotesPage.Shapes
        If ����.PlaceholderFormat.Type = ppPlaceholderBody Then
            With ����.TextFrame.TextRange
                If .Parent.HasText = True And �ʱ�ȭ���� = True Then
                    .text = Split(.text, "[", 2)(0)
                    'RTrim �� �ʿ䰡..?
                End If
                
                If .Parent.HasText = True And Not Right(.text, 1) = vbCr Then
                    .InsertAfter vbCr
                End If
                
                .InsertAfter ���Թ���
            End With
            Exit Sub
        End If
    Next ����
End Sub

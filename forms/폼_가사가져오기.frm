VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ��_���簡������ 
   Caption         =   "���� �Է�"
   ClientHeight    =   3850
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   4560
   OleObjectBlob   =   "��_���簡������.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "��_���簡������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    �ؽ�Ʈ_����.Text = ""
    �ؽ�Ʈ_����.Text = ""
End Sub

Private Sub ��ư_����_Click()
    '�ؽ�Ʈ ���� �۾�
    Dim ����и�() As String
    ����и� = Split(�ؽ�Ʈ_����.Text, vbCrLf & "//" & vbCrLf)
    
    With Application.ActivePresentation.Slides
        '���� ������ ���� ���� ���� �ε��� ���
        Dim ������ġ As Integer
        ������ġ = .Count + 1
    
        '�����̵� ����
        For i = LBound(����и�) To UBound(����и�)
            '�����̵� �߰�, "���� �� ����" ���̾ƿ�(#2)����
            .AddSlide (.Count + 1), ActivePresentation.SlideMaster.CustomLayouts(2)
            
            �����Է� .Item(.Count), ����и�(i)
        Next
        
        '������ �� �����̵�
        .AddSlide (.Count + 1), ActivePresentation.SlideMaster.CustomLayouts(2)
    End With
    
    '���� ���� �� ���� �Է�
    ActivePresentation.SectionProperties.AddBeforeSlide (������ġ), �ؽ�Ʈ_����.Text
    ActivePresentation.Slides(������ġ).Shapes(1).TextFrame.TextRange.Text = �ؽ�Ʈ_����.Text
    
    Unload Me
End Sub
Private Sub ��ư_���_Click()
    Unload Me
End Sub

Private Sub �����Է�(�����̵� As Slide, ���� As String)
    Dim ����и�() As String
    ����и�() = Split(����, vbCrLf & "&&" & vbCrLf)
    
    '��ư ������ ����� ���������� ���
    If Left(����и�(0), 1) = "[" Then
        '�̱���
        
    '�Ϲ����� ���
    Else
        '���� �Է�
        �����̵�.Shapes(2).TextFrame.TextRange.Text = ����и�(0)
        
        '��Ʈ �Է� (������)
        If UBound(����и�) = 1 Then
            Dim ���� As Shape
            
            '�Ʒ�ó�� For Each - If�� �ۿ� ������ �� ����. �ε���3 ���δ� ���� �Ұ�.
            For Each ���� In �����̵�.NotesPage.Shapes
                If ����.PlaceholderFormat.Type = ppPlaceholderBody Then
                    ����.TextFrame.TextRange.Text = ����и�(1)
                    Exit Sub
                End If
            Next ����
        End If
    End If
End Sub


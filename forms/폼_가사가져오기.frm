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
    �ؽ�Ʈ_����.text = ""
    �ؽ�Ʈ_����.text = ""
End Sub

Private Sub ��ư_���Ͽ���_Click()
    ���Ͽ���
End Sub
Private Sub ��ư_Ȯ��_Click()
    '�ؽ�Ʈ ���� �۾�
    Dim ����и�() As String
    ����и� = Split(�ؽ�Ʈ_����.text, vbCrLf & "//" & vbCrLf)
    
    With ActivePresentation.Slides
        '���� ������ ���� ���� ���� �ε��� ���
        Dim ������ġ As Integer
        ������ġ = .Count + 1
    
        '�����̵� ����
        Dim ���� As Variant
        For Each ���� In ����и�
            �����Է� ����
        Next ����
        
        '������ �� �����̵� �߰�
        .AddSlide (.Count + 1), ActivePresentation.SlideMaster.CustomLayouts(2)
    End With
    
    '���� ���� �� ���� �Է�
    ActivePresentation.SectionProperties.AddBeforeSlide ������ġ, �ؽ�Ʈ_����.text
    ActivePresentation.Slides(������ġ).Shapes(1).TextFrame.TextRange.text = �ؽ�Ʈ_����.text
    
    Unload Me
End Sub
Private Sub ��ư_���_Click()
    Unload Me
End Sub

Private Sub ���Ͽ���()
    Dim ����, ���� As String
    ���� = ""
    ���� = ""
    
    '���� ����
    Dim ���ϰ�� As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Clear
        .Filters.Add "�ؽ�Ʈ ����", "*.txt"
        
        If .Show = True Then
            ���ϰ�� = .SelectedItems(1)
            
            '���Խ� ����� ����(T) > ����(R)... > Microsoft VBScript Regular Expression 5.5 �߰�
            With CreateObject("VBScript.RegExp")
                'VBA���� �齽���� �̽��������� �� �Ǿ� Chr(92)�� ��ü
                '\\: �齽���� �̽�������
                '([^\\]+): �齽���ð� �ƴ� ���� 1�� �̻�, ĸó
                '\.txt$: ".txt"�� ������ Ȯ����
                .Pattern = Chr(92) & Chr(92) & _
                            "([^" & Chr(92) & Chr(92) & "]+)" & _
                            Chr(92) & ".txt$"
                .IgnoreCase = True
                
                ���� = .Execute(���ϰ��).Item(0).SubMatches(0)
            End With
            
            With CreateObject("ADODB.Stream")
                .Open
                .Type = 2 'adTypeText
                .Charset = "UTF-8"
                .LoadFromFile = ���ϰ��
                .LineSeparator = 13 'adCR
                
                Do Until .EOS
                    ���� = ���� & .ReadText(-2)
                Loop
            End With
    
            Me.�ؽ�Ʈ_����.text = ����
            Me.�ؽ�Ʈ_����.text = ����
        End If
    End With
End Sub

Private Sub �����Է�(���� As Variant)
    Dim ����и�() As String
    ����и� = Split(����, vbCrLf & "&&" & vbCrLf)
    
    '��ư ������ ����� ���������� ���
    If Left(����и�(0), 1) = "[" Then
        '�̱���
        
    '�Ϲ����� ���
    Else
        With ActivePresentation.Slides
            '�����̵� ����
            '���̾ƿ��� "���� �� ����(#2)"
            .AddSlide (.Count + 1), ActivePresentation.SlideMaster.CustomLayouts(2)
            
            With .Item(.Count)
                '����ĭ(#2)�� ���� �Է�
                .Shapes(2).TextFrame.TextRange.text = ����и�(0)
                
                '�����̵��Ʈ�� ��Ʈ �Է� (������)
                If UBound(����и�) = 1 Then
                    Dim ���� As Shape
                    
                    '�Ʒ�ó�� For Each - If�� �ۿ� ������ �� ����. �ε���3 ���δ� ���� �Ұ�.
                    For Each ���� In .NotesPage.Shapes
                        If ����.PlaceholderFormat.Type = ppPlaceholderBody Then
                            ����.TextFrame.TextRange.text = ����и�(1)
                            Exit Sub
                        End If
                    Next ����
                End If
            End With
        End With
    End If
End Sub


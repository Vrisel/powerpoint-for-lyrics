VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufImportText 
   Caption         =   "���� �Է�"
   ClientHeight    =   3855
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   4560
   OleObjectBlob   =   "ufImportText.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "ufImportText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    tbTitle.Text = ""
    tbLyrics.Text = ""
End Sub

Private Sub btnNext_Click()
    '�ؽ�Ʈ ���� �۾�
    Dim SplittedLyrics() As String
    SplittedLyrics() = Split(tbLyrics.Text, vbCrLf & "//" & vbCrLf)
    'With Application.ActivePresentation.Slides
    
    '���� ������ ���� �����̵� ī��Ʈ
    Dim iSlideCount As Integer
    iSlideCount = ActivePresentation.Slides.Count
    
    '�����̵� ����
    With ActivePresentation.Slides
        For i = LBound(SplittedLyrics) To UBound(SplittedLyrics)
            '�����̵� �߰�, "���� �� ����" ���̾ƿ�����
            .AddSlide (.Count + 1), ActivePresentation.SlideMaster.CustomLayouts(2)
            
            InputLyrics .Item(.Count), SplittedLyrics(i)
        Next
        
        '������ �� �����̵�
        .AddSlide (.Count + 1), ActivePresentation.SlideMaster.CustomLayouts(2)
    End With
    
    ''���� ���� �� ���� �Է�
    'Dim sTitle As String
    'sTitle = InputBox("������ �Է����ּ���.")
    ActivePresentation.SectionProperties.AddBeforeSlide (iSlideCount + 1), tbTitle.Text
    ActivePresentation.Slides(iSlideCount + 1).Shapes(1).TextFrame.TextRange.Text = tbTitle.Text
    
    Unload Me
End Sub
Private Sub btnCancel_Click()
    Unload Me
End Sub

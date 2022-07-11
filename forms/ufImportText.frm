VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufImportText 
   Caption         =   "가사 입력"
   ClientHeight    =   3855
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   4560
   OleObjectBlob   =   "ufImportText.frx":0000
   StartUpPosition =   1  '소유자 가운데
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
    '텍스트 가공 작업
    Dim SplittedLyrics() As String
    SplittedLyrics() = Split(tbLyrics.Text, vbCrLf & "//" & vbCrLf)
    'With Application.ActivePresentation.Slides
    
    '섹션 생성을 위해 슬라이드 카운트
    Dim iSlideCount As Integer
    iSlideCount = ActivePresentation.Slides.Count
    
    '슬라이드 생성
    With ActivePresentation.Slides
        For i = LBound(SplittedLyrics) To UBound(SplittedLyrics)
            '슬라이드 추가, "제목 및 내용" 레이아웃으로
            .AddSlide (.Count + 1), ActivePresentation.SlideMaster.CustomLayouts(2)
            
            InputLyrics .Item(.Count), SplittedLyrics(i)
        Next
        
        '마지막 빈 슬라이드
        .AddSlide (.Count + 1), ActivePresentation.SlideMaster.CustomLayouts(2)
    End With
    
    ''섹션 생성 및 제목 입력
    'Dim sTitle As String
    'sTitle = InputBox("제목을 입력해주세요.")
    ActivePresentation.SectionProperties.AddBeforeSlide (iSlideCount + 1), tbTitle.Text
    ActivePresentation.Slides(iSlideCount + 1).Shapes(1).TextFrame.TextRange.Text = tbTitle.Text
    
    Unload Me
End Sub
Private Sub btnCancel_Click()
    Unload Me
End Sub

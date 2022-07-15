VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 폼_가사가져오기 
   Caption         =   "가사 입력"
   ClientHeight    =   3850
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   4560
   OleObjectBlob   =   "폼_가사가져오기.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "폼_가사가져오기"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    텍스트_제목.Text = ""
    텍스트_가사.Text = ""
End Sub

Private Sub 버튼_다음_Click()
    '텍스트 가공 작업
    Dim 가사분리() As String
    가사분리 = Split(텍스트_가사.Text, vbCrLf & "//" & vbCrLf)
    
    With Application.ActivePresentation.Slides
        '구역 생성을 위해 최초 삽입 인덱스 기억
        Dim 삽입위치 As Integer
        삽입위치 = .Count + 1
    
        '슬라이드 생성
        For i = LBound(가사분리) To UBound(가사분리)
            '슬라이드 추가, "제목 및 내용" 레이아웃(#2)으로
            .AddSlide (.Count + 1), ActivePresentation.SlideMaster.CustomLayouts(2)
            
            가사입력 .Item(.Count), 가사분리(i)
        Next
        
        '마지막 빈 슬라이드
        .AddSlide (.Count + 1), ActivePresentation.SlideMaster.CustomLayouts(2)
    End With
    
    '구역 생성 및 제목 입력
    ActivePresentation.SectionProperties.AddBeforeSlide (삽입위치), 텍스트_제목.Text
    ActivePresentation.Slides(삽입위치).Shapes(1).TextFrame.TextRange.Text = 텍스트_제목.Text
    
    Unload Me
End Sub
Private Sub 버튼_취소_Click()
    Unload Me
End Sub

Private Sub 가사입력(슬라이드 As Slide, 가사 As String)
    Dim 가사분리() As String
    가사분리() = Split(가사, vbCrLf & "&&" & vbCrLf)
    
    '버튼 정보가 저장된 마지막줄의 경우
    If Left(가사분리(0), 1) = "[" Then
        '미구현
        
    '일반적인 경우
    Else
        '내용 입력
        슬라이드.Shapes(2).TextFrame.TextRange.Text = 가사분리(0)
        
        '노트 입력 (있으면)
        If UBound(가사분리) = 1 Then
            Dim 도형 As Shape
            
            '아래처럼 For Each - If로 밖에 접근할 수 없음. 인덱스3 으로는 접근 불가.
            For Each 도형 In 슬라이드.NotesPage.Shapes
                If 도형.PlaceholderFormat.Type = ppPlaceholderBody Then
                    도형.TextFrame.TextRange.Text = 가사분리(1)
                    Exit Sub
                End If
            Next 도형
        End If
    End If
End Sub


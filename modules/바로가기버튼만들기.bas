Attribute VB_Name = "바로가기버튼만들기"
Sub 바로가기버튼만들기()
'구버전에서 작동했던 단축키 할당 방법은 아래와 같음
'Sub Form_KeyDown(KeyCode As Integer, Ctrl As Integer, Alt As Integer)

    Set 현재슬라이드 = ActivePresentation.Windows.Item(1).Selection.SlideRange.Item(1)
    Dim 구역인덱스 As Long
    Dim 구역길이 As Long
    
    구역인덱스 = ActivePresentation.SectionProperties.FirstSlide(현재슬라이드.sectionIndex)
    구역길이 = ActivePresentation.SectionProperties.SlidesCount(현재슬라이드.sectionIndex) - 1
    
    '창 띄우기
    Load 폼_바로가기버튼만들기1
    With 폼_바로가기버튼만들기1
        '구역 내 슬라이드마다
        For i = 0 To 구역길이 - 1
            '목록에 항목 추가
            .리스트_슬라이드.AddItem i + 1
            
            '각 항목에 슬라이드노트 내용 추가
            Dim 도형 As Shape
            For Each 도형 In ActivePresentation.Slides(구역인덱스 + i).NotesPage.Shapes
                If 도형.PlaceholderFormat.Type = ppPlaceholderBody Then
                    If 도형.TextFrame.HasText Then
                        .리스트_슬라이드.List(i, 1) = Split(도형.TextFrame.TextRange.Text, vbCr, 2)(0)
                    Else
                        .리스트_슬라이드.List(i, 1) = ""
                    End If
                End If
            Next 도형
        Next
        
        '매개변수 저장 (...)
        .저장변수_대상구역인덱스.Caption = 현재슬라이드.sectionIndex
        
        .Show vbModal
    End With
End Sub

Sub 바로가기(도형 As Shape)
    Dim 도형텍스트, 구역인덱스, 구역첫슬라이드, 타겟 As Integer
    
    도형텍스트 = CInt(도형.TextFrame2.TextRange)
    구역인덱스 = SlideShowWindows(1).View.Slide.sectionIndex
    구역첫슬라이드 = ActivePresentation.SectionProperties.FirstSlide(구역인덱스)
    타겟 = 구역첫슬라이드 + 도형텍스트 - 1
    
    SlideShowWindows(1).View.GotoSlide 타겟
End Sub

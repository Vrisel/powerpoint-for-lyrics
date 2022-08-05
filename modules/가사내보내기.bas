Attribute VB_Name = "가사내보내기"
Sub 가사내보내기()
    Set 현재슬라이드 = ActivePresentation.Windows.Item(1).Selection.SlideRange.Item(1)
    Dim 구역인덱스, 구역첫슬라이드, 구역길이, 구역막슬라이드 As Integer
    
    구역인덱스 = 현재슬라이드.sectionIndex
    구역첫슬라이드 = ActivePresentation.SectionProperties.FirstSlide(구역인덱스)
    구역길이 = ActivePresentation.SectionProperties.SlidesCount(구역인덱스) - 1
        '마지막 -1은 빈슬라이드에 대한 것
    구역막슬라이드 = 구역첫슬라이드 + 구역길이 - 1
    
    Dim 제목, 텍스트, 슬라이드구분, 노트구분 As String
        제목 = ActivePresentation.SectionProperties.Name(구역인덱스)
        슬라이드구분 = vbCrLf & "//" & vbCrLf
        노트구분 = vbCrLf & Chr(38) & Chr(38) & vbCrLf
    Dim i As Integer
    Dim 슬라이드 As Slide
    Dim 도형 As Shape
    For i = 구역첫슬라이드 To 구역막슬라이드
        Set 슬라이드 = ActivePresentation.Slides(i)
        
        Dim 가사 As String
        가사 = ""
        For Each 도형 In 슬라이드.Shapes
            With 도형
                If .Type = msoPlaceholder _
                And .PlaceholderFormat.Type = ppPlaceholderObject _
                And .TextFrame.HasText Then
                    가사 = .TextFrame.TextRange.Text
                    Exit For
                End If
            End With
        Next 도형
        
        Dim 노트 As String
        노트 = ""
        For Each 도형 In 슬라이드.NotesPage.Shapes
            With 도형
                If .Type = msoPlaceholder _
                And .PlaceholderFormat.Type = ppPlaceholderBody _
                And .TextFrame.HasText Then
                    노트 = .TextFrame.TextRange.Text
                    Exit For
                End If
            End With
        Next 도형
        
        If i > 구역첫슬라이드 Then
            텍스트 = 텍스트 & 슬라이드구분
        End If
        
        텍스트 = 텍스트 & 가사
        
        If 노트 <> "" Then
            텍스트 = 텍스트 & 노트구분 & Split(노트, (vbCrLf & "["), 2)(0)
        End If
        
        If i = 구역막슬라이드 _
        And InStr(노트, (vbCrLf & "[")) Then
            텍스트 = 텍스트 & 노트구분 & "[" & Split(노트, (vbCrLf & "["), 2)(1)
        End If
    Next
    
    Dim 파일경로 As String
	'msoFileDialogSaveAs로는 txt파일을 저장할 수가 없어서 우회함
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "파일을 저장할 폴더를 선택해주세요."
        '.Filters.Clear
        '.Filters.Add "텍스트 파일", "*.txt"
        '.InitialFileName = 제목 & ".txt"
        
        If .Show = True Then
            파일경로 = .SelectedItems(1)
            
            Dim 파일이름 As String
파일이름입력:
            파일이름 = InputBox("저장할 파일 이름을 입력해주세요.", , 제목)
            If 파일이름 = "" Or LCase(파일이름) = ".txt" Then
                Select Case MsgBox(("파일이름이 입력되지 않았습니다." & vbCrLf & "다시 입력하시겠습니까?"), vbYesNo)
                    Case vbYes
                        GoTo 파일이름입력
                    Case vbNo
                        Exit Sub
                End Select
            End If
            
            If LCase(Right(파일이름, 4)) <> ".txt" Then
                파일이름 = 파일이름 & ".txt"
            End If
            
        Else
            Exit Sub
        End If
    End With
    
    With CreateObject("ADODB.Stream")
        .Open
        .Type = 2 'adTypeText
        .Charset = "UTF-8"
        .LineSeparator = -1 'adCrLf
        .WriteText 텍스트
        .SaveToFile (파일경로 & Chr(92) & 파일이름)
        .Close
    End With
End Sub

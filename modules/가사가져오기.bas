Attribute VB_Name = "가사가져오기"
Sub 가사가져오기()
    '텍스트파일을 직접 읽어들이면 좋겠지만 아직 기능을 못 찾아서;;
    '그냥 UserForm에다가 때려박기... -_-;;;
    Load 폼_가사가져오기
    With 폼_가사가져오기
        .Show vbModal
    End With
End Sub

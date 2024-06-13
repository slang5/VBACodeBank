Attribute VB_Name = "DoAnFormatedID"
Function DoAnID(texte1, texte2, wantedlen1, wantedlen2):
    Dim T1, T2 As String
    T1 = WantedSize(texte1, wantedlen1)
    T2 = WantedSize(texte2, wantedlen2)
     DoAnID = T1 & "_" & T2
End Function

Function WantedSize(Texte, Taille):
    Dim output As String
    output = ""
    If Len(Texte) < Taille Then
        For I = 1 To Taille - Len(Texte)
            output = output & "0"
        Next I
        output = output & Texte
    ElseIf Len(Texte) > Taille Then
        output = Left(Texte, Taille)
    Else
        output = Texte
    End If
    
    WantedSize = output
End Function


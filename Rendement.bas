Attribute VB_Name = "Rendement"
Function Rendement(Fin, Debut)
    Rendement = (Fin - Debut) / Debut
End Function

Function RendementLn(Fin, Debut)
    RendementLn = WorksheetFunction.Ln(Fin / Debut)
    
End Function


Attribute VB_Name = "AnalyseStatQuick"
Function StatBasique(Data As Range, Isnumber As Integer) As Variant
    Dim retour(1 To 4) As Variant
    Dim mean As Double, ecarttype As Double, skewness As Double, kurtosis As Double
    
    mean = WorksheetFunction.Average(Data): ecarttype = WorksheetFunction.StDev_P(Data): skewness = WorksheetFunction.Skew_p(Data): kurtosis = WorksheetFunction.Kurt(Data)
    
    If (Isnumber = 1) Then
        retour(1) = "Moyenne"
        retour(2) = "Ecart type"
        retour(3) = "Skewness"
        retour(4) = "Kurtosis"
        
    Else
        retour(1) = mean
        retour(2) = ecarttype
        retour(3) = skewness
        retour(4) = kurtosis
    End If
    
    StatBasique = WorksheetFunction.Transpose(retour)
End Function

Function MatriceCovVar(Matrix As Range)

    Dim numbercolumn As Integer
    numbercolumn = WorksheetFunction.Count(WorksheetFunction.Index(Matrix, 1, 0))
    
    
    ' par symétrie le nombre de row est pareil que column
    Dim MatFinale() As Variant
    ReDim MatFinale(1 To numbercolumn, 1 To numbercolumn)
    For I = 1 To numbercolumn
        For j = 1 To numbercolumn
            MatFinale(I, j) = WorksheetFunction.Covariance_P(WorksheetFunction.Index(Matrix, 0, I), WorksheetFunction.Index(Matrix, 0, j))
        Next
    Next
    MatriceCovVar = MatFinale
End Function


Function MatriceCorr(Matrix As Range)

    Dim numbercolumn As Integer
    numbercolumn = WorksheetFunction.Count(WorksheetFunction.Index(Matrix, 1, 0))
    
    
    ' par symétrie le nombre de row est pareil que column
    Dim MatFinale() As Variant
    ReDim MatFinale(1 To numbercolumn, 1 To numbercolumn)
    For I = 1 To numbercolumn
        For j = 1 To numbercolumn
            MatFinale(I, j) = WorksheetFunction.Correl(WorksheetFunction.Index(Matrix, 0, I), WorksheetFunction.Index(Matrix, 0, j))
        Next
    Next
    MatriceCorr = MatFinale
End Function


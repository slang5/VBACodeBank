Attribute VB_Name = "MonteCarloCholeskyCorrelation"
Function VecteurGaussienCorrelationByCholesky(moy1, ect1, moy2, ect2, correlation As Double) As Variant
    Dim retour(1 To 2) As Variant
    Dim random(1 To 2) As Double
    
    random(1) = Rnd: random(2) = Rnd
    
    retour(1) = WorksheetFunction.Norm_S_Inv(random(1))
    
    retour(2) = retour(1) * correlation + Sqr(1 - correlation * correlation) * WorksheetFunction.Norm_S_Inv(random(2))
    
    retour(1) = retour(1) * ect1 + moy1
    retour(2) = retour(2) * ect2 + moy2
    
    VecteurGaussienCorrelationByCholesky = retour
    
End Function



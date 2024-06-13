Attribute VB_Name = "VaR"
Function VaRNormal(volume, moyenne, ecarttype, alpha_p_quantile)
    alpha_p_quantile = WorksheetFunction.Norm_S_Inv(1 - alpha_p_quantile)
    VaRNormal = -volume * (moyenne + ecarttype * alpha_p_quantile)
End Function

Function VarLogNormal(volume, moyenne, ecarttype, alpha_p_quantile)
    alpha_p_quantile = WorksheetFunction.Norm_S_Inv(1 - alpha_p_quantile)
    VarLogNormal = volume * (1 - Exp(moyenne + ecarttype * alpha_p_quantile))
End Function

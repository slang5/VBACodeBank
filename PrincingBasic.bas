Attribute VB_Name = "PrincingBasic"
'Sujet : Prix Options

Function BlackAndScholes(Spot, Strike, Maturite, TauxInteret, Volatilite, CallOrPut)
    Dim D1, D2, C As Double
    
    D1 = (WorksheetFunction.Ln(Spot / Strike) + (TauxInteret + (Volatilite * Volatilite) / 2) * Maturite) / (Volatilite * Sqr(Maturite))
    D2 = D1 - Volatilite * Sqr(Maturite)
    
    If (CallOrPut) = "Call" Then
        C = Spot * WorksheetFunction.Norm_S_Dist(D1, True) - WorksheetFunction.Norm_S_Dist(D2, True) * Strike * Exp(-TauxInteret * Maturite)
    Else
        C = -Spot * WorksheetFunction.Norm_S_Dist(-D1, True) + WorksheetFunction.Norm_S_Dist(-D2, True) * Strike * Exp(-TauxInteret * Maturite)
    
    End If
    BlackAndScholes = C
        
End Function

Function BinomialTree(Spot, Strike, Time, Vol, TauxSR, Instrument, Netapes) As Double
    Dim Deltatime, P, Gain, Binomiale As Double: Gain = 0
    Dim N As Integer
    
    Deltatime = Time / Netapes
    N = Netapes
    P = (Exp(TauxSR * Deltatime) - Exp(-Vol * Sqr(Deltatime))) / (Exp(Vol * Sqr(Deltatime)) - Exp(-Vol * Sqr(Deltatime)))
    Dim St As Double

    If N > 0 Then
        Dim I As Integer
        Dim Temp As Double
        For I = 0 To N
            St = Spot * (Exp(Vol * Sqr(Deltatime)) ^ I) * ((Exp(-Vol * Sqr(Deltatime))) ^ (N - I))
            If Instrument = "Call" Then
                Temp = WorksheetFunction.Max(0, St - Strike)
            Else
                Temp = WorksheetFunction.Max(0, Strike - St)
            End If
            If Temp > 0 Then
                Binomiale = WorksheetFunction.BinomDist(I, N, P, False)
                Gain = Gain + Binomiale * Temp
            End If
            
        Next I
        
        End If

    Gain = Gain * Exp(-TauxSR * Time)
    BinomialTree = Gain
End Function


Function Monte_Carlo(Spot, Strike, Time, Vol, TauxSR, Instrument, Rep) As Double

    Dim Prix, Norm, St, Gain As Double: Prix = 0: Norm = 0: St = 0: Gain = 0
    
    For I = 1 To Rep
        Norm = WorksheetFunction.Norm_S_Inv(Rnd())
        St = Spot * Exp((TauxSR - Vol * Vol / 2) * Time + (Norm * Vol * Sqr(Time)))
        If Instrument = "Call" Then
            Gain = Gain + WorksheetFunction.Max(St - Strike, 0) * Exp(-TauxSR * Time)
        ElseIf Instrument = "Put" Then
            Gain = Gain + WorksheetFunction.Max(Strike - St, 0) * Exp(-TauxSR * Time)
        End If
    Next I
    
    Prix = Gain / Rep
    Monte_Carlo = Prix
End Function

'Sujet : Options

Function EquityCall(St, k)
    EquityCall = WorksheetFunction.Max(0, St - k)
End Function

Function EquityPut(St, k)
    EquityPut = WorksheetFunction.Max(0, k - St)
End Function

'Sujet : Obligation remboursement in fine a taux fixe

Function BondFixedYield(Nominal, Remboursement, Maturite, TxCoupon, TxInteret, Rythme)
    Part1 = Nominal * TxCoupon / Rythme
    Part_2 = Remboursement * (1 + TxInteret / Rythme) ^ (-Maturite * Rythme)
    Part_3_1 = 1 - (1 + TxInteret / Rythme) ^ (-Maturite * Rythme)
    Part_3_2 = TxInteret / Rythme
    Part = Part1 * (Part_3_1 / Part_3_2) + Part_2

    BondFixedYield = Part
End Function


Option Explicit
Public Function Regulateur(ByRef Kc As Double, ByRef Err As Double, ByRef ErrInt As Double, ByRef ErrDer As Double, ByRef tau_i As Double, ByRef tau_d As Double, ByRef MaxOpen As Double, ByRef MinOpen As Double) As Double
Dim Us As Double
If (Kc * (Err + (ErrInt / tau_i) + (ErrDer * tau_d))) > MaxOpen Then
    Us = MaxOpen
    Else
        If (Kc * (Err + (ErrInt / tau_i) + (ErrDer * tau_d))) < MinOpen Then
            Us = MinOpen
            Else: Us = Kc * (Err + (ErrInt / tau_i) + (ErrDer * tau_d))
        End If
End If
Regulateur = Us

End Function

Public Function Actionneur(ByRef tauValve As Double, ByRef Kvalve As Double, ByRef Regulateur As Double, ByRef deltat As Double, ByRef previousF As Double) As Double
Dim Fs As Double
If tauValve = 0 Then
    Fs = Kvalve * Regulateur
    Else: Fs = previousF + (deltat / tauValve) * ((Kvalve * Regulateur) - previousF)
End If
Actionneur = Fs

End Function

Public Function CReacteur1(ByRef previousCA1 As Double, ByRef deltat As Double, ByRef actionneurCA1 As Double, ByRef CA0 As Double, ByRef volume As Double, ByRef k As Double) As Double
Dim CAnew As Double
CAnew = previousCA1 + (deltat * ((actionneurCA1 * CA0 / volume) - (actionneurCA1 * previousCA1 / volume) - (k * previousCA1)))
CReacteur1 = CAnew

End Function

Public Function CReacteur2(ByRef previousCA2 As Double, ByRef deltat As Double, ByRef actionneurCA2 As Double, ByRef CA1 As Double, ByRef volume As Double, ByRef k As Double) As Double
Dim CAnew As Double
CAnew = previousCA2 + deltat * ((actionneurCA2 * CA1 / volume) - (actionneurCA2 * previousCA2 / volume) - (k * previousCA2))
CReacteur2 = CAnew

End Function

Public Function Capteur(ByRef previousCA2M As Double, ByRef tauCapt As Double, ByRef kCapt As Double, ByRef CA2 As Double, ByRef deltat As Double) As Double
Dim CA2Mnew
CA2Mnew = previousCA2M + (deltat / tauCapt) * ((kCapt * CA2) - previousCA2M)
Capteur = CA2Mnew

End Function

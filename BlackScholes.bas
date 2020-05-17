Attribute VB_Name = "BlackScholes"
Option Explicit

'Generalized Black-Scholes (para opções européias)
Public Function GBlackScholes(ByVal Calculo As String, ByVal S As Double, ByVal x As Double, ByVal Vol As Double, ByVal r As Double, ByVal T As Double, ByVal q As Double, ByVal Tipo As String, ByVal Modelo As Integer, Optional ByVal H As Double, Optional ByVal TipoBarreira As String = 0)

    Dim b As Double
    
    Dim d1 As Double
    Dim d1y As Double
    
    Dim d2 As Double
    Dim d2y As Double
    
    Dim h1 As Double
    Dim h2 As Double
    
    Dim lambda As Double
    Dim PremioRegular As Double
    
    If T = 0 Then
        If Calculo = "Premio" Then
            GBlackScholes = PayOff(S, x, Tipo)
        Else
            GBlackScholes = 0
        End If
        Exit Function
    End If
    
    'Opções com Barreira
    If TipoBarreira <> 0 Then
        PremioRegular = GBlackScholes(Calculo, S, x, Vol, r, T, q, Tipo, Modelo)
    End If
    
    Select Case Modelo
        Case 0 'Black-Scholes (Ações sem dividendos)
            b = r
        Case 1 'Merton (Ações com dividendos, sendo q os dividendos esperados) ou Garman (moedas, sendo q a taxa de juros externa)
            b = r - q
        Case 2 'Black (futuros)
            b = 0
    End Select
    
    'Opções com Barreira
    If TipoBarreira <> 0 Then
        
        Select Case Tipo
            
            Case "call"
            
                Select Case TipoBarreira
                
                    Case 1, 2 'Down-In e Down-Out respectivamente
                        
                        If H <= x Then
                            
                            'Down-In
                            lambda = (b + (Vol ^ 2 / 2)) / (Vol ^ 2)
                            
                            d1 = (Log(H ^ 2 / (S * x))) / (Vol * Sqr(T)) + lambda * Vol * Sqr(T)
                            d2 = d1 - Vol * Sqr(T)
                                                        
                            h1 = (H / S) ^ (2 * lambda)
                            h2 = (H / S) ^ (2 * lambda - 2)
                        
                            GBlackScholes = (S * h1 * Exp((b - r) * T) * NormalCDF(d1))
                            GBlackScholes = GBlackScholes - (x * h2 * Exp(-r * T) * NormalCDF(d2))
                            
                            'Down-Out
                            If TipoBarreira = 2 Then
                                GBlackScholes = PremioRegular - GBlackScholes
                            End If
                        
                        ElseIf H > x Then
                        
                            'Down-Out
                            lambda = (b + (Vol ^ 2 / 2)) / (Vol ^ 2)
                            
                            d1 = (Log(S / H) / (Vol * Sqr(T))) + lambda * Vol * Sqr(T)
                            d2 = d1 - Vol * Sqr(T)
                                                        
                            GBlackScholes = (S * Exp((b - r) * T) * NormalCDF(d1))
                            GBlackScholes = GBlackScholes - (x * Exp(-r * T) * NormalCDF(d2))
                            
                            d1 = (Log(H / S) / (Vol * Sqr(T))) + lambda * Vol * Sqr(T)
                            d2 = d1 - Vol * Sqr(T)
                                                        
                            h1 = (H / S) ^ (2 * lambda)
                            h2 = (H / S) ^ (2 * lambda - 2)
                            
                            GBlackScholes = GBlackScholes - (S * h1 * Exp((b - r) * T) * NormalCDF(d1))
                            GBlackScholes = GBlackScholes + (x * h2 * Exp(-r * T) * NormalCDF(d2))
                        
                            'Down-In
                            If TipoBarreira = 1 Then
                                GBlackScholes = PremioRegular - GBlackScholes
                            End If
                        
                        End If
                        
                    Case 3, 4 'Up-In e Up-Out respectivamente
                        
                        If H <= x Then
                        
                            'Up-In
                            If TipoBarreira = 3 Then
                                GBlackScholes = PremioRegular
                            Else 'Up-Out
                                GBlackScholes = 0
                            End If
                            
                        ElseIf H > x Then
                        
                            'Up-In
                            lambda = (b + (Vol ^ 2 / 2)) / (Vol ^ 2)
                            
                            d1 = (Log(S / H) / (Vol * Sqr(T))) + lambda * Vol * Sqr(T)
                            d2 = d1 - Vol * Sqr(T)
                                                        
                            GBlackScholes = (S * Exp((b - r) * T) * NormalCDF(d1))
                            GBlackScholes = GBlackScholes - (x * Exp(-r * T) * NormalCDF(d2))
                            
                            d1 = -((Log(H / S) / (Vol * Sqr(T))) + lambda * Vol * Sqr(T))
                            d2 = d1 + Vol * Sqr(T)
                            
                            d1y = -((Log(H ^ 2 / (S * x)) / (Vol * Sqr(T))) + lambda * Vol * Sqr(T))
                            d2y = d1y + Vol * Sqr(T)
                                                        
                            h1 = (H / S) ^ (2 * lambda)
                            h2 = (H / S) ^ (2 * lambda - 2)
                            
                            GBlackScholes = GBlackScholes - (S * h1 * Exp((b - r) * T) * (NormalCDF(d1y) - NormalCDF(d1)))
                            GBlackScholes = GBlackScholes + (x * h2 * Exp(-r * T) * (NormalCDF(d2y) - NormalCDF(d2)))
                        
                            'Up-Out
                            If TipoBarreira = 4 Then
                                GBlackScholes = PremioRegular - GBlackScholes
                            End If
                        
                        End If
                    
                End Select
                
            Case "put"
        
                Select Case TipoBarreira
                
                    Case 3, 4 'Up-In e Up-Out respectivamente
                        
                        If H >= x Then
                            
                            'Up-In
                            lambda = (b + (Vol ^ 2 / 2)) / (Vol ^ 2)
                            
                            d1 = ((Log(H ^ 2 / (S * x))) / (Vol * Sqr(T)) + lambda * Vol * Sqr(T))
                            d2 = d1 - Vol * Sqr(T)
                                                        
                            h1 = (H / S) ^ (2 * lambda)
                            h2 = (H / S) ^ (2 * lambda - 2)
                        
                            GBlackScholes = (-S * h1 * Exp((b - r) * T) * NormalCDF(-d1))
                            GBlackScholes = GBlackScholes + (x * h2 * Exp(-r * T) * NormalCDF(-d2))
                            
                            'Up-Out
                            If TipoBarreira = 4 Then
                                GBlackScholes = PremioRegular - GBlackScholes
                            End If
                        
                        ElseIf H < x Then
                        
                            'Up-Out
                            lambda = (b + (Vol ^ 2 / 2)) / (Vol ^ 2)
                            
                            d1 = (Log(S / H) / (Vol * Sqr(T))) + lambda * Vol * Sqr(T)
                            d2 = d1 - Vol * Sqr(T)
                                                        
                            GBlackScholes = (-S * Exp((b - r) * T) * NormalCDF(-d1))
                            GBlackScholes = GBlackScholes + (x * Exp(-r * T) * NormalCDF(-d2))
                            
                            d1 = (Log(H / S) / (Vol * Sqr(T))) + lambda * Vol * Sqr(T)
                            d2 = d1 - Vol * Sqr(T)
                                                        
                            h1 = (H / S) ^ (2 * lambda)
                            h2 = (H / S) ^ (2 * lambda - 2)
                            
                            GBlackScholes = GBlackScholes + (S * h1 * Exp((b - r) * T) * NormalCDF(-d1))
                            GBlackScholes = GBlackScholes - (x * h2 * Exp(-r * T) * NormalCDF(-d2))
                        
                            'Up-In
                            If TipoBarreira = 3 Then
                                GBlackScholes = PremioRegular - GBlackScholes
                            End If
                        
                        End If
                        
                    Case 1, 2 'Down-In e Down-Out respectivamente
                        
                        If H > x Then
                        
                            'Down-In
                            If TipoBarreira = 1 Then
                                GBlackScholes = PremioRegular
                            Else 'Down-Out
                                GBlackScholes = 0
                            End If
                            
                        ElseIf H <= x Then
                        
                            'Down-In
                            lambda = (b + (Vol ^ 2 / 2)) / (Vol ^ 2)
                            
                            d1 = (Log(S / H) / (Vol * Sqr(T))) + lambda * Vol * Sqr(T)
                            d2 = d1 - Vol * Sqr(T)
                                                        
                            GBlackScholes = (-S * Exp((b - r) * T) * NormalCDF(-d1))
                            GBlackScholes = GBlackScholes + (x * Exp(-r * T) * NormalCDF(-d2))
                            
                            d1 = (Log(H / S) / (Vol * Sqr(T))) + lambda * Vol * Sqr(T)
                            d2 = -d1 + Vol * Sqr(T)
                            
                            d1y = (Log(H ^ 2 / (S * x)) / (Vol * Sqr(T))) + lambda * Vol * Sqr(T)
                            d2y = -d1y + Vol * Sqr(T)
                                                        
                            h1 = (H / S) ^ (2 * lambda)
                            h2 = (H / S) ^ (2 * lambda - 2)
                            
                            GBlackScholes = GBlackScholes + (S * h1 * Exp((b - r) * T) * (NormalCDF(d1y) - NormalCDF(d1)))
                            GBlackScholes = GBlackScholes - (x * h2 * Exp(-r * T) * (NormalCDF(-d2y) - NormalCDF(-d2)))
                        
                            'Down-Out
                            If TipoBarreira = 2 Then
                                GBlackScholes = PremioRegular - GBlackScholes
                            End If
                        
                        End If
                    
                End Select
        
        End Select
        
    Else
        d1 = (Log(S / x) + (b + Vol ^ 2 / 2) * T) / (Vol * Sqr(T))
        d2 = d1 - Vol * Sqr(T)
    End If
    
    Select Case UCase(Calculo)
    
        Case "PREMIO"
        
            If TipoBarreira = 0 Then
                
                If UCase(Tipo) = "CALL" Then
                    GBlackScholes = (S * Exp((b - r) * T) * NormalCDF(d1))
                    GBlackScholes = GBlackScholes - (x * Exp(-r * T) * NormalCDF(d2))
                Else
                    GBlackScholes = (x * Exp(-r * T) * NormalCDF(-d2))
                    GBlackScholes = GBlackScholes - (S * Exp((b - r) * T) * NormalCDF(-d1))
                End If
            End If
            
        Case "DELTA"

            If UCase(Tipo) = "CALL" Then
                GBlackScholes = Exp((b - r) * T) * NormalCDF(d1)
            Else
                GBlackScholes = Exp((b - r) * T) * (NormalCDF(d1) - 1)
            End If
        
        Case "GAMA"
            
            GBlackScholes = NormalPDF(d1) * Exp((b - r) * T) / (S * Vol * Sqr(T))
            
        Case "THETA"
        
            If UCase(Tipo) = "CALL" Then
                GBlackScholes = (-S * Exp((b - r) * T) * NormalPDF(d1) * Vol) / (2 * Sqr(T))
                GBlackScholes = GBlackScholes - (S * (b - r) * Exp((b - r) * T) * NormalCDF(d1))
                GBlackScholes = GBlackScholes - (r * x * Exp(-r * T) * NormalCDF(d2))
            Else
                GBlackScholes = (-S * Exp((b - r) * T) * NormalPDF(d1) * Vol) / (2 * Sqr(T))
                GBlackScholes = GBlackScholes + (S * (b - r) * Exp((b - r) * T) * NormalCDF(-d1))
                GBlackScholes = GBlackScholes + (r * x * Exp(-r * T) * NormalCDF(-d2))
            End If
                
        Case "VEGA"
            
            GBlackScholes = S * Sqr(T) * NormalPDF(d1) * Exp((b - r) * T)
        
        Case "RHO"
        
            If UCase(Tipo) = "CALL" Then
                If Modelo = 0 Or Modelo = 1 Then
                    GBlackScholes = (T * x * Exp(-r * T) * NormalCDF(d2))
                Else
                    GBlackScholes = -T * Exp(-r * T) * (S * NormalCDF(d1) - x * NormalCDF(d2))
                End If
            Else
                If Modelo = 0 Or Modelo = 1 Then
                    GBlackScholes = -T * x * Exp(-r * T) * NormalCDF(-d2)
                Else
                    GBlackScholes = -T * Exp(-r * T) * (x * NormalCDF(-d2) - S * NormalCDF(-d1))
                End If
            End If
    
        Case "RHO2"
            If UCase(Tipo) = "CALL" Then
                GBlackScholes = -T * S * Exp((b - r) * T) * NormalCDF(d1)
            Else
                GBlackScholes = T * S * Exp((b - r) * T) * NormalCDF(-d1)
            End If
    
    End Select

End Function


Public Function NormalCDF(z)
    NormalCDF = Application.WorksheetFunction.NormSDist(z)
End Function


'Função de densidade normal padrão
Public Function NormalPDF(z)
    Dim d1, d2, Pi
    Pi = 3.14159265358979 'Aproximadamente (355 / 113)
    NormalPDF = 1 / Sqr(2 * Pi) * Exp(-z ^ 2 / 2)
End Function

Public Function VolImplicita(ByVal Premio As Double, ByVal S As Double, ByVal x As Double, ByVal r As Double, ByVal T As Double, ByVal q As Double, ByVal Tipo As String, ByVal Modelo As Integer)

    Dim n
    Dim a, b
    Dim eps
    Dim Vol
    Dim Dif
    
    a = 0
    b = 2
    eps = 10 ^ (-5)
    n = 1

    Vol = (a + b) / 2
    Dif = Premio - GBlackScholes("Premio", S, x, Vol, r, T, q, Tipo, Modelo)

    Do While Abs(Dif) > eps And n <= 100

        If Dif > 0 Then
            a = Vol
        Else
            b = Vol
        End If
    
        Vol = (a + b) / 2
        Dif = Premio - GBlackScholes("Premio", S, x, Vol, r, T, q, Tipo, Modelo)
    
        n = n + 1
    
    Loop
    
    If n > 100 Then
        VolImplicita = 0.1
    Else
        VolImplicita = Vol
    End If

End Function






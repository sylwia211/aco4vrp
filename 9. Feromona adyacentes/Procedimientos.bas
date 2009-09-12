Attribute VB_Name = "Procedimientos"
Option Explicit


Sub Abstract()

ReDim Preserve Summary(1 To GenerationNext)

Summary(GenerationNext).Covered = BestRoute(0).Covered
Summary(GenerationNext).Nv = BestNv
Summary(GenerationNext).Time = Timer - Time

End Sub


Sub AccDeltaTao(Upshot() As Integer, Path() As Issues, ObjFunct As Double, Vehicle As Integer)

Dim I1 As Integer
Dim I2 As Integer
Dim I3 As Integer

'Numero de vehiculos
For I1 = 1 To Vehicle
    'Se buscan todos los nodos de dicho vehiculo
    For I2 = Path(I1 - 1).Ctrl + 1 To Path(I1).Ctrl - 2
        DeltaTao(Upshot(I2), Upshot(I2 + 1)) = DeltaTao(Upshot(I2), Upshot(I2 + 1)) + 1 / ObjFunct
        DeltaTao(Upshot(I2 + 1), Upshot(I2)) = DeltaTao(Upshot(I2 + 1), Upshot(I2)) + 1 / ObjFunct
    Next I2
Next I1

End Sub


Sub Aco()

'Problemas(s)
Problem = Form1.Combo1.Text

If Problem = "Todos" Then
    
    For NProblem = 1 To 14
    
        Time = Timer
        
        Call MultiVRP
        
    Next NProblem

    'Generar resumen de resultados
    Call MultiPrint
    
Else
    Time = Timer
    Call SingleVRP
End If

MsgBox "I finished!!!!"
End Sub


Sub Ant()

'Contadores y variables locales
Dim I As Integer

'Inicializacion de la variable solucion
Nv = 1
DemandA = 0

ReDim Route(0 To MaxNv)
ReDim Solution(1 To Nodes + MaxNv + 1)
ReDim Assigned(1 To Nodes)
ReDim Nearest(1 To Nodes, 1 To 2)


Route(0).Cused = 0
Route(0).Ctrl = 1
Sol = 0
Solution(Route(Nv - 1).Ctrl) = 0

'Ciclo para generar la ruta de cada vehiculo
Do While Sol < Nodes

    'Redimension de las variables
    'ReDim Preserve Route(0 To Nv)    '*****MOVER PARA EL FINAL DEL CICLO CUANDO ESTE LISTO***** Color Naranja
    
    'Búsqueda del nodo más lejano
    Chosen = 0
    For I = 1 To Nodes
        If Assigned(I) = False Then
            If Dist(I, 0) > Dist(Chosen, 0) Then Chosen = I
        End If
    Next I
    'Asignacion del primer nodo de la ruta (el mas lejano posible)
    Solution(Route(Nv - 1).Ctrl + 1) = Chosen
    Furthest = Chosen
    'If assigned(chosen) = True Then
    '    chosen = chosen
    'end if
    Assigned(Chosen) = True
    Solution(Route(Nv - 1).Ctrl + 2) = 0
    Route(Nv).Ctrl = Route(Nv - 1).Ctrl + 2
    Sol = Sol + 1
    Route(Nv).Covered = Dist(0, Chosen) * 2
    Route(Nv).Tused = Route(Nv).Covered + TimeS(Chosen)
    Route(Nv).Cused = Dem(Chosen)
    Nearest(Chosen, 2) = Route(Nv - 1).Ctrl + 1
    Eta(Chosen, 0, 0) = Nearest(Chosen, 2)

    'Actualizacion del parametro ETA
    Call UpdateETA
    
    'Ciclo para completar la ruta del vehiculo
    Do While Route(Nv).Cused < Capv And Chosen <> 0   'Mientras haya capacidad en el vehiculo

        'Sum = denominador para el calculo de probabilidades
        Sum = 0
        For I = 1 To Nodes
            If Eta(First(I), Solution(Eta(I, 0, 0) - 1), Solution(Eta(I, 0, 0))) <> 0 Then
                Sum = Sum + ((Weight * Tao(Furthest, I)) + ((1 - Weight) * Eta(First(I), Solution(Eta(I, 0, 0) - 1), Solution(Eta(I, 0, 0)))))
            End If
        Next I
        
        Chosen = 0
        
        If Sum <> 0 Then
            Prob = 0
            Randomize
            Random = Rnd
            
            'Seleccion de un nodo para la ruta
            For I = 1 To Nodes
                If Eta(First(I), Solution(Eta(I, 0, 0) - 1), Solution(Eta(I, 0, 0))) <> 0 Then
                    Prob = Prob + (((Weight * Tao(Furthest, I)) + ((1 - Weight) * Eta(First(I), Solution(Eta(I, 0, 0) - 1), Solution(Eta(I, 0, 0))))) / Sum)     'Probabilidad acumulada
                    If Prob > Random Then
                        If Assigned(I) = False Then
                            If Route(Nv).Cused + Dem(I) <= Capv Then
                                Chosen = I
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next I
        End If
        
        'Actualizacion de la solucion
        If Chosen = 0 Then
            GoTo 10
        End If
        Sol = Sol + 1
        Route(Nv).Ctrl = Route(Nv).Ctrl + 1
        For I = Route(Nv).Ctrl To Nearest(Chosen, 2) + 1 Step -1
            Solution(I) = Solution(I - 1)
        Next I
        Solution(Nearest(Chosen, 2)) = Chosen
        Assigned(Chosen) = True
        Route(Nv).Cused = Route(Nv).Cused + Dem(Chosen)
        Route(Nv).Covered = Route(Nv).Covered + Dist(Chosen, Solution(Nearest(Chosen, 2) - 1)) + Dist(Chosen, Solution(Nearest(Chosen, 2) + 1)) - Dist(Solution(Nearest(Chosen, 2) - 1), Solution(Nearest(Chosen, 2) + 1))
        Route(Nv).Tused = Route(Nv).Tused + TimeS(Chosen) + Dist(Chosen, Solution(Nearest(Chosen, 2) - 1)) + Dist(Chosen, Solution(Nearest(Chosen, 2) + 1)) - Dist(Solution(Nearest(Chosen, 2) - 1), Solution(Nearest(Chosen, 2) + 1))
        
        'Actualizacion del parametro ETA
        Call UpdateETA
        
10
        
    Loop
    
    Route(0).Cused = Route(0).Cused + Route(Nv).Cused
    
    If Sol < Nodes Then
        If Nv >= MaxNv Then
            Do
                Call UnInfactibilization
                    If Worst(1) = 0 Then
                        K = K - 1
                        Sol = Nodes
                        Exit Do
                    End If
            Loop While Sol < Nodes
        Else
            If ((Demand - Route(0).Cused) / (Capv * (MaxNv - Nv))) > 1 Then
                Do
                    Call UnInfactibilization
                    If Worst(1) = 0 Then
                        K = K - 1
                        Sol = Nodes
                        Exit Do
                    End If
                Loop While ((Demand - Route(0).Cused) / (Capv * (MaxNv - Nv))) > 1
                DemandA = DemandA + Route(Nv).Cused
                Nv = Nv + 1
            Else
                DemandA = DemandA + Route(Nv).Cused
                Nv = Nv + 1
            End If
        End If
    End If
    
Loop

'Calculo funcion objetivo
Route(0).Covered = 0
For I = 1 To MaxNv
    Route(0).Covered = Route(0).Covered + Route(I).Covered
Next I


End Sub


Sub Plot()

'Contadores y Variables locales
Dim I As Integer
Dim J As Integer

Form2.Show
Form2.Cls
Form2.Scale (MinX - 0.1 * MaxX, MaxY + 0.1 * MaxX)-(MaxX + 0.1 * MaxX, MinY - 0.1 * MaxX)
Form2.DrawWidth = 5
Form2.PSet (X(0), Y(0)), vbRed
Form2.DrawWidth = 3
For I = 1 To Nodes
    If Assigned(I) = False Then
        Form2.PSet (X(I), Y(I)), vbBlue
    Else
        Form2.PSet (X(I), Y(I)), vbBlack
    End If
Next I
Form2.DrawWidth = 1
For I = 1 To UBound(BestSolution) - 1
    Form2.Line (X(BestSolution(I)), Y(BestSolution(I)))-(X(BestSolution(I + 1)), Y(BestSolution(I + 1))), vbYellow
Next I

Form3.Show
Form3.Cls
Form3.Scale (0, TheWorst)-(GenerationNext + 1, TheBest)
Form3.DrawWidth = 3
For I = 1 To GenerationNext
    For J = 1 To nAnts
        Form3.PSet (I, History(J, I)), vbRed
    Next J
Next I

End Sub


Sub Exchange()

Dim I1 As Integer
Dim I2 As Integer
Dim I3 As Integer
Dim I4 As Integer
Dim I5 As Integer
Dim Release As Boolean
Dim Temp As Neighbor

Temp.Covered = Route(0).Covered

Do

    Release = False

    For I1 = 1 To UBound(Route)
        For I3 = 1 To UBound(Route)
            If I1 <> I3 Then
                For I2 = Route(I1 - 1).Ctrl + 1 To Route(I1).Ctrl - 1
                    For I4 = Route(I3 - 1).Ctrl + 1 To Route(I3).Ctrl - 1
                        If (Route(I1).Cused + Dem(Solution(I4)) - Dem(Solution(I2))) <= Capv And (Route(I3).Cused + Dem(Solution(I2)) - Dem(Solution(I4))) <= Capv Then
                            If (Route(I1).Tused + TimeS(Solution(I4)) - TimeS(Solution(I2)) + Dist(Solution(I2 - 1), Solution(I4)) + Dist(Solution(I4), Solution(I2 + 1)) - Dist(Solution(I2), Solution(I2 + 1)) - Dist(Solution(I2), Solution(I2 - 1))) <= TimeC Or TimeC = 0 Then
                                If (Route(I3).Tused + TimeS(Solution(I2)) - TimeS(Solution(I4)) + Dist(Solution(I4 - 1), Solution(I2)) + Dist(Solution(I2), Solution(I4 + 1)) - Dist(Solution(I4), Solution(I4 + 1)) - Dist(Solution(I4), Solution(I4 - 1))) <= TimeC Or TimeC = 0 Then
                                    If (Route(0).Covered + Dist(Solution(I2 - 1), Solution(I4)) + Dist(Solution(I4), Solution(I2 + 1)) - Dist(Solution(I2), Solution(I2 + 1)) - Dist(Solution(I2), Solution(I2 - 1)) + Dist(Solution(I4 - 1), Solution(I2)) + Dist(Solution(I2), Solution(I4 + 1)) - Dist(Solution(I4), Solution(I4 + 1)) - Dist(Solution(I4), Solution(I4 - 1))) < Temp.Covered Then
                                        Temp.FV = I1
                                        Temp.SV = I3
                                        Temp.NFV = I2
                                        Temp.NSV = I4
                                        Temp.Covered = (Route(0).Covered + Dist(Solution(I2 - 1), Solution(I4)) + Dist(Solution(I4), Solution(I2 + 1)) - Dist(Solution(I2), Solution(I2 + 1)) - Dist(Solution(I2), Solution(I2 - 1)) + Dist(Solution(I4 - 1), Solution(I2)) + Dist(Solution(I2), Solution(I4 + 1)) - Dist(Solution(I4), Solution(I4 + 1)) - Dist(Solution(I4), Solution(I4 - 1)))
                                        Release = True
                                    End If
                                End If
                            End If
                        End If
                    Next I4
                Next I2
            End If
        Next I3
    Next I1
    
    If Release = True Then
        Route(Temp.FV).Cused = Route(Temp.FV).Cused + Dem(Solution(Temp.NSV)) - Dem(Solution(Temp.NFV))
        Route(Temp.SV).Cused = Route(Temp.SV).Cused - Dem(Solution(Temp.NSV)) + Dem(Solution(Temp.NFV))
        Route(Temp.FV).Tused = Route(Temp.FV).Tused + TimeS(Solution(Temp.NSV)) - TimeS(Solution(Temp.NFV)) + Dist(Solution(Temp.NFV - 1), Solution(Temp.NSV)) + Dist(Solution(Temp.NSV), Solution(Temp.NFV + 1)) - Dist(Solution(Temp.NFV), Solution(Temp.NFV + 1)) - Dist(Solution(Temp.NFV), Solution(Temp.NFV - 1))
        Route(Temp.FV).Covered = Route(Temp.FV).Covered + Dist(Solution(Temp.NFV - 1), Solution(Temp.NSV)) + Dist(Solution(Temp.NSV), Solution(Temp.NFV + 1)) - Dist(Solution(Temp.NFV), Solution(Temp.NFV + 1)) - Dist(Solution(Temp.NFV), Solution(Temp.NFV - 1))
        Route(Temp.SV).Tused = Route(Temp.SV).Tused - TimeS(Solution(Temp.NSV)) + TimeS(Solution(Temp.NFV)) + Dist(Solution(Temp.NSV - 1), Solution(Temp.NFV)) + Dist(Solution(Temp.NSV + 1), Solution(Temp.NFV)) - Dist(Solution(Temp.NSV), Solution(Temp.NSV + 1)) - Dist(Solution(Temp.NSV), Solution(Temp.NSV - 1))
        Route(Temp.SV).Covered = Route(Temp.SV).Covered + Dist(Solution(Temp.NSV - 1), Solution(Temp.NFV)) + Dist(Solution(Temp.NSV + 1), Solution(Temp.NFV)) - Dist(Solution(Temp.NSV), Solution(Temp.NSV + 1)) - Dist(Solution(Temp.NSV), Solution(Temp.NSV - 1))
        Route(0).Covered = Temp.Covered
            
        Temp.Help = Solution(Temp.NSV)
        Solution(Temp.NSV) = Solution(Temp.NFV)
        Solution(Temp.NFV) = Temp.Help
        
    End If
    
Loop While Release = True

If Int(Route(0).Covered * 100000) / 100000 < Int(RouteAnt(0).Covered * 100000) / 100000 Then
    NLS = NLS + 1
    Call AccDeltaTao(Solution(), Route(), Route(0).Covered / nAnts / 0.1, UBound(Route) - 1)
    
    For I1 = 1 To Nodes + MaxNv + 1
        BestAnt(I1) = Solution(I1)
    Next I1
    For I1 = 0 To MaxNv
        RouteAnt(I1) = Route(I1)
    Next I1
    
End If

If BestRoute(0).Covered > Route(0).Covered Then
    For I1 = 0 To UBound(Route)
        BestRoute(I1) = Route(I1)
    Next I1
    ReDim BestSolution(1 To Nodes + MaxNv + 1)
    For I1 = 1 To Nodes + MaxNv + 1
        BestSolution(I1) = Solution(I1)
    Next I1
    BestNv = UBound(Route)

    LastImprove = GenerationNext

End If


End Sub


'Sub Factibilization()

'Dim I1 As Integer   'Cantidad nodos no asignados
'Dim I2 As Integer   'Busca los nodos no asignados
'Dim I3 As Integer   'Busca en cada ruta
'Dim I4 As Integer   'Busca la ubicación en una solución que sea lo mejor (factible)

'For I1 = Sol + 1 To Nodes
'    For I2 = 1 To Nodes
'        If Assigned(I2) = False Then
'            Nearest(I2, 2) = 0
'            For I3 = 1 To MaxNv
'                If Capv >= Cused(I3) + Dem(I2) Then
'                    For I4 = Ctrl(I3) To Ctrl(I3 + 1) - 1
'                        If (Tused(I3) + TimeS(I2) + Dist(I2, Solution(I4)) + Dist(I2, Solution(I4 + 1))) < TimeC Or TimeC = 0 Then
'                            If Nearest(I2, 2) = 0 Then
'                                Nearest(I2, 1) = Solution(I4)
'                                Nearest(I2, 2) = I4
'                            Else
'                                If (Dist(I2, Solution(I4)) + Dist(I2, Solution(I4 + 1))) < (Dist(I2, Solution(Nearest(I2, 2))) + Dist(I2, Solution(Nearest(I2, 2) + 1))) Then
'                                    Nearest(I2, 1) = Solution(I4)
'                                    Nearest(I2, 2) = I4
'                                End If
'                            End If
'                        End If
'                    Next I4
'                End If
'            Next I3
'            If Nearest(I2, 2) = 0 Then
'                Sol = 0
'                I1 = Nodes
'                Exit For
'            Else
'                For I3 = 1 To MaxNv
'                    If Nearest(I2, 2) < Ctrl(I3 + 1) Then
'                        For I4 = Ctrl(Nv + 1) To Nearest(I2, 2) + 1 Step -1
'                            Solution(I4 + 1) = Solution(I4)
'                        Next I4
'                        Solution(Nearest(I2, 2) + 1) = I2
'                        For I4 = I3 + 1 To Nv + 1
'                            Ctrl(I4) = Ctrl(I4) + 1
'                        Next I4
'                        Cused(I3) = Cused(I3) + Dem(I2)
'                        Covered(I3) = Covered(I3) + Dist(I2, Solution(Nearest(I2, 2))) + Dist(I2, Solution(Nearest(I2, 2) + 2)) - Dist(Solution(Nearest(I2, 2)), Solution(Nearest(I2, 2) + 2))
'                        Tused(I3) = Tused(I3) + TimeS(I2) + Dist(I2, Solution(Nearest(I2, 2))) + Dist(I2, Solution(Nearest(I2, 2) + 2)) - Dist(Solution(Nearest(I2, 2)), Solution(Nearest(I2, 2) + 2))
'                        Covered(0) = Covered(0) + Dist(I2, Solution(Nearest(I2, 2))) + Dist(I2, Solution(Nearest(I2, 2) + 2)) - Dist(Solution(Nearest(I2, 2)), Solution(Nearest(I2, 2) + 2))
'                        I1 = I1 + 1
'                        Exit For
'                    End If
'                Next I3
'            End If
'        End If
'        If I1 > Nodes Then Exit For
'    Next I2
'Next I1

'End Sub


Sub Improved()

Dim I1 As Integer

ReDim Preserve BestRoute(0 To MaxNv)

If BestRoute(0).Covered > Route(0).Covered Or BestRoute(0).Covered = 0 Then
    ReDim RouteAnt(0 To MaxNv)
    For I1 = 0 To UBound(Route)
        BestRoute(I1) = Route(I1)
        RouteAnt(I1) = Route(I1)
    Next I1
    ReDim BestSolution(1 To Nodes + MaxNv + 1)
    ReDim BestAnt(0 To Nodes + MaxNv + 1)
    For I1 = 1 To Nodes + Nv + 1
        BestSolution(I1) = Solution(I1)
        BestAnt(I1) = Solution(I1)
    Next I1
    BestNv = UBound(Route)
    
    LastImprove = GenerationNext
    
ElseIf UBound(BestAnt) = 0 Then
    ReDim BestAnt(0 To Nodes + MaxNv + 1)
    ReDim RouteAnt(0 To MaxNv)
    For I1 = 1 To Nodes + MaxNv + 1
        BestAnt(I1) = Solution(I1)
    Next I1
    For I1 = 0 To MaxNv
        RouteAnt(I1) = Route(I1)
    Next I1
ElseIf RouteAnt(0).Covered > Route(0).Covered Then
    ReDim RouteAnt(0 To MaxNv)
    For I1 = 1 To Nodes + MaxNv + 1
        BestAnt(I1) = Solution(I1)
    Next I1
    For I1 = 0 To MaxNv
        RouteAnt(I1) = Route(I1)
    Next I1
End If

End Sub


Sub Insertion()

Dim I1 As Integer
Dim I2 As Integer
Dim I3 As Integer
Dim I4 As Integer
Dim I5 As Integer
Dim Release As Boolean
Dim Temp As Neighbor

Temp.Covered = Route(0).Covered

Do

    Release = False

    For I1 = 1 To UBound(Route)
        For I3 = 1 To UBound(Route)
            If I1 <> I3 Then
                For I2 = Route(I1 - 1).Ctrl + 1 To Route(I1).Ctrl - 1
                    For I4 = Route(I3 - 1).Ctrl + 1 To Route(I3).Ctrl - 1
                        If Route(I1).Cused + Dem(Solution(I4)) <= Capv Then
                            If (Route(I1).Tused + TimeS(Solution(I4)) + Dist(Solution(I2), Solution(I4)) + Dist(Solution(I4), Solution(I2 + 1)) - Dist(Solution(I2), Solution(I2 + 1))) <= TimeC Or TimeC = 0 Then
                                If (Route(0).Covered + Dist(Solution(I2), Solution(I4)) + Dist(Solution(I4), Solution(I2 + 1)) - Dist(Solution(I2), Solution(I2 + 1)) - Dist(Solution(I4), Solution(I4 - 1)) - Dist(Solution(I4), Solution(I4 + 1)) + Dist(Solution(I4 - 1), Solution(I4 + 1))) < Temp.Covered Then
                                    Temp.FV = I1
                                    Temp.SV = I3
                                    Temp.NFV = I2
                                    Temp.NSV = I4
                                    Temp.Covered = (Route(0).Covered + Dist(Solution(I2), Solution(I4)) + Dist(Solution(I4), Solution(I2 + 1)) - Dist(Solution(I2), Solution(I2 + 1)) - Dist(Solution(I4), Solution(I4 - 1)) - Dist(Solution(I4), Solution(I4 + 1)) + Dist(Solution(I4 - 1), Solution(I4 + 1)))
                                    Release = True
                                End If
                            End If
                        End If
                    Next I4
                Next I2
            End If
        Next I3
    Next I1
    
    If Release = True Then
        Route(Temp.FV).Cused = Route(Temp.FV).Cused + Dem(Solution(Temp.NSV))
        Route(Temp.SV).Cused = Route(Temp.SV).Cused - Dem(Solution(Temp.NSV))
        Route(Temp.FV).Tused = Route(Temp.FV).Tused + TimeS(Solution(Temp.NSV)) + Dist(Solution(Temp.NFV), Solution(Temp.NSV)) + Dist(Solution(Temp.NSV), Solution(Temp.NFV + 1)) - Dist(Solution(Temp.NFV), Solution(Temp.NFV + 1))
        Route(Temp.FV).Covered = Route(Temp.FV).Covered + Dist(Solution(Temp.NFV), Solution(Temp.NSV)) + Dist(Solution(Temp.NSV), Solution(Temp.NFV + 1)) - Dist(Solution(Temp.NFV), Solution(Temp.NFV + 1))
        Route(Temp.SV).Tused = Route(Temp.SV).Tused - TimeS(Solution(Temp.NSV)) + Dist(Solution(Temp.NSV - 1), Solution(Temp.NSV + 1)) - Dist(Solution(Temp.NSV), Solution(Temp.NSV + 1)) - Dist(Solution(Temp.NSV), Solution(Temp.NSV - 1))
        Route(Temp.SV).Covered = Route(Temp.SV).Covered + Dist(Solution(Temp.NSV - 1), Solution(Temp.NSV + 1)) - Dist(Solution(Temp.NSV), Solution(Temp.NSV + 1)) - Dist(Solution(Temp.NSV), Solution(Temp.NSV - 1))
        Route(0).Covered = Temp.Covered
        If Temp.FV < Temp.SV Then
            Temp.Help = Solution(Temp.NSV)
            For I1 = Temp.NSV To Temp.NFV + 2 Step -1
                Solution(I1) = Solution(I1 - 1)
            Next I1
            Solution(Temp.NFV + 1) = Temp.Help
            For I1 = Temp.FV To Temp.SV - 1
                Route(I1).Ctrl = Route(I1).Ctrl + 1
            Next I1
        Else
            Temp.Help = Solution(Temp.NSV)
            For I1 = Temp.NSV To Temp.NFV - 1
                Solution(I1) = Solution(I1 + 1)
            Next I1
            Solution(Temp.NFV) = Temp.Help
            For I1 = Temp.SV To Temp.FV - 1
                Route(I1).Ctrl = Route(I1).Ctrl - 1
            Next I1
        End If
    End If
    
Loop While Release = True

If Int(Route(0).Covered * 100000) / 100000 < Int(RouteAnt(0).Covered * 100000) / 100000 Then
    NLS = NLS + 1
    Call AccDeltaTao(Solution(), Route(), Route(0).Covered / nAnts / 0.1, UBound(Route) - 1)
    
    For I1 = 1 To Nodes + MaxNv + 1
        BestAnt(I1) = Solution(I1)
    Next I1
    For I1 = 0 To MaxNv
        RouteAnt(I1) = Route(I1)
    Next I1
    
End If

If BestRoute(0).Covered > Route(0).Covered Then
    For I1 = 0 To UBound(Route)
        BestRoute(I1) = Route(I1)
    Next I1
    ReDim BestSolution(1 To Nodes + MaxNv + 1)
    For I1 = 1 To Nodes + MaxNv + 1
        BestSolution(I1) = Solution(I1)
    Next I1
    BestNv = UBound(Route)
    
    LastImprove = GenerationNext
    
End If

End Sub


Sub MultiPrint()

Dim I1 As Integer

Open App.Path & "\Resultados\Summary Iter=" & nGen & " nAnts=" & nAnts & " Weight=" & Weight & " Rho=" & Rho & ".txt" For Output As #3
Print #3, vbTab & "Costo" & vbTab & "Iter." & vbTab & "Tiempo"
For I1 = 1 To 14
    Print #3, I1 & vbTab & Int(Final(I1).Covered * 100) / 100 & vbTab & Final(I1).Nv & vbTab & Int(Final(I1).Time * 1000) / 1000
Next I1
Close #3

End Sub


Sub MultiSinglePrint()

Dim I1 As Integer
Dim I2 As Integer
Dim Text As String

If NProblem < 10 Then
    Open App.Path & "\Resultados\Solve0" & NProblem & ".txt" For Output As #2
    Open App.Path & "\Resultados\Summary0" & NProblem & ".txt" For Output As #3
Else
    Open App.Path & "\Resultados\Solve" & NProblem & ".txt" For Output As #2
    Open App.Path & "\Resultados\Summary" & NProblem & ".txt" For Output As #3
End If

I2 = 1
Print #2, "Costo de la solución:"
Print #2, BestRoute(0).Covered
Print #2,
For I1 = 1 To MaxNv
    Text = "1" & vbTab & I1 & vbTab & Int(BestRoute(I1).Covered * 1000) / 1000 & vbTab & BestRoute(I1).Cused & vbTab
    Do
        Text = Text & BestSolution(I2) & vbTab
        I2 = I2 + 1
        If I2 > UBound(BestSolution) Then
            I2 = I2 - 1
            Exit Do
        End If
    Loop While BestSolution(I2) <> 0
    Text = Text & 0
    Print #2, Text
Next I1
Print #2,
Print #2, "Tiempo de ejecución:"
Print #2, Timer - Time
Close #2

Print #3, "Iteraciones: ", UBound(Summary)
Print #3,
Print #3, "Costo", "Vehículos", "Tiempo [seg]"
For I1 = 1 To UBound(Summary)
    Print #3, Int(Summary(I1).Covered * 1000) / 1000, Summary(I1).Nv, Int(Summary(I1).Time * 1000) / 1000
Next I1
Close #3

End Sub


Sub MultiVRP()

Dim K As Integer

'Lectura de datos
Call Reading_Multi

'Inicialización de variables y parámetros
Call Parameters

'Generaciones ' Iteraciones
GenerationNext = 0
LastImprove = 0
Do
    GenerationNext = GenerationNext + 1

    ReDim BestAnt(0 To 0)

    For K = 1 To nAnts
    
        ReDim TabuList(1 To Nodes, 1 To MaxNv)
        
        'Colonia
        Call Ant

        If Worst(1) <> 0 Then
            
            Call Improved

            'Acumular DeltaTao
            Call AccDeltaTao(Solution(), Route(), Route(0).Covered, Nv)
            
        End If
        
    Next K
    
    'BÚSQUEDA LOCAL
    NLS = 2
    Do While UBound(BestAnt) <> 0 And NLS > 0
        NLS = 0
        Call TwoOpt
        Call Insertion
        Call Exchange
    Loop
    
    Call AccDeltaTao(BestSolution(), BestRoute(), Route(0).Covered / nAnts, BestNv)
    
    'Actulización de Tao
    Call UpdateTAO

    Call Abstract

Loop While (GenerationNext - LastImprove) < nGen

Final(NProblem).Covered = BestRoute(0).Covered
Final(NProblem).Nv = GenerationNext
Final(NProblem).Time = Timer - Time

'Generar archivo de resultados
Call MultiSinglePrint

End Sub


Sub Parameters()

'Contadores propios del procedimiento
Dim I As Integer
Dim J As Integer
Dim K As Integer

'Parámetros
nGen = Val(Form1.Text1)
nAnts = Val(Form1.Text2)
Rho = Val(Form1.Text3)
Weight = Val(Form1.Text4)

'Dimensión de la solución
ReDim Eta(0 To Nodes, 0 To Nodes, 0 To Nodes)
'ReDim Save(1 To Nodes, 1 To Nodes)
ReDim Solution(0 To Nodes + MaxNv + 1)
ReDim Tao(1 To Nodes, 1 To Nodes)
ReDim DeltaTao(1 To Nodes, 1 To Nodes)
ReDim First(1 To Nodes)
ReDim Second(1 To Nodes)

ReDim BestRoute(0 To MaxNv)
ReDim BestSolution(1 To Nodes + MaxNv + 1)

'Savings - Ahorros
'For I = 1 To Nodes
'    For J = I + 1 To Nodes
'        Save(I, J) = Dist(0, I) + Dist(0, J) - Dist(I, J)
'    Next J
'Next I

'Eta - Información Heurística
Max = 0
For I = 1 To Nodes
    For J = 0 To Nodes
        For K = J + 1 To Nodes
            If Dist(I, J) = 0 And Dist(I, K) = 0 Then
                Eta(I, J, K) = 1
                Eta(I, K, J) = 1
            Else
                Eta(I, J, K) = 1 / (Dist(I, J) + Dist(I, K))
                Eta(I, K, J) = 1 / (Dist(I, J) + Dist(I, K))
                If Eta(I, J, K) > Max Then Max = Eta(I, J, K)
            End If
        Next K
    Next J
Next I
For I = 1 To Nodes
    First(I) = I
    For J = 0 To Nodes
        For K = 0 To Nodes
            Eta(I, J, K) = Eta(I, J, K) / Max
        Next K
    Next J
Next I

'Tao - Feromona
For I = 1 To Nodes
    Tao(I, I) = 0
    For J = I + 1 To Nodes
        Tao(I, J) = 1
        Tao(J, I) = 1
    Next J
Next I

End Sub


Sub Reading()

'Contadores propios del procedimiento
Dim I As Integer
Dim J As Integer

'Se abre el archivo de datos
Open App.Path & "\Datos\" & Problem & ".vrp" For Input As #1

    'Se lee el número de nodos y la capacidad
    Input #1, I, MaxNv, Nodes, J
    Input #1, TimeC, Capv
    
    'Se redimensionan los vectores de datos
    ReDim X(0 To Nodes)
    ReDim Y(0 To Nodes)
    ReDim Dem(0 To Nodes)
    ReDim Dist(0 To Nodes, 0 To Nodes)
    ReDim TimeS(1 To Nodes)
    
    I = 0
    Input #1, I, X(I), Y(I), J, J, J, J
            
    MinX = X(I)
    MaxX = X(I)
    MinY = Y(I)
    MaxY = Y(I)
    Demand = 0
    
    'Se leen las coordenadas y la demanda de cada nodo
    For I = 1 To Nodes
        
        Input #1, I, X(I), Y(I), TimeS(I), Dem(I), J, J, J
        
        If MinX > X(I) Then MinX = X(I)
        If MaxX < X(I) Then MaxX = X(I)
        If MinY > Y(I) Then MinY = Y(I)
        If MaxY < Y(I) Then MaxY = Y(I)
'STARTSTARTSTARTSTARTSTARTSTARTSTARTSTARTSTARTSTARTSTARTSTART
        Demand = Demand + Dem(I)
'ENDENDENDENDENDENDENDENDENDENDENDENDENDENDENDENDENDENDENDEND

    Next I

'Se cierra el archivo de datos
Close #1

'Se calcula la distancia entre nodos
For I = 0 To Nodes
    For J = 0 To Nodes
        Dist(I, J) = Sqr((X(I) - X(J)) ^ 2 + (Y(I) - Y(J)) ^ 2)
    Next J
Next I

End Sub


Sub Reading_Multi()

'Contadores propios del procedimiento
Dim I As Integer
Dim J As Integer

'Se abre el archivo de datos
If NProblem < 10 Then
    Open App.Path & "\Datos\p0" & NProblem & ".vrp" For Input As #1
Else
    Open App.Path & "\Datos\p" & NProblem & ".vrp" For Input As #1
End If


    'Se lee el número de nodos y la capacidad
    Input #1, I, MaxNv, Nodes, J
    Input #1, TimeC, Capv
    
    'Se redimensionan los vectores de datos
    ReDim X(0 To Nodes)
    ReDim Y(0 To Nodes)
    ReDim Dem(0 To Nodes)
    ReDim Dist(0 To Nodes, 0 To Nodes)
    ReDim TimeS(1 To Nodes)
    
    I = 0
    Input #1, I, X(I), Y(I), J, J, J, J
    
    Demand = 0
        
    'Se leen las coordenadas y la demanda de cada nodo
    For I = 1 To Nodes
        
        Input #1, I, X(I), Y(I), TimeS(I), Dem(I), J, J, J
        
        If I <> 0 Then
            If MinX > X(I) Then MinX = X(I)
            If MaxX < X(I) Then MaxX = X(I)
            If MinY > Y(I) Then MinY = Y(I)
            If MaxY < Y(I) Then MaxY = Y(I)
        Else
            MinX = X(I)
            MaxX = X(I)
            MinY = Y(I)
            MaxY = Y(I)
        End If
        
'STARTSTARTSTARTSTARTSTARTSTARTSTARTSTARTSTARTSTARTSTARTSTART
        Demand = Demand + Dem(I)
'ENDENDENDENDENDENDENDENDENDENDENDENDENDENDENDENDENDENDENDEND
        
    Next I

'Se cierra el archivo de datos
Close #1

'Se calcula la distancia entre nodos
For I = 0 To Nodes
    For J = 0 To Nodes
        Dist(I, J) = Sqr((X(I) - X(J)) ^ 2 + (Y(I) - Y(J)) ^ 2)
    Next J
Next I

End Sub


Sub SinglePrint()

Dim I1 As Integer
Dim I2 As Integer
Dim Text As String

Open App.Path & "\Resultados\Solve " & Problem & ".txt" For Output As #2
Open App.Path & "\Resultados\Summary " & Problem & ".txt" For Output As #3

I2 = 1
Print #2, "Costo de la solución:"
Print #2, BestRoute(0).Covered
Print #2,
For I1 = 1 To MaxNv
    Text = "1" & vbTab & I1 & vbTab & Int(BestRoute(I1).Covered * 1000) / 1000 & vbTab & BestRoute(I1).Cused & vbTab
    Do
        Text = Text & BestSolution(I2) & vbTab
        I2 = I2 + 1
        If I2 > UBound(BestSolution) Then Exit For
    Loop While BestSolution(I2) <> 0
    Text = Text & 0
    Print #2, Text
Next I1
Print #2,
Print #2, "Tiempo de ejecución:"
Print #2, Timer - Time
Close #2

Print #3, "Iteraciones: ", UBound(Summary)
Print #3,
Print #3, "Costo", "Vehículos", "Tiempo [seg]"
For I1 = 1 To UBound(Summary)
    Print #3, Int(Summary(I1).Covered * 1000) / 1000, Summary(I1).Nv, Int(Summary(I1).Time * 1000) / 1000
Next I1
Close #3

End Sub


Sub SingleVRP()

Dim I As Integer

'Lectura de datos
Call Reading

'Inicialización de variables y parámetros
Call Parameters

'Generaciones ' Iteraciones
GenerationNext = 0
LastImprove = 0
Do
    GenerationNext = GenerationNext + 1
    ReDim Preserve History(1 To nAnts, 1 To GenerationNext)

    ReDim BestAnt(0 To 0)

    For K = 1 To nAnts
    
        ReDim TabuList(1 To Nodes, 1 To MaxNv)
    
        'Colonia
        I = K
        Call Ant
        
        'Almacenamiento de historia
        If K = I Then
            History(K, GenerationNext) = Route(0).Covered
            If TheBest = 0 Then
                TheBest = History(K, GenerationNext)
            End If
            If History(K, GenerationNext) < TheBest Then
                TheBest = History(K, GenerationNext)
            ElseIf History(K, GenerationNext) > TheWorst Then
                TheWorst = History(K, GenerationNext)
            End If
        End If
            
        If Worst(1) <> 0 Then
            Call Improved

            'Acumular DeltaTao
            Call AccDeltaTao(Solution(), Route(), Route(0).Covered, Nv)

        End If
        
    Next K
    
    'BÚSQUEDA LOCAL
    NLS = 2
    Do While UBound(BestAnt) <> 0 And NLS > 0
        NLS = 0
        Call TwoOpt
        Call Insertion
        Call Exchange
    Loop
    
    Call Abstract
    
    Call AccDeltaTao(BestSolution(), BestRoute(), Route(0).Covered / nAnts, BestNv)
    
    'Actulización de Tao
    Call UpdateTAO

Loop While (GenerationNext - LastImprove) < nGen

'Mostrar resultados
MsgBox BestRoute(0).Covered & vbCrLf & Nv
Call Plot

'Generar archivo de resultados
Call SinglePrint

End Sub


Sub TwoOpt()

Dim I As Integer        'CONTADOR PARA EL NÚMERO DE VEHÍCULOS
Dim J As Integer        'CONTADOR PARA EL ENLACE 1 QUE SE ROMPE
Dim K As Integer        'CONTADOR PARA EL ENLACE 2 QUE SE ROMPE
Dim W As Integer        'CONTADOR AUXILIAR
Dim Temp As Integer     'VBLE TEMPORAL
Dim Release As Boolean  'Release

RouteAnt(0).Covered = 0
For I = 1 To UBound(RouteAnt)
    For J = RouteAnt(I - 1).Ctrl To RouteAnt(I).Ctrl - 1
        For K = J + 2 To RouteAnt(I).Ctrl - 1
            If Dist(BestAnt(J), BestAnt(J + 1)) + Dist(BestAnt(K), BestAnt(K + 1)) > Dist(BestAnt(J), BestAnt(K)) + Dist(BestAnt(J + 1), BestAnt(K + 1)) Then
                RouteAnt(I).Covered = RouteAnt(I).Covered + (-Dist(BestAnt(J), BestAnt(J + 1)) - Dist(BestAnt(K), BestAnt(K + 1)) + Dist(BestAnt(J), BestAnt(K)) + Dist(BestAnt(J + 1), BestAnt(K + 1)))
                RouteAnt(I).Tused = RouteAnt(I).Tused + (-Dist(BestAnt(J), BestAnt(J + 1)) - Dist(BestAnt(K), BestAnt(K + 1)) + Dist(BestAnt(J), BestAnt(K)) + Dist(BestAnt(J + 1), BestAnt(K + 1)))
                For W = 1 To Int((K - J) / 2)
                    Temp = BestAnt(J + W)
                    BestAnt(J + W) = BestAnt(K + 1 - W)
                    BestAnt(K + 1 - W) = Temp
                Next W
                Release = True
                K = J + 1
            End If
        Next K
    Next J
    RouteAnt(0).Covered = RouteAnt(0).Covered + RouteAnt(I).Covered
Next I

If Release = True Then
    NLS = NLS + 1
    
    Call AccDeltaTao(BestAnt(), RouteAnt(), Route(0).Covered / nAnts / 0.1, UBound(RouteAnt) - 1)
    
    For I = 0 To MaxNv
        Route(I) = RouteAnt(I)
    Next I
    
    ReDim Solution(1 To Nodes + MaxNv + 1)
    For I = 1 To Nodes + MaxNv + 1
        Solution(I) = BestAnt(I)
    Next I
    Nv = UBound(RouteAnt)
    
End If

ReDim Preserve BestRoute(0 To MaxNv)

If BestRoute(0).Covered > RouteAnt(0).Covered Then
    For I = 0 To UBound(RouteAnt)
        BestRoute(I) = RouteAnt(I)
    Next I
    ReDim BestSolution(1 To Nodes + MaxNv + 1)
    For I = 1 To Nodes + MaxNv + 1
        BestSolution(I) = BestAnt(I)
    Next I
    BestNv = UBound(RouteAnt)
    
    LastImprove = GenerationNext
    
End If

End Sub


Sub UnInfactibilization()

Dim I1 As Integer
Dim I2 As Integer
Dim I3 As Integer
Dim Release As Boolean
Dim Tabu As Integer
Dim Aux As Integer

ReDim Preserve TabuList(1 To Nodes, 1 To MaxNv)

'Buscar el peor
Worst(1) = 0
For I1 = 1 To Nv
    If Route(I1).Ctrl - Route(I1 - 1).Ctrl > 2 Then
        For I2 = Route(I1 - 1).Ctrl + 1 To Route(I1).Ctrl - 1
            If Worst(1) = 0 Then
                If TabuList(Solution(I2), I1) = False Then
                    Worst(1) = I1
                    Worst(2) = I2
                End If
            Else
                If ((Route(I1).Covered / (Route(I1).Ctrl - Route(I1 - 1).Ctrl)) - ((Route(I1).Covered - Dist(Solution(I2), Solution(I2 - 1)) - Dist(Solution(I2), Solution(I2 + 1)) + Dist(Solution(I2 - 1), Solution(I2 + 1))) / (Route(I1).Ctrl - Route(I1 - 1).Ctrl - 1))) > ((Route(Worst(1)).Covered / (Route(Worst(1)).Ctrl - Route(Worst(1) - 1).Ctrl)) - ((Route(Worst(1)).Covered - Dist(Solution(Worst(2)), Solution(Worst(2) - 1)) - Dist(Solution(Worst(2)), Solution(Worst(2) + 1)) + Dist(Solution(Worst(2) - 1), Solution(Worst(2) + 1))) / (Route(Worst(1)).Ctrl - Route(Worst(1) - 1).Ctrl - 1))) Then
                    If TabuList(Solution(I2), I1) = False Then
                        Worst(1) = I1
                        Worst(2) = I2
                    End If
                End If
            End If
        Next I2
    End If
Next I1

If Worst(1) <> 0 Then
    'Sacarlo de la solución
    Route(Worst(1)).Covered = Route(Worst(1)).Covered - Dist(Solution(Worst(2)), Solution(Worst(2) + 1)) - Dist(Solution(Worst(2)), Solution(Worst(2) - 1)) + Dist(Solution(Worst(2) - 1), Solution(Worst(2) + 1))
    Route(Worst(1)).Cused = Route(Worst(1)).Cused - Dem(Solution(Worst(2)))
    Route(Worst(1)).Tused = Route(Worst(1)).Tused - TimeS(Solution(Worst(2))) - Dist(Solution(Worst(2)), Solution(Worst(2) + 1)) - Dist(Solution(Worst(2)), Solution(Worst(2) - 1)) + Dist(Solution(Worst(2) - 1), Solution(Worst(2) + 1))
    Assigned(Solution(Worst(2))) = False
    Route(0).Cused = Route(0).Cused - Dem(Solution(Worst(2)))
    Tabu = Solution(Worst(2))
    TabuList(Solution(Worst(2)), Worst(1)) = True
    For I1 = Worst(2) To Route(Nv).Ctrl - 1
        Solution(I1) = Solution(I1 + 1)
    Next I1
    For I1 = Worst(1) + 1 To Nv + 1
        Route(I1 - 1).Ctrl = Route(I1 - 1).Ctrl - 1
    Next I1
    Sol = Sol - 1
Else
    Worst(1) = 1
    Aux = Nodes
End If

Do

'quién será mejor para insertar?
Worst(2) = 0
For I1 = 1 To Nodes
    For I2 = Route(Worst(1) - 1).Ctrl + 1 To Route(Worst(1)).Ctrl - 1
        If I1 = Solution(I2) Then Exit For
    Next I2
    If I2 = Route(Worst(1)).Ctrl Then
        For I2 = Route(Worst(1) - 1).Ctrl To Route(Worst(1)).Ctrl - 1
            If Worst(2) = 0 Then
                If Dem(I1) + Route(Worst(1)).Cused <= Capv Then
                    If (Route(Worst(1)).Tused + TimeS(I1) + Dist(I1, Solution(I2)) + Dist(I1, Solution(I2 + 1)) - Dist(Solution(I2), Solution(I2 + 1))) <= TimeC Or TimeC = 0 Then
                        If Release = False Or Assigned(I1) = False Then
                            If Tabu <> I1 Then
                                Worst(2) = I2
                                Worst(3) = I1
                            End If
                        End If
                    End If
                End If
            Else
                If Dist(Solution(Worst(2)), Worst(3)) + Dist(Solution(Worst(2) + 1), Worst(3)) > Dist(Solution(I2), I1) + Dist(Solution(I2 + 1), I1) Then
                    If Dem(I1) + Route(Worst(1)).Cused <= Capv Then
                        If (Route(Worst(1)).Tused + TimeS(I1) + Dist(I1, Solution(I2)) + Dist(I1, Solution(I2 + 1)) - Dist(Solution(I2), Solution(I2 + 1))) <= TimeC Or TimeC = 0 Then
                            If Release = False Or Assigned(I1) = False Then
                                If Tabu <> I1 Then
                                    Worst(2) = I2
                                    Worst(3) = I1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next I2
    End If
Next I1

'Insertémoslo
If Worst(2) <> 0 Then
If Assigned(Worst(3)) = False Then
    Route(Worst(1)).Covered = Route(Worst(1)).Covered - Dist(Solution(Worst(2)), Solution(Worst(2) + 1)) + Dist(Solution(Worst(2)), Worst(3)) + Dist(Solution(Worst(2) + 1), (Worst(3)))
    Route(0).Covered = Route(0).Covered - Dist(Solution(Worst(2)), Solution(Worst(2) + 1)) + Dist(Solution(Worst(2)), Worst(3)) + Dist(Solution(Worst(2) + 1), (Worst(3)))
    Route(Worst(1)).Tused = Route(Worst(1)).Tused + TimeS(Worst(3)) - Dist(Solution(Worst(2)), Solution(Worst(2) + 1)) + Dist(Solution(Worst(2)), Worst(3)) + Dist(Solution(Worst(2) + 1), (Worst(3)))
    Route(Worst(1)).Cused = Route(Worst(1)).Cused + Dem(Worst(3))
    For I1 = Route(Nv).Ctrl To Worst(2) + 1 Step -1
        Solution(I1 + 1) = Solution(I1)
    Next I1
    Solution(Worst(2) + 1) = Worst(3)
    For I1 = Worst(1) To Nv
        Route(I1).Ctrl = Route(I1).Ctrl + 1
    Next I1
    Assigned(Worst(3)) = True
    Route(0).Cused = Route(0).Cused + Dem(Worst(3))
    Sol = Sol + 1
    For I1 = Route(Worst(1) - 1).Ctrl To Route(Worst(1)).Ctrl
        Solution(I1) = Solution(I1)
    Next I1
    Tabu = 0
Else
    For I1 = Route(Worst(1) - 1).Ctrl + 1 To Route(Worst(1)).Ctrl - 1
        If Solution(I1) = Worst(3) Then Exit For
    Next I1
    If I1 = Route(Worst(1)).Ctrl Then
        'Insértelo donde es
        Route(Worst(1)).Covered = Route(Worst(1)).Covered - Dist(Solution(Worst(2)), Solution(Worst(2) + 1)) + Dist(Solution(Worst(2)), Worst(3)) + Dist(Solution(Worst(2) + 1), (Worst(3)))
        Route(0).Covered = Route(0).Covered - Dist(Solution(Worst(2)), Solution(Worst(2) + 1)) + Dist(Solution(Worst(2)), Worst(3)) + Dist(Solution(Worst(2) + 1), (Worst(3)))
        Route(Worst(1)).Tused = Route(Worst(1)).Tused + TimeS(Worst(3)) - Dist(Solution(Worst(2)), Solution(Worst(2) + 1)) + Dist(Solution(Worst(2)), Worst(3)) + Dist(Solution(Worst(2) + 1), (Worst(3)))
        Route(Worst(1)).Cused = Route(Worst(1)).Cused + Dem(Worst(3))
        For I1 = Route(Nv).Ctrl To Worst(2) + 1 Step -1
            Solution(I1 + 1) = Solution(I1)
        Next I1
        Solution(Worst(2) + 1) = Worst(3)
        For I1 = Worst(1) + 1 To Nv + 1
            Route(I1 - 1).Ctrl = Route(I1 - 1).Ctrl + 1
        Next I1
        Tabu = 0
        'Sáquelo de donde no es
        For I1 = 1 To Nv
            For I2 = Route(I1 - 1).Ctrl + 1 To Route(I1).Ctrl - 1
                If Solution(I2) = Worst(3) And I1 <> Worst(1) Then
                    Route(I1).Covered = Route(I1).Covered - Dist(Solution(I2), Solution(I2 + 1)) - Dist(Solution(I2), Solution(I2 - 1)) + Dist(Solution(I2 - 1), Solution(I2 + 1))
                    Route(0).Covered = Route(0).Covered - Dist(Solution(I2), Solution(I2 + 1)) - Dist(Solution(I2), Solution(I2 - 1)) + Dist(Solution(I2 - 1), Solution(I2 + 1))
                    Route(I1).Tused = Route(I1).Tused - TimeS(Solution(I2)) - Dist(Solution(I2), Solution(I2 + 1)) - Dist(Solution(I2), Solution(I2 - 1)) + Dist(Solution(I2 - 1), Solution(I2 + 1))
                    Route(I1).Cused = Route(I1).Cused - Dem(Solution(I2))
                    For I3 = I2 To Route(Nv).Ctrl - 1
                        Solution(I3) = Solution(I3 + 1)
                    Next I3
                    For I3 = I1 To Nv
                        Route(I3).Ctrl = Route(I3).Ctrl - 1
                    Next I3
                    I1 = Nv
                    Exit For
                End If
            Next I2
        Next I1
    End If
End If
Else
    If Release = False Then
        Release = True
        Worst(1) = 1
    Else
        Worst(1) = Worst(1) + 1
    End If
End If

If Release = True Then
    If MaxNv > Nv Then
        If ((Demand - Route(0).Cused) / (Capv * (MaxNv - Nv))) < 1 Then
            Exit Do
        End If
    End If
End If

Loop Until Worst(1) > Nv Or Sol = Nodes

If Aux = Nodes Then Worst(1) = 0

End Sub


Sub UpdateETA()

Dim I As Integer
Dim J As Integer

SumFirst = 0
'Ultima actividad seleccionada
'Eta(First(Chosen), 0, 0) = Ctrl(Nv + 1) + 1
First(Chosen) = 0
'Eta(First(Chosen), Solution(Eta(Chosen, 0, 0) - 1), Solution(Eta(Chosen, 0, 0))) = Eta(Chosen, Solution(Eta(Chosen, 0, 0) - 1), Solution(Eta(Chosen, 0, 0)))
'First(40) = First(12)
'I = (1 / (Dist(Chosen, 0) + Dist(Chosen, 1)) / Max)
For I = 1 To Nodes
    If Assigned(I) = False Then
        'Restriccion de capacidad
        If Dem(I) + Route(Nv).Cused <= Capv Then
            'Nearest() = punto en el que se va apegar el nodo seleccionado
            First(I) = I
            Nearest(I, 2) = Route(Nv).Ctrl
            For J = Route(Nv).Ctrl - 1 To Route(Nv - 1).Ctrl + 1 Step -1
                If (Dist(I, Solution(J)) + Dist(I, Solution(J - 1)) - Dist(Solution(J), Solution(J - 1))) < (Dist(I, Solution(Nearest(I, 2))) + Dist(I, Solution(Nearest(I, 2) - 1)) - Dist(Solution(Nearest(I, 2)), Solution(Nearest(I, 2) - 1))) Then
                    Nearest(I, 2) = J
                End If
            Next J
            Nearest(I, 1) = Solution(Nearest(I, 2))
            Eta(I, 0, 0) = Nearest(I, 2)
            
            If TimeC <> 0 And (Route(Nv).Tused + TimeS(I) + Dist(I, Solution(Nearest(I, 2) - 1)) + Dist(I, Solution(Nearest(I, 2) + 1)) - Dist(Solution(Nearest(I, 2) - 1), Solution(Nearest(I, 2) + 1))) > TimeC Then
                'Nodos infactibles por tiempo
                First(I) = 0
            End If
            
            Second(I) = First(I)
            
            If (Route(Nv).Covered / (Route(Nv).Ctrl - Route(Nv - 1).Ctrl)) < ((Route(Nv).Covered + Dist(I, Solution(Nearest(I, 2) - 1)) + Dist(I, Solution(Nearest(I, 2) + 1)) - Dist(Solution(Nearest(I, 2) - 1), Solution(Nearest(I, 2) + 1))) / (Route(Nv).Ctrl - Route(Nv - 1).Ctrl + 1)) Then
                First(I) = 0
            End If
            
            If First(I) <> 0 Then SumFirst = SumFirst + 1
            
        Else
            'Nodos infactibles por capacidad
            First(I) = 0
        End If
    End If
Next I

If SumFirst = 0 Then
    If MaxNv = Nv Then
        For I = 1 To Nodes
            If Second(I) <> 0 Then
                First(I) = Second(I)
                SumFirst = SumFirst + 1
            End If
        Next I
    ElseIf ((Demand - DemandA - Route(Nv).Cused) / (MaxNv - Nv)) > Capv Then
        For I = 1 To Nodes
            If Second(I) <> 0 Then
                First(I) = Second(I)
                SumFirst = SumFirst + 1
            End If
        Next I
    End If
End If

End Sub


Sub UpdateTAO()

Dim I1 As Integer
Dim I2 As Integer

Max = 0
For I1 = 1 To Nodes
    For I2 = I1 + 1 To Nodes
        If DeltaTao(I1, I2) > Max Then Max = DeltaTao(I1, I2)
    Next I2
Next I1
If Max <> 0 Then
    For I1 = 1 To Nodes
        For I2 = 1 To Nodes
            DeltaTao(I1, I2) = DeltaTao(I1, I2) / Max
        Next I2
    Next I1

    Max = 0
    For I1 = 1 To Nodes
        For I2 = I1 + 1 To Nodes
            Tao(I1, I2) = Rho * Tao(I1, I2) + DeltaTao(I1, I2)
            Tao(I2, I1) = Rho * Tao(I2, I1) + DeltaTao(I2, I1)
            If Tao(I1, I2) > Max Then Max = Tao(I1, I2)
        Next I2
    Next I1
    For I1 = 1 To Nodes
        For I2 = 1 To Nodes
            Tao(I1, I2) = Tao(I1, I2) / Max
        Next I2
    Next I1
End If

'Reinicializar DeltaTao
ReDim DeltaTao(1 To Nodes, 1 To Nodes)

End Sub

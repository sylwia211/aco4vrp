Attribute VB_Name = "Procedimientos"
Option Explicit


Sub AccDeltaTao(Upshot() As Integer, Check() As Integer, ObjFunct As Double, Vehicle As Integer)

Dim I1 As Integer
Dim I2 As Integer
Dim I3 As Integer

'Numero de vehiculos
For I1 = 1 To Vehicle
    'Se buscan todos los nodos de dicho vehiculo
    For I2 = Check(I1) + 1 To Check(I1 + 1) - 1
        For I3 = I2 + 1 To Check(I1 + 1) - 1
            DeltaTao(Upshot(I3), Upshot(I2)) = DeltaTao(Upshot(I3), Upshot(I2)) + 1 / ObjFunct
            DeltaTao(Upshot(I2), Upshot(I3)) = DeltaTao(Upshot(I2), Upshot(I3)) + 1 / ObjFunct
        Next I3
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
        
        'Colectar resultados
        Summary(NProblem, 1) = BestCovered(0)
        Summary(NProblem, 2) = Nv
        Summary(NProblem, 3) = Timer - Time
    
    Next NProblem
    
    'Generar resumen de resultados
    MultiPrint

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
'STARTSTARTSTARTSTARTSTARTSTARTSTARTSTARTSTARTSTARTSTARTSTART
CusedAcc = 0
'ENDENDENDENDENDENDENDENDENDENDENDENDENDENDENDENDENDENDENDEND

ReDim Ctrl(1 To Nv + 1)
ReDim Solution(1 To Nodes + MaxNv + 1)
ReDim Assigned(1 To Nodes)
ReDim Nearest(1 To Nodes, 1 To 2)
Ctrl(1) = 1
Sol = 0
Solution(Ctrl(Nv)) = 0

'Ciclo para generar la ruta de cada vehiculo
Do While Sol < Nodes

    'Redimension de las variables
    ReDim Preserve Ctrl(1 To Nv + 1)   '*****MOVER PARA EL FINAL DEL CICLO CUANDO ESTE LISTO***** Color Naranja
    ReDim Preserve Covered(0 To Nv)
    ReDim Preserve Cused(1 To Nv)
    ReDim Preserve Tused(1 To Nv)
    
    'Búsqueda del nodo más lejano
    Chosen = 0
    For I = 1 To Nodes
        If Assigned(I) = False Then
            If Dist(I, 0) > Dist(Chosen, 0) Then Chosen = I
        End If
    Next I
    'Asignacion del primer nodo de la ruta (el mas lejano posible)
    Solution(Ctrl(Nv) + 1) = Chosen
    Furthest = Chosen
    'If assigned(chosen) = True Then
    '    chosen = chosen
    'end if
    Assigned(Chosen) = True
    Solution(Ctrl(Nv) + 2) = 0
    Ctrl(Nv + 1) = Ctrl(Nv) + 2
    Sol = Sol + 1
    Covered(Nv) = Dist(0, Chosen) * 2
    Tused(Nv) = Covered(Nv) + TimeS(Chosen)
    Cused(Nv) = Dem(Chosen)
    Nearest(Chosen, 2) = Ctrl(Nv) + 1
    Eta(Chosen, 0, 0) = Nearest(Chosen, 2)

    'Actualizacion del parametro ETA
    Call UpdateETA
    
    'Ciclo para completar la ruta del vehiculo
    Do While Cused(Nv) < Capv And Chosen <> 0   'Mientras haya capacidad en el vehiculo

        'Sum = denominador para el calculo de probabilidades
        Sum = 0
        For I = 1 To Nodes
            Sum = Sum + ((Weight * Tao(Furthest, I)) * ((1 - Weight) * Eta(First(I), Solution(Eta(I, 0, 0) - 1), Solution(Eta(I, 0, 0)))))
        Next I
        
        Chosen = 0
        
        If Sum <> 0 Then
            Prob = 0
            Randomize
            Random = Rnd
            
            'Seleccion de un nodo para la ruta
            For I = 1 To Nodes
                Prob = Prob + (((Weight * Tao(Furthest, I)) * ((1 - Weight) * Eta(First(I), Solution(Eta(I, 0, 0) - 1), Solution(Eta(I, 0, 0))))) / Sum)     'Probabilidad acumulada
                If Prob > Random Then
                    If Assigned(I) = False Then
                        If Cused(Nv) + Dem(I) <= Capv Then
                            Chosen = I
                            Exit For
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
        Ctrl(Nv + 1) = Ctrl(Nv + 1) + 1
        For I = Ctrl(Nv + 1) To Nearest(Chosen, 2) + 1 Step -1
            Solution(I) = Solution(I - 1)
        Next I
        Solution(Nearest(Chosen, 2)) = Chosen
        Assigned(Chosen) = True
        Cused(Nv) = Cused(Nv) + Dem(Chosen)
        Covered(Nv) = Covered(Nv) + Dist(Chosen, Solution(Nearest(Chosen, 2) - 1)) + Dist(Chosen, Solution(Nearest(Chosen, 2) + 1)) - Dist(Solution(Nearest(Chosen, 2) - 1), Solution(Nearest(Chosen, 2) + 1))
        Tused(Nv) = Tused(Nv) + TimeS(Chosen) + Dist(Chosen, Solution(Nearest(Chosen, 2) - 1)) + Dist(Chosen, Solution(Nearest(Chosen, 2) + 1)) - Dist(Solution(Nearest(Chosen, 2) - 1), Solution(Nearest(Chosen, 2) + 1))
        
        'Actualizacion del parametro ETA
        Call UpdateETA
        
10
        
    Loop
    
'STARTSTARTSTARTSTARTSTARTSTARTSTARTSTARTSTARTSTARTSTARTSTART
    CusedAcc = CusedAcc + Cused(Nv)
    
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
            If ((Demand - CusedAcc) / (Capv * (MaxNv - Nv))) > 1 Then
                Do
                    Call UnInfactibilization
                    If Worst(1) = 0 Then
                        K = K - 1
                        Sol = Nodes
                        Exit Do
                    End If
                Loop While ((Demand - CusedAcc) / (Capv * (MaxNv - Nv))) > 1
                Nv = Nv + 1
            Else
                Nv = Nv + 1
            End If
        End If
    End If
'ENDENDENDENDENDENDENDENDENDENDENDENDENDENDENDENDENDENDENDEND
    
Loop

'Calculo funcion objetivo
Covered(0) = 0
For I = 1 To Nv
    Covered(0) = Covered(0) + Covered(I)
Next I

'For I1 = 1 To Nodes + Nv + 1
'    Solution(I1) = Solution(I1)
'Next I1

End Sub


Sub Plot()

'Contadores y Variables locales
Dim I As Integer

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

End Sub


Sub Exchange()



End Sub


Sub Factibilization()

Dim I1 As Integer   'Cantidad nodos no asignados
Dim I2 As Integer   'Busca los nodos no asignados
Dim I3 As Integer   'Busca en cada ruta
Dim I4 As Integer   'Busca la ubicación en una solución que sea lo mejor (factible)

For I1 = Sol + 1 To Nodes
    For I2 = 1 To Nodes
        If Assigned(I2) = False Then
            Nearest(I2, 2) = 0
            For I3 = 1 To MaxNv
                If Capv >= Cused(I3) + Dem(I2) Then
                    For I4 = Ctrl(I3) To Ctrl(I3 + 1) - 1
                        If (Tused(I3) + TimeS(I2) + Dist(I2, Solution(I4)) + Dist(I2, Solution(I4 + 1))) < TimeC Or TimeC = 0 Then
                            If Nearest(I2, 2) = 0 Then
                                Nearest(I2, 1) = Solution(I4)
                                Nearest(I2, 2) = I4
                            Else
                                If (Dist(I2, Solution(I4)) + Dist(I2, Solution(I4 + 1))) < (Dist(I2, Solution(Nearest(I2, 2))) + Dist(I2, Solution(Nearest(I2, 2) + 1))) Then
                                    Nearest(I2, 1) = Solution(I4)
                                    Nearest(I2, 2) = I4
                                End If
                            End If
                        End If
                    Next I4
                End If
            Next I3
            If Nearest(I2, 2) = 0 Then
                Sol = 0
                I1 = Nodes
                Exit For
            Else
                For I3 = 1 To MaxNv
                    If Nearest(I2, 2) < Ctrl(I3 + 1) Then
                        For I4 = Ctrl(Nv + 1) To Nearest(I2, 2) + 1 Step -1
                            Solution(I4 + 1) = Solution(I4)
                        Next I4
                        Solution(Nearest(I2, 2) + 1) = I2
                        For I4 = I3 + 1 To Nv + 1
                            Ctrl(I4) = Ctrl(I4) + 1
                        Next I4
                        Cused(I3) = Cused(I3) + Dem(I2)
                        Covered(I3) = Covered(I3) + Dist(I2, Solution(Nearest(I2, 2))) + Dist(I2, Solution(Nearest(I2, 2) + 2)) - Dist(Solution(Nearest(I2, 2)), Solution(Nearest(I2, 2) + 2))
                        Tused(I3) = Tused(I3) + TimeS(I2) + Dist(I2, Solution(Nearest(I2, 2))) + Dist(I2, Solution(Nearest(I2, 2) + 2)) - Dist(Solution(Nearest(I2, 2)), Solution(Nearest(I2, 2) + 2))
                        Covered(0) = Covered(0) + Dist(I2, Solution(Nearest(I2, 2))) + Dist(I2, Solution(Nearest(I2, 2) + 2)) - Dist(Solution(Nearest(I2, 2)), Solution(Nearest(I2, 2) + 2))
                        I1 = I1 + 1
                        Exit For
                    End If
                Next I3
            End If
        End If
        If I1 > Nodes Then Exit For
    Next I2
Next I1

End Sub


Sub Improved()

Dim I1 As Integer

ReDim Preserve BestCovered(0 To UBound(Covered))
ReDim Preserve BestCused(1 To UBound(Cused))

If BestCovered(0) > Covered(0) Or BestCovered(0) = 0 Then
    BestCovered(0) = Covered(0)
    For I1 = 1 To UBound(Covered)
        BestCovered(I1) = Covered(I1)
    Next I1
    For I1 = 1 To UBound(Cused)
        BestCused(I1) = Cused(I1)
    Next I1
    ReDim BestSolution(1 To Nodes + Nv + 1)
    ReDim BestAnt(0 To Nodes + Nv + 1)
    ReDim CtrlAnt(1 To Nv + 1)
    For I1 = 1 To Nodes + Nv + 1
        BestSolution(I1) = Solution(I1)
        BestAnt(I1) = Solution(I1)
    Next I1
    BestAnt(0) = BestCovered(0)
    For I1 = 1 To Nv + 1
        CtrlAnt(I1) = Ctrl(I1)
    Next I1
    BestNv = UBound(Ctrl)
ElseIf UBound(BestAnt) = 0 Then
    ReDim BestAnt(0 To Nodes + MaxNv + 1)
    ReDim CtrlAnt(1 To Nv + 1)
    For I1 = 1 To Nodes + Nv + 1
        BestAnt(I1) = Solution(I1)
    Next I1
    BestAnt(0) = Covered(0)
    For I1 = 1 To Nv + 1
        CtrlAnt(I1) = Ctrl(I1)
    Next I1
ElseIf BestAnt(0) > Covered(0) Then
    ReDim CtrlAnt(1 To Nv + 1)
    For I1 = 1 To Nodes + Nv + 1
        BestAnt(I1) = Solution(I1)
    Next I1
    BestAnt(0) = Covered(0)
    For I1 = 1 To Nv + 1
        CtrlAnt(I1) = Ctrl(I1)
    Next I1
End If

End Sub


Sub Insertion()

Dim I1 As Integer
Dim I2 As Integer
Dim I3 As Integer
Dim I4 As Integer

For I1 = 1 To UBound(CtrlAnt)
    For I2 = CtrlAnt(I1) To CtrlAnt(I1 + 1)
        For I3 = 1 To UBound(CtrlAnt)
            If I1 <> I3 Then
                For I4 = CtrlAnt(I3) To CtrlAnt(I3 + 1)
                    IF CUSED
                Next I4
            End If
        Next I3
    Next I2
Next I1

End Sub


Sub MultiPrint()

Dim I1 As Integer
Dim Text As String

Open App.Path & "\Resultados\Summary.txt" For Output As #3

Print #3, vbTab & "Costo" & vbTab & "Nv" & vbTab & "Tiempo"

For I1 = 1 To 14

    Print #3, I1 & vbTab & Int(Summary(I1, 1) * 100) / 100 & vbTab & Summary(I1, 2) & vbTab & Int(Summary(I1, 3) * 1000) / 1000
    
Next I1

Close #3

End Sub


Sub MultiSinglePrint()

Dim I1 As Integer
Dim I2 As Integer
Dim Text As String

If NProblem < 10 Then
    Open App.Path & "\Resultados\solve0" & NProblem & ".txt" For Output As #2
Else
    Open App.Path & "\Resultados\solve" & NProblem & ".txt" For Output As #2
End If

I2 = 1
Print #2, "Costo de la solución:"
Print #2, BestCovered(0)
Print #2,
For I1 = 1 To Nv
    Text = "1" & vbTab & I1 & vbTab & Int(BestCovered(I1) * 1000) / 1000 & vbTab & BestCused(I1) & vbTab
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

End Sub


Sub MultiVRP()

Dim K As Integer

'Lectura de datos
Call Reading_Multi

'Inicialización de variables y parámetros
Call Parameters

'Generaciones ' Iteraciones
For GenerationNext = 1 To nGen

    ReDim BestAnt(0 To 0)

    For K = 1 To nAnts
    
        ReDim TabuList(1 To Nodes, 1 To MaxNv)
        
        'Colonia
        Call Ant

        If Worst(1) <> 0 Then
            
            Call Improved

            'Acumular DeltaTao
            Call AccDeltaTao(Solution(), Ctrl(), Covered(0), Nv)
            
        End If
        
    Next K
    
'STARTSTARTSTARTSTARTSTARTSTARTSTARTSTARTSTARTSTARTSTARTSTART
    'BÚSQUEDA LOCAL
    If UBound(BestAnt) <> 0 Then
        Call Two_Opt
    End If
'ENDENDENDENDENDENDENDENDENDENDENDENDENDENDENDENDENDENDENDEND
    
    'Actulización de Tao
    Call UpdateTAO

Next GenerationNext

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
Weight = Val(Form1.Text3)

'Dimensión de la solución
ReDim Eta(0 To Nodes, 0 To Nodes, 0 To Nodes)
'ReDim Save(1 To Nodes, 1 To Nodes)
ReDim Solution(0 To Nodes + MaxNv + 1)
ReDim Tao(1 To Nodes, 1 To Nodes)
ReDim DeltaTao(1 To Nodes, 1 To Nodes)
ReDim First(1 To Nodes)

ReDim BestCused(1 To MaxNv)
ReDim BestCovered(0 To MaxNv)
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

Open App.Path & "\Resultados\solve" & Problem & ".txt" For Output As #2

I2 = 1
Print #2, "Costo de la solución:"
Print #2, BestCovered(0)
Print #2,
For I1 = 1 To MaxNv
    Text = "1" & vbTab & I1 & vbTab & Int(BestCovered(I1) * 1000) / 1000 & vbTab & BestCused(I1) & vbTab
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

End Sub


Sub SingleVRP()

'Lectura de datos
Call Reading

'Inicialización de variables y parámetros
Call Parameters

'Generaciones ' Iteraciones
For GenerationNext = 1 To nGen

    ReDim BestAnt(0 To 0)

    For K = 1 To nAnts
    
        ReDim TabuList(1 To Nodes, 1 To MaxNv)
    
        'Colonia
        Call Ant
            
        If Worst(1) <> 0 Then
            Call Improved

            'Acumular DeltaTao
            Call AccDeltaTao(Solution(), Ctrl(), Covered(0), Nv)

        End If
        
    Next K
    
    'BÚSQUEDA LOCAL
    If UBound(BestAnt) <> 0 Then
        Call TwoOpt
        Call Insertion
        Call Exchange
    End If
    
    'Actulización de Tao
    Call UpdateTAO

Next GenerationNext

'Mostrar resultados
MsgBox BestCovered(0) & vbCrLf & Nv
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

For I = 1 To UBound(CtrlAnt) - 1
    For J = CtrlAnt(I) To CtrlAnt(I + 1) - 1
        For K = J + 2 To CtrlAnt(I + 1) - 1
            If Dist(BestAnt(J), BestAnt(J + 1)) + Dist(BestAnt(K), BestAnt(K + 1)) > Dist(BestAnt(J), BestAnt(K)) + Dist(BestAnt(J + 1), BestAnt(K + 1)) Then
                BestAnt(0) = BestAnt(0) + (-Dist(BestAnt(J), BestAnt(J + 1)) - Dist(BestAnt(K), BestAnt(K + 1)) + Dist(BestAnt(J), BestAnt(K)) + Dist(BestAnt(J + 1), BestAnt(K + 1)))
                For W = 1 To Int((K - J + 1) / 2)
                    Temp = BestAnt(J + W)
                    BestAnt(J + W) = BestAnt(K + 1 - W)
                    BestAnt(K + 1 - W) = Temp
                    Release = True
                Next W
                K = J + 1
            End If
        Next K
    Next J
Next I

If Release = True Then
    Call AccDeltaTao(BestAnt(), CtrlAnt(), nAnts * 2 * BestAnt(0), UBound(CtrlAnt) - 1)
End If

ReDim Preserve BestCovered(0 To UBound(Covered))
ReDim Preserve BestCused(1 To UBound(Cused))

If BestCovered(0) > BestAnt(0) Then
    BestCovered(0) = BestAnt(0)
    For I = 1 To UBound(Covered)
        BestCovered(I) = Covered(I)
    Next I
    For I = 1 To UBound(Cused)
        BestCused(I) = Cused(I)
    Next I
    ReDim BestSolution(1 To Nodes + Nv + 1)
    For I = 1 To Nodes + Nv + 1
        BestSolution(I) = Solution(I)
    Next I
    BestNv = UBound(CtrlAnt) - 1
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
    If Ctrl(I1 + 1) - Ctrl(I1) > 2 Then
        For I2 = Ctrl(I1) + 1 To Ctrl(I1 + 1) - 1
            If Worst(1) = 0 Then
                If TabuList(Solution(I2), I1) = False Then
                    Worst(1) = I1
                    Worst(2) = I2
                End If
            Else
                If ((Covered(I1) / (Ctrl(I1 + 1) - Ctrl(I1))) - ((Covered(I1) - Dist(Solution(I2), Solution(I2 - 1)) - Dist(Solution(I2), Solution(I2 + 1)) + Dist(Solution(I2 - 1), Solution(I2 + 1))) / (Ctrl(I1 + 1) - Ctrl(I1) - 1))) > ((Covered(Worst(1)) / (Ctrl(Worst(1) + 1) - Ctrl(Worst(1)))) - ((Covered(Worst(1)) - Dist(Solution(Worst(2)), Solution(Worst(2) - 1)) - Dist(Solution(Worst(2)), Solution(Worst(2) + 1)) + Dist(Solution(Worst(2) - 1), Solution(Worst(2) + 1))) / (Ctrl(Worst(1) + 1) - Ctrl(Worst(1)) - 1))) Then
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
    Covered(Worst(1)) = Covered(Worst(1)) - Dist(Solution(Worst(2)), Solution(Worst(2) + 1)) - Dist(Solution(Worst(2)), Solution(Worst(2) - 1)) + Dist(Solution(Worst(2) - 1), Solution(Worst(2) + 1))
    Cused(Worst(1)) = Cused(Worst(1)) - Dem(Solution(Worst(2)))
    Tused(Worst(1)) = Tused(Worst(1)) - TimeS(Solution(Worst(2))) - Dist(Solution(Worst(2)), Solution(Worst(2) + 1)) - Dist(Solution(Worst(2)), Solution(Worst(2) - 1)) + Dist(Solution(Worst(2) - 1), Solution(Worst(2) + 1))
    Assigned(Solution(Worst(2))) = False
    CusedAcc = CusedAcc - Dem(Solution(Worst(2)))
    Tabu = Solution(Worst(2))
    TabuList(Solution(Worst(2)), Worst(1)) = True
    For I1 = Worst(2) To Ctrl(Nv + 1) - 1
        Solution(I1) = Solution(I1 + 1)
    Next I1
    For I1 = Worst(1) + 1 To Nv + 1
        Ctrl(I1) = Ctrl(I1) - 1
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
    For I2 = Ctrl(Worst(1)) + 1 To Ctrl(Worst(1) + 1) - 1
        If I1 = Solution(I2) Then Exit For
    Next I2
    If I2 = Ctrl(Worst(1) + 1) Then
        For I2 = Ctrl(Worst(1)) To Ctrl(Worst(1) + 1) - 1
            If Worst(2) = 0 Then
                If Dem(I1) + Cused(Worst(1)) <= Capv Then
                    If (Tused(Worst(1)) + TimeS(I1) + Dist(I1, Solution(I2)) + Dist(I1, Solution(I2 + 1)) - Dist(Solution(I2), Solution(I2 + 1))) <= TimeC Or TimeC = 0 Then
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
                    If Dem(I1) + Cused(Worst(1)) <= Capv Then
                        If (Tused(Worst(1)) + TimeS(I1) + Dist(I1, Solution(I2)) + Dist(I1, Solution(I2 + 1)) - Dist(Solution(I2), Solution(I2 + 1))) <= TimeC Or TimeC = 0 Then
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

'insertémoslo
If Worst(2) <> 0 Then
If Assigned(Worst(3)) = False Then
    Covered(Worst(1)) = Covered(Worst(1)) - Dist(Solution(Worst(2)), Solution(Worst(2) + 1)) + Dist(Solution(Worst(2)), Worst(3)) + Dist(Solution(Worst(2) + 1), (Worst(3)))
    Covered(0) = Covered(0) - Dist(Solution(Worst(2)), Solution(Worst(2) + 1)) + Dist(Solution(Worst(2)), Worst(3)) + Dist(Solution(Worst(2) + 1), (Worst(3)))
    Cused(Worst(1)) = Cused(Worst(1)) + Dem(Worst(3))
    Tused(Worst(1)) = Tused(Worst(1)) + TimeS(Worst(3)) - Dist(Solution(Worst(2)), Solution(Worst(2) + 1)) + Dist(Solution(Worst(2)), Worst(3)) + Dist(Solution(Worst(2) + 1), (Worst(3)))
    For I1 = Ctrl(Nv + 1) To Worst(2) + 1 Step -1
        Solution(I1 + 1) = Solution(I1)
    Next I1
    Solution(Worst(2) + 1) = Worst(3)
    For I1 = Worst(1) + 1 To Nv + 1
        Ctrl(I1) = Ctrl(I1) + 1
    Next I1
    Assigned(Worst(3)) = True
    CusedAcc = CusedAcc + Dem(Worst(3))
    Sol = Sol + 1
    For I1 = Ctrl(Worst(1)) To Ctrl(Worst(1) + 1)
        Solution(I1) = Solution(I1)
    Next I1
    Tabu = 0
Else
    For I1 = Ctrl(Worst(1)) + 1 To Ctrl(Worst(1) + 1) - 1
        If Solution(I1) = Worst(3) Then Exit For
    Next I1
    If I1 = Ctrl(Worst(1) + 1) Then
        'Insertelo donde es
        Covered(Worst(1)) = Covered(Worst(1)) - Dist(Solution(Worst(2)), Solution(Worst(2) + 1)) + Dist(Solution(Worst(2)), Worst(3)) + Dist(Solution(Worst(2) + 1), (Worst(3)))
        Covered(0) = Covered(0) - Dist(Solution(Worst(2)), Solution(Worst(2) + 1)) + Dist(Solution(Worst(2)), Worst(3)) + Dist(Solution(Worst(2) + 1), (Worst(3)))
        Cused(Worst(1)) = Cused(Worst(1)) + Dem(Worst(3))
        Tused(Worst(1)) = Tused(Worst(1)) + TimeS(Worst(3)) - Dist(Solution(Worst(2)), Solution(Worst(2) + 1)) + Dist(Solution(Worst(2)), Worst(3)) + Dist(Solution(Worst(2) + 1), (Worst(3)))
        For I1 = Ctrl(Nv + 1) To Worst(2) + 1 Step -1
            Solution(I1 + 1) = Solution(I1)
        Next I1
        Solution(Worst(2) + 1) = Worst(3)
        For I1 = Worst(1) + 1 To Nv + 1
            Ctrl(I1) = Ctrl(I1) + 1
        Next I1
        Tabu = 0
        'Sáquelo de donde no es
        For I1 = 1 To Nv
            For I2 = Ctrl(I1) + 1 To Ctrl(I1 + 1) - 1
                If Solution(I2) = Worst(3) And I1 <> Worst(1) Then
                    Covered(I1) = Covered(I1) - Dist(Solution(I2), Solution(I2 + 1)) - Dist(Solution(I2), Solution(I2 - 1)) + Dist(Solution(I2 - 1), Solution(I2 + 1))
                    Covered(0) = Covered(0) - Dist(Solution(I2), Solution(I2 + 1)) - Dist(Solution(I2), Solution(I2 - 1)) + Dist(Solution(I2 - 1), Solution(I2 + 1))
                    Cused(I1) = Cused(I1) - Dem(Solution(I2))
                    Tused(I1) = Tused(I1) - TimeS(Solution(I2)) - Dist(Solution(I2), Solution(I2 + 1)) - Dist(Solution(I2), Solution(I2 - 1)) + Dist(Solution(I2 - 1), Solution(I2 + 1))
                    For I3 = I2 To Ctrl(Nv + 1) - 1
                        Solution(I3) = Solution(I3 + 1)
                    Next I3
                    For I3 = I1 + 1 To Nv + 1
                        Ctrl(I3) = Ctrl(I3) - 1
                    Next I3
                    I1 = Nv
                    Exit For
                End If
            Next I2
        Next I1
    Else
    'TABU?????????????????????????????
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
        If ((Demand - CusedAcc) / (Capv * (MaxNv - Nv))) < 1 Then
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

'Ultima actividad seleccionada
'Eta(First(Chosen), 0, 0) = Ctrl(Nv + 1) + 1
First(Chosen) = 0
'Eta(First(Chosen), Solution(Eta(Chosen, 0, 0) - 1), Solution(Eta(Chosen, 0, 0))) = Eta(Chosen, Solution(Eta(Chosen, 0, 0) - 1), Solution(Eta(Chosen, 0, 0)))
'First(40) = First(12)
'I = (1 / (Dist(Chosen, 0) + Dist(Chosen, 1)) / Max)
For I = 1 To Nodes
    If Assigned(I) = False Then
        'Restriccion de capacidad
        If Dem(I) + Cused(Nv) <= Capv Then
            'Nearest() = punto en el que se va apegar el nodo seleccionado
            First(I) = I
            Nearest(I, 2) = Ctrl(Nv + 1)
            For J = Ctrl(Nv + 1) - 1 To Ctrl(Nv) + 1 Step -1
                If (Dist(I, Solution(J)) + Dist(I, Solution(J - 1)) - Dist(Solution(J), Solution(J - 1))) < (Dist(I, Solution(Nearest(I, 2))) + Dist(I, Solution(Nearest(I, 2) - 1)) - Dist(Solution(Nearest(I, 2)), Solution(Nearest(I, 2) - 1))) Then
                    Nearest(I, 2) = J
                End If
            Next J
            Nearest(I, 1) = Solution(Nearest(I, 2))
            Eta(I, 0, 0) = Nearest(I, 2)
            
            If TimeC <> 0 And TimeS(I) + Tused(Nv) + Dist(I, Solution(Nearest(I, 2) - 1)) + Dist(I, Solution(Nearest(I, 2) + 1)) - Dist(Solution(Nearest(I, 2) - 1), Solution(Nearest(I, 2) + 1)) > TimeC Then
                'Nodos infactibles por tiempo
                First(I) = 0
            End If
            
            If (Covered(Nv) / (Ctrl(Nv + 1) - Ctrl(Nv))) < ((Covered(Nv) + Dist(I, Solution(Nearest(I, 2) - 1)) + Dist(I, Solution(Nearest(I, 2) + 1)) - Dist(Solution(Nearest(I, 2) - 1), Solution(Nearest(I, 2) + 1))) / (Ctrl(Nv + 1) - Ctrl(Nv) + 1)) Then
                First(I) = 0
            End If
            
        Else
            'Nodos infactibles por capacidad
            First(I) = 0
        End If
    End If
Next I

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
            Tao(I1, I2) = Weight * Tao(I1, I2) + DeltaTao(I1, I2)
            Tao(I2, I1) = Weight * Tao(I2, I1) + DeltaTao(I2, I1)
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

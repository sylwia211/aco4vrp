Attribute VB_Name = "Procedimientos"
Option Explicit


Sub AccDeltaTao()

Dim I1 As Integer
Dim I2 As Integer
Dim I3 As Integer

'Numero de vehiculos
For I1 = 1 To Nv
    'Se buscan todos los nodos de dicho vehiculo
    For I2 = Ctrl(I1) + 1 To Ctrl(I1 + 1) - 1
        For I3 = I2 + 1 To Ctrl(I1 + 1) - 1
            DeltaTao(Solution(I3), Solution(I2)) = DeltaTao(Solution(I3), Solution(I2)) + 1 / Covered(0)
            DeltaTao(Solution(I2), Solution(I3)) = DeltaTao(Solution(I2), Solution(I3)) + 1 / Covered(0)
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
        Summary(NProblem, 1) = Covered(0)
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

    'Actualizacion del parametro ETA
    Call UpdateETA
    
    'Ciclo para completar la ruta del vehiculo
    Do While Cused(Nv) < Capv And Chosen <> 0   'Mientras haya capacidad en el vehiculo

        'Sum = denominador para el calculo de probabilidades
        Sum = 0
        For I = 1 To Nodes
            Sum = Sum + ((Weight * Tao(Eta(0, I), I)) * ((1 - Weight) * Eta(Eta(0, I), I)))
        Next I
        
        Chosen = 0
        
        If Sum <> 0 Then
            Prob = 0
            Randomize
            Random = Rnd
            
            'Seleccion de un nodo para la ruta
            For I = 1 To Nodes
                Prob = Prob + (((Weight * Tao(Eta(0, I), I)) * ((1 - Weight) * Eta(Eta(0, I), I))) / Sum)     'Probabilidad acumulada
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
        If Nearest(Chosen) = Solution(Ctrl(Nv + 1) - 1) Then
            Solution(Ctrl(Nv + 1) - 1) = Chosen
        Else
            For I = Ctrl(Nv + 1) - 2 To Ctrl(Nv) + 1 Step -1
                Solution(I + 1) = Solution(I)
            Next I
            Solution(Ctrl(Nv) + 1) = Chosen
        End If
        Assigned(Chosen) = True
        Cused(Nv) = Cused(Nv) + Dem(Chosen)
        Covered(Nv) = Covered(Nv) + Dist(Chosen, 0) + Dist(Chosen, Nearest(Chosen)) - Dist(Nearest(Chosen), 0)
        Tused(Nv) = Tused(Nv) + TimeS(Chosen) + Dist(Chosen, 0) + Dist(Chosen, Nearest(Chosen)) - Dist(Nearest(Chosen), 0)
        
        'Actualizacion del parametro ETA
        Call UpdateETA
        
10
        
    Loop
    
    If Sol < Nodes Then
        Nv = Nv + 1
        If Nv > MaxNv Then
            ReDim Preserve Solution(1 To Nodes + Nv + 1)
        End If
    End If
    
Loop

'Calculo funcion objetivo
Covered(0) = 0
For I = 1 To Nv
    Covered(0) = Covered(0) + Covered(I)
Next I

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
    For I1 = 1 To Nodes + Nv + 1
        BestSolution(I1) = Solution(I1)
    Next I1
End If

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
    Text = Text & vbTab & 0
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

    For K = 1 To nAnts
    
        'Colonia
        Call Ant
        Call Improved
        
        'Acumular DeltaTao
        Call AccDeltaTao
        
    Next K
    
    'Actulización de Tao
    Call Update_TAO

Next GenerationNext

'Generar archivo de resultados
Call MultiSinglePrint

End Sub


Sub Parameters()

'Contadores propios del procedimiento
Dim I As Integer
Dim J As Integer

'Variables Auxiliares
Dim Max As Double

'Parámetros
nGen = Val(Form1.Text1)
nAnts = Val(Form1.Text2)
Weight = Val(Form1.Text3)

'Dimensión de la solución
ReDim Eta(0 To Nodes, 1 To Nodes)
ReDim Save(1 To Nodes, 1 To Nodes)
ReDim Solution(0 To Nodes + MaxNv + 1)
ReDim Tao(1 To Nodes, 1 To Nodes)
ReDim DeltaTao(1 To Nodes, 1 To Nodes)

ReDim BestCused(1 To MaxNv)
ReDim BestCovered(0 To MaxNv)
ReDim BestSolution(1 To Nodes + MaxNv + 1)

'Savings - Ahorros
For I = 1 To Nodes
    For J = I + 1 To Nodes
        Save(I, J) = Dist(0, I) + Dist(0, J) - Dist(I, J)
    Next J
Next I

'Eta - Información Heurística
Max = 0
For I = 1 To Nodes
    For J = I + 1 To Nodes
        If Dist(I, J) = 0 Then
            Eta(I, J) = 1
            Eta(J, I) = 1
        Else
            Eta(I, J) = 1 / Dist(I, J)
            Eta(J, I) = 1 / Dist(J, I)
            If Eta(I, J) > Max Then Max = Eta(I, J)
        End If
    Next J
Next I
For I = 1 To Nodes
    For J = 1 To Nodes
        Eta(I, J) = Eta(I, J) / Max
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
    
    'Se leen las coordenadas y la demanda de cada nodo
    For I = 1 To Nodes
        
        Input #1, I, X(I), Y(I), TimeS(I), Dem(I), J, J, J
        
        If MinX > X(I) Then MinX = X(I)
        If MaxX < X(I) Then MaxX = X(I)
        If MinY > Y(I) Then MinY = Y(I)
        If MaxY < Y(I) Then MaxY = Y(I)
        
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
For I1 = 1 To Nv
    Text = "1" & vbTab & I1 & vbTab & Int(BestCovered(I1) * 1000) / 1000 & vbTab & BestCused(I1) & vbTab
    Do
        Text = Text & BestSolution(I2) & vbTab
        I2 = I2 + 1
    Loop While BestSolution(I2) <> 0
    Text = Text & vbTab & 0
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

    For K = 1 To nAnts
    
        'Colonia
        Call Ant
        Call Improved
        
        'Acumular DeltaTao
        Call AccDeltaTao
        
    Next K
    
    'Actulización de Tao
    Call Update_TAO

Next GenerationNext

'Mostrar resultados
MsgBox BestCovered(0) & vbCrLf & Nv
Call Plot

'Generar archivo de resultados
Call SinglePrint

End Sub

Sub UpdateETA()

Dim I As Integer
Dim J As Integer

'Ultima actividad seleccionada
Eta(0, Chosen) = Chosen
For I = 1 To Nodes
    If Assigned(I) = False Then
        'Restriccion de capacidad
        If Dem(I) + Cused(Nv) <= Capv Then
            'Nearest() = punto en el que se va apegar el nodo seleccionado
            Nearest(I, 2) = Ctrl(Nv + 1)
            For J = Ctrl(Nv + 1) - 1 To Ctrl(Nv) + 1 Step -1
                If (Dist(I, Solution(J)) + Dist(I, Solution(J - 1)) - Dist(Solution(J), Solution(J - 1))) < (Dist(I, Solution(Nearest(I, 2))) + Dist(I, Solution(Nearest(I, 2) - 1)) - Dist(Solution(Nearest(I, 2)), Solution(Nearest(I, 2) - 1))) Then
                    Nearest(I) = J
                End If
            Next J
            Nearest(I, 1) = Solution(Nearest(I, 2))
            Eta(0, I) = Nearest(I, 1)
            
            If TimeC <> 0 And TimeS(I) + Tused(Nv) + Dist(Solution(Ctrl(Nv + 1) - 1), I) + Dist(0, I) - Dist(0, Solution(Ctrl(Nv + 1) - 1)) > TimeC Then
                'Nodos infactibles por tiempo
                Eta(0, I) = I
            End If

                    
'            If Dist((Solution(Ctrl(Nv) + 1)), I) > Dist((Solution(Ctrl(Nv + 1) - 1)), I) Then
'                Nearest(I) = Solution(Ctrl(Nv + 1) - 1)
'                Eta(0, I) = Solution(Ctrl(Nv + 1) - 1)
                'Restriccion de longitud de la ruta (en unidades de tiempo)
'                If TimeC <> 0 And TimeS(I) + Tused(Nv) + Dist(Solution(Ctrl(Nv + 1) - 1), I) + Dist(0, I) - Dist(0, Solution(Ctrl(Nv + 1) - 1)) > TimeC Then
                    'Nodos infactibles por tiempo
'                    Eta(0, I) = I
'                End If
'            Else
'                Nearest(I) = Solution(Ctrl(Nv) + 1)
'                Eta(0, I) = Solution(Ctrl(Nv) + 1)
                'Restriccion de longitud de la ruta (en unidades de tiempo)
'                If TimeC <> 0 And (TimeS(I) + Tused(Nv) + Dist(Solution(Ctrl(Nv) + 1), I) + Dist(0, I) - Dist(0, Solution(Ctrl(Nv) + 1))) > TimeC Then
                    'Nodos infactibles por tiempo
'                    Eta(0, I) = I
'                End If
'            End If
        Else
            'Nodos infactibles por capacidad
            Eta(0, I) = I
        End If
    End If
Next I

End Sub


Sub Update_TAO()

Dim I1 As Integer
Dim I2 As Integer

Max = 0
For I1 = 1 To Nodes
    For I2 = I1 + 1 To Nodes
        Tao(I1, I2) = Tao(I1, I2) + DeltaTao(I1, I2)
        Tao(I2, I1) = Tao(I2, I1) + DeltaTao(I2, I1)
        If Tao(I1, I2) > Max Then Max = Tao(I1, I2)
    Next I2
Next I1
For I1 = 1 To Nodes
    For I2 = 1 To Nodes
        Tao(I1, I2) = Tao(I1, I2) / Max
    Next I2
Next I1

'Reinicializar DeltaTao
ReDim DeltaTao(1 To Nodes, 1 To Nodes)

End Sub

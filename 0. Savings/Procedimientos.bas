Attribute VB_Name = "Procedimientos"
Option Explicit


Sub DeleteRoute()

Dim I1 As Integer
Dim I2 As Integer

Dim Min As Integer

For I1 = MaxNv + 1 To Nv
    Min = Nv
    For I2 = Nv - 1 To 1 Step -1
        If Solution(I2).Solution(0) < Solution(Min).Solution(0) Then
            Min = I2
        End If
    Next I2
    
    For I2 = 2 To Solution(Min).Solution(0) - 1
        Asigned(Solution(Min).Solution(I2)) = True
    Next I2
    
    For I2 = Min To Nv - 1
        Solution(I2) = Solution(I2 + 1)
    Next I2
    
    Nv = Nv - 1
    
Next I1

Call UnInfactibilization

End Sub

Sub Exchange()

Dim I1 As Integer
Dim I2 As Integer
Dim I3 As Integer
Dim I4 As Integer
Dim I5 As Integer
Dim Release As Boolean
Dim Temp As Neighbor

Temp.Covered = Solution(0).Covered

Do

    Release = False

    For I1 = 1 To Nv
        For I3 = 1 To Nv
            If I1 <> I3 Then
                For I2 = 2 To Solution(I1).Solution(0) - 1
                    For I4 = 2 To Solution(I3).Solution(0) - 1
                        If (Solution(I1).Demanda + Dem(Solution(I3).Solution(I4)) - Dem(Solution(I1).Solution(I2))) <= Capv And (Solution(I3).Demanda + Dem(Solution(I1).Solution(I2)) - Dem(Solution(I3).Solution(I4))) <= Capv Then
                            If (Solution(I1).Time + TimeS(Solution(I3).Solution(I4)) - TimeS(Solution(I1).Solution(I2)) + Dist(Solution(I1).Solution(I2 - 1), Solution(I3).Solution(I4)) + Dist(Solution(I3).Solution(I4), Solution(I1).Solution(I2 + 1)) - Dist(Solution(I1).Solution(I2), Solution(I1).Solution(I2 + 1)) - Dist(Solution(I1).Solution(I2), Solution(I1).Solution(I2 - 1))) <= TimeC Or TimeC = 0 Then
                                If (Solution(I3).Time + TimeS(Solution(I1).Solution(I2)) - TimeS(Solution(I3).Solution(I4)) + Dist(Solution(I3).Solution(I4 - 1), Solution(I1).Solution(I2)) + Dist(Solution(I1).Solution(I2), Solution(I3).Solution(I4 + 1)) - Dist(Solution(I3).Solution(I4), Solution(I3).Solution(I4 + 1)) - Dist(Solution(I3).Solution(I4), Solution(I3).Solution(I4 - 1))) <= TimeC Or TimeC = 0 Then
                                    If (Solution(0).Covered + Dist(Solution(I1).Solution(I2 - 1), Solution(I3).Solution(I4)) + Dist(Solution(I3).Solution(I4), Solution(I1).Solution(I2 + 1)) - Dist(Solution(I1).Solution(I2), Solution(I1).Solution(I2 + 1)) - Dist(Solution(I1).Solution(I2), Solution(I1).Solution(I2 - 1)) + Dist(Solution(I3).Solution(I4 - 1), Solution(I1).Solution(I2)) + Dist(Solution(I1).Solution(I2), Solution(I3).Solution(I4 + 1)) - Dist(Solution(I3).Solution(I4), Solution(I3).Solution(I4 + 1)) - Dist(Solution(I3).Solution(I4), Solution(I3).Solution(I4 - 1))) < Temp.Covered Then
                                        Temp.FV = I1
                                        Temp.SV = I3
                                        Temp.NFV = I2
                                        Temp.NSV = I4
                                        Temp.Covered = (Solution(0).Covered + Dist(Solution(I1).Solution(I2 - 1), Solution(I3).Solution(I4)) + Dist(Solution(I3).Solution(I4), Solution(I1).Solution(I2 + 1)) - Dist(Solution(I1).Solution(I2), Solution(I1).Solution(I2 + 1)) - Dist(Solution(I1).Solution(I2), Solution(I1).Solution(I2 - 1)) + Dist(Solution(I3).Solution(I4 - 1), Solution(I1).Solution(I2)) + Dist(Solution(I1).Solution(I2), Solution(I3).Solution(I4 + 1)) - Dist(Solution(I3).Solution(I4), Solution(I3).Solution(I4 + 1)) - Dist(Solution(I3).Solution(I4), Solution(I3).Solution(I4 - 1)))
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
        Solution(Temp.FV).Demanda = Solution(Temp.FV).Demanda + Dem(Solution(Temp.SV).Solution(Temp.NSV)) - Dem(Solution(Temp.FV).Solution(Temp.NFV))
        Solution(Temp.SV).Demanda = Solution(Temp.SV).Demanda - Dem(Solution(Temp.SV).Solution(Temp.NSV)) + Dem(Solution(Temp.FV).Solution(Temp.NFV))
        Solution(Temp.FV).Time = Solution(Temp.FV).Time + TimeS(Solution(Temp.SV).Solution(Temp.NSV)) - TimeS(Solution(Temp.FV).Solution(Temp.NFV)) + Dist(Solution(Temp.FV).Solution(Temp.NFV - 1), Solution(Temp.SV).Solution(Temp.NSV)) + Dist(Solution(Temp.SV).Solution(Temp.NSV), Solution(Temp.FV).Solution(Temp.NFV + 1)) - Dist(Solution(Temp.FV).Solution(Temp.NFV), Solution(Temp.FV).Solution(Temp.NFV + 1)) - Dist(Solution(Temp.FV).Solution(Temp.NFV), Solution(Temp.FV).Solution(Temp.NFV - 1))
        Solution(Temp.SV).Time = Solution(Temp.SV).Time - TimeS(Solution(Temp.SV).Solution(Temp.NSV)) + TimeS(Solution(Temp.FV).Solution(Temp.NFV)) + Dist(Solution(Temp.SV).Solution(Temp.NSV - 1), Solution(Temp.FV).Solution(Temp.NFV)) + Dist(Solution(Temp.SV).Solution(Temp.NSV + 1), Solution(Temp.FV).Solution(Temp.NFV)) - Dist(Solution(Temp.SV).Solution(Temp.NSV), Solution(Temp.SV).Solution(Temp.NSV + 1)) - Dist(Solution(Temp.SV).Solution(Temp.NSV), Solution(Temp.SV).Solution(Temp.NSV - 1))
        Solution(Temp.FV).Covered = Solution(Temp.FV).Covered + Dist(Solution(Temp.FV).Solution(Temp.NFV - 1), Solution(Temp.SV).Solution(Temp.NSV)) + Dist(Solution(Temp.SV).Solution(Temp.NSV), Solution(Temp.FV).Solution(Temp.NFV + 1)) - Dist(Solution(Temp.FV).Solution(Temp.NFV), Solution(Temp.FV).Solution(Temp.NFV + 1)) - Dist(Solution(Temp.FV).Solution(Temp.NFV), Solution(Temp.FV).Solution(Temp.NFV - 1))
        Solution(Temp.SV).Covered = Solution(Temp.SV).Covered + Dist(Solution(Temp.SV).Solution(Temp.NSV - 1), Solution(Temp.FV).Solution(Temp.NFV)) + Dist(Solution(Temp.SV).Solution(Temp.NSV + 1), Solution(Temp.FV).Solution(Temp.NFV)) - Dist(Solution(Temp.SV).Solution(Temp.NSV), Solution(Temp.SV).Solution(Temp.NSV + 1)) - Dist(Solution(Temp.SV).Solution(Temp.NSV), Solution(Temp.SV).Solution(Temp.NSV - 1))
        Solution(0).Covered = Temp.Covered
            
        Temp.Help = Solution(Temp.SV).Solution(Temp.NSV)
        Solution(Temp.SV).Solution(Temp.NSV) = Solution(Temp.FV).Solution(Temp.NFV)
        Solution(Temp.FV).Solution(Temp.NFV) = Temp.Help
        
    End If
    
Loop While Release = True

End Sub


Sub Heuristico()

'Problemas(s)
Problem = Form1.Combo1.Text

If Problem = "Todos" Then
    
    For NProblem = 1 To 14
    
        Time = Timer
        
        Call MultiSavings
        
    Next NProblem

    'Generar resumen de resultados
    Call MultiPrint
    
Else
    Time = Timer
    Call SingleSavings
End If

MsgBox "I finished!!!!"
End Sub


Sub Insertion()

Dim I1 As Integer
Dim I2 As Integer
Dim I3 As Integer
Dim I4 As Integer
Dim I5 As Integer
Dim Release As Boolean
Dim Temp As Neighbor

Temp.Covered = Solution(0).Covered

Do

    Release = False

    For I1 = 1 To Nv
        For I3 = 1 To Nv
            If I1 <> I3 Then
                For I2 = 2 To Solution(I1).Solution(0) - 1
                    For I4 = 2 To Solution(I3).Solution(0) - 1
                        If Solution(I1).Demanda + Dem(Solution(I3).Solution(I4)) <= Capv Then
                            If (Solution(I1).Time + TimeS(Solution(I3).Solution(I4)) + Dist(Solution(I1).Solution(I2), Solution(I3).Solution(I4)) + Dist(Solution(I3).Solution(I4), Solution(I1).Solution(I2 + 1)) - Dist(Solution(I1).Solution(I2), Solution(I1).Solution(I2 + 1))) <= TimeC Or TimeC = 0 Then
                                If (Solution(0).Covered + Dist(Solution(I1).Solution(I2), Solution(I3).Solution(I4)) + Dist(Solution(I3).Solution(I4), Solution(I1).Solution(I2 + 1)) - Dist(Solution(I1).Solution(I2), Solution(I1).Solution(I2 + 1)) - Dist(Solution(I3).Solution(I4), Solution(I3).Solution(I4 - 1)) - Dist(Solution(I3).Solution(I4), Solution(I3).Solution(I4 + 1)) + Dist(Solution(I3).Solution(I4 - 1), Solution(I3).Solution(I4 + 1))) < Temp.Covered Then
                                    Temp.FV = I1
                                    Temp.SV = I3
                                    Temp.NFV = I2
                                    Temp.NSV = I4
                                    Temp.Covered = (Solution(0).Covered + Dist(Solution(I1).Solution(I2), Solution(I3).Solution(I4)) + Dist(Solution(I3).Solution(I4), Solution(I1).Solution(I2 + 1)) - Dist(Solution(I1).Solution(I2), Solution(I1).Solution(I2 + 1)) - Dist(Solution(I3).Solution(I4), Solution(I3).Solution(I4 - 1)) - Dist(Solution(I3).Solution(I4), Solution(I3).Solution(I4 + 1)) + Dist(Solution(I3).Solution(I4 - 1), Solution(I3).Solution(I4 + 1)))
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
        Solution(Temp.FV).Demanda = Solution(Temp.FV).Demanda + Dem(Solution(Temp.SV).Solution(Temp.NSV))
        Solution(Temp.SV).Demanda = Solution(Temp.SV).Demanda - Dem(Solution(Temp.SV).Solution(Temp.NSV))
        Solution(Temp.FV).Time = Solution(Temp.FV).Time + TimeS(Solution(Temp.SV).Solution(Temp.NSV)) + Dist(Solution(Temp.FV).Solution(Temp.NFV), Solution(Temp.SV).Solution(Temp.NSV)) + Dist(Solution(Temp.SV).Solution(Temp.NSV), Solution(Temp.FV).Solution(Temp.NFV + 1)) - Dist(Solution(Temp.FV).Solution(Temp.NFV), Solution(Temp.FV).Solution(Temp.NFV + 1))
        Solution(Temp.FV).Covered = Solution(Temp.FV).Covered + Dist(Solution(Temp.FV).Solution(Temp.NFV), Solution(Temp.SV).Solution(Temp.NSV)) + Dist(Solution(Temp.SV).Solution(Temp.NSV), Solution(Temp.FV).Solution(Temp.NFV + 1)) - Dist(Solution(Temp.FV).Solution(Temp.NFV), Solution(Temp.FV).Solution(Temp.NFV + 1))
        Solution(Temp.SV).Time = Solution(Temp.SV).Time - TimeS(Solution(Temp.SV).Solution(Temp.NSV)) + Dist(Solution(Temp.SV).Solution(Temp.NSV - 1), Solution(Temp.SV).Solution(Temp.NSV + 1)) - Dist(Solution(Temp.SV).Solution(Temp.NSV), Solution(Temp.SV).Solution(Temp.NSV + 1)) - Dist(Solution(Temp.SV).Solution(Temp.NSV), Solution(Temp.SV).Solution(Temp.NSV - 1))
        Solution(Temp.SV).Covered = Solution(Temp.SV).Covered + Dist(Solution(Temp.SV).Solution(Temp.NSV - 1), Solution(Temp.SV).Solution(Temp.NSV + 1)) - Dist(Solution(Temp.SV).Solution(Temp.NSV), Solution(Temp.SV).Solution(Temp.NSV + 1)) - Dist(Solution(Temp.SV).Solution(Temp.NSV), Solution(Temp.SV).Solution(Temp.NSV - 1))
        Solution(0).Covered = Temp.Covered
        Temp.Help = Solution(Temp.SV).Solution(Temp.NSV)
        Solution(Temp.FV).Solution(0) = Solution(Temp.FV).Solution(0) + 1
        ReDim Preserve Solution(Temp.FV).Solution(0 To Solution(Temp.FV).Solution(0))
        For I1 = Solution(Temp.FV).Solution(0) To Temp.NFV + 1 Step -1
            Solution(Temp.FV).Solution(I1) = Solution(Temp.FV).Solution(I1 - 1)
        Next I1
        Solution(Temp.FV).Solution(Temp.NFV + 1) = Temp.Help
        Solution(Temp.SV).Solution(0) = Solution(Temp.SV).Solution(0) - 1
        For I1 = Temp.NSV To Solution(Temp.SV).Solution(0)
            Solution(Temp.SV).Solution(I1) = Solution(Temp.SV).Solution(I1 + 1)
        Next I1
        ReDim Preserve Solution(Temp.SV).Solution(0 To Solution(Temp.SV).Solution(0))
    End If
    
Loop While Release = True

End Sub


Sub MultiPrint()

Dim I1 As Integer

Open App.Path & "\Resultados\Summary.txt" For Output As #3
Print #3, vbTab & "Costo" & vbTab & "Iter." & vbTab & "Tiempo"
For I1 = 1 To 14
    Print #3, I1 & vbTab & Int(Final(I1).Covered * 100) / 100 & vbTab & Final(I1).Demanda & vbTab & Int(Final(I1).Time * 1000) / 1000
Next I1
Close #3

End Sub


Sub MultiSavings()

'Lectura de datos
Call Reading_Multi

'Inicialización de variables y parámetros
Call Parameters

'Solución - Método de los ahorros
Call SavingsMethod

If Nv > MaxNv Then
    Call DeleteRoute
    MsgBox Solution(0).Covered & vbCrLf & Nv
    Call Plot
End If

'BÚSQUEDA LOCAL
'NLS = 2
'Do While UBound(BestAnt) <> 0 And NLS > 0
'    NLS = 0
'    Call TwoOpt
'    Call Insertion
'    Call Exchange
'Loop

ReDim Preserve Final(1 To NProblem)

Final(NProblem).Covered = Solution(0).Covered
Final(NProblem).Demanda = Nv
Final(NProblem).Time = Timer - Time

'Generar archivo de resultados
Call MultiSinglePrint

End Sub


Sub MultiSinglePrint()


End Sub


Sub Parameters()

'Contadores propios del procedimiento
Dim Aux As Saving
Dim I1 As Long
Dim I2 As Integer
Dim I3 As Integer
Dim I4 As Integer

'Dimensión de la solución
ReDim Solution(0 To Nodes)
ReDim Asigned(1 To Nodes)

ReDim Savings(1 To (Nodes * Nodes - Nodes) / 2)
I3 = 0
For I1 = 1 To Nodes - 1
    For I2 = I1 + 1 To Nodes
        I3 = I3 + 1
        Aux.I = I1
        Aux.J = I2
        Aux.S = Dist(0, I1) + Dist(0, I2) - Dist(I1, I2)
        For I4 = I3 - 1 To 1 Step -1
            If Aux.S > Savings(I4).S Then
                Savings(I4 + 1) = Savings(I4)
            Else
                Exit For
            End If
        Next I4
        Savings(I4 + 1) = Aux
    Next I2
Next I1

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
    If Asigned(I) = False Then
        Form2.PSet (X(I), Y(I)), vbBlue
    Else
        Form2.PSet (X(I), Y(I)), vbBlack
    End If
Next I
Form2.DrawWidth = 1
For I = 1 To Nv
    For J = 1 To Solution(I).Solution(0) - 1
        Form2.Line (X(Solution(I).Solution(J)), Y(Solution(I).Solution(J)))-(X(Solution(I).Solution(J + 1)), Y(Solution(I).Solution(J + 1))), vbYellow
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
    
    MinX = X(I)
    MaxX = X(I)
    MinY = Y(I)
    MaxY = Y(I)
    
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


Sub SavingsMethod()

Dim I1 As Integer
Dim I2 As Integer
Dim I3 As Integer
Dim I4 As Integer

ReDim Control(1 To Nodes, 1 To Nodes + 3)

For I1 = 1 To Nodes
    Control(I1, Nodes + 2) = I1
    Control(I1, Nodes + 3) = Dem(I1)
Next I1

For I2 = 1 To ((Nodes ^ 2 - Nodes) / 2)
    If Control(Savings(I2).I, Nodes + 1) < 2 And Control(Savings(I2).J, Nodes + 1) < 2 Then
        If Control(Savings(I2).I, Nodes + 2) <> Control(Savings(I2).J, Nodes + 2) Then
            If Control(Savings(I2).I, Nodes + 3) + Control(Savings(I2).J, Nodes + 3) <= Capv Then
                Call UpdateControl(I2)
            End If
        End If
    End If
Next I2

I1 = 1
Nv = 0
For I2 = I1 To Nodes
    If Control(I2, Nodes + 1) = 1 Then
        Nv = Nv + 1
        ReDim Preserve Solution(Nv).Solution(0 To 3)
        I3 = 2
        Solution(Nv).Solution(I3) = I2
        Solution(Nv).Solution(0) = 3
        Solution(Nv).Covered = Dist(Solution(Nv).Solution(I3 - 1), Solution(Nv).Solution(I3))
        Solution(Nv).Time = TimeS(Solution(Nv).Solution(I3)) + Dist(Solution(Nv).Solution(I3 - 1), Solution(Nv).Solution(I3))
        Control(I2, Nodes + 1) = 3
        Do
            For I4 = 1 To Nodes
                If Control(Solution(Nv).Solution(I3), I4) = 1 Or Control(I4, Solution(Nv).Solution(I3)) = 1 Then
                    If I4 <> Solution(Nv).Solution(I3 - 1) Then
                        I3 = I3 + 1
                        Solution(Nv).Solution(0) = Solution(Nv).Solution(0) + 1
                        ReDim Preserve Solution(Nv).Solution(0 To I3 + 1)
                        Solution(Nv).Solution(I3) = I4
                        Solution(Nv).Covered = Solution(Nv).Covered + Dist(Solution(Nv).Solution(I3 - 1), Solution(Nv).Solution(I3))
                        Solution(Nv).Time = Solution(Nv).Time + TimeS(Solution(Nv).Solution(I3)) + Dist(Solution(Nv).Solution(I3 - 1), Solution(Nv).Solution(I3))
                        Exit For
                    End If
                End If
            Next I4
        Loop While Control(I4, Nodes + 1) > 1
        Control(I4, Nodes + 1) = 3
        Solution(Nv).Covered = Solution(Nv).Covered + Dist(Solution(Nv).Solution(I3), Solution(Nv).Solution(I3 + 1))
        Solution(Nv).Time = Solution(Nv).Time + Dist(Solution(Nv).Solution(I3), Solution(Nv).Solution(I3 + 1))
        Solution(Nv).Demanda = Control(I4, Nodes + 3)
        Solution(0).Covered = Solution(0).Covered + Solution(Nv).Covered
    ElseIf Control(I2, Nodes + 1) = 0 Then
        Nv = Nv + 1
        ReDim Preserve Solution(Nv).Solution(0 To 3)
        Solution(Nv).Solution(I3) = I2
        Solution(Nv).Solution(0) = 3
        Solution(Nv).Covered = Dist(Solution(Nv).Solution(I3 - 1), Solution(Nv).Solution(I3)) * 2
        Solution(Nv).Time = TimeS(Solution(Nv).Solution(I3)) + Dist(Solution(Nv).Solution(I3 - 1), Solution(Nv).Solution(I3)) * 2
        Solution(Nv).Demanda = Control(I4, Nodes + 3)
        Solution(0).Covered = Solution(0).Covered + Solution(Nv).Covered
        Control(I4, Nodes + 1) = 3
    End If
Next I2

End Sub


Sub SinglePrint()

Dim I1 As Integer
Dim I2 As Integer
Dim Text As String

Open App.Path & "\Resultados\Solve " & Problem & ".txt" For Output As #2
'Open App.Path & "\Resultados\Summary " & Problem & ".txt" For Output As #3

Print #2, "Costo de la solución:"
Print #2, Solution(0).Covered
Print #2,
For I1 = 1 To Nv
    Text = "1" & vbTab & I1 & vbTab & Int(Solution(I1).Covered * 1000) / 1000 & vbTab & Solution(I1).Demanda & vbTab
    For I2 = 1 To Solution(I1).Solution(0)
        Text = Text & Solution(I1).Solution(I2) & vbTab
    Next I2
    Text = Text & 0
    Print #2, Text
Next I1
Print #2,
Print #2, "Tiempo de ejecución:"
Print #2, Timer - Time
Close #2

End Sub


Sub SingleSavings()

'Lectura de datos
Call Reading

'Inicialización de variables y parámetros
Call Parameters

'Solución - Método de los Ahorros
Call SavingsMethod
    
MsgBox Solution(0).Covered & vbCrLf & Nv
Call Plot

If Nv > MaxNv Then
    Call DeleteRoute
    MsgBox Solution(0).Covered & vbCrLf & Nv
    Call Plot
End If

'BÚSQUEDA LOCAL
NLS = 2
Do While NLS > 0
    NLS = 0
    Call TwoOpt
    Call Insertion
    Call Exchange
Loop

'Mostrar resultados
MsgBox "Final:" & vbCrLf & Solution(0).Covered & vbCrLf & Nv
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

Solution(0).Covered = 0
For I = 1 To Nv
    For J = 2 To Solution(I).Solution(0) - 2
        For K = J + 2 To Solution(I).Solution(0) - 1
            If Dist(Solution(I).Solution(J), Solution(I).Solution(J + 1)) + Dist(Solution(I).Solution(K), Solution(I).Solution(K + 1)) > Dist(Solution(I).Solution(J), Solution(I).Solution(K)) + Dist(Solution(I).Solution(J + 1), Solution(I).Solution(K + 1)) Then
                Solution(I).Covered = Solution(I).Covered + (-Dist(Solution(I).Solution(J), Solution(I).Solution(J + 1)) - Dist(Solution(I).Solution(K), Solution(I).Solution(K + 1)) + Dist(Solution(I).Solution(J), Solution(I).Solution(K)) + Dist(Solution(I).Solution(J + 1), Solution(I).Solution(K + 1)))
                Solution(I).Time = Solution(I).Time + (-Dist(Solution(I).Solution(J), Solution(I).Solution(J + 1)) - Dist(Solution(I).Solution(K), Solution(I).Solution(K + 1)) + Dist(Solution(I).Solution(J), Solution(I).Solution(K)) + Dist(Solution(I).Solution(J + 1), Solution(I).Solution(K + 1)))
                For W = 1 To Int((K - J) / 2)
                    Temp = Solution(I).Solution(J + W)
                    Solution(I).Solution(J + W) = Solution(I).Solution(K + 1 - W)
                    Solution(I).Solution(K + 1 - W) = Temp
                Next W
                Release = True
                K = J + 1
            End If
        Next K
    Next J
    Solution(0).Covered = Solution(0).Covered + Solution(I).Covered
Next I

If Release = True Then
    NLS = NLS + 1
End If

End Sub


Sub UnInfactibilization()

Dim I1 As Integer
Dim I2 As Integer
Dim I3 As Integer
Dim I4 As Integer

For I1 = 1 To (Nodes ^ 2 - Nodes) / 2
    If Asigned(Savings(I1).I) = True Or Asigned(Savings(I1).J) = True Then
        If Asigned(Savings(I1).I) = True And Asigned(Savings(I1).J) = True Then
        Else
            Savings(I1).Act = True
        End If
    End If
Next I1

For I1 = 1 To (Nodes ^ 2 - Nodes) / 2
    If Savings(I1).Act = True Then
        If Asigned(Savings(I1).I) = True Then
            For I2 = 1 To MaxNv
                For I3 = 2 To Solution(I2).Solution(0) - 1
                    If Savings(I1).J = Solution(I2).Solution(I3) Then
                        If (Solution(I2).Demanda + Dem(Savings(I1).I)) <= Capv Then
                            If (Dist(Solution(I2).Solution(I3 - 1), Savings(I1).I) + Dist(Solution(I2).Solution(I3), Savings(I1).I) - Dist(Solution(I2).Solution(I3 - 1), Solution(I2).Solution(I3))) < (Dist(Solution(I2).Solution(I3 + 1), Savings(I1).I) + Dist(Solution(I2).Solution(I3), Savings(I1).I) - Dist(Solution(I2).Solution(I3 + 1), Solution(I2).Solution(I3))) Then
                                If Solution(I2).Time + TimeS(Savings(I1).I) + (Dist(Solution(I2).Solution(I3 - 1), Savings(I1).I) + Dist(Solution(I2).Solution(I3), Savings(I1).I) - Dist(Solution(I2).Solution(I3 - 1), Solution(I2).Solution(I3))) < TimeC Or TimeC = 0 Then
                                    Solution(I2).Covered = Solution(I2).Covered + (Dist(Solution(I2).Solution(I3 - 1), Savings(I1).I) + Dist(Solution(I2).Solution(I3), Savings(I1).I) - Dist(Solution(I2).Solution(I3 - 1), Solution(I2).Solution(I3)))
                                    Solution(0).Covered = Solution(0).Covered + (Dist(Solution(I2).Solution(I3 - 1), Savings(I1).I) + Dist(Solution(I2).Solution(I3), Savings(I1).I) - Dist(Solution(I2).Solution(I3 - 1), Solution(I2).Solution(I3)))
                                    Solution(I2).Time = Solution(I2).Time + TimeS(Savings(I1).I) + (Dist(Solution(I2).Solution(I3 - 1), Savings(I1).I) + Dist(Solution(I2).Solution(I3), Savings(I1).I) - Dist(Solution(I2).Solution(I3 - 1), Solution(I2).Solution(I3)))
                                    Solution(I2).Demanda = Solution(I2).Demanda + Dem(Savings(I1).I)
                                    Solution(I2).Solution(0) = Solution(I2).Solution(0) + 1
                                    ReDim Preserve Solution(I2).Solution(0 To Solution(I2).Solution(0))
                                    For I4 = Solution(I2).Solution(0) - 1 To I3 + 1 Step -1
                                        Solution(I2).Solution(I4) = Solution(I2).Solution(I4 - 1)
                                    Next I4
                                    Solution(I2).Solution(I4) = Savings(I1).I
                                    Asigned(Savings(I1).I) = False
                                    I2 = MaxNv
                                    Exit For
                                Else
                                    I2 = MaxNv
                                    Exit For
                                End If
                            Else
                                If Solution(I2).Time + TimeS(Savings(I1).I) + (Dist(Solution(I2).Solution(I3 + 1), Savings(I1).I) + Dist(Solution(I2).Solution(I3), Savings(I1).I) - Dist(Solution(I2).Solution(I3 + 1), Solution(I2).Solution(I3))) < TimeC Or TimeC = 0 Then
                                    Solution(I2).Covered = Solution(I2).Covered + (Dist(Solution(I2).Solution(I3 + 1), Savings(I1).I) + Dist(Solution(I2).Solution(I3), Savings(I1).I) - Dist(Solution(I2).Solution(I3 + 1), Solution(I2).Solution(I3)))
                                    Solution(0).Covered = Solution(0).Covered + (Dist(Solution(I2).Solution(I3 + 1), Savings(I1).I) + Dist(Solution(I2).Solution(I3), Savings(I1).I) - Dist(Solution(I2).Solution(I3 + 1), Solution(I2).Solution(I3)))
                                    Solution(I2).Time = Solution(I2).Time + TimeS(Savings(I1).I) + (Dist(Solution(I2).Solution(I3 + 1), Savings(I1).I) + Dist(Solution(I2).Solution(I3), Savings(I1).I) - Dist(Solution(I2).Solution(I3 + 1), Solution(I2).Solution(I3)))
                                    Solution(I2).Demanda = Solution(I2).Demanda + Dem(Savings(I1).I)
                                    Solution(I2).Solution(0) = Solution(I2).Solution(0) + 1
                                    ReDim Preserve Solution(I2).Solution(0 To Solution(I2).Solution(0))
                                    For I4 = Solution(I2).Solution(0) - 1 To I3 + 2 Step -1
                                        Solution(I2).Solution(I4) = Solution(I2).Solution(I4 - 1)
                                    Next I4
                                    Solution(I2).Solution(I4) = Savings(I1).I
                                    Asigned(Savings(I1).I) = False
                                    I2 = MaxNv
                                    Exit For
                                Else
                                    I2 = MaxNv
                                    Exit For
                                End If
                            End If
                            Asigned(Savings(I1).I) = False
                        Else
                            I2 = MaxNv
                            Exit For
                        End If
                    End If
                Next I3
            Next I2
        ElseIf Asigned(Savings(I1).J) = True Then
            For I2 = 1 To MaxNv
                For I3 = 2 To Solution(I2).Solution(0) - 1
                    If Savings(I1).I = Solution(I2).Solution(I3) Then
                        If (Solution(I2).Demanda + Dem(Savings(I1).J)) <= Capv Then
                            If (Dist(Solution(I2).Solution(I3 - 1), Savings(I1).J) + Dist(Solution(I2).Solution(I3), Savings(I1).J) - Dist(Solution(I2).Solution(I3 - 1), Solution(I2).Solution(I3))) < (Dist(Solution(I2).Solution(I3 + 1), Savings(I1).J) + Dist(Solution(I2).Solution(I3), Savings(I1).J) - Dist(Solution(I2).Solution(I3 + 1), Solution(I2).Solution(I3))) Then
                                If Solution(I2).Time + TimeS(Savings(I1).J) + (Dist(Solution(I2).Solution(I3 - 1), Savings(I1).J) + Dist(Solution(I2).Solution(I3), Savings(I1).J) - Dist(Solution(I2).Solution(I3 - 1), Solution(I2).Solution(I3))) < TimeC Or TimeC = 0 Then
                                    Solution(I2).Covered = Solution(I2).Covered + (Dist(Solution(I2).Solution(I3 - 1), Savings(I1).J) + Dist(Solution(I2).Solution(I3), Savings(I1).J) - Dist(Solution(I2).Solution(I3 - 1), Solution(I2).Solution(I3)))
                                    Solution(0).Covered = Solution(0).Covered + (Dist(Solution(I2).Solution(I3 - 1), Savings(I1).J) + Dist(Solution(I2).Solution(I3), Savings(I1).J) - Dist(Solution(I2).Solution(I3 - 1), Solution(I2).Solution(I3)))
                                    Solution(I2).Time = Solution(I2).Time + TimeS(Savings(I1).J) + (Dist(Solution(I2).Solution(I3 - 1), Savings(I1).J) + Dist(Solution(I2).Solution(I3), Savings(I1).J) - Dist(Solution(I2).Solution(I3 - 1), Solution(I2).Solution(I3)))
                                    Solution(I2).Demanda = Solution(I2).Demanda + Dem(Savings(I1).J)
                                    Solution(I2).Solution(0) = Solution(I2).Solution(0) + 1
                                    ReDim Preserve Solution(I2).Solution(0 To Solution(I2).Solution(0))
                                    For I4 = Solution(I2).Solution(0) - 1 To I3 + 1 Step -1
                                        Solution(I2).Solution(I4) = Solution(I2).Solution(I4 - 1)
                                    Next I4
                                    Solution(I2).Solution(I4) = Savings(I1).J
                                    Asigned(Savings(I1).J) = False
                                    I2 = MaxNv
                                    Exit For
                                Else
                                    I2 = MaxNv
                                    Exit For
                                End If
                            Else
                                If Solution(I2).Time + TimeS(Savings(I1).J) + (Dist(Solution(I2).Solution(I3 + 1), Savings(I1).J) + Dist(Solution(I2).Solution(I3), Savings(I1).J) - Dist(Solution(I2).Solution(I3 + 1), Solution(I2).Solution(I3))) < TimeC Or TimeC = 0 Then
                                    Solution(I2).Covered = Solution(I2).Covered + (Dist(Solution(I2).Solution(I3 + 1), Savings(I1).J) + Dist(Solution(I2).Solution(I3), Savings(I1).J) - Dist(Solution(I2).Solution(I3 + 1), Solution(I2).Solution(I3)))
                                    Solution(0).Covered = Solution(0).Covered + (Dist(Solution(I2).Solution(I3 + 1), Savings(I1).J) + Dist(Solution(I2).Solution(I3), Savings(I1).J) - Dist(Solution(I2).Solution(I3 + 1), Solution(I2).Solution(I3)))
                                    Solution(I2).Time = Solution(I2).Time + TimeS(Savings(I1).J) + (Dist(Solution(I2).Solution(I3 + 1), Savings(I1).J) + Dist(Solution(I2).Solution(I3), Savings(I1).J) - Dist(Solution(I2).Solution(I3 + 1), Solution(I2).Solution(I3)))
                                    Solution(I2).Demanda = Solution(I2).Demanda + Dem(Savings(I1).J)
                                    Solution(I2).Solution(0) = Solution(I2).Solution(0) + 1
                                    ReDim Preserve Solution(I2).Solution(0 To Solution(I2).Solution(0))
                                    For I4 = Solution(I2).Solution(0) - 1 To I3 + 2 Step -1
                                        Solution(I2).Solution(I4) = Solution(I2).Solution(I4 - 1)
                                    Next I4
                                    Solution(I2).Solution(I4) = Savings(I1).J
                                    Asigned(Savings(I1).J) = False
                                    I2 = MaxNv
                                    Exit For
                                Else
                                    I2 = MaxNv
                                    Exit For
                                End If
                            End If
                            Asigned(Savings(I1).J) = False
                        Else
                            I2 = MaxNv
                            Exit For
                        End If
                    End If
                Next I3
            Next I2
        End If
    End If
Next I1

End Sub


Sub UpdateControl(I As Integer)

Dim I1 As Integer
Dim X As Integer

'Suma de demanda
X = Control(Savings(I).I, Nodes + 3) + Control(Savings(I).J, Nodes + 3)

'Asignación
Control(Savings(I).I, Savings(I).J) = 1
Control(Savings(I).J, Savings(I).I) = 1

'Número de asignaciones por nodo
Control(Savings(I).I, Nodes + 1) = Control(Savings(I).I, Nodes + 1) + 1
Control(Savings(I).J, Nodes + 1) = Control(Savings(I).J, Nodes + 1) + 1

'Igualar rutas
For I1 = 1 To Nodes
    If I1 <> Savings(I).J And Control(I1, Nodes + 2) = Control(Savings(I).J, Nodes + 2) Then
        Control(I1, Nodes + 2) = Control(Savings(I).I, Nodes + 2)
    End If
Next I1
Control(Savings(I).J, Nodes + 2) = Control(Savings(I).I, Nodes + 2)

'Demanda de la ruta
For I1 = 1 To Nodes
    If Control(I1, Nodes + 2) = Control(Savings(I).J, Nodes + 2) Then
        Control(I1, Nodes + 3) = X
    End If
Next I1

End Sub

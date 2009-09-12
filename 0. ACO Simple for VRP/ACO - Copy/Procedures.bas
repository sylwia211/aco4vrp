Attribute VB_Name = "Procedures"
Option Explicit

Sub ACO()

ReDim Summary(1 To 14)

Alpha = Val(Form1.Text3)
Ants = Val(Form1.Text2)
Beta = Val(Form1.Text4)
Generations = Val(Form1.Text1)
Problem = Form1.Combo1.Text
Rho = Val(Form1.Text5)

If Problem = "All" Then
    For i = 1 To 14
        Problem = i
        Call Initial_Information
        Call Colony
        Call MultiPrinting
        Summary(Problem) = Best_Solution(0, 0)
    Next i
    Call Print_Summary
Else
    Call Initial_Information
    Call Colony
    Call Printing
    Call Plot_Route
    Call Plot_Density
End If

End Sub

Sub Colony()

Dim ba As Integer
Dim bg As Integer

ReDim BaCapUsed(1 To Nv + 1) As Double
ReDim BaCovered(1 To Nv + 1) As Double
ReDim Best_Solution(0 To 2, 0 To Nodes + Nv + 1)
ReDim Solution(0 To Generations, 0 To Nodes + Nv + 1)

For g = 1 To Generations
    For a = 1 To Ants
        Call Construction_Solutions
        Call Update_Pheromone
        ReDim Preserve Solution(0 To Generations, 0 To Nodes + Nv + 1)
        ReDim Preserve BaCapUsed(1 To Nv + 1) As Double
        ReDim Preserve BaCovered(1 To Nv + 1) As Double
            If a <> 1 Then
                If Objective_Function(a) < Objective_Function(a - 1) Then
                    Solution(g, 0) = Objective_Function(a)
                        For ba = 1 To (Nodes + Nv + 1)
                            Solution(g, ba) = Route(ba - 1)
                        Next ba
                        For ba = 1 To Nv
                            BaCapUsed(ba) = CapUsed(ba)
                            BaCovered(ba) = Covered_Distance(ba)
                        Next ba
                End If
            Else
                Solution(g, 0) = Objective_Function(a)
                        For ba = 1 To (Nodes + Nv + 1)
                            Solution(g, ba) = Route(ba - 1)
                        Next ba
                        For ba = 1 To Nv
                            BaCapUsed(ba) = CapUsed(ba)
                            BaCovered(ba) = Covered_Distance(ba)
                        Next ba
            End If
    Next a
    ReDim Preserve Best_Solution(0 To 2, 0 To Nodes + Nv + 1)
            If g = 1 Then
                Best_Solution(0, 0) = Solution(g, 0)
                    For bg = 1 To Nodes + Nv + 1
                        Best_Solution(0, bg) = Solution(g, bg)
                    Next bg
                    For bg = 1 To Nv
                         Best_Solution(1, bg) = BaCapUsed(bg)
                         Best_Solution(2, bg) = BaCovered(bg)
                    Next bg
            Else
                If Solution(g, 0) < Best_Solution(0, 0) Then
                    Best_Solution(0, 0) = Solution(g, 0)
                        For bg = 1 To Nodes + Nv + 1
                            Best_Solution(0, bg) = Solution(g, bg)
                        Next bg
                        For bg = 1 To Nv
                            Best_Solution(1, bg) = BaCapUsed(bg)
                            Best_Solution(2, bg) = BaCovered(bg)
                        Next bg
                End If
            End If
Next g

End Sub

Sub Construction_Solutions()

Dim j As Integer
Dim k As Integer
Dim l As Integer
Dim Node() As Integer
Dim Random As Double
Dim t As Integer

Chosen = 0
NNodes = 1
Nv = 0
Probability = 0

ReDim Assigned(1 To Nodes)
ReDim CapUsed(1 To Nv + 1)
ReDim Covered_Distance(1 To Nv + 1)
ReDim Preserve Objective_Function(1 To Ants)
ReDim Route(0 To Nodes + Nv)
ReDim Node(0 To Nodes - NNodes)
j = 0
Route(j) = 0

Do
Randomize
Random = Rnd

For l = 1 To Nodes
    If Assigned(l) = False Then Sum = Sum + (((Tao(Chosen, l)) ^ Alpha) * ((Eta(Chosen, l)) ^ Beta))
Next l

Nv = Nv + 1
ReDim Preserve CapUsed(1 To Nv + 1)
ReDim Preserve Covered_Distance(1 To Nv + 1)
ReDim Preserve Route(0 To Nodes + Nv)


Do
k = 1
    For k = 1 To Nodes
        If Assigned(k) = False Then
            Probability = Probability + ((((Tao(Chosen, k)) ^ Alpha) * ((Eta(Chosen, k)) ^ Beta)) / Sum)
                If Probability > Random Then
                    Chosen = k
                    If (CapUsed(Nv) + Demand(Chosen)) <= Capv Then
                        j = j + 1
                        Route(j) = Chosen
                        Assigned(Chosen) = True
                        CapUsed(Nv) = CapUsed(Nv) + Demand(Chosen)
                        Covered_Distance(Nv) = Covered_Distance(Nv) + Distance(Route(j - 1), Chosen)
                        NNodes = NNodes + 1
                        Probability = 0
                        Sum = 0
                            For l = 1 To Nodes
                                If Assigned(l) = False Then Sum = Sum + (((Tao(Chosen, l)) ^ Alpha) * ((Eta(Chosen, l)) ^ Beta))
                            Next l
                           'If Sum <> 0 Then'End If
                                Randomize
                                Random = Rnd
                        Exit For
                    Else
                        For l = Node(0) To 1 Step -1
                            If Chosen = Node(l) Then
                                Chosen = Route(j)
                                Probability = 0
                                Randomize
                                Random = Rnd
                                Exit For
                            End If
                        Next l
                        Node(0) = Node(0) + 1
                        Node(Node(0)) = Chosen
                        If Node(0) = (Nodes - NNodes) Then
                            t = 1
                            Exit For
                        End If
                        Chosen = Route(j)
                        Probability = 0
                        Randomize
                        Random = Rnd
                        Exit For
                    End If
                End If
        End If
    If NNodes > Nodes Then
        t = 1
        Exit For
    End If
Next k
  
Loop Until t = 1

t = 0
j = j + 1
Route(j) = 0
Chosen = 0
Sum = 0
Probability = 0
Node(0) = 0
Covered_Distance(Nv) = Covered_Distance(Nv) + Distance(Route(j - 1), Route(j))
Objective_Function(a) = Objective_Function(a) + Covered_Distance(Nv)

Loop Until NNodes > Nodes

End Sub

Sub Initial_Information()

Dim j As Integer
Dim k As Integer

Open App.Path & "\Data\" & Problem & ".vrp" For Input As #1
Input #1, j, MaxNv, Nodes, j
Input #1, TimeC, Capv

ReDim X(0 To Nodes)
ReDim Y(0 To Nodes)
ReDim Demand(0 To Nodes)
ReDim Distance(0 To Nodes, 0 To Nodes)
ReDim Eta(0 To Nodes, 0 To Nodes)
ReDim Tao(0 To Nodes, 0 To Nodes)

Input #1, j, X(j), Y(j), k, Demand(j), k, k, k
    MinX = X(j)
    MaxX = X(j)
    MinY = Y(j)
    MaxY = Y(j)

For j = 0 To Nodes
    Input #1, j, X(j), Y(j), k, Demand(j), k, k, k
        If MinX > X(j) Then MinX = X(j)
        If MaxX < X(j) Then MaxX = X(j)
        If MinY > Y(j) Then MinY = Y(j)
        If MaxY < Y(j) Then MaxY = Y(j)
Next j

Close #1

'Eta-Heuristic Information
For j = 0 To Nodes
    For k = 0 To Nodes
        Distance(j, k) = (Sqr((X(j) - X(k)) ^ 2 + (Y(j) - Y(k)) ^ 2))
            If Distance(j, k) = 0 Then
                Eta(j, k) = 1
            Else
                Eta(j, k) = (1 / Distance(j, k))
            End If
    Next k
Next j

'Tao-Pheromone
For j = 0 To Nodes
    For k = 0 To Nodes
        If j = k Then
            Tao(j, k) = 0
        Else
            Tao(j, k) = 1
        End If
    Next k
Next j

End Sub
Sub MultiPrinting()
Dim i As Integer
Dim j As Integer
Dim Text As String

Open App.Path & "\Results\Solution" & " " & Problem & ".c&w" For Output As #2
j = 1

Print #2, "Vehicle:" & vbTab & "Covered Distance:" & vbTab & "Used Capacity:" & vbTab & vbTab & "Route:"

    For i = 1 To Nv
        Text = "   " & i & vbTab & vbTab & "  " & Int(Best_Solution(2, i) * 1000) / 1000 & vbTab & vbTab & "   " & Best_Solution(1, i) & vbTab & vbTab & vbTab
        Do
            Text = Text & Best_Solution(0, j - 1) & " "
            j = j + 1
        Loop While Best_Solution(0, j) <> 0
            Text = Text & 0
            Print #2, Text
            Text = ""
    Next i
        
    Print #2,
    Print #2, "Objective Function Value:"
    Print #2, Best_Solution(0, 0)
        
    Close #2

End Sub
Sub Plot_Density()
End Sub
Sub Plot_Route()

Dim i As Integer

Graphic.Show
Graphic.Cls
Graphic.Scale (MinX - 0.1 * MaxX, MaxY + 0.1 * MaxX)-(MaxX + 0.1 * MaxX, MinY - 0.1 * MaxX)
Graphic.DrawWidth = 5
Graphic.PSet (X(0), Y(0)), vbRed
Graphic.DrawWidth = 3
For i = 1 To Nodes
    If Assigned(i) = False Then
        Graphic.PSet (X(i), Y(i)), vbGreen
    Else
        Graphic.PSet (X(i), Y(i)), vbBlack
    End If
Next i
Graphic.DrawWidth = 1
For i = 1 To Nodes + Nv
    Graphic.Line (X(Best_Solution(0, i)), Y(Best_Solution(0, i)))-(X(Best_Solution(0, i + 1)), Y(Best_Solution(0, i + 1))), vbYellow
Next i

End Sub
Sub Printing()

Dim i As Integer
Dim j As Integer
Dim Text As String
ReDim Preserve Best_Solution(0 To 2, 0 To Nodes + Nv + 1)

Open App.Path & "\Results\Solution" & " " & Problem & ".c&w" For Output As #2
j = 1

Print #2, "Vehicle:" & vbTab & "Covered Distance:" & vbTab & "Used Capacity:" & vbTab & vbTab & "Route:"

    For i = 1 To Nv
        Text = "   " & i & vbTab & vbTab & "  " & Int(Best_Solution(2, i) * 1000) / 1000 & vbTab & vbTab & "   " & Best_Solution(1, i) & vbTab & vbTab & vbTab
        Do
            Text = Text & Best_Solution(0, j - 1) & " "
            j = j + 1
        Loop While Best_Solution(0, j) <> 0
            Text = Text & 0
            Print #2, Text
            Text = ""
    Next i
        
    Print #2,
    Print #2, "Objective Function Value:"
    Print #2, Best_Solution(0, 0)
        
    Close #2
    MsgBox "Solved Problem" & " " & Problem & " " & "(Solution=" & " " & Best_Solution(0, 0) & ")"

End Sub
Sub Print_Summary()
Dim i As Integer
Dim Text As String

Open App.Path & "\Results\Summary.txt" For Output As #3

Print #3, vbTab & "Problem" & vbTab & "Objective Function Value"

For i = 1 To 14

    Print #3, i & vbTab & Int(Summary(i) * 100) / 100
    
Next i

Close #3


End Sub

Sub Update_Pheromone()

Dim j As Integer
Dim k As Integer
ReDim Delta_Tao(0 To Nodes, 0 To Nodes)

For j = 1 To Nodes + Nv
    Delta_Tao(Route(j - 1), Route(j)) = Delta_Tao(Route(j - 1), Route(j)) + (1 / Distance(Route(j - 1), Route(j)))
 Next j

For j = 0 To Nodes
    For k = 0 To Nodes
        Tao(j, k) = ((1 - Rho) * Tao(j, k)) + Delta_Tao(j, k)
    Next k
Next j
    
End Sub

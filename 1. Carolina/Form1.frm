VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ACO for VRP"
   ClientHeight    =   6090
   ClientLeft      =   6960
   ClientTop       =   3150
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   4680
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "Form1.frx":0000
      Left            =   3360
      List            =   "Form1.frx":0031
      TabIndex        =   14
      Text            =   "Todos"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&END"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   11
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&START"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   10
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3360
      TabIndex        =   9
      Text            =   "0.9"
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3360
      TabIndex        =   8
      Text            =   "1"
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3360
      TabIndex        =   7
      Text            =   "1"
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3360
      TabIndex        =   6
      Text            =   "2"
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Text            =   "5"
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Problems:"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   4560
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "Parameter Alpha:"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label Label6 
      Caption         =   "Parameter Rho:"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "Parameter Beta:"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Ants"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Iterations  or Generations"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ACO for VRP (Parameters)"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'<>^

Private Sub Command1_Click()

'Problema(s)
Problem = Combo1.Text

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

MsgBox "I Finished!!!!"
End Sub


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


Sub Ant()

Dim I As Integer        'Contador

'Inicialización de la variable solución
Nv = 1
ReDim Ctrl(1 To Nv + 1)
ReDim Solution(1 To Nodes + MaxNv + 1)
ReDim Assigned(1 To Nodes)
ReDim Nearest(1 To Nodes, 1 To 2)
Ctrl(1) = 1
Sol = 0
Solution(Ctrl(Nv)) = 0

'Ciclo para generar la ruta de cada vehículo
Do While Sol < Nodes

    'Redimensión de la variables
    ReDim Preserve Ctrl(1 To Nv + 1) '***** MOVER PARA EL FINAL DEL CICLO CUANDO ESTE LISTO ***** Color Naranja
    ReDim Preserve Covered(0 To Nv)
    ReDim Preserve Cused(1 To Nv)
    ReDim Preserve Tused(1 To Nv)

    'Busqueda del nodo más lejano
    Chosen = 0
    For I = 1 To Nodes
        If Assigned(I) = False Then
            If Dist(I, 0) > Dist(Chosen, 0) Then Chosen = I
        End If
    Next I
    'Asignación del primer nodo de la ruta (el más lejano disponible)
    Solution(Ctrl(Nv) + 1) = Chosen
    Furthest = Chosen
'    If Assigned(Chosen) = True Then
'        Chosen = Chosen
'    End If
    Assigned(Chosen) = True
    Solution(Ctrl(Nv) + 2) = 0
    Ctrl(Nv + 1) = Ctrl(Nv) + 2
    Sol = Sol + 1
    Covered(Nv) = Dist(0, Chosen) * 2
    Tused(Nv) = Covered(Nv) + TimeS(Chosen)
    Cused(Nv) = Dem(Chosen)
    Nearest(Chosen, 2) = Ctrl(Nv) + 1
    Eta(Chosen, 0, 0) = Nearest(Chosen, 2)

    'Actualización del parámetro ETA
    Call UpdateETA
    
    'Ciclo para completar la ruta del vehículo
    Do While Cused(Nv) < Capv And Chosen <> 0     'Mientras haya capacidad en el vehículo
    
        'Sum = Denominador para el cálculo de probabilidades
        Sum = 0
        For I = 1 To Nodes
            Sum = Sum + (((Tao(Furthest, I)) ^ Alpha) * (Eta(First(I), Solution(Eta(I, 0, 0) - 1), Solution(Eta(I, 0, 0)))) ^ Beta)
        Next I
        
        Chosen = 0
        
        If Sum <> 0 Then
            Prob = 0
            Randomize
            Random = Rnd
        
            'Selección de un nodo para la ruta
            For I = 1 To Nodes
                Prob = Prob + ((((Tao(Furthest, I)) ^ Alpha) * (Eta(First(I), Solution(Eta(I, 0, 0) - 1), Solution(Eta(I, 0, 0)))) ^ Beta) / Sum)      'probabilidad acumulada
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
        
        'Actualización de la solución
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
        
        'Actualización del parámetro ETA
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

'Calculo Funcion Objetivo
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

    Print #3, I1 & vbTab & Int(Summary(I1, 1) * 100) / 100 & vbTab & Summary(I1, 2) & vbTab & Int(Summary(I1, 3) * 10000) / 10000

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
    Text = "1   " & I1 & "   " & Int(BestCovered(I1) * 1000) / 1000 & vbTab & BestCused(I1) & vbTab
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

    'Actualizacion de Tao
    Call UpdateTAO
    
Next GenerationNext

'Generar archivo de resultados
Call MultiSinglePrint

End Sub


Sub Reading()

Dim I As Integer
Dim J As Integer
 
 Open App.Path & "\Datos\" & Problem & ".vrp" For Input As #1
    Input #1, I, MaxNv, Nodes, J
    Input #1, TimeC, Capv
    
'Redimensiòn de los vectores que contendràn las coordenadas de cada nodo y su respectiva demanda.
    ReDim X(0 To Nodes)
    ReDim Y(0 To Nodes)
    ReDim Dem(1 To Nodes)
    ReDim Dist(0 To Nodes, 0 To Nodes)
    ReDim TimeS(1 To Nodes)
    
    I = 0
    Input #1, I, X(I), Y(I), J, J, J, J

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
Close #1

'Cálculo de la distancia entre nodos.
For I = 0 To Nodes
    For J = 0 To Nodes
        Dist(I, J) = Sqr(((X(J) - X(I)) ^ 2) + ((Y(J) - Y(I)) ^ 2))
    Next J
Next I
        
End Sub


Sub Reading_Multi()

Dim I As Integer
Dim J As Integer
 
If NProblem < 10 Then
    Open App.Path & "\Datos\p0" & NProblem & ".vrp" For Input As #1
Else
    Open App.Path & "\Datos\p" & NProblem & ".vrp" For Input As #1
End If
    Input #1, I, MaxNv, Nodes, J
    Input #1, TimeC, Capv
    
'Redimensiòn de los vectores que contendràn las coordenadas de cada nodo y su respectiva demanda.
    ReDim X(0 To Nodes)
    ReDim Y(0 To Nodes)
    ReDim Dem(1 To Nodes)
    ReDim Dist(0 To Nodes, 0 To Nodes)
    ReDim TimeS(1 To Nodes)
    
    I = 0
    Input #1, I, X(I), Y(I), J, J, J, J

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
Close #1

'Cálculo de la distancia entre nodos.
For I = 0 To Nodes
    For J = 0 To Nodes
        Dist(I, J) = Sqr(((X(J) - X(I)) ^ 2) + ((Y(J) - Y(I)) ^ 2))
    Next J
Next I
        
End Sub


Sub Parameters()

Dim I As Integer
Dim J As Integer

'Parámetros
    Alpha = Val(Form1.Text3)
    Beta = Val(Form1.Text4)
    Rho = Val(Form1.Text5)
    nGen = Val(Form1.Text1)
    nAnts = Val(Form1.Text2)
    
    ReDim Eta(0 To Nodes, 0 To Nodes, 0 To Nodes)
    ReDim Solution(1 To Nodes + MaxNv + 1)
    ReDim Tao(1 To Nodes, 1 To Nodes)
    ReDim DeltaTao(1 To Nodes, 1 To Nodes)
    
    ReDim BestCused(1 To MaxNv)
    ReDim BestCovered(0 To MaxNv)
    ReDim BestSolution(1 To Nodes + MaxNv + 1)
    
    ReDim First(1 To Nodes)

'Cálculo de Eta
For I = 1 To Nodes
    First(I) = I
    For J = 0 To Nodes
        For K = J + 1 To Nodes
            If Dist(I, J) = 0 And Dist(I, K) = 0 Then
                Eta(I, J, K) = 100000000000000#
                Eta(I, K, J) = 100000000000000#
            Else
                Eta(I, J, K) = 1 / (Dist(I, J) + Dist(I, K))
                Eta(I, K, J) = 1 / (Dist(I, J) + Dist(I, K))
            End If
        Next K
    Next J
Next I

'Cálculo de Tao, en la fase de inicialización Tao es 1 para cualquier arco i,j
    For I = 1 To Nodes
        For J = 1 To Nodes
            Tao(I, J) = 1
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

    'Actualizacion de Tao
    Call UpdateTAO
    
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

For I1 = 1 To Nodes
    For I2 = I1 + 1 To Nodes
        Tao(I1, I2) = Rho * Tao(I1, I2) + DeltaTao(I1, I2)
        Tao(I2, I1) = Rho * Tao(I2, I1) + DeltaTao(I2, I1)
    Next I2
Next I1

'Reinicializar DeltaTao
ReDim DeltaTao(1 To Nodes, 1 To Nodes)

End Sub


Private Sub Command2_Click()
End
End Sub

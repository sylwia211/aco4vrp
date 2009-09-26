VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ACO for VRP [0,1]"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4710
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   ScaleHeight     =   8745
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Muestras???"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   16
      Top             =   4920
      Width           =   4215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      ItemData        =   "ACO for VRP.frx":0000
      Left            =   3240
      List            =   "ACO for VRP.frx":0031
      MousePointer    =   3  'I-Beam
      TabIndex        =   11
      Text            =   "Todos"
      ToolTipText     =   "Lista de problemas"
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "End"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   5
      Top             =   7440
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parámetros"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   4215
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Viner Hand ITC"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3000
         TabIndex        =   15
         Text            =   "1"
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Viner Hand ITC"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3000
         TabIndex        =   12
         Text            =   "0.9"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Viner Hand ITC"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3000
         TabIndex        =   9
         Text            =   "0.9"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Viner Hand ITC"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3000
         TabIndex        =   7
         Text            =   "20"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Viner Hand ITC"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3000
         TabIndex        =   4
         Text            =   "5000"
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "p:"
         BeginProperty Font 
            Name            =   "Viner Hand ITC"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   2880
         Width           =   2895
      End
      Begin VB.Label Label6 
         Caption         =   "Weight:"
         BeginProperty Font 
            Name            =   "Viner Hand ITC"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Evaporation:"
         BeginProperty Font 
            Name            =   "Viner Hand ITC"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "Number of ants:"
         BeginProperty Font 
            Name            =   "Viner Hand ITC"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "Number of generations:"
         BeginProperty Font 
            Name            =   "Viner Hand ITC"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   2895
      End
   End
   Begin VB.CommandButton Execution 
      BackColor       =   &H80000013&
      Caption         =   "Solve"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      MaskColor       =   &H0080FF80&
      TabIndex        =   1
      Top             =   6480
      Width           =   3495
   End
   Begin VB.Label Label5 
      Caption         =   "Problems:"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   5760
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ACO for VRP"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = False
Else
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
End If
End Sub

Private Sub Command1_Click()
End
End Sub


Private Sub Execution_Click()

Dim I1 As Integer
Dim I2 As Integer
Dim I3 As Integer
Dim I4 As Integer
Dim I5 As Integer
Dim I6 As Integer

'Procedimiento Principal
Samples = 0
If Check1.Value = 1 Then
    For I1 = 1 To 6 'Ciclo para el número de generaciones
        If I1 = 1 Then
            nGen = 1
            nAnts = 500
        ElseIf I1 = 2 Then
            nGen = 500
            nAnts = 1
        ElseIf I1 = 3 Then
            nGen = 10
            nAnts = 50
        ElseIf I1 = 4 Then
            nGen = 50
            nAnts = 10
        ElseIf I1 = 5 Then
            nGen = 20
            nAnts = 25
        Else
            nGen = 25
            nAnts = 20
        End If
        For I2 = 1 To 1 'Ciclo para el número de hormigas
            For I3 = 1 To 5 'Ciclo para el peso (weight)
                If I3 = 1 Then
                    Weight = 1
                ElseIf I3 = 2 Then
                    Weight = 0.75
                ElseIf I3 = 3 Then
                    Weight = 0.5
                ElseIf I3 = 4 Then
                    Weight = 0.25
                Else
                    Weight = 0
                End If
                For I4 = 1 To 4 'Ciclo para la evaporación (Rho)
                    If I4 = 1 Then
                        Rho = 1
                    ElseIf I4 = 2 Then
                        Rho = 0.8
                    ElseIf I4 = 3 Then
                        Rho = 0.5
                    Else
                        Rho = 0
                    End If
                    For I5 = 1 To 1 ' Ciclo para la potencia del polinomio
                        p = I5
                        For I6 = 1 To 5 'Ciclo para el número de muestras
                            Samples = I6
                            Call Aco
                        Next I6
                    Next I5
                Next I4
            Next I3
        Next I2
    Next I1
Else
    Call Aco
End If
End Sub

VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ACO for VRP [0,1]"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4710
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
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
      Top             =   5040
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
      Top             =   6720
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
      Top             =   5760
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
      Top             =   5040
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
Private Sub Command1_Click()
End
End Sub


Private Sub Execution_Click()
'Procedimiento Principal
Call Aco
End Sub

VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4710
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
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
      ItemData        =   "Savings (vrp).frx":0000
      Left            =   2880
      List            =   "Savings (vrp).frx":0031
      TabIndex        =   3
      Text            =   "Todos"
      ToolTipText     =   "Lista de problemas"
      Top             =   1320
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
      TabIndex        =   2
      Top             =   3000
      Width           =   3375
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
      Top             =   2040
      Width           =   3375
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
      Left            =   600
      TabIndex        =   4
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Savings"
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
Call Heuristico
End Sub

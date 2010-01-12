VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Start programs a few minutes after Windows starts"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14085
   Icon            =   "StartUpLaterFrm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   14085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   10080
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   ".."
      Height          =   375
      Left            =   9240
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   840
      Width           =   9015
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lbls 
      Caption         =   "Finally press Add"
      Height          =   240
      Index           =   2
      Left            =   2520
      TabIndex        =   7
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label lbls 
      Caption         =   "Optionally: * add parameters * add additional wait time before starting this program"
      Height          =   240
      Index           =   1
      Left            =   2520
      TabIndex        =   6
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "Choose a program or paste its path. "
      Height          =   240
      Index           =   0
      Left            =   2520
      TabIndex        =   5
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Wait Seconds"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Debug.Print ""
End Sub

VERSION 5.00
Begin VB.Form FrmDebugger 
   Caption         =   "FrmDebugger"
   ClientHeight    =   9945
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6270
   LinkTopic       =   "Form2"
   ScaleHeight     =   9945
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   9255
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   9375
      Left            =   2280
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   3615
   End
   Begin VB.ListBox List1 
      Height          =   9225
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "FrmDebugger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

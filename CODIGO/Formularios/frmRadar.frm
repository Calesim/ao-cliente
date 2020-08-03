VERSION 5.00
Begin VB.Form frmRadar 
   Caption         =   "Radar"
   ClientHeight    =   9150
   ClientLeft      =   14850
   ClientTop       =   870
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   9150
   ScaleWidth      =   5880
   Begin VB.PictureBox picMapa 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   4000
      Left            =   120
      ScaleHeight     =   3945
      ScaleWidth      =   3945
      TabIndex        =   4
      Top             =   120
      Width           =   4000
   End
   Begin VB.Timer TimerRadar 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   5040
   End
   Begin VB.ListBox lstCharIndex 
      DataField       =   "Index"
      Height          =   4740
      Left            =   3480
      TabIndex        =   3
      Top             =   4200
      Width           =   615
   End
   Begin VB.ListBox lstY 
      DataField       =   "Index"
      Height          =   4740
      Left            =   3000
      TabIndex        =   2
      Top             =   4200
      Width           =   495
   End
   Begin VB.ListBox lstX 
      DataField       =   "Index"
      Height          =   4740
      Left            =   2520
      TabIndex        =   1
      Top             =   4200
      Width           =   495
   End
   Begin VB.ListBox lstPJs 
      DataField       =   "Index"
      Height          =   4740
      Left            =   120
      TabIndex        =   0
      Top             =   4200
      Width           =   2415
   End
End
Attribute VB_Name = "frmRadar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NHL Radar

Dim NumeroMapa As Integer


Private Sub TimerRadar_Timer()

Dim rx As Byte
Dim Ry As Byte
Dim Index As Integer
Dim I As Long

lstPJs.Clear
lstX.Clear
lstY.Clear
lstCharIndex.Clear
picMapa.Cls


If (UserMap <> NumeroMapa) Then
    picMapa.Picture = Nothing
    picMapa.DrawWidth = 3
    For rx = 1 To XMaxMapSize
        For Ry = 1 To YMaxMapSize
            If (MapData(rx, Ry).Blocked = 1) Then
                picMapa.PSet (rx * 40, Ry * 40), vbWhite
            End If
        Next
    Next
    NumeroMapa = UserMap
    picMapa.Picture = picMapa.Image
End If
    
picMapa.DrawWidth = 6

For I = 1 To LastChar
    If (charlist(I).Pos.X + charlist(I).Pos.Y > 0) Then
        If (charlist(I).Nombre <> vbNullString) Then
            lstPJs.AddItem (charlist(I).Nombre)
            lstX.AddItem (charlist(I).Pos.X)
            lstY.AddItem (charlist(I).Pos.Y)
            lstCharIndex.AddItem (I)
        End If
        
        If (charlist(I).Nombre = vbNullString) Then
            picMapa.PSet (charlist(I).Pos.X * 40, charlist(I).Pos.Y * 40), vbMagenta
        Else
                If charlist(I).invisible = True Then
                    picMapa.PSet (charlist(I).Pos.X * 40, charlist(I).Pos.Y * 40), vbRed
                Else
                    If (UserPos.X = charlist(I).Pos.X And UserPos.Y = charlist(I).Pos.Y) Then
                        picMapa.PSet (charlist(I).Pos.X * 40, charlist(I).Pos.Y * 40), vbGreen
                    Else
                        picMapa.PSet (charlist(I).Pos.X * 40, charlist(I).Pos.Y * 40), vbBlue
                    End If
                End If
            End If
    End If
Next I

End Sub



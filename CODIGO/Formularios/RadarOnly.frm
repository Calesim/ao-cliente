VERSION 5.00
Begin VB.Form RadarOnly 
   Caption         =   "RadarOnly"
   ClientHeight    =   11145
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4635
   LinkTopic       =   "RadarOnly"
   ScaleHeight     =   11145
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox ListaPrivi 
      Height          =   5130
      Left            =   3720
      TabIndex        =   5
      Top             =   5400
      Width           =   735
   End
   Begin VB.ListBox ListaChrIndex 
      Height          =   5130
      Left            =   3120
      TabIndex        =   4
      Top             =   5400
      Width           =   615
   End
   Begin VB.ListBox ListaY 
      Height          =   5130
      Left            =   2640
      TabIndex        =   3
      Top             =   5400
      Width           =   495
   End
   Begin VB.ListBox ListaX 
      Height          =   5130
      Left            =   2160
      TabIndex        =   2
      Top             =   5400
      Width           =   495
   End
   Begin VB.ListBox ListaPJs 
      Height          =   5130
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Width           =   2055
   End
   Begin VB.PictureBox RadarMap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      Height          =   5000
      Left            =   120
      ScaleHeight     =   4935
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label4 
      Caption         =   "Privile"
      Height          =   255
      Left            =   4080
      TabIndex        =   10
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Index"
      Height          =   255
      Left            =   3480
      TabIndex        =   9
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Y"
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "X"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label Lista 
      Caption         =   "LISTA"
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   5160
      Width           =   855
   End
End
Attribute VB_Name = "RadarOnly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Only Radar

Dim NumeroMapa As Integer

Public Sub RenderMap()
Dim rx As Byte
Dim Ry As Byte
Dim Index As Integer
Dim I As Long
Dim MyX As Byte
Dim MyY As Byte




ListaPJs.Clear
ListaX.Clear
ListaY.Clear
ListaChrIndex.Clear
ListaPrivi.Clear

RadarMap.Cls




If (UserMap <> NumeroMapa) Then
    RadarMap.Picture = Nothing
    RadarMap.DrawWidth = 3
    For rx = 1 To XMaxMapSize
        For Ry = 1 To YMaxMapSize
            If (MapData(rx, Ry).Blocked = 1) Then
                RadarMap.PSet (rx * 40, Ry * 40), vbWhite
            End If
        Next
    Next
    NumeroMapa = UserMap
    RadarMap.Picture = RadarMap.Image
End If

RadarMap.DrawWidth = 6

End Sub

Public Sub PopulateList()

'ListaPJs.Clear
'ListaX.Clear
'ListaY.Clear
'ListaChrIndex.Clear
'ListaPrivi.Clear


For I = 1 To LastChar
    If (charlist(I).Pos.X + charlist(I).Pos.Y > 0) Then
        If (charlist(I).Nombre <> vbNullString) Then
            ListaPJs.AddItem (charlist(I).Nombre)
            ListaX.AddItem (charlist(I).Pos.X)
            ListaY.AddItem (charlist(I).Pos.Y)
            ListaChrIndex.AddItem (I)
            ListaPrivi.AddItem (charlist(I).priv)
            
        End If
        
        If (charlist(I).Nombre = vbNullString) Then
            RadarMap.PSet (charlist(I).Pos.X * 40, charlist(I).Pos.Y * 40), vbMagenta
        Else
                If charlist(I).invisible = True Then
                    RadarMap.PSet (charlist(I).Pos.X * 40, charlist(I).Pos.Y * 40), vbRed
                Else
                    If (UserPos.X = charlist(I).Pos.X And UserPos.Y = charlist(I).Pos.Y) Then
                        RadarMap.PSet (charlist(I).Pos.X * 40, charlist(I).Pos.Y * 40), vbGreen
                    Else
                        RadarMap.PSet (charlist(I).Pos.X * 40, charlist(I).Pos.Y * 40), vbBlue
                    End If
                End If
            End If
    End If
Next I
End Sub

Public Sub UpdateMyCharPos(ByVal OldX As Byte, ByVal OldY As Byte, ByVal NewX As Byte, ByVal NewY As Byte)


RadarMap.PSet (OldX * 40, OldY * 40), vbBlack
RadarMap.PSet (NewX * 40, NewY * 40), vbYellow
End Sub



Public Sub CreateCharToList(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal Name As String, ByVal NickColor As Byte, ByVal Privileges As Byte)

Debug.Print ("CharIndex= " & CharIndex)
Debug.Print ("Privileges= " & Privileges)
If (Privileges = 0) Then

   If Name = vbNullString Then
    RadarMap.PSet (X * 40, Y * 40), RGB(255, 140, 205) 'Fucsia Claro
 Else
    RadarMap.PSet (X * 40, Y * 40), RGB(112, 132, 255) 'Celeste Oscuro  o Azul Claro(?)
 End If
    
End If
   If (Privileges = 1) Then
   RadarMap.PSet (X * 40, Y * 40), vbRed
   End If
    
   If (Privileges = 16) Then
   RadarMap.PSet (X * 40, Y * 40), RGB(255, 198, 24)
   End If
    
    If Name = vbNullString Then
        ListaPJs.AddItem (Name & " BODY: " & Body)
    Else
        ListaPJs.AddItem (Name)
    
    
   
   
   
End If
     ListaX.AddItem (X)
    ListaY.AddItem (Y)
    ListaChrIndex.AddItem (CharIndex)
    ListaPrivi.AddItem (Privileges)
    
End Sub

Public Sub DeleteCharToList(ByVal CharIndex As Integer, ByVal X As Long, ByVal Y As Long)
Dim Borrar As Integer


If ListaChrIndex.ListCount <> 0 Then

For I = 0 To ListaChrIndex.ListCount - 1
'Debug.Print "ListaChrIndex.ListCount: " & I
If ListaChrIndex.List(I) = CharIndex Then
Borrar = I
RadarMap.PSet (X * 40, Y * 40), vbBlack
Exit For
End If
Next

ListaPJs.RemoveItem (Borrar)
ListaX.RemoveItem (Borrar)
ListaY.RemoveItem (Borrar)
ListaChrIndex.RemoveItem (Borrar)
ListaPrivi.RemoveItem (Borrar)

End If
End Sub

Public Sub HandleCharacterMove(ByVal CharIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal OldX As Integer, ByVal OldY As Integer)
Dim str As String
Dim I As Integer
Dim Borrar As Integer

'Debug.Print ("CharIndex= " & CharIndex)
'Debug.Print ("X= " & X)
'Debug.Print ("Y= " & Y)
'Debug.Print ("OldX= " & OldX)
'Debug.Print ("OldY= " & OldY)



If ListaChrIndex.ListCount <> 0 Then
For I = 0 To ListaChrIndex.ListCount - 1


'Debug.Print "QUE ES STO!:  " & ListaChrIndex.List(I)
If Val(ListaChrIndex.List(I)) = CharIndex Then
Borrar = I
'Debug.Print "LLEGO A BORRAR EL CHAR ES:  " & I
Exit For
End If
Next





'ListaPJs.RemoveItem (Borrar)
'ListaX.RemoveItem (Borrar)
'ListaY.RemoveItem (Borrar)
'ListaChrIndex.RemoveItem (Borrar)
'ListaPrivi.RemoveItem (Borrar)
RadarMap.PSet (OldX * 40, OldY * 40), vbBlack

If (InStr(ListaPrivi.List(Borrar), "BODY:") > 0) Then
    RadarMap.PSet (X * 40, Y * 40), RGB(255, 140, 205) 'Fucsia Claro
Else
    RadarMap.PSet (X * 40, Y * 40), RGB(112, 132, 255) 'Celeste Oscuro  o Azul Claro(?)
End If

If ListaPrivi.List(Borrar) = 0 Then
 Else
    RadarMap.PSet (X * 40, Y * 40), RGB(112, 132, 255) 'Celeste Oscuro  o Azul Claro(?)
 End If
 
    If (ListaPrivi.List(Borrar) = 1) Then
   RadarMap.PSet (X * 40, Y * 40), vbRed
   End If
    
   If (ListaPrivi.List(Borrar) = 16) Then
   RadarMap.PSet (X * 40, Y * 40), RGB(255, 198, 24)
   End If

ListaX.List(Borrar) = X
ListaY.List(Borrar) = Y


End If


End Sub



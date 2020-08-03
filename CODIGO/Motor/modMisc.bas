Attribute VB_Name = "modMisc"
Option Explicit

Public Function IsArrayInitialized(ByVal Array_Pointer As Long) As Boolean
    
    Dim Destination_Pointer As Long
    
    IsArrayInitialized = False
    CopyMemory Destination_Pointer, ByVal Array_Pointer, 4
    If Destination_Pointer = False Then
        IsArrayInitialized = False
    Else
        IsArrayInitialized = True
    End If
        
End Function


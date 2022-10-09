VERSION 5.00
Begin VB.UserControl Docking 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "Docking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private Sub UserControl_Initialize()
    
    If Not Parent Is Nothing Then
        Debug.Print TypeName(Parent)
        
        
        'if form.autoshowchildren
        
    End If
        
End Sub

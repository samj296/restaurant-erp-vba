VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RestSelection 
   Caption         =   "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   5700
   OleObjectBlob   =   "RestSelection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RestSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub UserForm_Initialize()
    Cbox_RestSelection.Clear
    Cbox_RestSelection.ColumnCount = 1
    Cbox_RestSelection.ColumnWidths = 120
    Cbox_RestSelection.AddItem "ArabianHub"
    Cbox_RestSelection.AddItem "ALARASI"
        
End Sub

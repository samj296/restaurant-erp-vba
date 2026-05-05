VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM_Table_Slection 
   Caption         =   "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
   ClientHeight    =   5985
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   12290
   OleObjectBlob   =   "FRM_Table_Slection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FRM_Table_Slection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BTN_Table_Select_Click()
    Dim TableNo As Byte
    TableNo = CByte(LBox_Table_Selection.Value)
    MsgBox "The Table number-" & TableNo & " selected", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
    
End Sub

Private Sub UserForm_Initialize()
    Dim TableConnection As New ADODB.Connection
    Dim TabRec As New ADODB.Recordset
    Dim SQL As String
    
    
    SQL = "Select * From TBL_Table_Status"
    
    Call Connection(TableConnection)
    
    TabRec.Open SQL, TableConnection, adOpenKeyset, adLockOptimistic
    
    With TabRec
    
    
    Dim i As Byte
    
    
    LBox_Table_Selection.Clear
    LBox_Table_Selection.ColumnCount = 4
    LBox_Table_Selection.ColumnWidths = "120,120,120,120"
    LBox_Table_Selection.AddItem
    LBox_Table_Selection.List(0, 0) = "Table Number"
    LBox_Table_Selection.List(0, 1) = "Status"
    LBox_Table_Selection.List(0, 2) = "KOT HOLD"
    LBox_Table_Selection.List(0, 3) = "Name"
    
        i = 1
        .MoveFirst
    Do Until .EOF
        
        LBox_Table_Selection.AddItem
        LBox_Table_Selection.List(i, 0) = .Fields("Table Number")
        LBox_Table_Selection.List(i, 1) = .Fields("Status")
        If Not .Fields("Status") = "Free" Then
            LBox_Table_Selection.List(i, 1) = .Fields("KOT Hold Number")
            LBox_Table_Selection.List(i, 2) = .Fields("Customer Name")
        End If
        .MoveNext
        i = i + 1
        
    Loop
    
    End With
    
    TabRec.Close
    TableConnection.Close
    
End Sub

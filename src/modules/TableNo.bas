Attribute VB_Name = "TableNo"
Sub Table_Selection(T As Byte, KOT As String, CustName As String, orderlist As clsMenuList)
    Dim Con As New ADODB.Connection
    Dim TableRec As New ADODB.Recordset
    Dim SQL As String
    
    Call Connection(Con)
    SQL = "Select * From TBL_Table_Status Where  TBL_Table_Status.[Table Number]=" & T
                    
    TableRec.Open SQL, Con, adOpenKeyset, adLockOptimistic
    
    With TableRec
        
        If .RecordCount = 0 Then
            MsgBox "No table is seleted", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
            Exit Sub
        End If
        
        .MoveFirst
        .Fields("Status") = "Occupied"
        .Fields("KOT Hold Number") = KOT
        .Fields("Customer Name") = CustName
           .Update
           
    End With
    
    TableRec.Close
    
    

    
    
    
End Sub

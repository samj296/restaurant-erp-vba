Attribute VB_Name = "Menu"
Option Explicit

Sub Menu_Connection(MenuCon As ADODB.Connection)
   
   Dim Fpath As String
   Dim FCheck As String
        
        On Error GoTo ErrorHandler
Reprocess:

   Fpath = ShSupport.Range("I8").Value

        FCheck = Dir(Fpath)
   
   If FCheck = VBA.Constants.vbNullString Or Fpath = "" Then
        
        MsgBox "File Location changed. Please select the file.", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
        

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
        .ButtonName = "Pick Folder"
        If .Show = 0 Then
            MsgBox "Please Select the folder to continue", , "ALA RASI FOODS AND BEVERAGES PRIIVATE LIMITED"
            GoTo Reprocess
        Else
            ShSupport.Range("I7").Value = .SelectedItems(1)
            ShSupport.Range("I6").Value = ShSupport.Range("I7").Value & "\Income_&_Exp.accdb"
            ShSupport.Range("I3").Value = ShSupport.Range("I7").Value & "\AlaRasi.accdb"
            ShSupport.Range("I8").Value = ShSupport.Range("I7").Value & "\Menu.accdb"
            GoTo Reprocess
        End If
        
    End With
        
'        Fpath = Application.GetOpenFilename("Access File,*.acc*")
'        ShSupport.Range("I3").Value = Fpath
'            GoTo Reprocess
        
   
    End If
    
    MenuCon.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Fpath & ";Jet OLEDB:DataBase Password= admin123"
          
    Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case Is = 52
        
            MsgBox "Please Connect to the server.", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
        
        Case Is = -2147217843
            
            MsgBox "Database Password changed.", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
            
            
        Case Else
        
            MsgBox "Some error has occured. Error Number is - " & Err.Number & vbNewLine & "Error discription is - " & Err.Description, , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
    
    End Select
    
    
End Sub

Sub DishCode_Print()
    Dim DataCon As New ADODB.Connection
    Dim DataRec As New ADODB.Recordset
    Dim SQL As String
    
        Call Menu_Connection(DataCon)
        SQL = "Select * From TBL_Arabian_Menu"
        
        ShPrint.Range("A1").CurrentRegion.Clear

        ShPrint.Range("C1").CurrentRegion.Clear
        
        DataRec.Open SQL, DataCon, adOpenKeyset, adLockOptimistic
        
        Dim R As Long
        R = 1
        With DataRec
            If Not .RecordCount = 0 Then
                .MoveFirst
                Do Until .EOF
                    ShPrint.Cells(R, 1).Value = .Fields("Print Name")
                    ShPrint.Cells(R, 3).Value = "A " & .Fields("Dish Code")
                    R = R + 1
                    .MoveNext
                Loop
            End If
        
        End With
        
        DataRec.Close
                
        SQL = "Select * From TBL_Manna_Menu"
        
        DataRec.Open SQL, DataCon, adOpenKeyset, adLockOptimistic
        
        With DataRec
            If Not .RecordCount = 0 Then
                .MoveFirst
                Do Until .EOF
                    ShPrint.Cells(R, 1).Value = .Fields("Print Name")
                    ShPrint.Cells(R, 3).Value = "M " & .Fields("Dish Code")
                    R = R + 1
                    .MoveNext
                Loop
            End If
        
        End With
        
        DataRec.Close
        SQL = "Select * From TBL_AlaRasi_Menu"
        
        DataRec.Open SQL, DataCon, adOpenKeyset, adLockOptimistic
        
        With DataRec
            If Not .RecordCount = 0 Then
                .MoveFirst
                Do Until .EOF
                    ShPrint.Cells(R, 1).Value = .Fields("Print Name")
                    ShPrint.Cells(R, 3).Value = "P " & .Fields("Dish Code")
                    R = R + 1
                    .MoveNext
                Loop
            End If
        End With
                    
            ShPrint.Range("A1").CurrentRegion.Select
            Selection.RowHeight = 15
            ShPrint.Range("A1").Select
        
        DataRec.Close
        
        DataCon.Close


End Sub

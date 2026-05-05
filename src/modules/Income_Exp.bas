Attribute VB_Name = "Income_Exp"
Option Explicit
'############################################################################################################################
'############################################################################################################################
'############################################################################################################################
'######  @@@@@@@@@@@@@@   @@@              @@@@@@@@@@@@@@     %%%%%%%%%%%      %%%%%%%%%%%%%%      %%%%%%%%%%%    %%%  ######
'######  @@@@@@@@@@@@@@   @@@              @@@@@@@@@@@@@@     %%%%%%%%%%%%%    %%%%%%%%%%%%%%     %%%%%%%%%%%%%   %%%  ######
'######  @@@        @@@   @@@              @@@        @@@     %%%        %%%   %%%        %%%    %%%              %%%  ######
'######  @@@        @@@   @@@              @@@        @@@     %%%        %%    %%%        %%%    %%%              %%%  ######
'######  @@@        @@@   @@@              @@@        @@@     %%%        %     %%%        %%%     %%%%%%%%%%%%    %%%  ######
'######  @@@@@@@@@@@@@@   @@@              @@@@@@@@@@@@@@     %%%%%%%%%%%%%    %%%%%%%%%%%%%%      %%%%%%%%%%%%   %%%  ######
'######  @@@@@@@@@@@@@@   @@@              @@@@@@@@@@@@@@     %%%%%%%%%%%%%%   %%%%%%%%%%%%%%               %%%   %%%  ######
'######  @@@        @@@   @@@              @@@         @@     %%%         %%   %%%        %%%               %%%   %%%  ######
'######  @@@        @@@   @@@@@@@@@@@@@@   @@@         @@     %%%         %%   %%%        %%%     %%%%%%%%%%%%    %%%  ######
'######  @@@        @@@   @@@@@@@@@@@@@@   @@@         @@     %%%         %%   %%%        %%%      %%%%%%%%%%     %%%  ######
'############################################################################################################################
'############################################################################################################################
'############################################################################################################################


Sub Inc_Exp_Con(Con As ADODB.Connection)
   
   Dim Fpath As String
   Dim FCheck As String
        
        On Error GoTo ErrorHandler
        
Reprocess:
   Fpath = ShSupport.Range("I6").Value
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
'        ShSupport.Range("I6").Value = Fpath
'            GoTo Reprocess
        
   
    End If
   
        Con.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Fpath & ";Jet OLEDB:DataBase Password= admin123"
        
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

Sub Export_Data()
    'Dim WB As Workbook
    Dim SH As Worksheet
    Dim Con As New ADODB.Connection
    Dim DataRec As New ADODB.Recordset
    Dim SQL As String
    Dim i As Long

    
    
    Set SH = ShPrint
    SQL = "Select * From TBL_Expense"
    Call Inc_Exp_Con(Con)
    
    DataRec.Open SQL, Con, adOpenKeyset, adLockOptimistic
    i = 1
    
    With DataRec
        .MoveFirst
    
        Do Until .EOF
            SH.Cells(i, 1) = .Fields("ExpDate")
            SH.Cells(i, 2) = .Fields("Particular")
            SH.Cells(i, 3) = .Fields("Ledger Cat")
            SH.Cells(i, 4) = .Fields("Payment Mode")
            SH.Cells(i, 5) = .Fields("Amount")
            SH.Cells(i, 6) = .Fields("Narration")
            i = i + 1
            .MoveNext
        Loop
        
    End With
    
    
End Sub

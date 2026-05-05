Attribute VB_Name = "AlaRasi"
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


Sub Connection(Con As ADODB.Connection)
   
   Dim Fpath As String
   Dim FCheck As String
        
        On Error GoTo ErrorHandler
Reprocess:

   Fpath = ShSupport.Range("I3").Value

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
    
    'Application.Wait (Now + TimeValue("00:00:20"))
   
        Con.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Fpath & ";Jet OLEDB:DataBase Password= admin123"
        
        
        
    Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case Is = 52
        
            MsgBox "Please Connect to the server.", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
        
        Case Is = -2147217843
            
            MsgBox "DataBase Password Changed", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
            
            
        Case Else
        
            MsgBox "Some error has occured. Error Number is - " & Err.Number & vbNewLine & "Error discription is - " & Err.Description, , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
    
    End Select


End Sub

Sub Quit_App()
    Dim AlaCon As New ADODB.Connection
    Dim LogRec As New ADODB.Recordset
    Dim sqlLog As String
    
    On Error GoTo ErrorHandler
    sqlLog = "Select * From TBL_Login_Detail"
    Call Connection(AlaCon)
    LogRec.Open sqlLog, AlaCon, adOpenKeyset, adLockOptimistic
    
    With LogRec
        .MoveLast
        
        Do While Not .BOF
            If .Fields("UserID") = ShSupport.Range("I5").Value And IsNull(.Fields("LogOut Date")) Then
                .Fields("LogOut Date") = Now()
                .Update
            GoTo FinalTask
        Exit Sub
        
        
            End If
            .MovePrevious
        Loop
        
    Exit Sub
        
    End With
    
    LogRec.Close
    AlaCon.Close
    MsgBox "Unable to update logout time, please contact Sam.", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
      'Call Logout_Error_Msg
    
      
FinalTask:
        LogRec.Close
        AlaCon.Close
        Application.CommandBars("Task Pane").Visible = True
        ThisWorkbook.Save
        Application.DisplayAlerts = False
        Excel.Application.Quit
        
    Exit Sub
    
ErrorHandler:
    Select Case Err.Number
        
        Case Is = 52
        
        MsgBox "Please Connect to the Server", , "ALA RASI FOODS AND BEVRAGES PRIVATE LIMITED"
        
        Case Is = 2147217843
            
            MsgBox "Database Password changed.", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
        
        Case Is = 3709
        
            MsgBox "Unable to connect the Database", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
        
        Case Else
        
        MsgBox "Some Error has occured, and the Error number is " & Err.Number & ". Error description is " & Err.Description, , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
        
    
    End Select
End Sub

Sub Logout_Error_Msg()
     MsgBox "Unable to LogOut, Please Contact Sam", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
        
End Sub


Sub Excel_Mode()
    Call Entry_Point
    ShMenu.Visible = xlSheetVisible
    ShSupport.Visible = xlSheetVisible
    ShSupport.Visible = xlSheetVisible
        If Not ActiveSheet.Name = ShSupport.Name Then
            ShSupport.Activate
        End If

    Call Exit_Point

    ThisWorkbook.Application.Visible = True
    ThisWorkbook.Application.WindowState = xlMaximized

    End Sub

Sub App_Mode()
    Call Entry_Point
    ShMenu.Visible = xlSheetVeryHidden
    ShSupport.Visible = xlSheetVeryHidden
    ThisWorkbook.Application.Visible = False
    'ShSupport.Visible = xlSheetVeryHidden
    Call Exit_Point
    FRM_DashBoard.Show

End Sub

Sub Entry_Point()
    
    
End Sub

Sub Exit_Point()
    

End Sub


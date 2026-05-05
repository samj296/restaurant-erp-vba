Attribute VB_Name = "Temp"
Option Explicit
'############################################################################################################################
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
'############################################################################################################################


Sub Date_Check()
    If ShSupport.Range("L9").Value > ShSupport.Range("M9").Value Then
        Debug.Print "Date Is earlier"
    
    Else
        
        Debug.Print "Date is after"
        
    
    End If
End Sub


Sub Export_Data()
    Dim WB As Workbook
    Dim SH As Worksheet
    Dim DataCon As New ADODB.Connection
    Dim DataRec As New ADODB.Recordset
    Dim DataSql As String
    Dim Inc_Exp As Range
    Dim LRange As Range
    
    Set WB = Workbooks("Alarasi expense 2020-21.xlsx")
    Set SH = WB.Sheets("Watch")
    
    Call Inc_Exp_Con(DataCon)
    DataSql = "Select * from TBL_Validation_Income"
    DataRec.Open DataSql, DataCon, adOpenKeyset, adLockOptimistic
    
    
    Set Inc_Exp = SH.ListObjects("RecPart").ListColumns("Rec Particular").DataBodyRange
        
        
    For Each LRange In Inc_Exp
        With DataRec
            .AddNew
            .Fields("Rec Particular") = LRange.Value
'            .Fields("Particular") = LRange.Value
'            .Fields("Bank/Cash") = LRange.Offset(, 2).Value
'            .Fields("Rec From") = LRange.Offset(, 3).Value
'            .Fields("Amount") = LRange.Offset(, 4).Value
'            .Fields("Narration") = LRange.Offset(, 5).Value
            .Update
            
        End With
    
    Next LRange
    
    DataRec.Close
    
'    DataSql = "Select * from TBl_Expense"
'
'    DataRec.Open DataSql, DataCon, adOpenKeyset, adLockOptimistic
'
'    Set SH = WB.Sheets("Paid")
'
'    Set Inc_Exp = SH.ListObjects("Table3").ListColumns("Particular").DataBodyRange
'
'    For Each LRange In Inc_Exp
'        With DataRec
'            .AddNew
'            .Fields("ExpDate") = LRange.Offset(, -1).Value
'            .Fields("Particular") = LRange.Value
'            .Fields("Ledger Cat") = LRange.Offset(, 1).Value
'            .Fields("Payment Mode") = LRange.Offset(, 2).Value
'            .Fields("Amount") = LRange.Offset(, 3).Value
'            .Update
'
'        End With
'
'    Next LRange
'    DataRec.Close
    DataCon.Close
    
    
End Sub

Sub Menu_Export()
    Dim MenuRange As Range
    Dim R As Range
    Dim Code As Byte
    Dim DataCon As New ADODB.Connection
    Dim MenuRec As New ADODB.Recordset
    Dim MenuSql As String
    
    MenuSql = "Select * From TBL_Arabian_Menu"
    
    Call Menu_Connection(DataCon)
    
    MenuRec.Open MenuSql, DataCon, adOpenKeyset, adLockOptimistic
    
    Set MenuRange = ShMenu.ListObjects("Arabian_Menu_List").ListColumns("Cat").DataBodyRange
    
    Code = 1
    For Each R In MenuRange
    With MenuRec
        .AddNew
        .Fields("Dish Code") = R.Offset(, -1).Value
        .Fields("Dish Cat") = R.Value
        .Fields("Variant    ") = R.Offset(, 1).Value
        .Fields("Online Price") = R.Offset(, 2).Value
        .Fields("Take Away Price") = R.Offset(, 3).Value
        .Fields("DineIn Price") = R.Offset(, 4).Value
        .Update
        
    End With
    
        Code = Code + 1
        
    Next R
    

End Sub

Sub testing()
        Dim DataCon As New ADODB.Connection
        Dim DataRec As New ADODB.Recordset
        Dim SQL As String
        Dim FilePath As String
        Dim DataRange As Range
        Dim Dish As Range  'For Looping inside the data range

            FilePath = ShSupport.Range("I7").Value & "\Menu_Variant.accdb"
            SQL = "Select * From TBL_Shawarma"
            
            DataCon.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & FilePath & "; Jet OLEDB:Database Password=AlaRasi12062019"
            
            DataRec.Open SQL, DataCon, adOpenKeyset, adLockOptimistic
            
            Set DataRange = ShMenu.ListObjects("Arabian_Menu_list").ListColumns("Dishes").DataBodyRange
                    
                    With DataRec
                    
                    
                            For Each Dish In DataRange
                            .AddNew
                            .Fields("Dish") = Dish.Value
                            .Fields("Amount") = Dish.Offset(, 1).Value
                        
                    
            
                            Next Dish
            
                    
                    
                    
                    
                    
                End With






End Sub

Sub Creating_Table()
    Dim DataCon As New ADODB.Connection
    Dim DataRec As New ADODB.Recordset
    Dim Catalog As Object
    Dim FilePath As String
    Dim SQL As String
    Dim Xlspeak As SpeechLib.SpVoice
    
    Set Xlspeak = New SpVoice
    
        FilePath = ShSupport.Range("I7").Value & "\Menu_Variant.accdb"
        Application.Speech.Speak "This is the product of ALARASI FOODS AND BEVERAGES PRIVATE LIMITED", True
        Application.Speech.Speak "The login Time is" & Format(TimeValue(Now), "hh:mm AM/PM"), True
        
         Set Xlspeak.Voice = Xlspeak.GetVoices.Item(0)
         Xlspeak.Speak "This is the product of ALARASI FOODS AND BEVERAGES PRIVATE LIMITED"
         
         Set Xlspeak.Voice = Xlspeak.GetVoices.Item(2)
         Xlspeak.Speak "The login Time is" & Format(TimeValue(Now), "hh:mm AM/PM")


End Sub

Sub SQL_LOGIN()

    Dim LogCon As New ADODB.Connection, LogRec As New ADODB.Recordset
    Dim SQL As String
        
        On Error GoTo Handler
        SQL = "Select * From TBL_User Where TBL_User.[User Name]='samj' and TBL_User.Password='9893671902'"
        
        Call Connection(LogCon)
        
        LogRec.Open SQL, LogCon, adOpenKeyset, adLockOptimistic
        
        
        With LogRec
            
            .MoveFirst
            
            MsgBox "Total Record Count is: " & .RecordCount & vbNewLine & "User Name is: " & .Fields("User Name") & vbNewLine & "Password is:- " _
            & .Fields("Password"), , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
            
            
        
        End With
        
        LogRec.Close
        LogCon.Close
        
        Exit Sub
        
Handler:    Select Case Err.Number
            
            Case Is = 3021
                
                MsgBox "Invalid Credential", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
            
            End Select
        
        LogRec.Close
        LogCon.Close

End Sub

Sub percent()
    Dim M As Variant
    
    M = "Sam"
    
        If IsNumeric(M) Then
            MsgBox "Is numeric"
        
        Else
            MsgBox "Is not Numeric"
        End If

End Sub

Sub gosub_checking()
    
    GoSub Sam
    
    MsgBox "this is the 1st return statment", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
    
    GoSub Sub2
    
    MsgBox "This is the 2nd return statment"
    
    GoSub Sub3
    
    MsgBox "This is the 3rd return statement"
    Exit Sub
    
   
Sam:
    MsgBox "This is the 1st GoSub statement", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
    
    Return
Sub2:
    MsgBox "This is the 2nd GoSub Statement", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
    
    Return

Sub3:
    MsgBox "This is the 3rd GoSub Statement", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
    
    Return
   
End Sub


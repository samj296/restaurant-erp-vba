VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Expense 
   Caption         =   "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED (Expense)"
   ClientHeight    =   11020
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   20390
   OleObjectBlob   =   "Frm_Expense.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frm_Expense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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


Private Sub BTN_Back_Click()
    
    Me.Hide
    FRM_DashBoard.Show
    
End Sub



Private Sub BTN_Clear_Click()
            TXT_Particular.Text = ""
            CBox_Ledger.Text = ""
            CBox_Payment_Mode.Text = ""
            TXT_Amount.Text = ""
            TXT_Narration.Text = ""
            TXT_Calc.Text = ""
            TXT_Particular.SetFocus
            
End Sub

Private Sub BTN_Update_Click()
    Dim DataCon As New ADODB.Connection
    Dim DataRec As New ADODB.Recordset
    Dim SQL As String
    
        SQL = "Select * From TBL_Expense"
        Call Inc_Exp_Con(DataCon)
        DataRec.Open SQL, DataCon, adOpenKeyset, adLockOptimistic
        
        With DataRec
            .AddNew
            .Fields("ExpDate") = TXT_Date.Text
            .Fields("Particular") = TXT_Particular.Text
            .Fields("Ledger Cat") = CBox_Ledger.Text
            .Fields("Payment Mode") = CBox_Payment_Mode.Text
            .Fields("Amount") = TXT_Amount.Text
            .Fields("Narration") = TXT_Narration.Text
            .Update
        End With
        
            TXT_Particular.Text = ""
            CBox_Ledger.Text = ""
            CBox_Payment_Mode.Text = ""
            TXT_Amount.Text = ""
            TXT_Narration.Text = ""
        
        DataRec.Close
        DataCon.Close
    
End Sub





Private Sub TXT_Calc_AfterUpdate()
    Dim amnt() As String
    Dim NumItem As Byte
    Dim i As Byte 'for looping
    Dim rawNum As String
    
    On Error GoTo ErrorHandler
    If Left(TXT_Calc.Text, 1) = "=" Then
        rawNum = Right(TXT_Calc.Text, Len(TXT_Calc.Text) - 1)
    Else
        rawNum = TXT_Calc.Text
    
    End If
    
    Dim Result As Double
    amnt = Split(rawNum, "+")
    NumItem = UBound(amnt())
    For i = 0 To NumItem
        Result = Result + CLng(amnt(i))
    Next i
    TXT_Amount.Text = Result
    TXT_Narration.Text = TXT_Narration.Text & " (" & rawNum & ")"
    Exit Sub
    
ErrorHandler:
    Select Case Err.Number
    
    Case Is = 13
        
        MsgBox "Please correct the amount wriiten in the text box", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
        TXT_Calc.SetFocus
        Exit Sub
    Case Is = 6
        Exit Sub
    
    Case Else
        MsgBox "Some error has occured and the error number is: " & Err.Number & "." & vbNewLine & "The error discription is: " & Err.Description, , _
        "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
        
        
    End Select
    
End Sub



Private Sub UserForm_Initialize()
    TXT_Particular.SetFocus
    TXT_Date.Text = Format(Now(), "DD/MM/YYYY")
    
    CBox_Ledger.Clear
    CBox_Ledger.ColumnCount = 1
    CBox_Ledger.ColumnWidths = "120"
    CBox_Payment_Mode.Clear
    CBox_Payment_Mode.ColumnCount = 1
    CBox_Payment_Mode.ColumnWidths = "120"
    
    Dim DataCon As New ADODB.Connection
    Dim DataRec As New ADODB.Recordset
    Dim SQL As String
    SQL = "Select * From TBL_Ledger"
    
    Call Inc_Exp_Con(DataCon)
    
    DataRec.Open SQL, DataCon, adOpenKeyset, adLockOptimistic
    
    With DataRec
        .MoveFirst
        Do Until .EOF
            CBox_Ledger.AddItem .Fields("Ledger Cat")
        .MoveNext
        Loop
    
    End With
    
    DataRec.Close
    
    SQL = "Select * From TBL_Payment_Mode"
    
    DataRec.Open SQL, DataCon, adOpenKeyset, adLockOptimistic
    
    With DataRec
        .MoveFirst
            Do Until .EOF
                CBox_Payment_Mode.AddItem .Fields("Mode")
                .MoveNext
            Loop
        
    
    End With

End Sub

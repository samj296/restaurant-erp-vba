VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Salary 
   Caption         =   "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED (Salary Detail)"
   ClientHeight    =   11025
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   20390
   OleObjectBlob   =   "Frm_Salary.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frm_Salary"
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


Private Sub ComboBox1_Change()

End Sub




Private Sub BTN_Add_Gift_Deduction_Click()
    If CBox_Name.Value = "" Then
        MsgBox "Please select the Name from the DropDown List", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
        Exit Sub
    End If
    
    If CBox_Month_Deduction_Gift.Value = "" Then
        MsgBox "Please Select the month from the DropDown List", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
        Exit Sub
    End If
    
    If Cbox_Cat_Gift_Deduction.Value = "" Then
        MsgBox "Please Choose from the Gift and deduction from the DropDown List", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
        Exit Sub
    End If
    
    If TXT_Gift_Deduction.Text = "" Then
        MsgBox "Please Enter the value in the TEXT BOX to proceed", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
        Exit Sub
    End If
    
    Dim DataCon As New ADODB.Connection
    Dim DataRec As New ADODB.Recordset
    Dim StaffSql As String
    Dim ID As Byte
    
    If Cbox_Cat_Gift_Deduction.Value = "Gift" Then
        StaffSql = "Select * From TBL_Staff_Gift_Bonus"
    
    ElseIf Cbox_Cat_Gift_Deduction.Value = "Deduction" Then
        StaffSql = "Select * From TBL_Staff_Deduction"
    
    Else
        MsgBox "Please Select Gift or Deduction from the DropDown list"
        Exit Sub
        
    End If
    
        Call Inc_Exp_Con(DataCon)
        DataRec.Open StaffSql, DataCon, adOpenKeyset, adLockOptimistic
        ID = ShSupport.Range("T1").Value
        
        With DataRec
            .MoveFirst
            Do Until .EOF
                If .Fields("Id no") = ID Then
                    .Fields(CBox_Month_Deduction_Gift.Value) = CLng(TXT_Gift_Deduction.Text)
                    .Update
                    Cbox_Cat_Gift_Deduction.Value = ""
                    CBox_Month_Deduction_Gift.Value = ""
                    TXT_Gift_Deduction.Text = ""
                    DataRec.Clone
                    DataCon.Close
                    Call CBox_Name_Change
                    Exit Sub
                End If
                .MoveNext
            Loop
        End With
        
    DataRec.Close
    DataCon.Close
    
End Sub

Private Sub BTN_Back_Click()
    Me.Hide
    FRM_DashBoard.Show
End Sub

Private Sub BTN_Print_Click()
    Dim SalArray As Variant
    Dim List As Double
    Dim Rloop As Double
    
    List = Lbox_Payment_given.ListCount - 1
    ShPrint.Range("A1").CurrentRegion.Clear
    ShPrint.Range("A1", "B1").Merge
    ShPrint.Range("A1", "B1").HorizontalAlignment = xlCenter
    ShPrint.Range("A1", "B1").VerticalAlignment = xlCenter
    
    ReDim SalArray(1 To List, 1 To 2)
    
    For Rloop = 0 To List - 1
    
        
        SalArray(Rloop + 1, 1) = Lbox_Payment_given.List(Rloop + 1, 1)
        SalArray(Rloop + 1, 2) = Lbox_Payment_given.List(Rloop + 1, 4)
        
    Next Rloop
    
    For Rloop = 1 To List
        
        ShPrint.Cells(Rloop + 1, 1) = SalArray(Rloop, 1)
        ShPrint.Cells(Rloop + 1, 2) = SalArray(Rloop, 2)
        
        
    Next Rloop
        ShPrint.Range("A1").Value = "ALA RASI FOODS AND" & vbNewLine & "BEVERAGES PRIVATE" & vbNewLine & "LIMITED"
        ShPrint.Range("A1").CurrentRegion.PrintOut
        ShPrint.Range("A1").CurrentRegion.Clear
End Sub

Private Sub Cbox_Cat_Gift_Deduction_Change()
    If Cbox_Cat_Gift_Deduction.Value = "Gift" Then
        LBL_Gift_Inst = "Enter the Amount"
    ElseIf Cbox_Cat_Gift_Deduction.Value = "Deduction" Then
        LBL_Gift_Inst = "Enter the Value in Days"
    Else
        LBL_Gift_Inst = "Please Select from the DropDown list"
    
    End If
End Sub

Private Sub CBox_Name_Change()
    Dim DataCon As New ADODB.Connection
    Dim DataRec As New ADODB.Recordset
    Dim StaffSql As String
    Dim Seq As Long
    Dim TPaid As Double
    
    If CBox_Name.Value = "" Then
        LBL_Total_Amount = "Please select the Name"
        Lbox_Payment_given.Visible = False
        LBox_Salary_Monthly.Visible = False
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    StaffSql = "Select TBL_Expense.ExpDate, TBL_Expense.Particular, TBL_Ledger.[Ledger Cat]," _
    & "TBL_Expense.[Payment Mode], TBL_Expense.Amount From TBL_Ledger Inner Join TBL_Expense on TBL_Ledger.[Ledger Cat]=TBL_Expense.[Ledger Cat]" _
    & "Where ((TBL_Ledger.[Ledger Cat])='" & CBox_Name & "')"
    
    Call Inc_Exp_Con(DataCon)
    DataRec.Open StaffSql, DataCon, adOpenKeyset, adLockOptimistic
    
    Lbox_Payment_given.Visible = True
    LBox_Salary_Monthly.Visible = True
    
        LBox_Salary_Monthly.Clear
        Lbox_Payment_given.Clear
        Lbox_Payment_given.ColumnCount = 5
        Lbox_Payment_given.ColumnWidths = "60,140,120,60,60"
        Lbox_Payment_given.AddItem
        Lbox_Payment_given.List(0, 0) = "Date"
        Lbox_Payment_given.List(0, 1) = "Particular"
        Lbox_Payment_given.List(0, 2) = "Ledger Cat"
        Lbox_Payment_given.List(0, 3) = "Mode"
        Lbox_Payment_given.List(0, 4) = "Amount"
        
        Seq = 1
        'Run time error 3021 if the user type the letter or name which is not present in the list
    With DataRec
        .MoveFirst
            Do Until .EOF
            
                Lbox_Payment_given.AddItem
                Lbox_Payment_given.List(Seq, 0) = .Fields("ExpDate")
                Lbox_Payment_given.List(Seq, 1) = .Fields("Particular")
                Lbox_Payment_given.List(Seq, 2) = .Fields("Ledger Cat")
                Lbox_Payment_given.List(Seq, 3) = .Fields("Payment Mode")
                Lbox_Payment_given.List(Seq, 4) = .Fields("Amount")
            
                Seq = Seq + 1
                .MoveNext
            
            
            Loop
    
    
    End With
    
    DataRec.Close
    
    StaffSql = "Select sum(TBL_Expense.Amount) AS Amount From TBL_Expense Where TBL_Expense.[Ledger Cat]= ('" & CBox_Name.Value & "')"
    
    
    DataRec.Open StaffSql, DataCon, adOpenKeyset, adLockOptimistic

    DataRec.MoveFirst
        TPaid = DataRec.Fields("Amount")
    LBL_Total_Amount = "Total Aamount Paid= " & TPaid

    DataRec.Close
        
    Dim ID_No As Byte 'For Id number for selecting the user from another table in Access
    Dim StaffName As String 'For Name of the staff for doing the calculation
        
    StaffSql = "Select * From TBL_Staff_Detail"
    
    DataRec.Open StaffSql, DataCon, adOpenKeyset, adLockOptimistic
    
    With DataRec
        .MoveFirst
        
        Do Until .EOF
            If .Fields("Staff Name") = CBox_Name.Value Then
                
                ID_No = .Fields("Id")
                StaffName = .Fields("Staff Name")
                GoTo SalaryCalculation
                
            End If
            .MoveNext
        Loop
    
    
    End With
    
    DataRec.Close
    
    
SalaryCalculation:
    If Not DataRec Is Nothing Then
        DataRec.Close
    End If
    
    Dim Paid_Amount As Variant
    'For calculation of amount paid
    
    Dim Salary_Calculation As Variant
    'For Calculation of salary on monthly basis
    
    Dim Gift As Variant
    'For Calculation of Gift and bonus
    Dim Deduction As Variant
    'For Calculation of Deductable amount
    Dim ActualSalary As Variant
    'For Salary Decided amount
    
    Dim DueAmount As Variant
    'For Calculation of amount remaining per month
    Dim Result As Variant
    'For calculating the result in array directly
    Dim MonthAr As Variant
    'For writing month with array in one go
    
    ReDim Gift(1, 1 To 12)
    ReDim Deduction(1, 1 To 12)
    ReDim ActualSalary(1, 1 To 12)
    ReDim DueAmount(1, 1 To 12)
    ReDim Salary_Calculation(1, 1 To 12)
    ReDim Result(1, 1 To 12)
    ReDim Paid_Amount(1, 1 To 12)
    ReDim MonthAr(1, 1 To 12)
    
    'Calculation of the Gift and bonus in array for whole year
    StaffSql = "Select * From TBL_Staff_Gift_Bonus"
    
    DataRec.Open StaffSql, DataCon, adOpenKeyset, adLockOptimistic
    
        With DataRec
        .MoveFirst
            Do Until .EOF
                If .Fields("Id no") = ID_No Then
                    ShSupport.Range("T1").Value = ID_No
                    ShSupport.Range("R1").Value = StaffName
                    Gift(1, 1) = .Fields("April")
                    Gift(1, 2) = .Fields("May")
                    Gift(1, 3) = .Fields("June")
                    Gift(1, 4) = .Fields("July")
                    Gift(1, 5) = .Fields("August")
                    Gift(1, 6) = .Fields("September")
                    Gift(1, 7) = .Fields("October")
                    Gift(1, 8) = .Fields("November")
                    Gift(1, 9) = .Fields("December")
                    Gift(1, 10) = .Fields("January")
                    Gift(1, 11) = .Fields("February")
                    Gift(1, 12) = .Fields("March")
                    GoTo ActualSalary
                End If
            .MoveNext
            Loop
        End With
        
ActualSalary:
            DataRec.Close
            

            'Salary Decided Amount in array
            
            StaffSql = "Select * From TBL_Staff_Salary_Detail"
            
            DataRec.Open StaffSql, DataCon, adOpenKeyset, adLockOptimistic
            
            With DataRec
                .MoveFirst
                Do Until .EOF
                If .Fields("Staff Id") = ID_No Then
                       
                       ActualSalary(1, 1) = .Fields("April")
                       ActualSalary(1, 2) = .Fields("May")
                       ActualSalary(1, 3) = .Fields("June")
                       ActualSalary(1, 4) = .Fields("July")
                       ActualSalary(1, 5) = .Fields("August")
                       ActualSalary(1, 6) = .Fields("September")
                       ActualSalary(1, 7) = .Fields("October")
                       ActualSalary(1, 8) = .Fields("November")
                       ActualSalary(1, 9) = .Fields("December")
                       ActualSalary(1, 10) = .Fields("January")
                       ActualSalary(1, 11) = .Fields("February")
                       ActualSalary(1, 12) = .Fields("March")
                        GoTo Deduction
                End If
                .MoveNext
                Loop
            
            End With
Deduction:
           DataRec.Close
           
           'For calculation of Deduction amount
           
           StaffSql = "Select * From TBL_Staff_Deduction"
           
           DataRec.Open StaffSql, DataCon, adOpenKeyset, adLockOptimistic
           
           
           Dim DeductionDays As Variant
           
           ReDim DeductionDays(1, 1 To 12)
           
           With DataRec
            .MoveFirst
            Do Until .EOF
                If .Fields("Id no") = ID_No Then
                    
                    Deduction(1, 1) = (ActualSalary(1, 1) / 30) * .Fields("April")
                    Deduction(1, 2) = (ActualSalary(1, 2) / 30) * .Fields("May")
                    Deduction(1, 3) = (ActualSalary(1, 3) / 30) * .Fields("June")
                    Deduction(1, 4) = (ActualSalary(1, 4) / 30) * .Fields("July")
                    Deduction(1, 5) = (ActualSalary(1, 5) / 30) * .Fields("August")
                    Deduction(1, 6) = (ActualSalary(1, 6) / 30) * .Fields("September")
                    Deduction(1, 7) = (ActualSalary(1, 7) / 30) * .Fields("October")
                    Deduction(1, 8) = (ActualSalary(1, 8) / 30) * .Fields("November")
                    Deduction(1, 9) = (ActualSalary(1, 9) / 30) * .Fields("December")
                    Deduction(1, 10) = (ActualSalary(1, 10) / 30) * .Fields("January")
                    Deduction(1, 11) = (ActualSalary(1, 11) / 30) * .Fields("February")
                    Deduction(1, 12) = (ActualSalary(1, 12) / 30) * .Fields("March")
                    
                    DeductionDays(1, 1) = .Fields("April")
                    DeductionDays(1, 2) = .Fields("May")
                    DeductionDays(1, 3) = .Fields("June")
                    DeductionDays(1, 4) = .Fields("July")
                    DeductionDays(1, 5) = .Fields("August")
                    DeductionDays(1, 6) = .Fields("September")
                    DeductionDays(1, 7) = .Fields("October")
                    DeductionDays(1, 8) = .Fields("November")
                    DeductionDays(1, 9) = .Fields("December")
                    DeductionDays(1, 10) = .Fields("January")
                    DeductionDays(1, 11) = .Fields("February")
                    DeductionDays(1, 12) = .Fields("March")
                        GoTo Salary_Calculation
                End If
                .MoveNext
            Loop
           
           End With
            'DataRec.Close
            
            'For Salary Calculation (Salary_Calculation)
Salary_Calculation:
              Salary_Calculation(1, 1) = (ActualSalary(1, 1) + Gift(1, 1)) - Deduction(1, 1)
              Salary_Calculation(1, 2) = (ActualSalary(1, 2) + Gift(1, 2)) - Deduction(1, 2)
              Salary_Calculation(1, 3) = (ActualSalary(1, 3) + Gift(1, 3)) - Deduction(1, 3)
              Salary_Calculation(1, 4) = (ActualSalary(1, 4) + Gift(1, 4)) - Deduction(1, 4)
              Salary_Calculation(1, 5) = (ActualSalary(1, 5) + Gift(1, 5)) - Deduction(1, 5)
              Salary_Calculation(1, 6) = (ActualSalary(1, 6) + Gift(1, 6)) - Deduction(1, 6)
              Salary_Calculation(1, 7) = (ActualSalary(1, 7) + Gift(1, 7)) - Deduction(1, 7)
              Salary_Calculation(1, 8) = (ActualSalary(1, 8) + Gift(1, 8)) - Deduction(1, 8)
              Salary_Calculation(1, 9) = (ActualSalary(1, 9) + Gift(1, 9)) - Deduction(1, 9)
              Salary_Calculation(1, 10) = (ActualSalary(1, 10) + Gift(1, 10)) - Deduction(1, 10)
              Salary_Calculation(1, 11) = (ActualSalary(1, 11) + Gift(1, 11)) - Deduction(1, 11)
              Salary_Calculation(1, 12) = (ActualSalary(1, 12) + Gift(1, 12)) - Deduction(1, 12)
        
            'Calculation of Paid amount
            Dim Remaining As Double
            'for Calculation of remaining paid amount
            
             If TPaid > Salary_Calculation(1, 1) Or TPaid = Salary_Calculation(1, 1) Then
             
                Paid_Amount(1, 1) = Salary_Calculation(1, 1)
                Remaining = TPaid - Salary_Calculation(1, 1)
                
             Else
                Paid_Amount(1, 1) = TPaid
                Paid_Amount(1, 2) = 0
                Paid_Amount(1, 3) = 0
                Paid_Amount(1, 4) = 0
                Paid_Amount(1, 5) = 0
                Paid_Amount(1, 6) = 0
                Paid_Amount(1, 7) = 0
                Paid_Amount(1, 8) = 0
                Paid_Amount(1, 9) = 0
                Paid_Amount(1, 10) = 0
                Paid_Amount(1, 11) = 0
                Paid_Amount(1, 12) = 0
                    GoTo FinishingCode
             End If
             
            If Remaining > Salary_Calculation(1, 2) Or Remaining = Salary_Calculation(1, 2) Then
                
                Paid_Amount(1, 2) = Salary_Calculation(1, 2)
                Remaining = Remaining - Salary_Calculation(1, 2)
            
            Else
                Paid_Amount(1, 2) = Remaining
                Paid_Amount(1, 3) = 0
                Paid_Amount(1, 4) = 0
                Paid_Amount(1, 5) = 0
                Paid_Amount(1, 6) = 0
                Paid_Amount(1, 7) = 0
                Paid_Amount(1, 8) = 0
                Paid_Amount(1, 9) = 0
                Paid_Amount(1, 10) = 0
                Paid_Amount(1, 11) = 0
                Paid_Amount(1, 12) = 0
                    GoTo FinishingCode
            End If
    
            If Remaining > Salary_Calculation(1, 3) Or Remaining = Salary_Calculation(1, 3) Then
                Paid_Amount(1, 3) = Salary_Calculation(1, 3)
                Remaining = Remaining - Salary_Calculation(1, 3)
            
            Else
                Paid_Amount(1, 3) = Remaining
                Paid_Amount(1, 4) = 0
                Paid_Amount(1, 5) = 0
                Paid_Amount(1, 6) = 0
                Paid_Amount(1, 7) = 0
                Paid_Amount(1, 8) = 0
                Paid_Amount(1, 9) = 0
                Paid_Amount(1, 10) = 0
                Paid_Amount(1, 11) = 0
                Paid_Amount(1, 12) = 0
                    GoTo FinishingCode
            End If
            
            If Remaining > Salary_Calculation(1, 4) Or Remaining = Salary_Calculation(1, 4) Then
                Paid_Amount(1, 4) = Salary_Calculation(1, 4)
                Remaining = Remaining - Salary_Calculation(1, 4)
            
            Else
                Paid_Amount(1, 4) = Remaining
                Paid_Amount(1, 5) = 0
                Paid_Amount(1, 6) = 0
                Paid_Amount(1, 7) = 0
                Paid_Amount(1, 8) = 0
                Paid_Amount(1, 9) = 0
                Paid_Amount(1, 10) = 0
                Paid_Amount(1, 11) = 0
                Paid_Amount(1, 12) = 0

                    GoTo FinishingCode
            
            End If
            
            If Remaining > Salary_Calculation(1, 5) Or Remaining = Salary_Calculation(1, 5) Then
                Paid_Amount(1, 5) = Salary_Calculation(1, 5)
                Remaining = Remaining - Salary_Calculation(1, 5)
            Else
                Paid_Amount(1, 5) = Remaining
                Paid_Amount(1, 6) = 0
                Paid_Amount(1, 7) = 0
                Paid_Amount(1, 8) = 0
                Paid_Amount(1, 9) = 0
                Paid_Amount(1, 10) = 0
                Paid_Amount(1, 11) = 0
                Paid_Amount(1, 12) = 0
                    GoTo FinishingCode
            End If
            
            If Remaining > Salary_Calculation(1, 6) Or Remaining = Salary_Calculation(1, 6) Then
                Paid_Amount(1, 6) = Salary_Calculation(1, 6)
                Remaining = Remaining - Salary_Calculation(1, 6)
            Else
                Paid_Amount(1, 6) = Remaining
                Paid_Amount(1, 7) = 0
                Paid_Amount(1, 8) = 0
                Paid_Amount(1, 9) = 0
                Paid_Amount(1, 10) = 0
                Paid_Amount(1, 11) = 0
                Paid_Amount(1, 12) = 0
                    GoTo FinishingCode
            End If
            
            If Remaining > Salary_Calculation(1, 7) Or Remaining = Salary_Calculation(1, 7) Then
                Paid_Amount(1, 7) = Salary_Calculation(1, 7)
                Remaining = Remaining - Salary_Calculation(1, 7)
            
            Else
                Paid_Amount(1, 7) = Remaining
                Paid_Amount(1, 8) = 0
                Paid_Amount(1, 9) = 0
                Paid_Amount(1, 10) = 0
                Paid_Amount(1, 11) = 0
                Paid_Amount(1, 12) = 0
                    GoTo FinishingCode
                
            End If
            
            If Remaining > Salary_Calculation(1, 8) Or Remaining = Salary_Calculation(1, 8) Then
                Paid_Amount(1, 8) = Salary_Calculation(1, 8)
                Remaining = Remaining - Salary_Calculation(1, 8)
                
            Else
                Paid_Amount(1, 8) = Remaining
                Paid_Amount(1, 9) = 0
                Paid_Amount(1, 10) = 0
                Paid_Amount(1, 11) = 0
                Paid_Amount(1, 12) = 0
                    GoTo FinishingCode
                    
            End If
            
            If Remaining > Salary_Calculation(1, 9) Or Remaining = Salary_Calculation(1, 9) Then
                Paid_Amount(1, 9) = Salary_Calculation(1, 9)
                Remaining = Remaining - Salary_Calculation(1, 9)
                
            Else
                Paid_Amount(1, 9) = Remaining
                Paid_Amount(1, 10) = 0
                Paid_Amount(1, 11) = 0
                Paid_Amount(1, 12) = 0
                    GoTo FinishingCode
                    
            End If
            
            If Remaining > Salary_Calculation(1, 10) Or Remaining = Salary_Calculation(1, 10) Then
                Paid_Amount(1, 10) = Salary_Calculation(1, 10)
                Remaining = Remaining - Salary_Calculation(1, 10)
                
            Else
                Paid_Amount(1, 10) = Remaining
                Paid_Amount(1, 11) = 0
                Paid_Amount(1, 12) = 0
                    GoTo FinishingCode
                
            End If
            
            If Remaining > Salary_Calculation(1, 11) Or Remaining = Salary_Calculation(1, 11) Then
                Paid_Amount(1, 11) = Salary_Calculation(1, 11)
                Remaining = Remaining - Salary_Calculation(1, 11)
                
            Else
                Paid_Amount(1, 11) = Remaining
                Paid_Amount(1, 12) = 0
                
            End If
            
            If Remaining > Salary_Calculation(1, 12) Or Remaining = Salary_Calculation(1, 12) Then
                Paid_Amount(1, 12) = Salary_Calculation(1, 12)
                Remaining = Remaining - Salary_Calculation(1, 12)
                
            Else
                Paid_Amount(1, 12) = Remaining
                
            End If
            
FinishingCode:
            
            DueAmount(1, 1) = Salary_Calculation(1, 1) - Paid_Amount(1, 1)
            DueAmount(1, 2) = Salary_Calculation(1, 2) - Paid_Amount(1, 2)
            DueAmount(1, 3) = Salary_Calculation(1, 3) - Paid_Amount(1, 3)
            DueAmount(1, 4) = Salary_Calculation(1, 4) - Paid_Amount(1, 4)
            DueAmount(1, 5) = Salary_Calculation(1, 5) - Paid_Amount(1, 5)
            DueAmount(1, 6) = Salary_Calculation(1, 6) - Paid_Amount(1, 6)
            DueAmount(1, 7) = Salary_Calculation(1, 7) - Paid_Amount(1, 7)
            DueAmount(1, 8) = Salary_Calculation(1, 8) - Paid_Amount(1, 8)
            DueAmount(1, 9) = Salary_Calculation(1, 9) - Paid_Amount(1, 9)
            DueAmount(1, 10) = Salary_Calculation(1, 10) - Paid_Amount(1, 10)
            DueAmount(1, 11) = Salary_Calculation(1, 11) - Paid_Amount(1, 11)
            DueAmount(1, 12) = Salary_Calculation(1, 12) - Paid_Amount(1, 12)
            
            MonthAr(1, 1) = "April"
            MonthAr(1, 2) = "May"
            MonthAr(1, 3) = "June"
            MonthAr(1, 4) = "July"
            MonthAr(1, 5) = "August"
            MonthAr(1, 6) = "September"
            MonthAr(1, 7) = "October"
            MonthAr(1, 8) = "November"
            MonthAr(1, 9) = "December"
            MonthAr(1, 10) = "January"
            MonthAr(1, 11) = "February"
            MonthAr(1, 12) = "March"
            
            
            
            'for entry in Lbox_Payment_given
            
            LBox_Salary_Monthly.Clear
            LBox_Salary_Monthly.ColumnCount = 8
            LBox_Salary_Monthly.ColumnWidths = "120,120,120,120,120,120,120,120"
            
            LBox_Salary_Monthly.AddItem
            LBox_Salary_Monthly.List(0, 0) = "Month"
            LBox_Salary_Monthly.List(0, 1) = "Salary Decided"
            LBox_Salary_Monthly.List(0, 2) = "Salary Amount"
            LBox_Salary_Monthly.List(0, 3) = "Paid"
            LBox_Salary_Monthly.List(0, 4) = "Due"
            LBox_Salary_Monthly.List(0, 5) = "Deduction (Amnt)"
            LBox_Salary_Monthly.List(0, 6) = "Deduction (Days)"
            LBox_Salary_Monthly.List(0, 7) = "Gift/Bonus"
            
            Dim C As Byte

            For C = 1 To 12
                ShSupport.Cells(C + 2, 18).Value = ActualSalary(1, C)
                ShSupport.Cells(C + 2, 19).Value = Gift(1, C)
                ShSupport.Cells(C + 2, 20).Value = Deduction(1, C)
                ShSupport.Cells(C + 2, 21).Value = Salary_Calculation(1, C)
                ShSupport.Cells(C + 2, 22).Value = Paid_Amount(1, C)
                ShSupport.Cells(C + 2, 23).Value = DueAmount(1, 12)
                
                LBox_Salary_Monthly.AddItem
                LBox_Salary_Monthly.List(C, 0) = MonthAr(1, C)
                LBox_Salary_Monthly.List(C, 1) = ActualSalary(1, C)
                LBox_Salary_Monthly.List(C, 2) = Salary_Calculation(1, C)
                LBox_Salary_Monthly.List(C, 3) = Paid_Amount(1, C)
                LBox_Salary_Monthly.List(C, 4) = DueAmount(1, C)
                LBox_Salary_Monthly.List(C, 5) = Deduction(1, C)
                LBox_Salary_Monthly.List(C, 6) = DeductionDays(1, C)
                LBox_Salary_Monthly.List(C, 7) = Gift(1, C)
                
            Next

            
            

    DataRec.Close
    DataCon.Close
    
    Exit Sub



    
ErrorHandler:
    Select Case Err.Number
        Case Is = 3021
        MsgBox "Please select the Name from the drop down List", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
        
        Case Else
        MsgBox "Some Error has occured and the error number is:-" & Err.Number & ". The Error Description is :-" & Err.Description
        
        
    
    End Select
    
End Sub

Private Sub UserForm_Activate()
    Call UserForm_Initialize
End Sub

Private Sub UserForm_Initialize()
    Dim DataCon As New ADODB.Connection
    Dim DataRec As New ADODB.Recordset
    Dim StaffSql As String
    
    StaffSql = "Select * From TBL_Staff_Detail"
    
    Call Inc_Exp_Con(DataCon)
    CBox_Name.Clear
    CBox_Name.ColumnCount = 1
    CBox_Name.ColumnWidths = "80"
    Cbox_Cat_Gift_Deduction.Clear
    CBox_Month_Deduction_Gift.Clear
    Lbox_Payment_given.Clear
    LBox_Salary_Monthly.Clear
    Lbox_Payment_given.Visible = False
    LBox_Salary_Monthly.Visible = False
    
    DataRec.Open StaffSql, DataCon, adOpenKeyset, adLockOptimistic
    
    'For Addig Ustaff in the CBox_Name dropdown list
    
    With DataRec
        .MoveFirst
        
        Do Until .EOF
            
            CBox_Name.AddItem .Fields("Staff Name")
           .MoveNext
        Loop
    End With
        
        'For Adding month for gift & deduction Combobox
        
        Dim GMonth As Range
        Dim C As Range  'For looping inside the month range in TBL_SalDetail
        
        Set GMonth = ShSupport.ListObjects("TBL_SalDetail").ListColumns("Month").DataBodyRange
        
        For Each C In GMonth
            
            CBox_Month_Deduction_Gift.AddItem C.Value
            
        
        Next C
        
        Cbox_Cat_Gift_Deduction.AddItem "Gift"
        Cbox_Cat_Gift_Deduction.AddItem "Deduction"
        
        
        
    If CBox_Name.Value = "" Then
        LBL_Total_Amount = "Please Select the Name"
    End If
            
            
    DataRec.Close
    DataCon.Close

End Sub



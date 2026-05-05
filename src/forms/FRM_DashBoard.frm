VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM_DashBoard 
   Caption         =   "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
   ClientHeight    =   11020
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   20390
   OleObjectBlob   =   "FRM_DashBoard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FRM_DashBoard"
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


Private Sub BTN_Attendence_Click()
    FRM_Attendence.Show
End Sub

Private Sub BTN_Billing_Click()
    
    Me.Hide
    FRM_Billing.Show
    
End Sub

Private Sub BTN_Excel_Click()
    Select Case ShSupport.Range("I2").Value
        
    Case Is = ShSupport.Range("L2").Value
            Me.Hide
            Call Excel_Mode
    
    Case Else
        MsgBox "You are not Athorise to open This", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
        
    End Select
    
End Sub


Private Sub BTN_Exit_Click()
    Call Quit_App
End Sub

Private Sub BTN_Expense_Click()
    Select Case ShSupport.Range("I2").Value
    
    Case Is = ShSupport.Range("L2").Value
        Me.Hide
        Frm_Expense.Show
        
    Case Else
        MsgBox "You are not authorise for this action", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
    
    End Select
    
End Sub

Private Sub BTN_Income_Click()
    Select Case ShSupport.Range("I2").Value
    
    Case Is = ShSupport.Range("L2").Value
        Me.Hide
        Frm_Income.Show
    Case Else
        MsgBox "You are not authorise for this action", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
        
    End Select
End Sub


Private Sub BTN_Salary_Detail_Click()
    Select Case ShSupport.Range("I2").Value
    
    Case Is = ShSupport.Range("L2").Value
        Me.Hide
        Frm_Salary.Show
    Case Else
        MsgBox "You are not authorise for this action", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
        
    End Select
End Sub

Private Sub BTN_Switch_User_Click()
    Dim AlaCon As New ADODB.Connection
    Dim LogRec As New ADODB.Recordset
    Dim sqlLog As String
    
    On Error GoTo ErrorHandler
    sqlLog = "Select * From TBL_Login_Detail"
    Call Connection(AlaCon)
    LogRec.Open sqlLog, AlaCon, adOpenKeyset, adLockOptimistic
    
    With LogRec
        .MoveFirst
        
        Do Until .EOF
            If .Fields("UserID") = ShSupport.Range("I5").Value And IsNull(.Fields("LogOut Date")) Then
                .Fields("LogOut Date") = Now()
                .Update
            LogRec.Close
        AlaCon.Close
        ThisWorkbook.Save
        Me.Hide
        Frm_LogIn.Show
        Exit Sub
        
        
            End If
            .MoveNext
        Loop
        
    End With
    
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


Private Sub IMG_Total_Sale_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal x As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    Dim Chrt As Chart
    
    Set Chrt = ShSupport.ChartObjects("Chrt_Sale").Chart
    
    
    
    Chrt.Export Filename:=VBA.Environ("Temp") & Application.PathSeparator & "Sale_Chrt.Jpg", filtername:="Jpg"
    
    IMG_Total_Sale.Picture = LoadPicture(VBA.Environ("Temp") & Application.PathSeparator & "Sale_Chrt.Jpg")
    
End Sub

Private Sub UserForm_Activate()
    Call UserForm_Initialize
End Sub



Private Sub UserForm_Click()
    IMG_Total_Sale.Visible = False
    LBox_Attendance.Visible = False
End Sub


Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    IMG_Total_Sale.Visible = True
    Dim Chrt As Chart
    

            Set Chrt = ShSupport.ChartObjects("Chrt_Sale").Chart
        
        
        
            Chrt.Export Filename:=VBA.Environ("Temp") & Application.PathSeparator & "Sale_Chrt.Jpg", filtername:="Jpg"
        
            IMG_Total_Sale.Picture = LoadPicture(VBA.Environ("Temp") & Application.PathSeparator & "Sale_Chrt.Jpg")
    
    Select Case ShSupport.Range("I2").Value
        
        Case Is = "Admin"
            
            LBox_Attendance.Visible = True
        
        Case Else
        
            LBox_Attendance.Visible = False
    
    End Select
    
    
End Sub

Private Sub UserForm_Initialize()
    IMG_Total_Sale.Visible = False
    LBox_Attendance.Visible = False
    
'    Dim Chrt As Chart
'
'    Set Chrt = ShSupport.ChartObjects("Chrt_Sale").Chart
'
'
'
'    Chrt.Export Filename:=VBA.Environ("Temp") & Application.PathSeparator & "Sale_Chrt.Jpg", filtername:="Jpg"
'
'    IMG_Total_Sale.Picture = LoadPicture(VBA.Environ("Temp") & Application.PathSeparator & "Sale_Chrt.Jpg")
    
End Sub




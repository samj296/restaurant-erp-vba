VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_LogIn 
   Caption         =   "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
   ClientHeight    =   11040
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   20370
   OleObjectBlob   =   "Frm_LogIn.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frm_LogIn"
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


Private Sub BTN_cancel_Click()
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Excel.Application.Quit
    
End Sub

Private Sub BTN_LogIn_Click()
    Dim AlaCon As New ADODB.Connection
    Dim UserRec As New ADODB.Recordset
    Dim sqlLogDetail As String
    Dim Xlspeak As SpeechLib.SpVoice
    
    Set Xlspeak = New SpVoice
    Set Xlspeak.Voice = Xlspeak.GetVoices.Item(2)
    
    Dim sqlUser As String
    Dim LoginId As Byte
    
    sqlUser = "select"
    
    
    
    
    
    
    sqlUser = "Select * From TBL_User"
    sqlLogDetail = "Select * From TBL_Login_Detail"
Connection:


    Call Connection(AlaCon)

        On Error GoTo ErrorHandler

    UserRec.Open sqlUser, AlaCon, adOpenKeyset, adLockOptimistic

    With UserRec
        .MoveFirst
            Do Until .EOF
                If TXT_UserID.Text = .Fields("User Name") And TXT_Password.Text = .Fields("Password") Then

                    MsgBox "Welcome " & .Fields("Full Name"), , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
                    ShSupport.Range("I4").Value = Now()
                    Xlspeak.Speak "Welcome " & .Fields("Full Name") & " Have a nice day!" & "This the product of ALARASI FOODS AND BEVERAGES PRIVATE LIMITED."
                    LoginId = .Fields("UserID_NO")
                    ShSupport.Range("I5").Value = LoginId
                    ShSupport.Range("I2").Value = .Fields("Admin Rights")
                    Application.CommandBars("Task Pane").Visible = False
                        Call BTN_Reset_Click
                    


                    GoTo Final
                End If

                .MoveNext
            Loop
    End With
        UserRec.Close

    
                Call BTN_Reset_Click
                TXT_UserID.SetFocus
                MsgBox "Login Failed"
                AlaCon.Close

                    Exit Sub
                    
Final:
            UserRec.Close
        UserRec.Open sqlLogDetail, AlaCon, adOpenKeyset, adLockOptimistic

    With UserRec
        .AddNew
            .Fields("Login Date") = ShSupport.Range("I4").Value
            .Fields("UserID") = LoginId
            .Update
    End With
            AlaCon.Close
            Me.Hide
      FRM_DashBoard.Show
    


        Unload Me
    Exit Sub

ErrorHandler:
    If Err.Number = 3709 Then
        MsgBox "Unable to connect to the DataBase. Please Connect to the main computer to proceed", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
    Else
        MsgBox "Some Error has occured, and the error number is:-", vbInformation, "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"

    End If
    
    
End Sub


Private Sub BTN_Reset_Click()
    TXT_UserID.Text = ""
    TXT_Password.Text = ""
    
End Sub








'
'Private Sub TXT_Password_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'     If KeyAscii.Value = vbKeyReturn Then
'        Call BTN_LogIn_Click
'     ElseIf KeyAscii.Value = vbKeyEscape Then
'        Call BTN_cancel_Click
'     End If
'
'
'End Sub
'
'
'
'Private Sub TXT_UserID_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'     If KeyAscii.Value = vbKeyReturn Then
'        Call BTN_LogIn_Click
'     ElseIf KeyAscii.Value = vbKeyEscape Then
'        Call BTN_cancel_Click
'     End If
'
'End Sub
'
'Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'
'End Sub
'
'Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'     If KeyCode = vbKeyReturn Then
'        Call BTN_LogIn_Click
'    ElseIf KeyCode = vbKeyEscape Then
'        Call BTN_cancel_Click
'    End If
'
'End Sub
Private Sub UserForm_Click()

End Sub

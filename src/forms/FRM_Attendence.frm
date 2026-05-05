VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM_Attendence 
   Caption         =   "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
   ClientHeight    =   7250
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   10140
   OleObjectBlob   =   "FRM_Attendence.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FRM_Attendence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As Boolean
Private Sub BTN_close_Click()
    x = True
    Unload Me
End Sub

Private Sub BTN_LogIn_Click()
    Dim RightT As String
    Dim UserCon As New ADODB.Connection
    Dim UserRec As New ADODB.Recordset
    Dim SQL As String
    Dim LogStatus As String
    
                    
        SQL = "Select * From TBL_User Where TBL_User.[User Name]= '" & TXT_UserID.Text & "' AND TBL_User.Password= '" & TXT_Password.Text & "'"
            
        
        Call Connection(UserCon)
        UserRec.Open SQL, UserCon, adOpenKeyset, adLockOptimistic
        
        If UserRec.RecordCount > 0 Then
        
            With UserRec
                .MoveFirst
                    
                Dim UserName As String
                UserName = .Fields("Full Name")
                
                
            End With
            Select Case OBTN_LogIn
                Case Is = True
                    
                    RightT = TimeValue(Format(ShSupport.Range("I9").Value, "HH:MM AM/PM")) 'Starting time as per decided
                
                Case Is = False
                    If OBTN_Logout = True Then
                    
                    RightT = TimeValue(Format(ShSupport.Range("I10").Value, "HH:MM AM/PM")) 'Closing time as per decided

                    ElseIf OBTN_Break_Start = True Or OBTN_Break_Start = True Or OBTN_Break_end = True Then
                        RightT = TimeValue(Format(Now, "HH:MM AM/PM"))
                    End If
                
            End Select
            
            
            
            Dim LogTime As String
            Dim EntryField As String
            
            If TimeValue(RightT) < TimeValue(Now) And OBTN_LogIn = True Then    'if Login is late
                    
                    LogTime = Format(TimeValue(RightT) - TimeValue(Now), "hh:mm:ss")
                    MsgBox "You are late by " & LogTime, , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
                    LogStatus = "Late by " & LogTime
                    EntryField = "LogIn Date-Time"
                    
            ElseIf TimeValue(RightT) > TimeValue(Now) And OBTN_LogIn = True Or TimeValue(RightT) = TimeValue(Now) And OBTN_LogIn = True Then 'If log in is on time
                    
                    LogStatus = "On Time"
                    EntryField = "LogIn Date-Time"
                    

            ElseIf TimeValue(RightT) > TimeValue(Now) And OBTN_Logout = True Then       'if logout is early before time
                    
                    LogTime = Format(TimeValue(RightT) - TimeValue(Now), "hh:mm:ss")
                    MsgBox "You are early by " & LogTime, , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
                    LogStatus = "Early"
                    EntryField = "Logout Date-Time"
                    
            ElseIf TimeValue(RightT) < TimeValue(Now) And OBTN_Logout = True Then       'if logout is on time
                    
                    LogStatus = "on time"
                    EntryField = "Logout Date-Time"
            
            End If
            
                UserRec.Close
            
                SQL = "Select * From TBL_Attendance"
                
                UserRec.Open SQL, UserCon, adOpenKeyset, adLockOptimistic
                    
                    If OBTN_LogIn = True Then
                        
                        GoTo LogIn
                        
                    ElseIf OBTN_Logout = True Then
                        
                        GoTo LogOut
                    ElseIf OBTN_Break_Start = True Then
                        
                        GoTo BreakStart
                    
                    ElseIf OBTN_Break_end = True Then
                        
                        GoTo BreakEnd
                    
                    End If
                
LogIn:
                With UserRec
                    .AddNew
                    .Fields("User Name (Full Name)") = UserName
                    .Fields("LogIn Date-Time") = Date & " " & TimeValue(Now())
                    .Fields("Status") = LogStatus
                    .Update
                    
                End With
                    
                    GoTo FinalTask
                    

LogOut:
                UserRec.Close
                SQL = "Select * From TBL_Attendance where TBL_Attendance.[User Name (Full Name)]= '" & UserName & "' AND TBL_Attendance.[Logout Date-Time] is null"
                
                UserRec.Open SQL, UserCon, adOpenKeyset, adLockOptimistic
                
                If UserRec.RecordCount > 0 Then
                
                    UserRec.Fields("Logout Date-Time") = Now()
                    UserRec.Fields("Status") = UserRec.Fields("Status") & "/" & LogStatus
                    UserRec.Update
                ElseIf UserRec.RecordCount = 0 Then
                    
                    MsgBox "Loging Out without LoggedIn. This will be counted as Delayed", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
                        
                        UserRec.Close
                        SQL = "Select * From TBL_Attendance"
                        UserRec.Open SQL, UserCon, adOpenKeyset, adLockOptimistic
                        
                        UserRec.AddNew
                        UserRec.Fields("User Name (Full Name)") = UserName
                        UserRec.Fields("Logout Date-Time") = Date & " " & TimeValue(Now)
                        UserRec.Fields("Status") = "Skipped/" & LogStatus
                        UserRec.Update
                    
                End If
                GoTo FinalTask
                
BreakStart:
            UserRec.Close
            SQL = "Select * From TBL_Break"
            UserRec.Open SQL, UserCon, adOpenKeyset, adLockOptimistic
            
            With UserRec
                .AddNew
                .Fields("User Name") = UserName
                .Fields("Break Date") = Format(Date, "DD/MM/YYYY")
                .Fields("Break Start (Time)") = RightT
                .Update
                
            End With
            
            GoTo FinalTask

BreakEnd:
            UserRec.Close
            SQL = "Select * From TBL_Break Where TBL_Break.[User Name]= '" & UserName & "' AND TBL_Break.[Break End (Time)] is Null "
            UserRec.Open SQL, UserCon, adOpenKeyset, adLockOptimistic
            
            With UserRec
                If Not .RecordCount = 0 Then
                    .Fields("Break End (Time)") = RightT
                    .Update
                    GoTo FinalTask
                Else
                    MsgBox "Break Start time is not recorded, This will considered as half Day", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
                    SQL = "Select * From TBL_Break"
                    
                End If
                
            End With
BreakEndNew:
            UserRec.Close
            UserRec.Open SQL, UserCon, adOpenKeyset, adLockOptimistic
            
            With UserRec
                .AddNew
                .Fields("User Name") = UserName
                .Fields("Break Date") = Format(Date, "DD/MM/YYYY")
                .Fields("Break End (Time)") = RightT
                .Update
                
            End With
            
            UserRec.Close
                
FinalTask:
            TXT_UserID.Text = ""
            TXT_Password.Text = ""
            'UserRec.Close
            UserCon.Close
        
        ElseIf UserRec.RecordCount = 0 Then
            MsgBox "Invalid credential. Unable to Login", , "ALA RASI FOODS AND BEVERAGES PRIVATE LIMITED"
            
            UserRec.Close
            UserCon.Close
            
            Exit Sub
        End If
        
        'UserRec.Close
        'UserCon.Close

End Sub

Private Sub BTN_Reset_Click()

    TXT_UserID.Text = ""
    TXT_Password.Text = ""

End Sub

Private Sub LBL_Clock_Click()
    Call clock
End Sub

Private Sub UserForm_Click()

    If x = True Then
        x = False
    End If
    
    Call clock
End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If x = False Then
        x = True
    End If
End Sub

Private Sub UserForm_Deactivate()
    x = True
    
End Sub



Private Sub UserForm_Initialize()
    OBTN_LogIn = True
    x = False
    
    
End Sub

Private Sub UserForm_Terminate()
    x = True
End Sub

Sub clock()
    
            
        
        Do
                DoEvents
            If x = True Then Exit Sub
                LBL_Clock = Format(Now, "DD-MM-YYYY h:MM:ss am/pm")
                
            
        Loop
    
End Sub

